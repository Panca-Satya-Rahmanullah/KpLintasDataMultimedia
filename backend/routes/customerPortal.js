var express = require('express');
var router = express.Router();
var multer = require('multer');
var path = require('path');
var fs = require('fs');
var db = require('../config/db');
var verifyCustomerToken = require('../middleware/customerAuth');
var SocketService = require('../services/socket');

// Protect all portal routes with customer JWT token
router.use(verifyCustomerToken);

// Multer Upload Configuration
var uploadDir = path.join(__dirname, '../public/uploads/bukti');
if (!fs.existsSync(uploadDir)) {
  fs.mkdirSync(uploadDir, { recursive: true });
}

var storage = multer.diskStorage({
  destination: function(req, file, cb) {
    cb(null, uploadDir);
  },
  filename: function(req, file, cb) {
    var uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9);
    cb(null, 'bukti-' + uniqueSuffix + path.extname(file.originalname));
  }
});

var upload = multer({
  storage: storage,
  fileFilter: function(req, file, cb) {
    if (!file.mimetype.startsWith('image/')) {
      return cb(new Error('Hanya file gambar yang diperbolehkan!'), false);
    }
    cb(null, true);
  },
  limits: {
    fileSize: 5 * 1024 * 1024 // Max 5MB
  }
});

/* GET /api/customer/portal/billing - Get active bill for customer */
router.get('/billing', function(req, res) {
  var id_pelanggan = req.customerId;

  function sendBillingResponse(billingData) {
    var lastPaymentSql = `
      SELECT pem.*, t.periode, t.nominal 
      FROM pembayaran pem 
      JOIN tagihan t ON pem.id_tagihan = t.id_tagihan 
      WHERE t.id_pelanggan = ? 
      ORDER BY pem.tanggal_upload DESC 
      LIMIT 1
    `;
    db.query(lastPaymentSql, [id_pelanggan], function(payErr, payResults) {
      var lastPayment = null;
      if (!payErr && payResults && payResults[0]) {
        lastPayment = payResults[0];
      }
      res.json({
        success: true,
        data: billingData,
        lastPayment: lastPayment
      });
    });
  }

  var sql = `
    SELECT t.*, p.nama, p.paket, p.status_tagihan 
    FROM tagihan t 
    JOIN pelanggan p ON t.id_pelanggan = p.id_pelanggan 
    WHERE t.id_pelanggan = ? AND t.status != 'lunas'
    ORDER BY t.due_date ASC 
    LIMIT 1
  `;

  db.query(sql, [id_pelanggan], function(err, results) {
    if (err) {
      return res.status(500).json({ success: false, message: 'Database error', error: err.message });
    }

    if (results.length === 0) {
      // Check if we need to generate the first/missing tagihan for their current due_date
      var checkCustSql = `
        SELECT p.*, pl.harga 
        FROM pelanggan p
        LEFT JOIN paket_layanan pl ON p.paket = pl.nama_paket
        WHERE p.id_pelanggan = ?
      `;
      db.query(checkCustSql, [id_pelanggan], function(custErr, custResults) {
        if (custErr || custResults.length === 0) {
          return sendBillingResponse(null);
        }

        var customer = custResults[0];
        if (!customer.due_date || !customer.harga) {
          return sendBillingResponse(null);
        }

        // Format due_date to YYYY-MM period
        var d = new Date(customer.due_date);
        var year = d.getFullYear();
        var month = d.getMonth() + 1;
        var period = year + '-' + (month < 10 ? '0' + month : month);

        // Check if a bill for this period already exists
        var checkBillSql = 'SELECT * FROM tagihan WHERE id_pelanggan = ? AND periode = ?';
        db.query(checkBillSql, [id_pelanggan, period], function(billErr, billResults) {
          if (billErr || billResults.length > 0) {
            return sendBillingResponse(null);
          }

          console.log(`[Billing Service] Generating missing bill for customer ${customer.nama} (${period})`);
          var Tagihan = require('../models/Tagihan');
          Tagihan.create({
            id_pelanggan: id_pelanggan,
            periode: period,
            nominal: customer.harga,
            status: 'belum_bayar',
            due_date: customer.due_date
          }, function(createBillErr, newBill) {
            if (createBillErr) {
              console.error('[Billing Service] Failed to generate bill:', createBillErr.message);
              return res.status(500).json({ success: false, message: 'Gagal membuat tagihan otomatis', error: createBillErr.message });
            }

            sendBillingResponse({
              id_tagihan: newBill.id_tagihan,
              id_pelanggan: id_pelanggan,
              periode: period,
              nominal: customer.harga,
              status: 'belum_bayar',
              due_date: customer.due_date,
              nama: customer.nama,
              paket: customer.paket,
              status_tagihan: customer.status_tagihan
            });
          });
        });
      });
      return;
    }

    sendBillingResponse(results[0]);
  });
});

/* POST /api/customer/portal/pay - Upload payment proof */
router.post('/pay', function(req, res) {
  upload.single('bukti')(req, res, function(err) {
    if (err) {
      return res.status(400).json({ success: false, message: err.message });
    }

    var { id_tagihan } = req.body;
    var customerId = req.customerId;

    if (!id_tagihan) {
      return res.status(400).json({ success: false, message: 'ID Tagihan wajib disertakan.' });
    }

    if (!req.file) {
      return res.status(400).json({ success: false, message: 'Bukti transfer wajib diunggah.' });
    }

    // Verify that the bill belongs to the logged-in customer
    var verifySql = 'SELECT * FROM tagihan WHERE id_tagihan = ? AND id_pelanggan = ?';
    db.query(verifySql, [id_tagihan, customerId], function(err, results) {
      if (err) {
        return res.status(500).json({ success: false, message: 'Database error' });
      }

      if (results.length === 0) {
        if (req.file) {
          fs.unlink(req.file.path, () => {});
        }
        return res.status(403).json({ success: false, message: 'Akses ditolak. Tagihan bukan milik Anda.' });
      }

      var tagihan = results[0];
      if (tagihan.status === 'lunas') {
        if (req.file) {
          fs.unlink(req.file.path, () => {});
        }
        return res.status(400).json({ success: false, message: 'Tagihan ini sudah lunas.' });
      }
      if (tagihan.status === 'menunggu_verifikasi') {
        if (req.file) {
          fs.unlink(req.file.path, () => {});
        }
        return res.status(400).json({ success: false, message: 'Pembayaran untuk tagihan ini sedang menunggu verifikasi admin.' });
      }

      var relativePath = '/uploads/bukti/' + req.file.filename;

      // 1. Insert into pembayaran table
      var insertSql = `
        INSERT INTO pembayaran (id_tagihan, bukti_file, status, tanggal_upload) 
        VALUES (?, ?, 'pending', NOW())
      `;
      db.query(insertSql, [id_tagihan, relativePath], function(err, paymentResult) {
        if (err) {
          return res.status(500).json({ success: false, message: 'Gagal mencatat pembayaran.', error: err.message });
        }

        // 2. Update tagihan status to 'menunggu_verifikasi'
        var updateSql = "UPDATE tagihan SET status = 'menunggu_verifikasi' WHERE id_tagihan = ?";
        db.query(updateSql, [id_tagihan], function(err) {
          if (err) {
            return res.status(500).json({ success: false, message: 'Gagal memperbarui status tagihan.' });
          }

          // Broadcast notification to Admin dashboard (Tahap 6)
          SocketService.broadcast('pembayaran_masuk', {
            id_pembayaran: paymentResult.insertId,
            id_tagihan: id_tagihan,
            nama_pelanggan: results[0].nama,
            tanggal_upload: new Date()
          });

          res.json({
            success: true,
            message: 'Bukti transfer berhasil diunggah! Pembayaran Anda sedang menunggu verifikasi admin.',
            data: {
              id_pembayaran: paymentResult.insertId,
              bukti_file: relativePath
            }
          });
        });
      });
    });
  });
});

module.exports = router;
