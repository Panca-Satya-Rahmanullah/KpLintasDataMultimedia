var express = require('express');
var router = express.Router();
var db = require('../config/db');
var Pengeluaran = require('../models/Pengeluaran');
var verifyToken = require('../middleware/auth');
var ExcelJS = require('exceljs');
var path = require('path'); // Wajib ditambahkan untuk memanggil path template

// Protect all report routes with admin token
router.use(verifyToken);

/* GET /api/reports/summary - Get financial summary for a period */
/* GET /api/reports/export-excel - EXPORT GABUNGAN TEMPLATE & TABEL BERDAMPINGAN */
/* GET /api/reports/export-excel - EXPORT GABUNGAN TEMPLATE & TABEL BERDAMPINGAN */
router.get('/export-excel', function (req, res) {
  var { periode } = req.query;
  if (!periode) {
    var today = new Date();
    var y = today.getFullYear();
    var m = today.getMonth() + 1;
    periode = y + '-' + (m < 10 ? '0' + m : m);
  }

  var incomeSql = `
    SELECT t.*, p.nama, p.email, p.no_hp 
    FROM tagihan t 
    JOIN pelanggan p ON t.id_pelanggan = p.id_pelanggan 
    WHERE t.status = 'lunas' AND t.periode = ?
    ORDER BY t.updated_at DESC
  `;

  db.query(incomeSql, [periode], function (err, incomes) {
    if (err) return res.status(500).send('Gagal mengambil data pemasukan.');

    Pengeluaran.getAll(periode, async function (err, expenses) {
      if (err) return res.status(500).send('Gagal mengambil data pengeluaran.');

      try {
        var totalPemasukan = incomes.reduce((sum, item) => sum + parseFloat(item.nominal), 0);
        var totalPengeluaran = expenses.reduce((sum, item) => sum + parseFloat(item.nominal), 0);
        var labaBersih = totalPemasukan - totalPengeluaran;

        const templatePath = path.join(__dirname, '../public/templates/template_laporan.xlsx');
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(templatePath);

        const sheet = workbook.getWorksheet('Dashboard');
        if (!sheet) return res.status(500).send('Sheet "Dashboard" tidak ditemukan.');

        // Tembakkan nilai Total (Font akan mengikuti template Excel)
        sheet.getCell('B8').value = totalPemasukan;
        sheet.getCell('D8').value = totalPengeluaran;
        sheet.getCell('E8').value = labaBersih;

        var borderThin = {
          top: { style: 'thin' }, bottom: { style: 'thin' },
          left: { style: 'thin' }, right: { style: 'thin' }
        };

        // Variable khusus untuk Font Times New Roman
        var fontTNR = { name: 'Times New Roman', size: 11 };
        var fontTNRBold = { name: 'Times New Roman', size: 11, bold: true };
        var fontTNRHeader = { name: 'Times New Roman', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };

        // ==========================================
        // TABEL DETAIL PEMASUKAN (Sisi Kiri)
        // ==========================================
        var incRow = 11; 
        
        sheet.getCell(`B${incRow}`).value = 'DETAIL PEMASUKAN (TAGIHAN LUNAS)';
        sheet.getCell(`B${incRow}`).font = { name: 'Times New Roman', bold: true, size: 12, color: { argb: 'FF1F497D' } };
        incRow++;

        var incomeHeaders = ['Tanggal Bayar', 'Nama Pelanggan', 'Email Pelanggan', 'Nominal'];
        incomeHeaders.forEach((h, idx) => {
          var cell = sheet.getCell(incRow, idx + 2); 
          cell.value = h;
          cell.font = fontTNRHeader;
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E78' } }; 
          cell.border = borderThin;
        });
        incRow++;

        incomes.forEach((item) => {
          var row = sheet.getRow(incRow);
          row.getCell(2).value = new Date(item.updated_at).toLocaleDateString('id-ID'); 
          row.getCell(3).value = item.nama;                                             
          row.getCell(4).value = item.email || '-';                                     
          
          var nominalCell = row.getCell(5);                                             
          nominalCell.value = parseFloat(item.nominal);
          nominalCell.numFormat = '"Rp"#,##0'; 
          
          // Terapkan Font Times New Roman ke setiap sel data
          [2, 3, 4, 5].forEach(col => {
              row.getCell(col).border = borderThin;
              row.getCell(col).font = fontTNR;
          });
          incRow++;
        });

        // ==========================================
        // TABEL DETAIL PENGELUARAN (Sisi Kanan)
        // ==========================================
        var expRow = 11; 
        
        sheet.getCell(`G${expRow}`).value = 'DETAIL PENGELUARAN OPERASIONAL';
        sheet.getCell(`G${expRow}`).font = { name: 'Times New Roman', bold: true, size: 12, color: { argb: 'FFC65911' } };
        expRow++;

        var expenseHeaders = ['Tanggal Pengeluaran', 'Keterangan / Tujuan', 'Nominal Pengeluaran'];
        expenseHeaders.forEach((h, idx) => {
          var cell = sheet.getCell(expRow, idx + 7); 
          cell.value = h;
          cell.font = fontTNRHeader;
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC65911' } }; 
          cell.border = borderThin;
        });
        expRow++;

        expenses.forEach((item) => {
          var row = sheet.getRow(expRow);
          row.getCell(7).value = new Date(item.tanggal).toLocaleDateString('id-ID');     
          row.getCell(8).value = item.keterangan || item.kategori;                       
          
          var nominalCell = row.getCell(9);                                              
          nominalCell.value = parseFloat(item.nominal);
          nominalCell.numFormat = '"Rp"#,##0'; 

          // Terapkan Font Times New Roman ke setiap sel data
          [7, 8, 9].forEach(col => {
              row.getCell(col).border = borderThin;
              row.getCell(col).font = fontTNR;
          });
          expRow++;
        });

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename=Laporan_Keuangan_${periode}.xlsx`);

        await workbook.xlsx.write(res);
        res.end();

      } catch (error) {
        console.error('Error:', error);
        res.status(500).send('Gagal mengekspor laporan.');
      }
    });
  });
});

module.exports = router;