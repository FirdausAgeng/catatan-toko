// ============================================================
// APP SCRIPT BACKEND: TOKO MANAGEMENT SYSTEM (V3 - WITH DETAIL & LOADING)
// ============================================================

function doGet() {
  var output = HtmlService.createTemplateFromFile('Index').evaluate();
  output.setTitle('Aplikasi Catatan Toko');
  output.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  return output;
}

function showSidebar() {
  var html = HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Catatan Toko');
  SpreadsheetApp.getUi().showSidebar(html);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(' App Toko')
    .addItem('Buka Sidebar App', 'showSidebar')
    .addToUi();
}

function getOrCreateSheet(sheetName, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f3f4f6');
  }
  return sheet;
}

// ------------------------------------------------------------
// 1. MODUL HUTANG
// ------------------------------------------------------------
function getDataHutang(keyword, statusFilter) {
  var sheet = getOrCreateSheet('Hutang', ['ID', 'Tanggal', 'Nama Pelanggan', 'Nominal', 'Keterangan', 'Status', 'Tanggal Lunas']);
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  keyword = (keyword || '').toLowerCase();
  statusFilter = statusFilter || 'Semua';

  var results = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var id = row[0];
    var tanggal = row[1] instanceof Date ? Utilities.formatDate(row[1], Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm") : row[1];
    var nama = String(row[2] || '');
    var nominal = row[3] || 0;
    var ket = String(row[4] || '');
    var status = String(row[5] || 'Belum Lunas');
    var tglLunas = row[6] instanceof Date ? Utilities.formatDate(row[6], Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm") : (row[6] || '-');

    if (statusFilter !== 'Semua' && status !== statusFilter) continue;
    if (keyword && !nama.toLowerCase().includes(keyword) && !ket.toLowerCase().includes(keyword) && !String(id).toLowerCase().includes(keyword)) continue;

    results.push({
      rowIndex: i + 1,
      id: id,
      tanggal: tanggal,
      nama: nama,
      nominal: nominal,
      keterangan: ket,
      status: status,
      tanggalLunas: tglLunas
    });
  }
  return results;
}

function updateStatusHutang(rowIndex, newStatus) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Hutang');
  if (!sheet) return { success: false, message: 'Sheet Hutang tidak ditemukan' };
  
  sheet.getRange(rowIndex, 6).setValue(newStatus);
  if (newStatus === 'Lunas') {
    sheet.getRange(rowIndex, 7).setValue(new Date());
  }
  return { success: true, message: 'Status hutang berhasil diubah menjadi ' + newStatus };
}

function deleteHutang(rowIndex) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Hutang');
  if (!sheet) return { success: false, message: 'Sheet Hutang tidak ditemukan' };
  sheet.deleteRow(rowIndex);
  return { success: true, message: 'Data hutang berhasil dihapus' };
}

function tambahHutang(form) {
  var sheet = getOrCreateSheet('Hutang', ['ID', 'Tanggal', 'Nama Pelanggan', 'Nominal', 'Keterangan', 'Status', 'Tanggal Lunas']);
  var id = 'HTG-' + new Date().getTime();
  sheet.appendRow([id, new Date(), form.nama, Number(form.nominal), form.keterangan, 'Belum Lunas', '-']);
  return { success: true, message: 'Catatan hutang berhasil ditambahkan' };
}

// ------------------------------------------------------------
// 2. MODUL KEMBALIAN
// ------------------------------------------------------------
function getDataKembalian(keyword, statusFilter) {
  var sheet = getOrCreateSheet('Kembalian', ['ID', 'Tanggal', 'Nama Pelanggan', 'Nominal Kembalian', 'Keterangan', 'Status', 'Tanggal Diambil']);
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  keyword = (keyword || '').toLowerCase();
  statusFilter = statusFilter || 'Semua';

  var results = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var id = row[0];
    var tanggal = row[1] instanceof Date ? Utilities.formatDate(row[1], Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm") : row[1];
    var nama = String(row[2] || '');
    var nominal = row[3] || 0;
    var ket = String(row[4] || '');
    var status = String(row[5] || 'Belum Diambil');
    var tglDiambil = row[6] instanceof Date ? Utilities.formatDate(row[6], Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm") : (row[6] || '-');

    if (statusFilter !== 'Semua' && status !== statusFilter) continue;
    if (keyword && !nama.toLowerCase().includes(keyword) && !ket.toLowerCase().includes(keyword) && !String(id).toLowerCase().includes(keyword)) continue;

    results.push({
      rowIndex: i + 1,
      id: id,
      tanggal: tanggal,
      nama: nama,
      nominal: nominal,
      keterangan: ket,
      status: status,
      tanggalDiambil: tglDiambil
    });
  }
  return results;
}

function updateStatusKembalian(rowIndex, newStatus) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Kembalian');
  if (!sheet) return { success: false, message: 'Sheet Kembalian tidak ditemukan' };
  
  sheet.getRange(rowIndex, 6).setValue(newStatus);
  if (newStatus === 'Sudah Diambil') {
    sheet.getRange(rowIndex, 7).setValue(new Date());
  }
  return { success: true, message: 'Status kembalian berhasil diubah menjadi ' + newStatus };
}

function deleteKembalian(rowIndex) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Kembalian');
  if (!sheet) return { success: false, message: 'Sheet Kembalian tidak ditemukan' };
  sheet.deleteRow(rowIndex);
  return { success: true, message: 'Data kembalian berhasil dihapus' };
}

function tambahKembalian(form) {
  var sheet = getOrCreateSheet('Kembalian', ['ID', 'Tanggal', 'Nama Pelanggan', 'Nominal Kembalian', 'Keterangan', 'Status', 'Tanggal Diambil']);
  var id = 'KMB-' + new Date().getTime();
  sheet.appendRow([id, new Date(), form.nama, Number(form.nominal), form.keterangan, 'Belum Diambil', '-']);
  return { success: true, message: 'Catatan kembalian berhasil ditambahkan' };
}

// ------------------------------------------------------------
// 3. MODUL SELISIH STOK
// ------------------------------------------------------------
function getDataSelisihStok(keyword) {
  var sheet = getOrCreateSheet('Selisih_Stok', ['ID', 'Tanggal', 'Nama Barang', 'Stok PC', 'Stok Real (Fisik)', 'Selisih', 'Keterangan / Catatan']);
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  keyword = (keyword || '').toLowerCase();

  var results = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var id = row[0];
    var tanggal = row[1] instanceof Date ? Utilities.formatDate(row[1], Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm") : row[1];
    var namaBarang = String(row[2] || '');
    var stokPc = row[3] || 0;
    var stokReal = row[4] || 0;
    var selisih = row[5] || (stokReal - stokPc);
    var ket = String(row[6] || '');

    if (keyword && !namaBarang.toLowerCase().includes(keyword) && !ket.toLowerCase().includes(keyword) && !String(id).toLowerCase().includes(keyword)) continue;

    results.push({
      rowIndex: i + 1,
      id: id,
      tanggal: tanggal,
      namaBarang: namaBarang,
      stokPc: stokPc,
      stokReal: stokReal,
      selisih: selisih,
      keterangan: ket
    });
  }
  return results;
}

function deleteSelisihStok(rowIndex) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Selisih_Stok');
  if (!sheet) return { success: false, message: 'Sheet Selisih_Stok tidak ditemukan' };
  sheet.deleteRow(rowIndex);
  return { success: true, message: 'Data stok berhasil dihapus' };
}

function tambahSelisihStok(form) {
  var sheet = getOrCreateSheet('Selisih_Stok', ['ID', 'Tanggal', 'Nama Barang', 'Stok PC', 'Stok Real (Fisik)', 'Selisih', 'Keterangan / Catatan']);
  var id = 'STK-' + new Date().getTime();
  var pc = Number(form.stokPc || 0);
  var real = Number(form.stokReal || 0);
  var diff = real - pc;
  sheet.appendRow([id, new Date(), form.namaBarang, pc, real, diff, form.keterangan]);
  return { success: true, message: 'Catatan selisih stok berhasil disimpan' };
}

// ------------------------------------------------------------
// 4. MODUL RIWAYAT / HISTORY
// ------------------------------------------------------------
function getRiwayatSemua(keyword) {
  var listHutang = getDataHutang(keyword, 'Lunas');
  var listKembalian = getDataKembalian(keyword, 'Sudah Diambil');
  
  var combined = [];
  
  listHutang.forEach(function(item) {
    combined.push({
      tipe: 'Hutang Lunas',
      id: item.id,
      tanggalAwal: item.tanggal,
      tanggalSelesai: item.tanggalLunas,
      nama: item.nama,
      nominal: item.nominal,
      keterangan: item.keterangan
    });
  });

  listKembalian.forEach(function(item) {
    combined.push({
      tipe: 'Kembalian Diambil',
      id: item.id,
      tanggalAwal: item.tanggal,
      tanggalSelesai: item.tanggalDiambil,
      nama: item.nama,
      nominal: item.nominal,
      keterangan: item.keterangan
    });
  });

  combined.sort(function(a, b) {
    return new Date(b.tanggalSelesai) - new Date(a.tanggalSelesai);
  });

  return combined;
}
