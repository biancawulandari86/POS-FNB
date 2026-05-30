// ============================================
// Cafe POS - Google Apps Script Backend
// Database: Google Sheets
// Versi: Lengkap + Fitur Pesanan dari Meja
// ============================================

const SPREADSHEET_ID = '1tUmCEmS40muR9St6PJwfjmpSzGd3P3emE3oNxad-PUc';

// Kategori yang TIDAK ditampilkan di menu customer (bahan baku / operasional)
const KATEGORI_TERSEMBUNYI = ['Bahan Baku', 'Kemasan', 'Bahan Mentah', 'Operasional'];

// ============================================
// HELPER FUNCTIONS
// ============================================

function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function getSheet(sheetName) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    if (sheetName === 'Data_Barang') {
      sheet.appendRow(['ID_Barang', 'Barcode', 'Nama', 'Satuan', 'Harga_Beli', 'Harga_Rata', 'Harga_Jual', 'Stok', 'Total_Nilai', 'Terjual', 'Kategori', 'Deskripsi']);
    } else if (sheetName === 'Masuk_Gudang') {
      sheet.appendRow(['ID_Log', 'Tanggal', 'ID_Barang', 'Nama_Barang', 'Supplier', 'Jumlah_Masuk', 'Harga_Beli', 'Harga_Rata_Baru']);
    } else if (sheetName === 'Penjualan') {
      sheet.appendRow(['ID_Transaksi', 'Tanggal', 'Waktu', 'List_Barang', 'Total', 'Keuntungan']);
    }
  }
  return sheet;
}

function generateId(prefix, data, idColumn = 1) {
  if (data.length <= 1) return prefix + '001';
  const ids = data.slice(1).map(row => row[idColumn - 1]);
  const maxNum = Math.max(...ids.map(id => {
    const num = parseInt(String(id).replace(prefix, ''));
    return isNaN(num) ? 0 : num;
  }));
  return prefix + String(maxNum + 1).padStart(3, '0');
}

function getTodayDate() {
  const now = new Date();
  // Eksplisit WIB (Asia/Jakarta = UTC+7) agar konsisten antara simpan & baca
  return Utilities.formatDate(now, 'Asia/Jakarta', 'yyyy-MM-dd');
}

function getCurrentTime() {
  const now = new Date();
  return Utilities.formatDate(now, 'Asia/Jakarta', 'HH:mm');
}

// DEBUG: jalankan di Apps Script editor untuk cek isi sheet Pesanan_Meja
function debugPesananMeja() {
  const sheet = getSheetPesananMeja();
  const data = sheet.getDataRange().getValues();
  const today = getTodayDate();
  console.log('=== DEBUG PESANAN MEJA ===');
  console.log('Today (getTodayDate):', today);
  console.log('Total rows di sheet:', data.length - 1);
  if (data.length > 1) {
    data.slice(1).forEach((row, i) => {
      const tgl = row[1];
      const tglNorm = tgl instanceof Date
        ? Utilities.formatDate(tgl, 'Asia/Jakarta', 'yyyy-MM-dd')
        : String(tgl).substring(0, 10);
      console.log('Row ' + (i+1) + ': id=' + row[0] + ', tanggal_raw=' + tgl + ', tanggal_norm=' + tglNorm + ', match=' + (tglNorm === today) + ', status=' + row[7]);
    });
  }
  const hasil = getPesananMeja();
  console.log('Hasil getPesananMeja():', JSON.stringify(hasil));
}

// ============================================
// WEB APP ENDPOINTS
// ============================================

function doGet(e) {
  const page = e && e.parameter && e.parameter.page;

  if (page === 'order') {
    return HtmlService.createHtmlOutputFromFile('CustomerOrder')
      .setTitle('Pesan Menu')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Cafe POS - Sistem Kasir & Inventaris')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// doPost: handle request POST (opsional, untuk integrasi eksternal)
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    let result;
    if (action === 'simpanPesananMeja') {
      result = simpanPesananMeja(body.params);
    } else if (action === 'getPesananMeja') {
      result = getPesananMeja();
    } else {
      result = { success: false, message: 'Action tidak dikenal: ' + action };
    }
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================
// API: DATA BARANG
// Kolom 11 = Kategori, Kolom 12 = Deskripsi
// ============================================

function getDataBarang() {
  console.log('getDataBarang - Start');
  const sheet = getSheet('Data_Barang');
  const data = sheet.getDataRange().getValues();
  console.log('getDataBarang - Total rows:', data.length);
  if (data.length <= 1) return [];

  const headers = data[0];
  const colCount = headers.length;

  const result = data.slice(1).map(row => {
    const base = {
      id: row[0], barcode: row[1], nama: row[2], satuan: row[3],
      hargaBeli: row[4],
      hargaRata: row[5] || row[4],
      hargaJual: row[6],
      stok: row[7] || 0,
      totalNilai: row[8] || 0,
      terjual: row[9] || 0
    };
    if (colCount >= 11) base.kategori = row[10] || '';
    if (colCount >= 12) base.deskripsi = row[11] || '';
    return base;
  });

  console.log('getDataBarang - Returning', result.length, 'records');
  return result;
}

// ============================================
// API: MENU UNTUK CUSTOMER (halaman order meja)
// Filter: hanya kategori menu (bukan bahan baku/kemasan),
//         hanya yang stok > 0
// ============================================
function getMenuForCustomer() {
  const semua = getDataBarang();
  return semua.filter(b => {
    const kat = (b.kategori || '').trim();
    const tersembunyi = KATEGORI_TERSEMBUNYI.some(k => k.toLowerCase() === kat.toLowerCase());
    return !tersembunyi && b.stok > 0;
  });
}

function getBarangByBarcode(barcode) {
  const data = getDataBarang();
  return data.find(b => b.barcode === barcode);
}

function getBarangById(id) {
  const data = getDataBarang();
  return data.find(b => b.id === id);
}

function tambahBarang(params) {
  const sheet = getSheet('Data_Barang');
  const data = sheet.getDataRange().getValues();
  const newId = generateId('BRG', data);
  const hargaBeli = parseInt(params.hargaBeli);
  const stok = parseInt(params.stok);
  sheet.appendRow([
    newId, params.barcode, params.nama, params.satuan,
    hargaBeli, hargaBeli, parseInt(params.hargaJual),
    stok, hargaBeli * stok, 0,
    params.kategori || '', params.deskripsi || ''
  ]);
  return { success: true, id: newId, message: 'Barang berhasil ditambahkan' };
}

function updateHargaBarang(params) {
  const sheet = getSheet('Data_Barang');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === params.id) {
      sheet.getRange(i + 1, 7).setValue(parseInt(params.hargaJual));
      // Pastikan header kolom Kategori dan Deskripsi ada
      if (data[0].length < 11) sheet.getRange(1, 11).setValue('Kategori');
      if (data[0].length < 12) sheet.getRange(1, 12).setValue('Deskripsi');
      // Update kategori dan deskripsi jika dikirim
      if (params.kategori !== undefined) {
        sheet.getRange(i + 1, 11).setValue(params.kategori || '');
      }
      if (params.deskripsi !== undefined) {
        sheet.getRange(i + 1, 12).setValue(params.deskripsi || '');
      }
      return { success: true, message: 'Harga jual berhasil diupdate' };
    }
  }
  return { success: false, message: 'Barang tidak ditemukan' };
}

function updateStokBarang(params) {
  const sheet = getSheet('Data_Barang');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === params.id) {
      const stokLama = data[i][7] || 0;
      const hargaRata = data[i][5] || data[i][4] || 0;
      const stokBaru = stokLama + params.jumlah;
      sheet.getRange(i + 1, 8).setValue(stokBaru);
      const totalNilaiBaru = stokBaru * hargaRata;
      sheet.getRange(i + 1, 9).setValue(totalNilaiBaru);
      return { success: true, stokBaru, totalNilai: totalNilaiBaru };
    }
  }
  return { success: false, message: 'Barang tidak ditemukan' };
}

function updateTerjualBarang(id, jumlah) {
  const sheet = getSheet('Data_Barang');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      const terjualBaru = (data[i][9] || 0) + jumlah;
      sheet.getRange(i + 1, 10).setValue(terjualBaru);
      return { success: true };
    }
  }
  return { success: false };
}

// ============================================
// API: MASUK GUDANG
// ============================================

function getMasukGudang() {
  const sheet = getSheet('Masuk_Gudang');
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  return data.slice(1).map(row => {
    let tanggalStr = row[1];
    if (tanggalStr instanceof Date) {
      tanggalStr = Utilities.formatDate(tanggalStr, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    } else if (typeof tanggalStr !== 'string') {
      tanggalStr = String(tanggalStr);
    }
    return {
      id: row[0], tanggal: tanggalStr, idBarang: row[2], namaBarang: row[3],
      supplier: row[4], jumlah: row[5], hargaBeli: row[6] || 0, hargaRata: row[7] || 0
    };
  });
}

function tambahStokMasuk(params) {
  const sheet = getSheet('Masuk_Gudang');
  const sheetData = sheet.getDataRange().getValues();
  const newId = generateId('LOG', sheetData);
  const barangSheet = getSheet('Data_Barang');
  const barangData = barangSheet.getDataRange().getValues();
  let barangRow = -1;
  for (let i = 1; i < barangData.length; i++) {
    if (barangData[i][0] === params.idBarang) { barangRow = i; break; }
  }
  if (barangRow === -1) return { success: false, message: 'Barang tidak ditemukan' };

  const barang = barangData[barangRow];
  const stokLama = barang[7] || 0;
  const hargaRataLama = barang[5] || barang[4];
  const totalNilaiLama = barang[8] || (stokLama * hargaRataLama);
  const jumlahMasuk = parseInt(params.jumlah);
  const hargaBeliBaru = parseInt(params.hargaBeli) || hargaRataLama;
  const totalNilaiBaru = totalNilaiLama + (jumlahMasuk * hargaBeliBaru);
  const stokBaru = stokLama + jumlahMasuk;
  const hargaRataBaru = Math.round(totalNilaiBaru / stokBaru);

  barangSheet.getRange(barangRow + 1, 5).setValue(hargaBeliBaru);
  barangSheet.getRange(barangRow + 1, 6).setValue(hargaRataBaru);
  barangSheet.getRange(barangRow + 1, 8).setValue(stokBaru);
  barangSheet.getRange(barangRow + 1, 9).setValue(totalNilaiBaru);
  sheet.appendRow([newId, getTodayDate(), params.idBarang, barang[2], params.supplier, jumlahMasuk, hargaBeliBaru, hargaRataBaru]);
  return { success: true, id: newId, hargaRata: hargaRataBaru, message: 'Stok masuk berhasil dicatat. Harga rata: ' + hargaRataBaru };
}

// ============================================
// API: PENJUALAN
// ============================================

function getPenjualan() {
  try {
    const sheet = getSheet('Penjualan');
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    return data.slice(1).reverse().map(row => {
      try {
        let tanggalStr = row[1];
        // Selalu gunakan Asia/Jakarta agar konsisten dengan getTodayDate()
        if (tanggalStr instanceof Date) {
          tanggalStr = Utilities.formatDate(tanggalStr, 'Asia/Jakarta', 'yyyy-MM-dd');
        } else {
          // Normalisasi string: pastikan format yyyy-MM-dd dengan zero-padding
          tanggalStr = String(tanggalStr).trim();
          // Handle format yyyy-M-d → yyyy-MM-dd
          const parts = tanggalStr.split('-');
          if (parts.length === 3) {
            tanggalStr = parts[0] + '-' + parts[1].padStart(2,'0') + '-' + parts[2].padStart(2,'0');
          }
        }
        const waktuStr = row[2] instanceof Date
          ? Utilities.formatDate(row[2], 'Asia/Jakarta', 'HH:mm')
          : String(row[2] || '');
        return {
          id: String(row[0] || ''),
          tanggal: tanggalStr || '',
          waktu: waktuStr,
          listBarang: String(row[3] || ''),
          total: Number(row[4]) || 0,
          keuntungan: Number(row[5]) || 0
        };
      } catch (e) {
        console.error('getPenjualan row error:', e, row);
        return null;
      }
    }).filter(item => item !== null);
  } catch (error) {
    console.error('getPenjualan - ERROR:', error);
    return [];
  }
}

function getPenjualanByDate(params) {
  return getPenjualan().filter(p => p.tanggal === params.tanggal);
}

function simpanPenjualan(params) {
  const sheet = getSheet('Penjualan');
  const sheetData = sheet.getDataRange().getValues();
  const newId = generateId('TRX', sheetData);
  const items = JSON.parse(params.items);
  let totalKeuntungan = 0;

  for (const item of items) {
    const barang = getBarangById(item.id);
    if (!barang) return { success: false, message: 'Barang tidak ditemukan: ' + item.nama };
    totalKeuntungan += (item.harga - barang.hargaRata) * item.qty;
    const result = updateStokBarang({ id: item.id, jumlah: -item.qty });
    if (!result.success) return { success: false, message: 'Gagal mengupdate stok untuk ' + item.nama };
    updateTerjualBarang(item.id, item.qty);
  }

  sheet.appendRow([newId, getTodayDate(), getCurrentTime(), params.listBarang, parseInt(params.total), totalKeuntungan]);
  return { success: true, id: newId, tanggal: getTodayDate(), waktu: getCurrentTime(), keuntungan: totalKeuntungan, message: 'Transaksi berhasil disimpan' };
}

// ============================================
// API: PENJUALAN DARI PESANAN MEJA
// Kasir klik "Selesai & Bayar" → langsung masuk ke sheet Penjualan
// ============================================
function selesaikanPesananMeja(params) {
  try {
    // 1. Update status pesanan jadi SELESAI
    updateStatusPesananMeja({ id: params.pesananId, status: 'SELESAI' });

    // 2. Parse item-item dari listBarang pesanan
    //    Format: "Cappuccino (2), Nasi Goreng, Teh Susu (3)"
    const listBarang = params.listBarang || '';
    const total = parseInt(params.total) || 0;

    // 3. Simpan ke sheet Penjualan (tanpa update stok — stok diupdate manual oleh kasir)
    //    Jika ingin otomatis kurangi stok, gunakan simpanPenjualan dengan items JSON
    const sheet = getSheet('Penjualan');
    const sheetData = sheet.getDataRange().getValues();
    const newId = generateId('TRX', sheetData);

    // Hitung keuntungan kasar (opsional — 0 jika item tidak dikenali)
    let keuntungan = 0;
    const semuaBarang = getDataBarang();
    const parts = listBarang.split(',').map(s => s.trim());
    for (const part of parts) {
      const match = part.match(/^(.+?)\s*\((\d+)\)$/);
      const nama = match ? match[1].trim() : part.trim();
      const qty  = match ? parseInt(match[2]) : 1;
      const found = semuaBarang.find(b => b.nama === nama);
      if (found) {
        keuntungan += (found.hargaJual - (found.hargaRata || found.hargaBeli)) * qty;
        // Kurangi stok otomatis
        updateStokBarang({ id: found.id, jumlah: -qty });
        updateTerjualBarang(found.id, qty);
      }
    }

    sheet.appendRow([newId, getTodayDate(), getCurrentTime(), listBarang, total, keuntungan]);
    return { success: true, id: newId, message: 'Pesanan selesai & masuk ke transaksi!' };
  } catch (err) {
    console.error('selesaikanPesananMeja error:', err);
    return { success: false, message: err.message };
  }
}

// ============================================
// API: DASHBOARD STATS
// ============================================

function getDashboardStats() {
  const barang = getDataBarang();
  const penjualan = getPenjualan();
  const today = getTodayDate();
  const penjualanHariIni = penjualan.filter(p => String(p.tanggal) === today);
  return {
    totalBarang: barang.length || 0,
    stokMenipis: barang.filter(b => (b.stok || 0) < 10).length,
    penjualanHariIni: penjualanHariIni.reduce((sum, p) => sum + (p.total || 0), 0),
    keuntunganHariIni: penjualanHariIni.reduce((sum, p) => sum + (p.keuntungan || 0), 0),
    transaksiHariIni: penjualanHariIni.length
  };
}

// ============================================
// API: STATISTIK BARANG
// ============================================

function getStatistikBarang() {
  const barang = getDataBarang();
  if (!barang || barang.length === 0) return { terlaris: [], kurangLaku: [] };
  const terlaris = [...barang]
    .filter(b => (b.terjual || 0) > 0)
    .sort((a, b) => (b.terjual || 0) - (a.terjual || 0))
    .slice(0, 5)
    .map(b => ({ idBarang: b.id, nama: b.nama, terjual: b.terjual || 0, stok: b.stok || 0 }));
  const kurangLaku = [...barang]
    .filter(b => (b.terjual || 0) < 50)
    .sort((a, b) => (a.terjual || 0) - (b.terjual || 0))
    .slice(0, 5)
    .map(b => ({ idBarang: b.id, nama: b.nama, terjual: b.terjual || 0, stok: b.stok || 0 }));
  return { terlaris, kurangLaku };
}

// ============================================
// API: GRAFIK KEUNTUNGAN
// ============================================

function getGrafikKeuntungan(params) {
  const tahun = params.tahun || '2026';
  const bulan = params.bulan || 'all';
  const penjualan = getPenjualan();
  let filtered = penjualan.filter(p => String(p.tanggal).startsWith(tahun));
  if (bulan !== 'all') {
    filtered = filtered.filter(p => String(p.tanggal).split('-')[1] === bulan.padStart(2, '0'));
    const daysInMonth = new Date(parseInt(tahun), parseInt(bulan), 0).getDate();
    const labels = [], data = [];
    for (let i = 1; i <= daysInMonth; i++) {
      labels.push(i.toString());
      const dateStr = `${tahun}-${bulan.padStart(2,'0')}-${i.toString().padStart(2,'0')}`;
      data.push(filtered.filter(p => p.tanggal === dateStr).reduce((sum, p) => sum + p.keuntungan, 0));
    }
    return { labels, data };
  } else {
    const namaBulan = ['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Agu','Sep','Okt','Nov','Des'];
    const data = [];
    for (let i = 1; i <= 12; i++) {
      const monthStr = i.toString().padStart(2, '0');
      data.push(filtered.filter(p => p.tanggal.startsWith(`${tahun}-${monthStr}`)).reduce((sum, p) => sum + p.keuntungan, 0));
    }
    return { labels: namaBulan, data };
  }
}

// ============================================
// API: PESANAN DARI MEJA
// Sheet: Pesanan_Meja
// Kolom: ID_Pesanan | Tanggal | Waktu | No_Meja | List_Barang | Total | Catatan | Status
// ============================================

function getSheetPesananMeja() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName('Pesanan_Meja');
  if (!sheet) {
    sheet = ss.insertSheet('Pesanan_Meja');
    sheet.appendRow(['ID_Pesanan', 'Tanggal', 'Waktu', 'No_Meja', 'List_Barang', 'Total', 'Catatan', 'Status', 'Bukti_Transfer_URL', 'Metode_Bayar']);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 10).setBackground('#c45d1a').setFontColor('white').setFontWeight('bold');
  } else {
    // Migrasi otomatis: tambah/rename kolom jika belum ada
    const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    // Rename Bukti_Transfer lama → Bukti_Transfer_URL jika masih pakai nama lama
    const idxLama = header.indexOf('Bukti_Transfer');
    if (idxLama >= 0 && !header.includes('Bukti_Transfer_URL')) {
      sheet.getRange(1, idxLama + 1).setValue('Bukti_Transfer_URL');
    }
    if (!header.includes('Bukti_Transfer_URL') && !header.includes('Bukti_Transfer')) {
      const col = sheet.getLastColumn() + 1;
      sheet.getRange(1, col).setValue('Bukti_Transfer_URL').setBackground('#c45d1a').setFontColor('white').setFontWeight('bold');
    }
    if (!header.includes('Metode_Bayar')) {
      const col = sheet.getLastColumn() + 1;
      sheet.getRange(1, col).setValue('Metode_Bayar').setBackground('#c45d1a').setFontColor('white').setFontWeight('bold');
    }
  }
  // Paksa kolom Tanggal (B) dan Waktu (C) sebagai plain text
  sheet.getRange('B:B').setNumberFormat('@STRING@');
  sheet.getRange('C:C').setNumberFormat('@STRING@');
  return sheet;
}

// Ambil atau buat folder Google Drive untuk bukti transfer
function getFolderBuktiTransfer() {
  const FOLDER_NAME = 'Cafe_POS_Bukti_Transfer';
  const folders = DriveApp.getFoldersByName(FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(FOLDER_NAME);
}

// Dipanggil oleh CustomerOrder.html saat pelanggan menekan "Kirim Pesanan"
function simpanPesananMeja(params) {
  try {
    const sheet = getSheetPesananMeja();
    const data = sheet.getDataRange().getValues();
    const newId = generateId('PMJ', data);
    const tanggalStr = getTodayDate();   // 'yyyy-MM-dd' string
    const waktuStr   = getCurrentTime(); // 'HH:mm' string

    const metodeBayar = params.metodeBayar || 'cash';
    // Jika transfer → status MENUNGGU_BUKTI sampai customer upload foto
    // Jika cash → langsung BARU (muncul di kasir)
    const statusAwal = (metodeBayar === 'transfer') ? 'MENUNGGU_BUKTI' : 'BARU';

    const lastRow = sheet.getLastRow() + 1;
    const rowData = [
      newId,
      tanggalStr,
      waktuStr,
      String(params.noMeja || '?'),
      params.listBarang || '',
      parseInt(params.total) || 0,
      params.catatan || '',
      statusAwal,
      '',   // Bukti_Transfer_URL (diisi nanti via simpanBuktiTransfer)
      metodeBayar
    ];
    sheet.getRange(lastRow, 1, 1, rowData.length).setValues([rowData]);

    // Pastikan kolom tanggal & waktu baris ini tetap text
    sheet.getRange(lastRow, 2).setNumberFormat('@STRING@');
    sheet.getRange(lastRow, 3).setNumberFormat('@STRING@');

    console.log('simpanPesananMeja OK:', newId, tanggalStr, waktuStr, 'metode:', metodeBayar, 'status:', statusAwal);
    return { success: true, id: newId, message: 'Pesanan berhasil dikirim ke kasir!' };
  } catch (err) {
    console.error('simpanPesananMeja error:', err);
    return { success: false, message: 'Terjadi kesalahan: ' + err.message };
  }
}

// Dipanggil setelah pesanan tersimpan — customer upload bukti transfer (base64 image)
// Foto disimpan ke Google Drive → hanya URL yang disimpan ke Sheets (hindari limit 50k karakter)
function simpanBuktiTransfer(params) {
  try {
    // params: { pesananId, buktiBase64, mimeType }
    if (!params.pesananId) return { success: false, message: 'pesananId tidak dikirim' };
    if (!params.buktiBase64) return { success: false, message: 'Data foto tidak dikirim' };

    // 1. Simpan foto ke Google Drive
    const folder = getFolderBuktiTransfer();
    const mimeType = params.mimeType || 'image/jpeg';
    const ext = mimeType.includes('png') ? 'png' : 'jpg';
    const namaFile = 'bukti_' + params.pesananId + '_' + new Date().getTime() + '.' + ext;

    // Strip prefix data:image/...;base64, jika ada
    const base64Clean = params.buktiBase64.replace(/^data:[^;]+;base64,/, '');
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Clean), mimeType, namaFile);
    const file = folder.createFile(blob);

    // 2. Set permission agar bisa dilihat siapa saja yang punya link (untuk kasir)
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const fileUrl = file.getUrl();
    const fileId  = file.getId();
    // URL thumbnail langsung (embed di halaman kasir tanpa redirect)
    const thumbUrl = 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w800';

    // 3. Update sheet: simpan URL + ubah status jadi BARU
    const sheet = getSheetPesananMeja();
    const data = sheet.getDataRange().getValues();
    const header = data[0];
    const colBukti  = header.indexOf('Bukti_Transfer_URL') >= 0
                      ? header.indexOf('Bukti_Transfer_URL') + 1
                      : header.indexOf('Bukti_Transfer') + 1; // fallback nama lama
    const colStatus = 8; // kolom Status (1-based, kolom H)

    if (colBukti <= 0) return { success: false, message: 'Kolom Bukti_Transfer_URL tidak ditemukan' };

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === params.pesananId) {
        sheet.getRange(i + 1, colBukti).setValue(thumbUrl);
        // Ubah status → BARU supaya kasir baru melihat pesanan ini
        sheet.getRange(i + 1, colStatus).setValue('BARU');
        console.log('simpanBuktiTransfer OK: row', i + 1, 'url:', thumbUrl);
        return { success: true, message: 'Bukti transfer berhasil dikirim ke kasir!', url: thumbUrl };
      }
    }
    return { success: false, message: 'Pesanan tidak ditemukan: ' + params.pesananId };
  } catch (err) {
    console.error('simpanBuktiTransfer error:', err);
    return { success: false, message: err.message };
  }
}

// Helper: normalisasi nilai tanggal dari Sheets ke string 'yyyy-MM-dd'
function normalizeTanggal(tgl) {
  if (tgl instanceof Date) {
    return Utilities.formatDate(tgl, 'Asia/Jakarta', 'yyyy-MM-dd');
  }
  // Jika sudah string, ambil 10 karakter pertama (yyyy-MM-dd)
  return String(tgl).substring(0, 10);
}

// Helper: normalisasi waktu dari Sheets ke string 'HH:mm'
function normalizeWaktu(waktu) {
  if (waktu instanceof Date) {
    return Utilities.formatDate(waktu, 'Asia/Jakarta', 'HH:mm');
  }
  // Jika sudah string (misal "22:43"), kembalikan apa adanya
  return String(waktu);
}

// Dipanggil oleh halaman kasir untuk polling notifikasi pesanan masuk
function getPesananMeja() {
  try {
    const sheet = getSheetPesananMeja();
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      console.log('getPesananMeja: sheet kosong');
      return [];
    }
    const today = getTodayDate(); // format: 'yyyy-MM-dd'
    console.log('getPesananMeja: today=' + today + ', total rows=' + (data.length - 1));

    const semua = data.slice(1);
    semua.forEach((row, i) => {
      const tglRaw = row[1];
      const tglNorm = normalizeTanggal(tglRaw);
      console.log('  row' + (i+1) + ': id=' + row[0] + ' tgl_raw=' + tglRaw + ' (' + typeof tglRaw + ') tgl_norm=' + tglNorm + ' match=' + (tglNorm === today) + ' status=' + row[7]);
    });

    const hasil = semua
      .filter(row => normalizeTanggal(row[1]) === today)
      .reverse()
      .map(row => ({
        id: row[0],
        tanggal: normalizeTanggal(row[1]),
        waktu: normalizeWaktu(row[2]),
        noMeja: row[3],
        listBarang: row[4],
        total: row[5] || 0,
        catatan: row[6] || '',
        status: row[7] || 'BARU',
        buktiTransfer: row[8] || '',   // ini sekarang URL Google Drive (thumbnail)
        metodeBayar: row[9] || 'cash'
      }));

    console.log('getPesananMeja: hasil=' + hasil.length + ' pesanan lolos filter');
    return hasil;
  } catch (err) {
    console.error('getPesananMeja error:', err);
    return [];
  }
}

// Kasir klik "Proses" atau "Selesai"
function updateStatusPesananMeja(params) {
  try {
    const sheet = getSheetPesananMeja();
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === params.id) {
        sheet.getRange(i + 1, 8).setValue(params.status);
        return { success: true };
      }
    }
    return { success: false, message: 'Pesanan tidak ditemukan' };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

// ============================================
// API: PENGATURAN TRANSFER / QRIS
// Disimpan di sheet "Settings" baris dengan key "transfer_settings"
// ============================================

function getSheetSettings() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName('Settings');
  if (!sheet) {
    sheet = ss.insertSheet('Settings');
    sheet.appendRow(['Key', 'Value']);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 2).setBackground('#c45d1a').setFontColor('white').setFontWeight('bold');
  }
  return sheet;
}

function getTransferSettings() {
  try {
    const sheet = getSheetSettings();
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === 'transfer_settings') {
        const parsed = JSON.parse(data[i][1] || '{}');
        return parsed;
      }
    }
    return { aktif: false };
  } catch (err) {
    console.error('getTransferSettings error:', err);
    return { aktif: false };
  }
}

function simpanTransferSettings(params) {
  try {
    const sheet = getSheetSettings();
    const data = sheet.getDataRange().getValues();

    // Pertahankan qrisImageUrl yang sudah ada jika params tidak kirim field tsb
    let qrisImageUrlLama = '';
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === 'transfer_settings') {
        try {
          const existing = JSON.parse(data[i][1] || '{}');
          qrisImageUrlLama = existing.qrisImageUrl || '';
        } catch(e) {}
        break;
      }
    }

    const jsonVal = JSON.stringify({
      aktif: params.aktif === true || params.aktif === 'true',
      namaBank: params.namaBank || '',
      noRekening: params.noRekening || '',
      atasNama: params.atasNama || '',
      catatan: params.catatan || '',
      // Jika params.qrisImageUrl dikirim (termasuk string kosong untuk hapus),
      // gunakan nilai itu; jika tidak dikirim sama sekali, pakai nilai lama.
      qrisImageUrl: (params.qrisImageUrl !== undefined) ? (params.qrisImageUrl || '') : qrisImageUrlLama
    });

    // Cari baris yang sudah ada, update jika ketemu
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === 'transfer_settings') {
        sheet.getRange(i + 1, 2).setValue(jsonVal);
        return { success: true, message: 'Pengaturan transfer berhasil disimpan' };
      }
    }
    // Belum ada, tambah baris baru
    sheet.appendRow(['transfer_settings', jsonVal]);
    return { success: true, message: 'Pengaturan transfer berhasil disimpan' };
  } catch (err) {
    console.error('simpanTransferSettings error:', err);
    return { success: false, message: err.message };
  }
}

// ============================================
// API: SIMPAN & AMBIL GAMBAR QRIS
// Gambar disimpan ke Google Drive, URL-nya disimpan di Settings
// ============================================

// Folder khusus gambar QRIS (terpisah dari bukti transfer)
function getFolderQris() {
  const FOLDER_NAME = 'Cafe_POS_QRIS';
  const folders = DriveApp.getFoldersByName(FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(FOLDER_NAME);
}

/**
 * Dipanggil dari halaman Settings saat kasir upload gambar QRIS.
 * params: { imageBase64: string, mimeType: string }
 * Mengembalikan: { success, thumbUrl, downloadUrl }
 */
function simpanQrisImage(params) {
  try {
    if (!params.imageBase64) return { success: false, message: 'Data gambar tidak dikirim' };

    const folder = getFolderQris();
    const mimeType = params.mimeType || 'image/jpeg';
    const ext = mimeType.includes('png') ? 'png' : 'jpg';
    const namaFile = 'qris_' + new Date().getTime() + '.' + ext;

    // Hapus file QRIS lama jika ada (agar tidak menumpuk di Drive)
    const filesLama = folder.getFiles();
    while (filesLama.hasNext()) {
      try { filesLama.next().setTrashed(true); } catch(e) {}
    }

    // Strip prefix data:image/...;base64, jika ada
    const base64Clean = params.imageBase64.replace(/^data:[^;]+;base64,/, '');
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Clean), mimeType, namaFile);
    const file = folder.createFile(blob);

    // Izinkan siapa saja dengan link (untuk ditampilkan ke customer)
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const fileId = file.getId();
    const thumbUrl = 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w800';
    const downloadUrl = 'https://drive.google.com/uc?export=download&id=' + fileId;

    // Simpan URL ke Settings agar bisa diambil via getQrisImageUrl()
    simpanTransferSettings({
      // Ambil data transfer yang sudah ada, tambahkan qrisImageUrl
      ...getTransferSettings(),
      qrisImageUrl: thumbUrl,
      qrisDownloadUrl: downloadUrl
    });

    console.log('simpanQrisImage OK: fileId=' + fileId);
    return { success: true, thumbUrl, downloadUrl, message: 'Gambar QRIS berhasil disimpan' };
  } catch (err) {
    console.error('simpanQrisImage error:', err);
    return { success: false, message: err.message };
  }
}

/**
 * Dipanggil dari CustomerOrder.html dan halaman Settings
 * untuk mendapatkan URL gambar QRIS yang tersimpan.
 * Mengembalikan: { thumbUrl, downloadUrl } atau {} jika belum ada
 */
function getQrisImageUrl() {
  try {
    const settings = getTransferSettings();
    if (settings && settings.qrisImageUrl) {
      return {
        thumbUrl: settings.qrisImageUrl,
        downloadUrl: settings.qrisDownloadUrl || settings.qrisImageUrl
      };
    }
    return {};
  } catch (err) {
    console.error('getQrisImageUrl error:', err);
    return {};
  }
}

// ============================================
// API: QR CODE MEJA
// ============================================

// Kembalikan base URL aplikasi (tanpa parameter meja)
function getBaseUrl() {
  return ScriptApp.getService().getUrl();
}

// Kembalikan URL lengkap untuk meja tertentu
function generateUrlMeja(noMeja) {
  const baseUrl = ScriptApp.getService().getUrl();
  return baseUrl + '?page=order&meja=' + noMeja;
}

// Generate URL untuk semua meja (1-20), log ke console
function generateSemuaUrlMeja() {
  const baseUrl = ScriptApp.getService().getUrl();
  for (let i = 1; i <= 20; i++) {
    console.log('Meja ' + i + ': ' + baseUrl + '?page=order&meja=' + i);
  }
}

// ============================================
// SEED DATA
// ============================================

function seedData() {
  const ss = getSpreadsheet();
  let sheetBarang = ss.getSheetByName('Data_Barang');
  let sheetMasuk = ss.getSheetByName('Masuk_Gudang');
  let sheetPenjualan = ss.getSheetByName('Penjualan');
  if (sheetBarang) ss.deleteSheet(sheetBarang);
  if (sheetMasuk) ss.deleteSheet(sheetMasuk);
  if (sheetPenjualan) ss.deleteSheet(sheetPenjualan);

  sheetBarang = ss.insertSheet('Data_Barang');
  sheetBarang.appendRow(['ID_Barang', 'Barcode', 'Nama', 'Satuan', 'Harga_Beli', 'Harga_Rata', 'Harga_Jual', 'Stok', 'Total_Nilai', 'Terjual', 'Kategori', 'Deskripsi']);
  sheetMasuk = ss.insertSheet('Masuk_Gudang');
  sheetMasuk.appendRow(['ID_Log', 'Tanggal', 'ID_Barang', 'Nama_Barang', 'Supplier', 'Jumlah_Masuk', 'Harga_Beli', 'Harga_Rata_Baru']);
  sheetPenjualan = ss.insertSheet('Penjualan');
  sheetPenjualan.appendRow(['ID_Transaksi', 'Tanggal', 'Waktu', 'List_Barang', 'Total', 'Keuntungan']);

  const produkData = generate500Produk();
  for (let i = 0; i < produkData.length; i++) {
    const p = produkData[i];
    sheetBarang.appendRow([p.id, p.barcode, p.nama, p.satuan, p.hargaBeli, p.hargaBeli, p.hargaJual, p.stok, p.hargaBeli * p.stok, 0, p.kategori || '', '']);
  }

  const transaksiData = generate100Transaksi(produkData);
  for (const t of transaksiData) {
    sheetPenjualan.appendRow([t.id, t.tanggal, t.waktu, t.listBarang, t.total, t.keuntungan]);
  }
  updateTerjaldDariTransaksi(transaksiData, sheetBarang);
  return { success: true, produk: produkData.length, transaksi: transaksiData.length };
}

function generate500Produk() {
  const kategoriMap = {
    'Kopi': ['Espresso','Americano','Cappuccino','Latte','Flat White','Macchiato','Mocha','Cold Brew','Iced Latte','Vietnamese Coffee'],
    'Non-Kopi': ['Matcha Latte','Teh Tarik','Teh Susu','Coklat Panas','Coklat Dingin','Milo','Susu Murni','Smoothie','Jus Segar','Lemonade'],
    'Minuman Soda': ['Soda Gembira','Italian Soda','Sparkling Water','Teh Soda','Kopi Soda','Yuzu Soda','Lychee Soda','Strawberry Soda'],
    'Makanan': ['Nasi Goreng','Mie Goreng','Pasta','Sandwich','Burger','Wrap','Rice Bowl','Nasi Ayam','Mie Ayam','Soto'],
    'Pastry': ['Croissant','Roti Bakar','Waffle','Pancake','Donat','Muffin','Cookies','Brownie','Cheesecake','Toast'],
    'Dessert': ['Es Krim','Pudding','Crepes','Affogato','Ice Blended','Banana Foster','Tiramisu','Panna Cotta'],
    'Bahan Baku': ['Biji Kopi Arabica','Biji Kopi Robusta','Kopi Blend','Matcha Powder','Coklat Bubuk','Sirup Vanilla','Susu Full Cream','Susu Oat'],
    'Kemasan': ['Cup Plastik','Cup Kertas Hot','Sedotan Kertas','Kantong Kertas','Paper Bag','Tisu','Stirrer','Lid Cup']
  };
  const satuanList = ['cup','pcs','porsi','slice','box','pack'];
  const produk = [];
  let idCounter = 1;
  const katKeys = Object.keys(kategoriMap);

  for (let i = 0; i < 500; i++) {
    const kat = katKeys[Math.floor(Math.random() * katKeys.length)];
    const itemList = kategoriMap[kat];
    const item = itemList[Math.floor(Math.random() * itemList.length)];
    const size = Math.floor(Math.random() * 10) + 1;
    const sizeUnit = ['gr','kg','ml','L','pcs'][Math.floor(Math.random() * 5)];
    const barcode = '899' + String(Math.floor(Math.random() * 100000000000)).padStart(9, '0');
    const hargaBeli = Math.floor(Math.random() * 100000) + 2000;
    const margin = 0.1 + (Math.random() * 0.3);
    const hargaJual = Math.ceil(hargaBeli * (1 + margin) / 100) * 100;
    const stok = Math.floor(Math.random() * 200) + 10;
    produk.push({
      id: 'BRG' + String(idCounter).padStart(3, '0'),
      barcode, nama: `${item} ${size}${sizeUnit}`,
      satuan: satuanList[Math.floor(Math.random() * satuanList.length)],
      hargaBeli, hargaJual, stok, kategori: kat
    });
    idCounter++;
  }
  return produk;
}

function generate100Transaksi(produkData) {
  const transaksi = [];
  const today = new Date('2026-03-13');
  for (let i = 0; i < 100; i++) {
    const daysAgo = Math.floor(Math.random() * 365);
    const tgl = new Date(today);
    tgl.setDate(tgl.getDate() - daysAgo);
    const tanggal = Utilities.formatDate(tgl, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const jam = String(Math.floor(Math.random() * 14) + 8).padStart(2, '0');
    const menit = String(Math.floor(Math.random() * 60)).padStart(2, '0');
    const jumlahItem = Math.floor(Math.random() * 5) + 1;
    const items = [];
    let total = 0, keuntungan = 0;
    for (let j = 0; j < jumlahItem; j++) {
      const barang = produkData[Math.floor(Math.random() * produkData.length)];
      const qty = Math.floor(Math.random() * 5) + 1;
      items.push({ nama: barang.nama, qty });
      total += barang.hargaJual * qty;
      keuntungan += (barang.hargaJual - barang.hargaBeli) * qty;
    }
    transaksi.push({
      id: 'TRX' + String(i + 1).padStart(3, '0'),
      tanggal, waktu: `${jam}:${menit}`,
      listBarang: items.map(it => it.qty > 1 ? `${it.nama} (${it.qty})` : it.nama).join(', '),
      total, keuntungan
    });
  }
  return transaksi;
}

function updateTerjaldDariTransaksi(transaksiData, sheetBarang) {
  const terjualCount = {};
  for (const t of transaksiData) {
    const items = t.listBarang.split(', ');
    for (const item of items) {
      const match = item.match(/(.+)\s*\((\d+)\)$/);
      if (match) { terjualCount[match[1]] = (terjualCount[match[1]] || 0) + parseInt(match[2]); }
      else { terjualCount[item] = (terjualCount[item] || 0) + 1; }
    }
  }
  const data = sheetBarang.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const terjual = terjualCount[data[i][2]] || 0;
    if (terjual > 0) sheetBarang.getRange(i + 1, 10).setValue(terjual);
  }
}

function setupDatabase() {
  return seedData();
}