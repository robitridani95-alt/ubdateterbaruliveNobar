/**
 * ============================================================
 * LIVE SPORT
 * Backend Google Apps Script (JSON API)
 * Database: Google Sheets
 * Frontend: di-hosting terpisah (misalnya di Vercel), lihat index.html
 * ============================================================
 * CARA SETUP:
 * 1. Buat Google Spreadsheet baru, salin ID-nya (bagian di URL antara /d/ dan /edit)
 * 2. Tempel ID tersebut ke variabel SHEET_ID di bawah ini
 * 3. Jalankan fungsi setupSheets() sekali lewat menu Run > setupSheets
 *    (ini akan otomatis membuat semua sheet & header yang dibutuhkan)
 * 4. Deploy > New deployment > Web app > Execute as: Me, Who has access: Anyone
 * 5. Salin URL Web App-nya, tempel ke variabel API_URL di index.html
 * 6. Deploy index.html ke Vercel (atau hosting statis lain)
 *
 * ALUR LOGIN:
 * - Satu form login dipakai bersama oleh Admin, Pendaftar, dan Donatur.
 * - Pendaftar login pakai No. HP + Password yang diisi saat mendaftar.
 *   Setelah diverifikasi, pendaftar melihat link streaming sesuai cabang
 *   olahraga yang didaftarkan.
 * - Donatur login pakai No. HP + Password yang diisi saat donasi.
 *   Setelah donasinya diverifikasi, donatur bisa melihat & memilih link
 *   streaming dari semua cabang olahraga yang tersedia.
 *
 * AKUN ADMIN DEFAULT (bisa diganti di sheet "Admin"):
 * Username: admin
 * Password: admin123
 * ============================================================
 */

const SHEET_ID = 'GANTI_DENGAN_ID_SPREADSHEET_ANDA'; // <-- WAJIB DIISI
const TZ = 'Asia/Jakarta'; // WIB

// ------------------------------------------------------------
// ENTRY POINT - GET (baca data): ?action=namaFungsi
// ------------------------------------------------------------
function doGet(e) {
  try {
    const action = e.parameter.action;
    let result;
    switch (action) {
      case 'getLiveData': result = getLiveData(); break;
      case 'getConfig': result = getConfig(); break;
      case 'getAllPertandingan': result = getAllPertandingan(); break;
      case 'getAllPendaftar': result = getAllPendaftar(); break;
      case 'getAllDonasi': result = getAllDonasi(); break;
      case 'getAllLinkCabang': result = getAllLinkCabang(); break;
      default: result = { error: 'Aksi GET tidak dikenal: ' + action };
    }
    return jsonOutput_(result);
  } catch (err) {
    return jsonOutput_({ error: err.message });
  }
}

// ------------------------------------------------------------
// ENTRY POINT - POST (tulis/ubah data): body { action, data }
// ------------------------------------------------------------
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    const data = body.data || {};
    let result;
    switch (action) {
      case 'daftarPeserta': result = daftarPeserta(data); break;
      case 'kirimDonasi': result = kirimDonasi(data); break;
      case 'adminLogin': result = adminLogin(data.username, data.password); break;
      case 'pendaftarLogin': result = pendaftarLogin(data.noHp, data.password); break;
      case 'donorLogin': result = donorLogin(data.noHp, data.password); break;
      case 'addPertandingan': result = addPertandingan(data); break;
      case 'updateSkor': result = updateSkor(data.id, data.skorA, data.skorB, data.status); break;
      case 'deletePertandingan': result = deletePertandingan(data.id); break;
      case 'verifikasiPendaftar': result = verifikasiPendaftar(data.id, data.status); break;
      case 'verifikasiDonasi': result = verifikasiDonasi(data.id, data.status); break;
      case 'updateConfig': result = updateConfig(data); break;
      case 'upsertLinkCabang': result = upsertLinkCabang(data); break;
      case 'deleteLinkCabang': result = deleteLinkCabang(data.cabang); break;
      default: result = { error: 'Aksi POST tidak dikenal: ' + action };
    }
    return jsonOutput_(result);
  } catch (err) {
    return jsonOutput_({ error: err.message });
  }
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ------------------------------------------------------------
// UTIL
// ------------------------------------------------------------
function getSS_() {
  return SpreadsheetApp.openById(SHEET_ID);
}

function nowWIBString_() {
  return Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm:ss');
}

function generateId_(prefix) {
  return prefix + '-' + Utilities.formatDate(new Date(), TZ, 'yyMMddHHmmss') + '-' + Math.floor(Math.random() * 900 + 100);
}

// Pastikan URL selalu punya skema (http/https), kalau tidak, link jadi
// relatif terhadap halaman web app sendiri dan "tidak bisa dibuka".
function normalizeUrl_(url) {
  const bersih = String(url || '').trim();
  if (!bersih) return bersih;
  if (/^https?:\/\//i.test(bersih)) return bersih;
  return 'https://' + bersih;
}

// Ubah link share Google Drive (format apa pun: /file/d/ID/view, ?id=ID, uc?id=ID,
// atau langsung ID filenya saja) jadi link gambar langsung yang bisa dipakai di <img src>.
// Kalau bukan link Drive / ID tidak ketemu, link aslinya dikembalikan apa adanya.
function driveLinkKeGambar_(url) {
  const bersih = String(url || '').trim();
  if (!bersih) return bersih;
  let fileId = '';
  let m = bersih.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (m) fileId = m[1];
  if (!fileId) {
    m = bersih.match(/[?&]id=([a-zA-Z0-9_-]+)/);
    if (m) fileId = m[1];
  }
  if (!fileId && /^[a-zA-Z0-9_-]{20,}$/.test(bersih)) fileId = bersih; // admin tempel ID-nya saja
  if (!fileId) return normalizeUrl_(bersih); // bukan link Drive, biarkan seperti biasa
  return 'https://lh3.googleusercontent.com/d/' + fileId;
}

function getOrCreateSheet_(name, headers) {
  const ss = getSS_();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#0B1120').setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1000, headers.length).setNumberFormat('@');
  }
  return sheet;
}

// Migrasi otomatis: pastikan sebuah kolom header ada di sheet (dipakai supaya
// sheet lama yang sudah terlanjur dipakai/deploy tetap otomatis kebagian
// kolom baru tanpa perlu setup ulang manual). Mengembalikan nomor kolom (1-based).
function ensureHeaderColumn_(sheet, headerName) {
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  let idx = headers.indexOf(headerName);
  if (idx === -1) {
    idx = lastCol; // kolom baru ditambah persis setelah kolom terakhir yang ada
    sheet.getRange(1, idx + 1).setValue(headerName)
      .setFontWeight('bold').setBackground('#0B1120').setFontColor('#FFFFFF');
    sheet.getRange(1, idx + 1, 1000, 1).setNumberFormat('@');
  }
  return idx + 1;
}

// Bikin logo club otomatis dari inisial nama cabang/club kalau admin tidak
// mengunggah logo sendiri, jadi kolom logo tidak pernah kosong di mata penonton.
function buatLogoOtomatis_(namaCabang) {
  const nama = String(namaCabang || 'Club').trim() || 'Club';
  return 'https://ui-avatars.com/api/?name=' + encodeURIComponent(nama) +
    '&background=FFB800&color=0B1120&bold=true&size=128&length=2';
}

function sheetToObjects_(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i].join('') === '') continue;
    const obj = {};
    headers.forEach((h, idx) => { obj[h] = data[i][idx]; });
    obj._row = i + 1;
    rows.push(obj);
  }
  return rows;
}

// ------------------------------------------------------------
// SETUP (jalankan sekali secara manual dari editor Apps Script)
// ------------------------------------------------------------
function setupSheets() {
  // LogoA/LogoB = logo club Tim A & Tim B (opsional, diisi admin lewat form skor,
  // kalau kosong dibuatkan otomatis dari inisial nama tim - lihat index.html).
  getOrCreateSheet_('Pertandingan', ['ID', 'Cabang', 'TimA', 'TimB', 'SkorA', 'SkorB', 'Status', 'Venue', 'WaktuMulai', 'Catatan', 'LogoA', 'LogoB']);
  // Pendaftar sekarang punya Password supaya bisa login setelah diverifikasi
  getOrCreateSheet_('Pendaftar', ['ID', 'Nama', 'KelasTim', 'Cabang', 'NoHP', 'Password', 'Status', 'Timestamp']);
  // Donasi sekarang punya NoHP + Password supaya donatur bisa login lihat link streaming
  getOrCreateSheet_('Donasi', ['ID', 'Nama', 'Nominal', 'NoHP', 'Password', 'Metode', 'BuktiURL', 'Pesan', 'Status', 'Timestamp']);
  // Link streaming per cabang olahraga, diatur admin, dilihat pendaftar terverifikasi
  // LogoA/LogoB = logo untuk sisi A & B pada cabang ini, otomatis dibuatkan dari
  // inisial nama cabang kalau admin tidak isi sendiri
  getOrCreateSheet_('LinkCabang', LINKCABANG_HEADERS_);
  const config = getOrCreateSheet_('Config', ['Key', 'Value']);
  const cfgRows = sheetToObjects_(config);
  const defaults = {
    NAMA_ACARA: 'Pekan Olahraga',
    LOGO_ACARA: '' // link foto logo acara (diisi admin lewat panel, boleh link share Google Drive)
  };
  Object.keys(defaults).forEach(key => {
    const exists = cfgRows.some(r => r.Key === key);
    if (!exists) config.appendRow([key, defaults[key]]);
  });
  const admin = getOrCreateSheet_('Admin', ['Username', 'Password']);
  const adminRows = sheetToObjects_(admin);
  if (adminRows.length === 0) {
    admin.appendRow(['admin', 'admin123']);
  }
  SpreadsheetApp.flush();
  return 'Setup selesai. Sheet Pertandingan, Pendaftar, Donasi, LinkCabang, Config, dan Admin sudah siap.';
}

// ------------------------------------------------------------
// PUBLIC: LIVE SCORE
// ------------------------------------------------------------
function getLiveData() {
  const sheet = getOrCreateSheet_('Pertandingan', PERTANDINGAN_HEADERS_);
  ensureHeaderColumn_(sheet, 'LogoA'); // migrasi otomatis untuk sheet lama yang belum punya kolom ini
  ensureHeaderColumn_(sheet, 'LogoB');
  const rows = sheetToObjects_(sheet);
  const result = { live: [], akanDatang: [], selesai: [] };
  rows.forEach(r => {
    const status = String(r.Status || '').toLowerCase();
    if (status === 'live') result.live.push(r);
    else if (status === 'selesai') result.selesai.push(r);
    else result.akanDatang.push(r);
  });
  return result;
}

function getConfig() {
  const sheet = getOrCreateSheet_('Config', ['Key', 'Value']);
  const rows = sheetToObjects_(sheet);
  const cfg = {};
  rows.forEach(r => { cfg[r.Key] = r.Value; });
  return cfg;
}

// ------------------------------------------------------------
// PUBLIC: PENDAFTARAN PESERTA (kini termasuk Password untuk login)
// ------------------------------------------------------------
function daftarPeserta(data) {
  if (!data || !data.nama || !data.cabang || !data.password) {
    return { success: false, message: 'Nama, cabang olahraga, dan password wajib diisi.' };
  }
  const sheet = getOrCreateSheet_('Pendaftar', ['ID', 'Nama', 'KelasTim', 'Cabang', 'NoHP', 'Password', 'Status', 'Timestamp']);
  const rows = sheetToObjects_(sheet);
  const nomorSudahAda = rows.some(r => String(r.NoHP) === String(data.noHp) && data.noHp);
  if (nomorSudahAda) {
    return { success: false, message: 'No. HP ini sudah pernah mendaftar. Gunakan No. HP lain atau langsung login.' };
  }
  const id = generateId_('DAF');
  const rowIndex = sheet.getLastRow() + 1;
  sheet.getRange(rowIndex, 1, 1, 8).setNumberFormat('@');
  sheet.getRange(rowIndex, 1, 1, 8).setValues([[
    id, data.nama, data.kelasTim || '-', data.cabang, String(data.noHp || '-').trim(), String(data.password).trim(), 'Menunggu', nowWIBString_()
  ]]);
  return { success: true, message: 'Pendaftaran berhasil! Simpan No. HP & password untuk login setelah diverifikasi panitia.', id: id };
}

// ------------------------------------------------------------
// PUBLIC: DONASI KOPI
// ------------------------------------------------------------
function kirimDonasi(data) {
  // Pendaftaran nonton sekarang GRATIS - tidak butuh nominal/bukti transfer,
  // dan otomatis langsung berstatus "Terverifikasi" (tidak perlu approval admin).
  if (!data || !data.nama || !data.noHp || !data.password) {
    return { success: false, message: 'Nama, No. HP, dan password wajib diisi.' };
  }
  const sheet = getOrCreateSheet_('Donasi', ['ID', 'Nama', 'Nominal', 'NoHP', 'Password', 'Metode', 'BuktiURL', 'Pesan', 'Status', 'Timestamp']);
  const id = generateId_('DON');
  const rowIndex = sheet.getLastRow() + 1;
  sheet.getRange(rowIndex, 1, 1, 10).setNumberFormat('@');
  sheet.getRange(rowIndex, 1, 1, 10).setValues([[
    id, data.nama, 0, String(data.noHp).trim(), String(data.password).trim(), 'Gratis', '', data.pesan || '-', 'Terverifikasi', nowWIBString_()
  ]]);
  return { success: true, message: 'Pendaftaran berhasil! Kamu bisa langsung Masuk pakai No. HP & password ini untuk nonton semua cabang, gratis.', id: id };
}

function simpanFileKeDrive_(base64Data, fileName, folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
  const matches = base64Data.match(/^data:(.+);base64,(.+)$/);
  const contentType = matches ? matches[1] : 'image/jpeg';
  const pureBase64 = matches ? matches[2] : base64Data;
  const bytes = Utilities.base64Decode(pureBase64);
  const blob = Utilities.newBlob(bytes, contentType, fileName);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  // file.getUrl() mengembalikan link viewer Drive (.../view), yang TIDAK bisa dipakai
  // langsung sebagai src <img> — hasilnya gambar tidak muncul tanpa error apapun.
  // Gunakan format link direct-image dari googleusercontent agar bisa dirender <img>.
  return 'https://lh3.googleusercontent.com/d/' + file.getId();
}

// ------------------------------------------------------------
// LOGIN: ADMIN
// ------------------------------------------------------------
function adminLogin(username, password) {
  const sheet = getOrCreateSheet_('Admin', ['Username', 'Password']);
  const rows = sheetToObjects_(sheet);
  const userBersih = String(username || '').trim();
  const passBersih = String(password || '').trim();
  const found = rows.find(r => String(r.Username || '').trim() === userBersih && String(r.Password || '').trim() === passBersih);
  return { success: !!found };
}

// ------------------------------------------------------------
// LOGIN: PENDAFTAR (pakai No. HP + Password)
// ------------------------------------------------------------
function pendaftarLogin(noHp, password) {
  const sheet = getOrCreateSheet_('Pendaftar', ['ID', 'Nama', 'KelasTim', 'Cabang', 'NoHP', 'Password', 'Status', 'Timestamp']);
  const rows = sheetToObjects_(sheet);
  const noHpBersih = String(noHp || '').trim();
  const passBersih = String(password || '').trim();
  const found = rows.find(r => String(r.NoHP || '').trim() === noHpBersih && String(r.Password || '').trim() === passBersih);
  if (!found) {
    return { success: false, message: 'No. HP atau password salah.' };
  }
  return {
    success: true,
    id: found.ID,
    nama: found.Nama,
    kelasTim: found.KelasTim,
    cabang: found.Cabang,
    status: found.Status
  };
}

// ------------------------------------------------------------
// LOGIN: DONATUR (pakai No. HP + Password yang diisi saat donasi)
// Jika donatur pernah donasi lebih dari sekali, diutamakan baris yang
// sudah Terverifikasi; kalau belum ada, dipakai donasi terakhir.
// ------------------------------------------------------------
function donorLogin(noHp, password) {
  const sheet = getOrCreateSheet_('Donasi', ['ID', 'Nama', 'Nominal', 'NoHP', 'Password', 'Metode', 'BuktiURL', 'Pesan', 'Status', 'Timestamp']);
  const rows = sheetToObjects_(sheet);
  const noHpBersih = String(noHp || '').trim();
  const passBersih = String(password || '').trim();
  const cocok = rows.filter(r => String(r.NoHP || '').trim() === noHpBersih && String(r.Password || '').trim() === passBersih);
  if (cocok.length === 0) {
    return { success: false, message: 'No. HP atau password salah.' };
  }
  const terverifikasi = cocok.filter(r => r.Status === 'Terverifikasi');
  const found = terverifikasi.length > 0 ? terverifikasi[terverifikasi.length - 1] : cocok[cocok.length - 1];
  return {
    success: true,
    id: found.ID,
    nama: found.Nama,
    nominal: found.Nominal,
    status: found.Status
  };
}

// ------------------------------------------------------------
// ADMIN: KELOLA PERTANDINGAN / SKOR LIVE
// ------------------------------------------------------------
const PERTANDINGAN_HEADERS_ = ['ID', 'Cabang', 'TimA', 'TimB', 'SkorA', 'SkorB', 'Status', 'Venue', 'WaktuMulai', 'Catatan', 'LogoA', 'LogoB'];
const LINKCABANG_HEADERS_ = ['Cabang', 'LinkURL', 'LogoA', 'LogoB', 'Keterangan', 'Timestamp'];

function getAllPertandingan() {
  const sheet = getOrCreateSheet_('Pertandingan', PERTANDINGAN_HEADERS_);
  ensureHeaderColumn_(sheet, 'LogoA'); // migrasi otomatis untuk sheet lama yang belum punya kolom ini
  ensureHeaderColumn_(sheet, 'LogoB');
  return sheetToObjects_(sheet).reverse();
}

function addPertandingan(data) {
  const sheet = getOrCreateSheet_('Pertandingan', PERTANDINGAN_HEADERS_);
  const colLogoA = ensureHeaderColumn_(sheet, 'LogoA'); // migrasi otomatis untuk sheet lama
  const colLogoB = ensureHeaderColumn_(sheet, 'LogoB');
  const id = generateId_('PTD');
  const rowIndex = sheet.getLastRow() + 1;
  // Logo tim: pakai yang diisi admin kalau ada, kalau tidak dibuatkan otomatis
  // dari inisial nama tim (sama seperti logo cabang di LinkCabang), supaya
  // logo tidak pernah kosong di mata penonton.
  const logoA = data.logoA ? driveLinkKeGambar_(data.logoA) : buatLogoOtomatis_(data.timA);
  const logoB = data.logoB ? driveLinkKeGambar_(data.logoB) : buatLogoOtomatis_(data.timB);
  sheet.getRange(rowIndex, 1, 1, 10).setNumberFormat('@');
  sheet.getRange(rowIndex, 1, 1, 10).setValues([[
    id, data.cabang, data.timA, data.timB, data.skorA || 0, data.skorB || 0,
    data.status || 'Akan Datang', data.venue || '-', data.waktuMulai || nowWIBString_(), data.catatan || '-'
  ]]);
  sheet.getRange(rowIndex, colLogoA).setNumberFormat('@').setValue(logoA);
  sheet.getRange(rowIndex, colLogoB).setNumberFormat('@').setValue(logoB);
  return { success: true, message: 'Pertandingan berhasil ditambahkan.' };
}

function updateSkor(id, skorA, skorB, status) {
  const sheet = getOrCreateSheet_('Pertandingan', PERTANDINGAN_HEADERS_);
  const rows = sheetToObjects_(sheet);
  const target = rows.find(r => String(r.ID) === String(id));
  if (!target) return { success: false, message: 'Pertandingan tidak ditemukan.' };
  sheet.getRange(target._row, 5).setNumberFormat('@').setValue(String(skorA));
  sheet.getRange(target._row, 6).setNumberFormat('@').setValue(String(skorB));
  sheet.getRange(target._row, 7).setNumberFormat('@').setValue(String(status));
  return { success: true, message: 'Skor berhasil diperbarui.' };
}

function deletePertandingan(id) {
  const sheet = getOrCreateSheet_('Pertandingan', PERTANDINGAN_HEADERS_);
  const rows = sheetToObjects_(sheet);
  const target = rows.find(r => String(r.ID) === String(id));
  if (!target) return { success: false, message: 'Data tidak ditemukan.' };
  sheet.deleteRow(target._row);
  return { success: true, message: 'Pertandingan berhasil dihapus.' };
}

// ------------------------------------------------------------
// ADMIN: VERIFIKASI PENDAFTAR
// ------------------------------------------------------------
function getAllPendaftar() {
  const sheet = getOrCreateSheet_('Pendaftar', ['ID', 'Nama', 'KelasTim', 'Cabang', 'NoHP', 'Password', 'Status', 'Timestamp']);
  // Password sengaja tidak dikirim ke tampilan admin demi keamanan sederhana
  return sheetToObjects_(sheet).reverse().map(r => ({
    ID: r.ID, Nama: r.Nama, KelasTim: r.KelasTim, Cabang: r.Cabang, NoHP: r.NoHP, Status: r.Status, Timestamp: r.Timestamp
  }));
}

function verifikasiPendaftar(id, statusBaru) {
  const sheet = getOrCreateSheet_('Pendaftar', ['ID', 'Nama', 'KelasTim', 'Cabang', 'NoHP', 'Password', 'Status', 'Timestamp']);
  const rows = sheetToObjects_(sheet);
  const target = rows.find(r => String(r.ID) === String(id));
  if (!target) return { success: false, message: 'Pendaftar tidak ditemukan.' };
  sheet.getRange(target._row, 7).setNumberFormat('@').setValue(String(statusBaru)); // kolom Status
  return { success: true, message: 'Status pendaftar diperbarui menjadi ' + statusBaru + '.' };
}

// ------------------------------------------------------------
// ADMIN: VERIFIKASI DONASI
// ------------------------------------------------------------
function getAllDonasi() {
  const sheet = getOrCreateSheet_('Donasi', ['ID', 'Nama', 'Nominal', 'NoHP', 'Password', 'Metode', 'BuktiURL', 'Pesan', 'Status', 'Timestamp']);
  // Password sengaja tidak dikirim ke tampilan admin demi keamanan sederhana
  return sheetToObjects_(sheet).reverse().map(r => ({
    ID: r.ID, Nama: r.Nama, Nominal: r.Nominal, NoHP: r.NoHP, Metode: r.Metode,
    BuktiURL: r.BuktiURL, Pesan: r.Pesan, Status: r.Status, Timestamp: r.Timestamp
  }));
}

function verifikasiDonasi(id, statusBaru) {
  const sheet = getOrCreateSheet_('Donasi', ['ID', 'Nama', 'Nominal', 'NoHP', 'Password', 'Metode', 'BuktiURL', 'Pesan', 'Status', 'Timestamp']);
  const rows = sheetToObjects_(sheet);
  const target = rows.find(r => String(r.ID) === String(id));
  if (!target) return { success: false, message: 'Donasi tidak ditemukan.' };
  sheet.getRange(target._row, 9).setNumberFormat('@').setValue(String(statusBaru)); // kolom Status
  return { success: true, message: 'Status donasi diperbarui menjadi ' + statusBaru + '.' };
}

// ------------------------------------------------------------
// ADMIN: UPDATE CONFIG (nama acara, dll)
// ------------------------------------------------------------
function updateConfig(data) {
  const sheet = getOrCreateSheet_('Config', ['Key', 'Value']);
  const rows = sheetToObjects_(sheet);
  Object.keys(data).forEach(key => {
    // LOGO_ACARA diisi admin berupa link share Google Drive -> diubah otomatis
    // jadi link gambar langsung supaya logo tampil di halaman, bukan link viewer Drive.
    const nilai = key === 'LOGO_ACARA' ? driveLinkKeGambar_(data[key]) : String(data[key]);
    const target = rows.find(r => r.Key === key);
    if (target) sheet.getRange(target._row, 2).setNumberFormat('@').setValue(nilai);
    else sheet.appendRow([key, nilai]);
  });
  return { success: true, message: 'Pengaturan berhasil disimpan.' };
}

// ------------------------------------------------------------
// ADMIN + PENDAFTAR: LINK STREAMING PER CABANG OLAHRAGA
// ------------------------------------------------------------
function getAllLinkCabang() {
  const sheet = getOrCreateSheet_('LinkCabang', LINKCABANG_HEADERS_);
  ensureHeaderColumn_(sheet, 'LogoA'); // migrasi otomatis untuk sheet lama yang belum punya kolom ini
  ensureHeaderColumn_(sheet, 'LogoB');
  return sheetToObjects_(sheet);
}

// Tambah baru jika cabang belum ada, atau update jika sudah ada (upsert by Cabang)
// Logo A & Logo B wajib diisi admin (tidak lagi dibuat otomatis).
function upsertLinkCabang(data) {
  if (!data || !data.cabang || !data.linkUrl || !data.logoA || !data.logoB) {
    return { success: false, message: 'Cabang, link, Logo A, dan Logo B wajib diisi.' };
  }
  const sheet = getOrCreateSheet_('LinkCabang', LINKCABANG_HEADERS_);
  const colCabang = ensureHeaderColumn_(sheet, 'Cabang');
  const colLink = ensureHeaderColumn_(sheet, 'LinkURL');
  const colLogoA = ensureHeaderColumn_(sheet, 'LogoA');
  const colLogoB = ensureHeaderColumn_(sheet, 'LogoB');
  const colKet = ensureHeaderColumn_(sheet, 'Keterangan');
  const colTime = ensureHeaderColumn_(sheet, 'Timestamp');
  const linkUrlBersih = normalizeUrl_(data.linkUrl);
  const logoABersih = driveLinkKeGambar_(data.logoA);
  const logoBBersih = driveLinkKeGambar_(data.logoB);
  const rows = sheetToObjects_(sheet);
  const target = rows.find(r => String(r.Cabang).toLowerCase() === String(data.cabang).toLowerCase());
  if (target) {
    sheet.getRange(target._row, colLink).setNumberFormat('@').setValue(linkUrlBersih);
    sheet.getRange(target._row, colLogoA).setNumberFormat('@').setValue(logoABersih);
    sheet.getRange(target._row, colLogoB).setNumberFormat('@').setValue(logoBBersih);
    sheet.getRange(target._row, colKet).setNumberFormat('@').setValue(data.keterangan || '-');
    sheet.getRange(target._row, colTime).setNumberFormat('@').setValue(nowWIBString_());
    return { success: true, message: 'Link untuk ' + data.cabang + ' berhasil diperbarui.' };
  }
  const rowIndex = sheet.getLastRow() + 1;
  sheet.getRange(rowIndex, colCabang).setNumberFormat('@').setValue(data.cabang);
  sheet.getRange(rowIndex, colLink).setNumberFormat('@').setValue(linkUrlBersih);
  sheet.getRange(rowIndex, colLogoA).setNumberFormat('@').setValue(logoABersih);
  sheet.getRange(rowIndex, colLogoB).setNumberFormat('@').setValue(logoBBersih);
  sheet.getRange(rowIndex, colKet).setNumberFormat('@').setValue(data.keterangan || '-');
  sheet.getRange(rowIndex, colTime).setNumberFormat('@').setValue(nowWIBString_());
  return { success: true, message: 'Link untuk ' + data.cabang + ' berhasil ditambahkan.' };
}

function deleteLinkCabang(cabang) {
  const sheet = getOrCreateSheet_('LinkCabang', LINKCABANG_HEADERS_);
  const rows = sheetToObjects_(sheet);
  const target = rows.find(r => String(r.Cabang).toLowerCase() === String(cabang).toLowerCase());
  if (!target) return { success: false, message: 'Link tidak ditemukan.' };
  sheet.deleteRow(target._row);
  return { success: true, message: 'Link berhasil dihapus.' };
}
