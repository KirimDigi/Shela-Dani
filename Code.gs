/* ============================================================
 * RSVP + Ucapan -> Google Spreadsheet (Shela & Dani)
 * Cara pakai: tinggal copas seluruh isi file ini ke
 * Extensions > Apps Script, lalu ikuti 3 langkah di bawah.
 *
 * 1) Ganti SHEET_ID dengan ID spreadsheet Anda
 *    (dari URL: docs.google.com/spreadsheets/d/INI_ID_... )
 * 2) Project Settings > Time zone = Asia/Jakarta
 * 3) Deploy > New deployment > Web app
 *    Execute as: Me | Who has access: Anyone
 *    -> copy URL /exec -> tempel ke index.html
 *       const GOOGLE_SHEET_API = 'URL_TADI';
 * Sheet: tab "Ucapan", header baris 1:
 * Timestamp | Nama | Kehadiran | JumlahTamu | Ucapan
 * ============================================================ */

const SHEET_ID = 'PASTE_SPREADSHEET_ID_ANDA';
const SHEET_NAME = 'Ucapan';

function getSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['Timestamp', 'Nama', 'Kehadiran', 'JumlahTamu', 'Ucapan']);
  }
  return sh;
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}

// GET: dibaca index.html untuk tampil di bawah tombol Kirim
function doGet() {
  const vals = getSheet().getDataRange().getValues();
  const data = [];
  for (let i = vals.length - 1; i >= 1; i--) {
    const r = vals[i];
    if (!r[1] && !r[4]) continue;
    const ts = r[0];
    let waktu = '';
    try {
      const d = (ts instanceof Date) ? ts : new Date(ts);
      waktu = isNaN(d.getTime()) ? String(ts || '') : Utilities.formatDate(d, 'Asia/Jakarta', 'dd MMM yyyy, HH:mm');
    } catch (err) { waktu = String(ts || ''); }
    const kh = String(r[2] || '');
    data.push({
      nama: String(r[1] || ''),
      kehadiran: (kh === 'present' || kh === 'Hadir') ? 'present' : 'notpresent',
      tamu: String(r[3] == null ? '' : r[3]),
      ucapan: String(r[4] || ''),
      waktu: waktu
    });
  }
  return json({ status: 'ok', data: data });
}

// POST: dipanggil saat tamu klik Kirim (jam & tanggal dicatat otomatis)
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    const p = (e && e.parameter) || {};
    const nama = String(p.nama || '').trim();
    const kehadiran = String(p.kehadiran || '').trim();
    const tamu = String(p.tamu || '').trim();
    const ucapan = String(p.ucapan || '').trim();
    if (nama.length < 2 || !kehadiran || ucapan.length < 2)
      return json({ status: 'error', message: 'Data belum lengkap.' });
    if (kehadiran === 'present' && !tamu)
      return json({ status: 'error', message: 'Pilih jumlah tamu.' });
    const now = new Date();
    getSheet().appendRow([now, nama, kehadiran === 'present' ? 'Hadir' : 'Tidak Hadir', tamu, ucapan]);
    return json({
      status: 'ok', nama: nama, kehadiran: kehadiran, tamu: tamu, ucapan: ucapan,
      waktu: Utilities.formatDate(now, 'Asia/Jakarta', 'dd MMM yyyy, HH:mm')
    });
  } catch (err) {
    return json({ status: 'error', message: String(err) });
  } finally {
    lock.releaseLock();
  }
}
