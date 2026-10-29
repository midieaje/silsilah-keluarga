/**
 * SILSILAH KELUARGA — Backend Google Apps Script (+ penyaji aplikasi untuk Google Sites)
 * ----------------------------------------------------------------------------------
 * CARA PASANG (ringkas; panduan lengkap ada di dokumen "Panduan Silsilah Keluarga"):
 * 1. Buat Google Sheet baru (kosong). Menu: Ekstensi > Apps Script.
 * 2. Hapus isi default, tempel seluruh file ini sebagai Code.gs, simpan.
 * 3. Tombol "+" di samping "File" > HTML > beri nama "index". Tempel SELURUH isi family-tree.html
 *    ke file index tersebut (biarkan  SCRIPT_URL: ""  kosong, URL diisi otomatis oleh doGet di bawah).
 * 4. Pilih fungsi "setup" > Jalankan > setujui izin (Spreadsheet + Drive untuk foto).
 * 5. Setelan Proyek > Properti skrip: ubah PASSWORD_VIEW (lihat saja) dan PASSWORD_EDIT (mengedit).
 * 6. Terapkan > Deployment baru > Aplikasi web:
 *      Jalankan sebagai : Saya
 *      Yang memiliki akses: Siapa saja
 *    Salin "URL aplikasi web" (berakhiran /exec). URL ini bisa langsung dibuka atau
 *    disematkan di Google Sites: Sisipkan > Sematkan > Dengan URL.
 *
 * Setiap kali Anda mengubah kode atau file index, buat versi baru:
 *   Terapkan > Kelola deployment > ikon pensil > Versi: Versi baru > Terapkan  (URL tetap sama).
 */

const SHEETS = {
  orang: { name: 'Orang',      cols: ['id', 'nama', 'gender', 'lahir', 'wafat', 'tempat', 'catatan', 'foto', 'parent_nikah_id', 'kode'] },
  nikah: { name: 'Pernikahan', cols: ['id', 'p1', 'p2', 'tahun', 'status', 'catatan'] }
};
const FOTO_FOLDER = 'Silsilah Keluarga - Foto';

/** Jalankan sekali dari editor. */
function setup() {
  const p = PropertiesService.getScriptProperties();
  if (!p.getProperty('PASSWORD_VIEW')) p.setProperty('PASSWORD_VIEW', 'keluarga123');
  if (!p.getProperty('PASSWORD_EDIT')) p.setProperty('PASSWORD_EDIT', 'admin123');
  Object.keys(SHEETS).forEach(sheet_);
  folder_();
}

/** Membuka URL aplikasi web = menampilkan aplikasi (file "index"). Tanpa file index: menampilkan info API. */
function doGet() {
  let html;
  try { html = HtmlService.createHtmlOutputFromFile('index').getContent(); }
  catch (err) { return out_({ ok: true, info: 'Silsilah Keluarga API aktif. Tambahkan file HTML bernama "index" untuk menyajikan aplikasi.' }); }
  // Isi alamat API secara otomatis dengan URL aplikasi web ini
  html = html.replace(/SCRIPT_URL:\s*""/, 'SCRIPT_URL: ' + JSON.stringify(ScriptApp.getService().getUrl()));
  return HtmlService.createHtmlOutput(html)
    .setTitle('Silsilah Keluarga')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  let res;
  try {
    const req = JSON.parse(e.postData.contents);
    const role = auth_(req.password);
    if (!role) return out_({ ok: false, error: 'Password salah' });

    switch (req.action) {
      case 'login':
      case 'getAll':
        res = { ok: true, role: role, orang: read_('orang'), nikah: read_('nikah') };
        break;
      case 'batch':
        needEdit_(role);
        res = batch_(req.ops || []);
        break;
      case 'uploadPhoto':
        needEdit_(role);
        res = upload_(req);
        break;
      case 'deletePhoto':
        needEdit_(role);
        trash_(req.url);
        res = { ok: true };
        break;
      default:
        res = { ok: false, error: 'Aksi tidak dikenal' };
    }
  } catch (err) {
    res = { ok: false, error: String(err && err.message ? err.message : err) };
  }
  return out_(res);
}

/* ---------- Auth ---------- */
function auth_(pw) {
  pw = String(pw || '');
  if (!pw) return null;
  const p = PropertiesService.getScriptProperties();
  if (pw === p.getProperty('PASSWORD_EDIT')) return 'edit';
  if (pw === p.getProperty('PASSWORD_VIEW')) return 'view';
  return null;
}
function needEdit_(role) {
  if (role !== 'edit') throw new Error('Password ini hanya untuk melihat. Gunakan password edit.');
}

/* ---------- Sheet helpers ---------- */
function sheet_(key) {
  const def = SHEETS[key];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(def.name);
  if (!sh) {
    sh = ss.insertSheet(def.name);
    sh.getRange(1, 1, sh.getMaxRows(), def.cols.length).setNumberFormat('@'); // simpan sbg teks (tanggal tidak berubah format)
    sh.getRange(1, 1, 1, def.cols.length).setValues([def.cols]).setFontWeight('bold');
    sh.setFrozenRows(1);
  } else if (sh.getLastColumn() < def.cols.length) {
    // Migrasi: sheet lama belum punya kolom baru (mis. "kode")
    if (sh.getMaxColumns() < def.cols.length) sh.insertColumnsAfter(sh.getMaxColumns(), def.cols.length - sh.getMaxColumns());
    sh.getRange(1, 1, sh.getMaxRows(), def.cols.length).setNumberFormat('@');
    sh.getRange(1, 1, 1, def.cols.length).setValues([def.cols]).setFontWeight('bold');
  }
  return sh;
}

function read_(key) {
  const def = SHEETS[key];
  const sh = sheet_(key);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const tz = Session.getScriptTimeZone();
  const vals = sh.getRange(2, 1, last - 1, def.cols.length).getValues();
  return vals
    .filter(function (r) { return r[0] !== ''; })
    .map(function (r) {
      const o = {};
      def.cols.forEach(function (c, i) {
        const v = r[i];
        o[c] = v === null || v === undefined ? '' :
          (v instanceof Date ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : String(v));
      });
      return o;
    });
}

function findRow_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}

/** ops: [{op:'upsert', t:'orang'|'nikah', row:{...}}, {op:'delete', t:'orang'|'nikah', id:'...'}] */
function batch_(ops) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    ops.forEach(function (op) {
      const def = SHEETS[op.t];
      if (!def) throw new Error('Tabel tidak valid');
      const sh = sheet_(op.t);
      const id = op.op === 'delete' ? op.id : op.row.id;
      if (!id) throw new Error('ID kosong');
      const r = findRow_(sh, id);

      if (op.op === 'delete') {
        if (r > 0) sh.deleteRow(r);
      } else {
        const vals = def.cols.map(function (c) { return op.row[c] == null ? '' : String(op.row[c]); });
        let target = r;
        if (target < 0) {
          target = sh.getLastRow() + 1;
          if (target > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 50);
        }
        const rg = sh.getRange(target, 1, 1, def.cols.length);
        rg.setNumberFormat('@');
        rg.setValues([vals]);
      }
    });
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- Foto (Google Drive) ---------- */
function folder_() {
  const it = DriveApp.getFoldersByName(FOTO_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOTO_FOLDER);
}

function idFromUrl_(u) {
  const m = String(u || '').match(/[?&]id=([\w-]+)/);
  return m ? m[1] : '';
}

function trash_(url) {
  const id = idFromUrl_(url);
  if (!id) return;
  try { DriveApp.getFileById(id).setTrashed(true); } catch (e) { /* abaikan */ }
}

function upload_(req) {
  if (!req.base64) throw new Error('Data foto kosong');
  const bytes = Utilities.base64Decode(req.base64);
  if (bytes.length > 6 * 1024 * 1024) throw new Error('Foto terlalu besar');
  const blob = Utilities.newBlob(bytes, req.mime || 'image/jpeg', (req.personId || 'foto') + '_' + Date.now() + '.jpg');
  const f = folder_().createFile(blob);
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  if (req.old) trash_(req.old);
  return { ok: true, url: 'https://drive.google.com/thumbnail?id=' + f.getId() + '&sz=w600' };
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
