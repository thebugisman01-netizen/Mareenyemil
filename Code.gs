// Catat Jualan - backend Google Sheets (Apps Script)
// Pasang di: Spreadsheet > Ekstensi > Apps Script. Lalu Deploy > Aplikasi web.
const SH = {
  Users: ['id', 'username', 'hash', 'toko', 'pin'],
  Barang: ['id', 'nama', 'harga', 'modal', 'stok', 'min'],
  Transaksi: ['id', 'waktu', 'items', 'total', 'bayar', 'kembali', 'customer', 'hp', 'metode', 'status', 'tgllunas'],
  Keuangan: ['id', 'tanggal', 'jenis', 'ket', 'jumlah'],
  Pengaturan: ['key', 'value']
};

function sheet(n) {
  if (!SH[n]) throw new Error('Fungsi sheet() tidak dijalankan langsung. Pilih fungsi "setup" di menu dropdown lalu klik Jalankan.');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(n);
  if (!s) {
    s = ss.insertSheet(n);
    s.getRange(1, 1, 1, SH[n].length).setValues([SH[n]]).setFontWeight('bold');
    s.getRange('A:Z').setNumberFormat('@'); // simpan sebagai teks agar tanggal tidak berubah
    s.setFrozenRows(1);
  } else if (s.getLastColumn() < SH[n].length) {
    // migrasi otomatis: sheet lama ditambah kolom baru (mis. customer & hp)
    s.getRange(1, 1, 1, SH[n].length).setValues([SH[n]]).setFontWeight('bold');
    s.getRange('A:Z').setNumberFormat('@');
  }
  return s;
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function rows(n) {
  const v = sheet(n).getDataRange().getValues();
  const h = v.shift();
  return v.filter(r => r[0] !== '').map(r => Object.fromEntries(h.map((k, i) => [k, r[i]])));
}

// Ambil semua data
function doGet(e) {
  if (e && e.parameter && e.parameter.qris) return qrisPage(e.parameter);
  const d = {};
  Object.keys(SH).forEach(n => d[n] = rows(n));
  return out(d);
}

// Simpan / ubah / hapus satu baris
function doPost(e) {
  const p = JSON.parse(e.postData.contents);
  if (!SH[p.sheet]) return out({ ok: false, error: 'sheet tidak dikenal' });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (p.op === 'bulk') { // simpan banyak baris sekaligus (pulihkan cadangan)
      const bs = sheet(p.sheet), bh = SH[p.sheet];
      const bv = bs.getDataRange().getValues(); bv.shift();
      const all = bv.filter(r => r[0] !== '').map(r => bh.map((k, i) => r[i] === undefined ? '' : r[i]));
      const pos = {}; all.forEach((r, i) => pos[String(r[0])] = i);
      (p.rows || []).forEach(row => {
        const arr = bh.map(k => (row[k] === undefined || row[k] === null) ? '' : row[k]);
        const i = pos[String(arr[0])];
        if (i === undefined) { pos[String(arr[0])] = all.length; all.push(arr); } else all[i] = arr;
      });
      const last = bs.getLastRow();
      if (last > 1) bs.getRange(2, 1, last - 1, bh.length).clearContent();
      if (all.length) bs.getRange(2, 1, all.length, bh.length).setValues(all);
      return out({ ok: true });
    }
    const s = sheet(p.sheet), h = SH[p.sheet], key = h[0];
    const v = s.getDataRange().getValues();
    const idx = v.findIndex((r, i) => i > 0 && String(r[0]) === String(p.row[key]));
    if (p.op === 'del') {
      if (idx > 0) s.deleteRow(idx + 1);
    } else {
      const arr = h.map(k => (p.row[k] === undefined || p.row[k] === null) ? '' : p.row[k]);
      if (idx > 0) s.getRange(idx + 1, 1, 1, h.length).setValues([arr]);
      else s.appendRow(arr);
    }
    return out({ ok: true });
  } finally {
    lock.release();
  }
}

// Jalankan sekali (opsional) untuk membuat semua sheet sekaligus
function setup() { Object.keys(SH).forEach(n => sheet(n)); }

// Halaman QRIS yang bisa dibuka customer dari tautan di pesan WhatsApp
function qrisPage(q) {
  const c = {};
  rows('Pengaturan').forEach(r => c[r.key] = r.value);
  const esc = s => String(s || '').replace(/[&<>"']/g, m => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[m]));
  const ok = /^data:image\/(png|jpeg);base64,[A-Za-z0-9+\/=]+$/.test(c.qris || '');
  if (!ok) return HtmlService.createHtmlOutput('QRIS belum tersedia.');
  const n = Number(q.n) || 0;
  const rp = String(Math.round(n)).replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  const html = '<div style="font-family:sans-serif;text-align:center;padding:16px"><h2>' + esc(c.toko) + '</h2><p>Scan QRIS untuk membayar</p>' +
    (n > 0 ? '<p style="font-size:26px"><b>Rp' + rp + '</b></p>' : '') +
    '<img src="' + c.qris + '" style="max-width:100%;width:320px;border:1px solid #ccc;border-radius:8px;background:#fff"></div>';
  return HtmlService.createHtmlOutput(html).setTitle('QRIS ' + (c.toko || '')).addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
