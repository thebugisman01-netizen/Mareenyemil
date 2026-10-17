// Catat Jualan - backend Google Sheets (Apps Script)
// Pasang di: Spreadsheet > Ekstensi > Apps Script. Lalu Deploy > Aplikasi web.
const SH = {
  Users: ['id', 'username', 'hash', 'toko'],
  Barang: ['id', 'nama', 'harga', 'modal', 'stok', 'min'],
  Transaksi: ['id', 'waktu', 'items', 'total', 'bayar', 'kembali', 'customer', 'hp'],
  Keuangan: ['id', 'tanggal', 'jenis', 'ket', 'jumlah'],
  Pengaturan: ['key', 'value']
};

function sheet(n) {
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
function doGet() {
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
function setup() { Object.keys(SH).forEach(sheet); }
