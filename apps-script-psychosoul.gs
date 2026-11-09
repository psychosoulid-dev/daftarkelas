/**
 * Backend pendaftaran & admin Psychosoul (Google Apps Script)
 * Spreadsheet tujuan: 1bZ_5Rs7glTHjgFuXwwBUizl5IbJtFka-ggCDHIli38A
 *
 * SETUP:
 * 1. Buka spreadsheet tujuan > Extensions > Apps Script, tempel seluruh kode ini.
 * 2. Project Settings > Script properties > tambah  ADMIN_PIN = (PIN rahasia admin).
 * 3. Deploy > New deployment > Web app
 *      Execute as: Me   |   Who has access: Anyone
 * 4. Izinkan akses (Sheets + Drive), lalu salin URL "/exec" ke API_URL di file HTML.
 * Setiap ubah kode, buat deployment baru (Manage deployments > Edit > New version).
 */
const SS_ID = "1bZ_5Rs7glTHjgFuXwwBUizl5IbJtFka-ggCDHIli38A";
const FOLDER_NAME = "Bukti Pembayaran Psychosoul";
const HEAD = ["Waktu","ID Daftar","Kode Kelas","Kelas","Pembiayaan","Total","Peserta ke","Nama","Usia",
  "Jenis Kelamin","No WA","Alamat","Pekerjaan","Jurusan","Semester","Tempat Pendidikan",
  "Lama Praktik (thn)","Tempat Kerja","Bukti Bayar","Status Bayar",
  "PHQ-4 Jawaban","PHQ-4 Cemas","PHQ-4 Depresi","PHQ-4 Total",
  "Hasil Konseling","Nama Konselor"];   // 2 kolom terakhir diisi admin lewat halaman admin   // 4 kolom PHQ-4 hanya terisi untuk kelas yang memakai skrining

const out = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
const safe = v => (typeof v === "string" && /^[=+\-@]/.test(v)) ? "'" + v : (v == null ? "" : v); // cegah formula injection

function doGet() { return out({ ok: true, msg: "Psychosoul API aktif" }); }

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.action === "daftar") return out(daftar(d));
    if (d.action === "admin") return out(admin(d));
    if (d.action === "konseling") return out(konseling(d));
    return out({ ok: false, error: "Aksi tidak dikenal" });
  } catch (err) { return out({ ok: false, error: String(err.message || err) }); }
}

function folder() {
  const it = DriveApp.getFoldersByName(FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
}

// Skor PHQ-4 dihitung ulang di server dari 4 jawaban (0-3). Cemas = item 1+2, Depresi = item 3+4.
function phq(p) {
  const a = [1, 2, 3, 4].map(i => p["phq" + i]);
  if (a.every(x => x === undefined || x === "" || x === null)) return ["", "", "", ""];
  const n = a.map(Number);
  if (n.some(x => !(x >= 0 && x <= 3))) throw new Error("Jawaban PHQ-4 tidak lengkap");
  return [n.join(","), n[0] + n[1], n[2] + n[3], n[0] + n[1] + n[2] + n[3]];
}

function ensureHeader(sh) {                                // sheet lama: tambah header kolom baru bila kurang
  const have = sh.getLastColumn();
  if (have < HEAD.length) sh.getRange(1, have + 1, 1, HEAD.length - have).setValues([HEAD.slice(have)]).setFontWeight("bold");
}

function cekPin(d) {
  const pin = PropertiesService.getScriptProperties().getProperty("ADMIN_PIN");
  if (!pin) return "ADMIN_PIN belum diatur di Script properties";
  if (String(d.pin) !== pin) { Utilities.sleep(800); return "PIN salah"; }
  return null;
}

// Admin mengisi hasil konseling & nama konselor untuk satu peserta (dicari lewat ID Daftar + Peserta ke)
function konseling(d) {
  const err = cekPin(d); if (err) return { ok: false, error: err };
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = SpreadsheetApp.openById(SS_ID).getSheetByName(String(d.kode || ""));
    if (!sh || sh.getRange(1, 1).getValue() !== "Waktu") throw new Error("Sheet kelas tidak ditemukan");
    ensureHeader(sh);
    const v = sh.getRange(1, 1, sh.getLastRow(), HEAD.length).getValues(), h = v[0];
    const cId = h.indexOf("ID Daftar"), cKe = h.indexOf("Peserta ke"), cH = h.indexOf("Hasil Konseling"), cN = h.indexOf("Nama Konselor");
    for (let i = 1; i < v.length; i++) {
      if (String(v[i][cId]) === String(d.id) && String(v[i][cKe]) === String(d.ke)) {
        sh.getRange(i + 1, cH + 1, 1, 2).setValues([[safe(String(d.hasil || "").slice(0, 3000)), safe(String(d.konselor || "").slice(0, 120))]]);
        return { ok: true };
      }
    }
    throw new Error("Data peserta tidak ditemukan");
  } finally { lock.releaseLock(); }
}

function daftar(d) {
  const kode = String(d.kode || "").replace(/[\[\]*?:\/\\]/g, "_").slice(0, 90);
  if (!kode || !d.peserta || !d.peserta.length) throw new Error("Data pendaftaran tidak lengkap");
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    let sh = ss.getSheetByName(kode);                       // satu sheet per kode kelas/event
    if (!sh) {
      sh = ss.insertSheet(kode);
      sh.appendRow(HEAD); sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, HEAD.length).setFontWeight("bold");
    } else ensureHeader(sh);
    const id = "REG-" + Utilities.formatDate(new Date(), "Asia/Jakarta", "yyMMdd-HHmmss") + "-" + Math.floor(100 + Math.random() * 900);
    let link = "";
    if (d.bukti && d.bukti.data) {
      const blob = Utilities.newBlob(Utilities.base64Decode(d.bukti.data), d.bukti.mime || "image/jpeg",
        id + "_" + String(d.peserta[0].nama || "peserta").slice(0, 40) + ".jpg");
      link = folder().createFile(blob).getUrl();
    }
    const now = new Date(), start = sh.getLastRow() + 1;
    const rows = d.peserta.map((p, i) => [now, id, kode, safe(d.kelas), safe(d.pembiayaan), Number(d.total) || 0, i + 1,
      safe(p.nama), safe(p.usia), safe(p.jk), safe(p.wa), safe(p.alamat), safe(p.kerja || "-"),
      safe(p.jurusan || ""), safe(p.semester || ""), safe(p.tempat_pendidikan || ""),
      safe(p.lama_praktik || ""), safe(p.tempat_kerja || ""), link, "Menunggu verifikasi"].concat(phq(p), ["", ""]));
    sh.getRange(start, 11, rows.length, 1).setNumberFormat("@");   // No WA tetap teks (0 depan aman)
    sh.getRange(start, 1, rows.length, HEAD.length).setValues(rows);
    return { ok: true, id: id };
  } finally { lock.releaseLock(); }
}

function admin(d) {
  const err = cekPin(d); if (err) return { ok: false, error: err };
  const data = {};
  SpreadsheetApp.openById(SS_ID).getSheets().forEach(sh => {
    if (sh.getLastRow() < 1 || sh.getRange(1, 1).getValue() !== "Waktu") return;   // hanya sheet kelas
    const v = sh.getDataRange().getValues(), h = v.shift();
    data[sh.getName()] = v.map(r => {
      const o = {}; h.forEach((k, i) => o[k] = r[i] instanceof Date
        ? Utilities.formatDate(r[i], "Asia/Jakarta", "dd MMM yyyy HH:mm") : r[i]);
      return o;
    });
  });
  return { ok: true, data: data };
}
