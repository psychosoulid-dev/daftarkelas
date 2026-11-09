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
  "Hasil Konseling","Nama Konselor","Status Konseling","No Urut Voucher","Voucher Token","Voucher Scan","Voucher Dipakai","Kode Voucher"];   // 3 kolom terakhir diisi admin lewat halaman admin; 4 kolom PHQ-4 hanya terisi untuk kelas yang memakai skrining

const out = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
const safe = v => (typeof v === "string" && /^[=+\-@]/.test(v)) ? "'" + v : (v == null ? "" : v); // cegah formula injection

function doGet() { return out({ ok: true, msg: "Psychosoul API aktif" }); }

// ===== Voucher diskon lewat QR =====
// Voucher diterbitkan di sheet kelas "asal" (peserta berstatus konseling Selesai) dan dipakai di kelas "tujuan".
// QR berisi link ?v=<no urut>-<token>. Token acak mencegah orang menebak kode urut.
const VCH = { prefix: "2611.DE09", pad: 3, asal: "C11.3.26.03.1", tujuan: "C11.3.26.03.2", diskon: 10 };
const kodeVoucher = no => VCH.prefix + "." + String(no).padStart(VCH.pad, "0");
function acak(n) {
  const c = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"; let s = "";
  for (let i = 0; i < n; i++) s += c.charAt(Math.floor(Math.random() * c.length));
  return s;
}
// Beri nomor urut + token ke satu baris, hanya sekali dan hanya bila status konseling "Selesai"; setelah itu tidak berubah.
function terbit(sh, v, i) {
  const h = v[0], cS = h.indexOf("Status Konseling"), cV = h.indexOf("No Urut Voucher"), cT = h.indexOf("Voucher Token");
  let no = Number(v[i][cV]) || 0, tok = String(v[i][cT] || "");
  if (String(v[i][cS]) !== "Selesai") return { no: no || "", token: tok };
  if (!no) { no = v.slice(1).reduce((m, r) => Math.max(m, Number(r[cV]) || 0), 0) + 1; sh.getRange(i + 1, cV + 1).setValue(no); v[i][cV] = no; }
  if (!tok) { tok = acak(6); sh.getRange(i + 1, cT + 1).setValue(tok); v[i][cT] = tok; }
  return { no: no, token: tok };
}
function cariVoucher(no, token) {
  const n = Number(no), t = String(token || "");
  if (!n || !t) return null;
  const sh = SpreadsheetApp.openById(SS_ID).getSheetByName(VCH.asal);
  if (!sh || sh.getLastRow() < 2) return null;
  ensureHeader(sh);
  const v = sh.getRange(1, 1, sh.getLastRow(), HEAD.length).getValues(), h = v[0], c = k => h.indexOf(k);
  for (let i = 1; i < v.length; i++) {
    if (Number(v[i][c("No Urut Voucher")]) !== n) continue;
    if (String(v[i][c("Voucher Token")]) !== t || String(v[i][c("Status Konseling")]) !== "Selesai") return null;
    return { sh: sh, v: v, i: i, c: c };
  }
  return null;
}
// Publik: dipanggil saat QR dibuka. Memeriksa voucher dan menandai "sudah di-scan".
function voucherCek(d) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const vc = cariVoucher(d.no, d.t);
    if (!vc) { Utilities.sleep(700); return { ok: true, valid: false, alasan: "Voucher tidak valid" }; }
    const cS = vc.c("Voucher Scan"), dipakai = !!vc.v[vc.i][vc.c("Voucher Dipakai")];
    if (!vc.v[vc.i][cS]) vc.sh.getRange(vc.i + 1, cS + 1).setValue(new Date());
    return { ok: true, valid: !dipakai, alasan: dipakai ? "Voucher sudah dipakai" : "", diskon: VCH.diskon, tujuan: VCH.tujuan,
      kode: kodeVoucher(d.no), nama: String(vc.v[vc.i][vc.c("Nama")] || "").split(" ")[0] };
  } finally { lock.releaseLock(); }
}
// Admin: terbitkan (atau ambil) nomor + token untuk peserta yang konselingnya sudah Selesai.
function voucherTerbit(d) {
  const err = cekPin(d); if (err) return { ok: false, error: err };
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = SpreadsheetApp.openById(SS_ID).getSheetByName(String(d.kode || ""));
    if (!sh || sh.getRange(1, 1).getValue() !== "Waktu") throw new Error("Sheet kelas tidak ditemukan");
    ensureHeader(sh);
    const v = sh.getRange(1, 1, sh.getLastRow(), HEAD.length).getValues(), h = v[0];
    const cId = h.indexOf("ID Daftar"), cKe = h.indexOf("Peserta ke"), cS = h.indexOf("Status Konseling");
    for (let i = 1; i < v.length; i++) {
      if (String(v[i][cId]) === String(d.id) && String(v[i][cKe]) === String(d.ke)) {
        if (String(v[i][cS]) !== "Selesai") throw new Error("Konseling belum berstatus Selesai");
        const t = terbit(sh, v, i);
        return { ok: true, no: t.no, token: t.token };
      }
    }
    throw new Error("Data peserta tidak ditemukan");
  } finally { lock.releaseLock(); }
}

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.action === "daftar") return out(daftar(d));
    if (d.action === "admin") return out(admin(d));
    if (d.action === "konseling") return out(konseling(d));
    if (d.action === "verifikasi") return out(verifikasi(d));
    if (d.action === "bukti") return out(bukti(d));
    if (d.action === "voucher_cek") return out(voucherCek(d));
    if (d.action === "voucher_terbit") return out(voucherTerbit(d));
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
    const cId = h.indexOf("ID Daftar"), cKe = h.indexOf("Peserta ke"), cH = h.indexOf("Hasil Konseling"), cN = h.indexOf("Nama Konselor"), cS = h.indexOf("Status Konseling"), cV = h.indexOf("No Urut Voucher");
    const st = String(d.status || "");
    if (["", "Selesai", "Eskalasi", "Risiko tinggi"].indexOf(st) < 0) throw new Error("Status konseling tidak valid");
    for (let i = 1; i < v.length; i++) {
      if (String(v[i][cId]) === String(d.id) && String(v[i][cKe]) === String(d.ke)) {
        sh.getRange(i + 1, cH + 1).setValue(safe(String(d.hasil || "").slice(0, 3000)));
        sh.getRange(i + 1, cN + 1).setValue(safe(String(d.konselor || "").slice(0, 120)));
        sh.getRange(i + 1, cS + 1).setValue(st);
        v[i][cS] = st;
        const t = terbit(sh, v, i);   // nomor urut + token voucher: diberikan sekali saat status jadi "Selesai"
        return { ok: true, no: t.no, token: t.token };
      }
    }
    throw new Error("Data peserta tidak ditemukan");
  } finally { lock.releaseLock(); }
}

// Admin menandai bukti bayar "Sesuai" / "Tidak sesuai" untuk satu pendaftaran (semua baris dengan ID Daftar yang sama)
function verifikasi(d) {
  const err = cekPin(d); if (err) return { ok: false, error: err };
  const status = d.status === "Sesuai" ? "Sesuai" : d.status === "Tidak sesuai" ? "Tidak sesuai" : null;
  if (!status) return { ok: false, error: "Status tidak valid" };
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = SpreadsheetApp.openById(SS_ID).getSheetByName(String(d.kode || ""));
    if (!sh || sh.getRange(1, 1).getValue() !== "Waktu") throw new Error("Sheet kelas tidak ditemukan");
    const v = sh.getDataRange().getValues(), h = v[0], cId = h.indexOf("ID Daftar"), cS = h.indexOf("Status Bayar");
    let n = 0;
    for (let i = 1; i < v.length; i++) if (String(v[i][cId]) === String(d.id)) { sh.getRange(i + 1, cS + 1).setValue(status); n++; }
    if (!n) throw new Error("Data pendaftaran tidak ditemukan");
    return { ok: true, updated: n };
  } finally { lock.releaseLock(); }
}

// Ambil gambar bukti bayar dari Drive (hanya file di folder bukti) dan kirim sebagai base64 agar bisa tampil di halaman admin
function bukti(d) {
  const err = cekPin(d); if (err) return { ok: false, error: err };
  const id = String(d.id || "");
  if (!/^[\w-]{10,}$/.test(id)) return { ok: false, error: "ID file tidak valid" };
  const f = DriveApp.getFileById(id), par = f.getParents();
  let dalamFolder = false;
  while (par.hasNext()) if (par.next().getName() === FOLDER_NAME) dalamFolder = true;
  if (!dalamFolder) return { ok: false, error: "File bukan bukti pembayaran" };
  const b = f.getBlob();
  return { ok: true, mime: b.getContentType(), data: Utilities.base64Encode(b.getBytes()) };
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
    let vch = null;                                         // voucher diskon (hanya untuk kelas tujuan, sekali pakai)
    if (d.voucher && d.voucher.no) {
      if (kode !== VCH.tujuan) throw new Error("Voucher tidak berlaku untuk kelas ini");
      vch = cariVoucher(d.voucher.no, d.voucher.t);
      if (!vch) throw new Error("Voucher tidak valid");
      if (vch.v[vch.i][vch.c("Voucher Dipakai")]) throw new Error("Voucher sudah dipakai");
    }
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
      safe(p.lama_praktik || ""), safe(p.tempat_kerja || ""), link, "Menunggu verifikasi"].concat(phq(p)));
    rows.forEach(r => { while (r.length < HEAD.length) r.push(""); if (vch) r[HEAD.indexOf("Kode Voucher")] = kodeVoucher(d.voucher.no); });
    sh.getRange(start, 11, rows.length, 1).setNumberFormat("@");   // No WA tetap teks (0 depan aman)
    sh.getRange(start, 1, rows.length, HEAD.length).setValues(rows);
    if (vch) vch.sh.getRange(vch.i + 1, vch.c("Voucher Dipakai") + 1).setValue(id);   // voucher terpakai
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
        ? Utilities.formatDate(r[i], "Asia/Jakarta", "dd/MM/yyyy HH:mm") : r[i]);
      return o;
    });
  });
  return { ok: true, data: data };
}
