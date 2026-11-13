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
const FOLDER_SMART = "Dokumen SMART Goal Psychosoul";
const VCH_SEMUA = "C11.3.26.03.2";   // kelas Reflektif (C11.3.26.03.2): voucher 10% semua kelas, tanpa syarat status konseling
const SMART_MAX_B64 = 7 * 1024 * 1024;                  // batas ~5 MB file (base64 lebih besar)
const SMART_MIME = {pdf: "application/pdf", doc: "application/msword", docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", jpg: "image/jpeg", jpeg: "image/jpeg", png: "image/png"};
const HEAD = ["Waktu","ID Daftar","Kode Kelas","Kelas","Pembiayaan","Total","Peserta ke","Nama","Usia",
  "Jenis Kelamin","No WA","Alamat","Pekerjaan","Jurusan","Semester","Tempat Pendidikan",
  "Lama Praktik (thn)","Tempat Kerja","Bukti Bayar","Status Bayar",
  "PHQ-4 Jawaban","PHQ-4 Cemas","PHQ-4 Depresi","PHQ-4 Total",
  "Hasil Konseling","Nama Konselor","Status Konseling","No Urut Voucher","Voucher Diserahkan","Kode Voucher","Dokumen SMART Goal","Tema Curhat","Persetujuan","Waktu Persetujuan","Versi Persetujuan"];   // 3 kolom terakhir diisi admin lewat halaman admin; 4 kolom PHQ-4 hanya terisi untuk kelas yang memakai skrining

const out = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
const safe = v => (typeof v === "string" && /^[=+\-@]/.test(v)) ? "'" + v : (v == null ? "" : v); // cegah formula injection

function doGet() { return out({ ok: true, msg: "Psychosoul API aktif" }); }

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.action === "status_daftar") return out({ ok: true, tutup: daftarTutup() });   // publik: kelas yang pendaftarannya ditutup admin
    if (d.action === "tutup") return out(tutupDaftar(d));
    if (d.action === "daftar") return out(daftar(d));
    if (d.action === "admin") return out(admin(d));
    if (d.action === "konseling") return out(konseling(d));
    if (d.action === "voucher_serah") return out(voucherSerah(d));
    if (d.action === "verifikasi") return out(verifikasi(d));
    if (d.action === "bukti") return out(bukti(d));
    if (d.action === "smartfile") return out(smartfile(d));
    return out({ ok: false, error: "Aksi tidak dikenal" });
  } catch (err) { return out({ ok: false, error: String(err.message || err) }); }
}

function folderSmart() {
  const it = DriveApp.getFoldersByName(FOLDER_SMART);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_SMART);
}

// Ambil dokumen SMART Goal (dilindungi PIN) untuk preview/download di halaman admin
function smartfile(d) {
  const err = cekPin(d); if (err) return { ok: false, error: err };
  const sh = SpreadsheetApp.openById(SS_ID).getSheetByName(String(d.kode || ""));
  if (!sh || sh.getRange(1, 1).getValue() !== "Waktu") throw new Error("Sheet kelas tidak ditemukan");
  const v = sh.getDataRange().getValues(), h = v[0];
  const cId = h.indexOf("ID Daftar"), cKe = h.indexOf("Peserta ke"), cSG = h.indexOf("Dokumen SMART Goal");
  if (cSG < 0) throw new Error("Belum ada dokumen SMART Goal");
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][cId]) === String(d.id) && String(v[i][cKe]) === String(d.ke)) {
      const m = String(v[i][cSG] || "").match(/[-\w]{25,}/);
      if (!m) throw new Error("Belum ada dokumen SMART Goal");
      const f = DriveApp.getFileById(m[0]), blob = f.getBlob();
      return { ok: true, name: f.getName().replace(/^[^_]*_/, ""), mime: blob.getContentType(), data: Utilities.base64Encode(blob.getBytes()) };
    }
  }
  throw new Error("Data peserta tidak ditemukan");
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
    const cId = h.indexOf("ID Daftar"), cKe = h.indexOf("Peserta ke"), cH = h.indexOf("Hasil Konseling"), cN = h.indexOf("Nama Konselor"), cS = h.indexOf("Status Konseling"), cV = h.indexOf("No Urut Voucher"), cSG = h.indexOf("Dokumen SMART Goal"), cT = h.indexOf("Tema Curhat");
    const st = String(d.status || "");
    if (["", "Selesai", "Eskalasi", "Risiko tinggi"].indexOf(st) < 0) throw new Error("Status konseling tidak valid");
    const tema = String(d.tema || "");
    if (["", "Heart Talks", "Social Circus", "Career Crunch", "What’s Next?", "Lain-lain"].indexOf(tema) < 0) throw new Error("Tema curhat tidak valid");
    for (let i = 1; i < v.length; i++) {
      if (String(v[i][cId]) === String(d.id) && String(v[i][cKe]) === String(d.ke)) {
        sh.getRange(i + 1, cT + 1).setValue(tema);
        sh.getRange(i + 1, cH + 1).setValue(safe(String(d.hasil || "").slice(0, 3000)));
        sh.getRange(i + 1, cN + 1).setValue(safe(String(d.konselor || "").slice(0, 120)));
        sh.getRange(i + 1, cS + 1).setValue(st);            // status kosong = draft: isian tetap tersimpan
        let linkSG = String(v[i][cSG] || "");
        if (d.smart && d.smart.data) {
          const nm = String(d.smart.name || "smart-goal").replace(/[\\\/:*?"<>|]/g, "_").slice(0, 80);
          const ext = (nm.split(".").pop() || "").toLowerCase();
          if (!SMART_MIME[ext]) throw new Error("Format dokumen harus PDF, Word, JPG, atau PNG");
          if (String(d.smart.data).length > SMART_MAX_B64) throw new Error("Ukuran dokumen melebihi 5 MB");
          const blob = Utilities.newBlob(Utilities.base64Decode(d.smart.data), SMART_MIME[ext], String(d.id) + "_" + nm);
          linkSG = folderSmart().createFile(blob).getUrl();
          sh.getRange(i + 1, cSG + 1).setValue(linkSG);
        }
        // No urut voucher = urutan sesi konseling selesai; diberikan sekali saat status catatan pertama kali diisi, lalu tidak berubah
        let no = Number(v[i][cV]) || 0;
        if (st && !no) {   // status follow-up apa pun = sesi selesai
          no = v.slice(1).reduce((m, r) => Math.max(m, Number(r[cV]) || 0), 0) + 1;
          sh.getRange(i + 1, cV + 1).setValue(no);
        }
        return { ok: true, no: no || "", smart: linkSG };
      }
    }
    throw new Error("Data peserta tidak ditemukan");
  } finally { lock.releaseLock(); }
}

// Admin menandai voucher sudah/belum diserahkan ke peserta (hanya bila konseling berstatus Selesai)
function voucherSerah(d) {
  const err = cekPin(d); if (err) return { ok: false, error: err };
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = SpreadsheetApp.openById(SS_ID).getSheetByName(String(d.kode || ""));
    if (!sh || sh.getRange(1, 1).getValue() !== "Waktu") throw new Error("Sheet kelas tidak ditemukan");
    ensureHeader(sh);
    const v = sh.getRange(1, 1, sh.getLastRow(), HEAD.length).getValues(), h = v[0];
    const cId = h.indexOf("ID Daftar"), cKe = h.indexOf("Peserta ke"), cS = h.indexOf("Status Konseling"), cD = h.indexOf("Voucher Diserahkan");
    for (let i = 1; i < v.length; i++) {
      if (String(v[i][cId]) === String(d.id) && String(v[i][cKe]) === String(d.ke)) {
        if (d.serah && String(d.kode) !== VCH_SEMUA && String(v[i][cS]).trim() === "") throw new Error("Sesi konseling belum selesai (status catatan belum diisi)");
        const w = d.serah ? Utilities.formatDate(new Date(), "Asia/Jakarta", "dd MMM yyyy, HH.mm") : "";
        sh.getRange(i + 1, cD + 1).setValue(w);
        return { ok: true, waktu: w };
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

// Voucher diskon: kode = 2611.DE09.NNN, NNN = "No Urut Voucher" peserta Jajan Curhat. Harus dikenal dan belum pernah dipakai.
const VCH_PREFIX = "2611.DE09", VCH_SUMBER = "C11.3.26.03.1";
function cekVoucher(ss, sh, kode) {
  const kv = String(kode || "").trim().toUpperCase(), m = kv.match(/^2611\.DE09\.(\d{3})$/);
  if (!m) throw new Error("Format kode voucher tidak valid");
  const src = ss.getSheetByName(VCH_SUMBER);
  let dikenal = false;
  if (src && src.getLastRow() > 1) {
    const v = src.getDataRange().getValues(), c = v[0].indexOf("No Urut Voucher");
    dikenal = c >= 0 && v.slice(1).some(r => Number(r[c]) === Number(m[1]));
  }
  if (!dikenal) throw new Error("Kode voucher tidak dikenali");
  if (sh.getLastRow() > 1) {
    const v = sh.getDataRange().getValues(), c = v[0].indexOf("Kode Voucher");
    if (c >= 0 && v.slice(1).some(r => String(r[c]).toUpperCase() === kv)) throw new Error("Kode voucher sudah digunakan");
  }
  return kv;
}

// ID pendaftaran. Jajan Curhat: regcurhat2611.001, .002, ... (nomor urut pendaftaran, bukan per peserta). Kelas lain: REG-tanggal-jam-acak.
const ID_PREFIX = { "C11.3.26.03.1": "regcurhat2611." };
function buatId(sh, kode) {
  const pre = ID_PREFIX[kode];
  if (!pre) return "REG-" + Utilities.formatDate(new Date(), "Asia/Jakarta", "yyMMdd-HHmmss") + "-" + Math.floor(100 + Math.random() * 900);
  let maks = 0;
  if (sh.getLastRow() > 1) {
    const c = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].indexOf("ID Daftar");
    if (c >= 0) sh.getRange(2, c + 1, sh.getLastRow() - 1, 1).getValues().forEach(r => {
      const x = String(r[0]); if (x.indexOf(pre) === 0) maks = Math.max(maks, parseInt(x.slice(pre.length), 10) || 0);
    });
  }
  return pre + ("00" + (maks + 1)).slice(-3);
}

// Jam buka pendaftaran per kode kelas (WIB). Sebelum waktu ini, server menolak pendaftaran. Kosongkan/hapus entri bila tidak dipakai.
const BUKA_JAM = { "C11.3.26.03.1": "2026-11-14T13:00:00+07:00", "C11.3.26.03.2": "2026-11-15T13:00:00+07:00" };

function daftar(d) {
  const kode = String(d.kode || "").replace(/[\[\]*?:\/\\]/g, "_").slice(0, 90);
  if (!kode || !d.peserta || !d.peserta.length) throw new Error("Data pendaftaran tidak lengkap");
  if (daftarTutup().indexOf(kode) >= 0) throw new Error("Pendaftaran kelas ini sudah ditutup.");
  if (BUKA_JAM[kode] && Date.now() < Date.parse(BUKA_JAM[kode])) throw new Error("Pendaftaran belum dibuka. Silakan coba lagi sesuai jadwal pembukaan.");
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.openById(SS_ID);
    let sh = ss.getSheetByName(kode);                       // satu sheet per kode kelas/event
    if (!sh) {
      sh = ss.insertSheet(kode);
      sh.appendRow(HEAD); sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, HEAD.length).setFontWeight("bold");
    } else ensureHeader(sh);
    const kv = String(d.pembiayaan) === "voucher" ? cekVoucher(ss, sh, d.voucher) : "";
    const id = buatId(sh, kode);
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
    const pj = d.persetujuan || {};
    rows.forEach(r => { while (r.length < HEAD.length) r.push(""); r[HEAD.indexOf("Kode Voucher")] = kv;
      if (pj.setuju) {   // bukti persetujuan (informed consent), hanya kelas yang memakainya
        r[HEAD.indexOf("Persetujuan")] = "Setuju" + (pj.izinPeserta ? " + izin peserta lain" : "");
        r[HEAD.indexOf("Waktu Persetujuan")] = Utilities.formatDate(new Date(), "Asia/Jakarta", "dd/MM/yyyy HH:mm");
        r[HEAD.indexOf("Versi Persetujuan")] = safe(pj.versi);
      } });   // lengkapi kolom admin (konseling, voucher) yang kosong
    sh.getRange(start, 11, rows.length, 1).setNumberFormat("@");   // No WA tetap teks (0 depan aman)
    sh.getRange(start, 1, rows.length, HEAD.length).setValues(rows);
    return { ok: true, id: id };
  } finally { lock.releaseLock(); }
}

// Tutup/buka pendaftaran manual dari laman admin. Disimpan di Script properties (kunci TUTUP_DAFTAR).
function daftarTutup() {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty("TUTUP_DAFTAR") || "[]"); } catch (e) { return []; }
}
function tutupDaftar(d) {
  const err = cekPin(d); if (err) return { ok: false, error: err };
  const kode = String(d.kode || "").slice(0, 90); if (!kode) return { ok: false, error: "Kode kelas kosong" };
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const set = daftarTutup().filter(x => x !== kode);
    if (d.tutup) set.push(kode);
    PropertiesService.getScriptProperties().setProperty("TUTUP_DAFTAR", JSON.stringify(set));
    return { ok: true, tutup: set };
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
