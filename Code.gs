/**
 * ============================================================================
 * INSAN STATISTIK TELADAN — BPS Kabupaten Buleleng
 * Google Apps Script — dipakai sebagai API JSON saja.
 * ============================================================================
 *
 * Antarmuka (index.html, dashboard.html) di-hosting di GitHub Pages dan
 * memanggil berkas ini lewat fetch. Spreadsheet TIDAK PERNAH diakses langsung
 * oleh peramban, jadi kredensialnya tetap di sisi server.
 *
 * === PASANG ===
 *  1. Tempel berkas ini ke proyek Apps Script.
 *  2. Jalankan siapkanSheetVote() sekali.
 *  3. Deploy → Web app → Execute as: Me, Who has access: Anyone.
 *  4. Salin URL /exec, tempel ke API_URL di index.html dan dashboard.html.
 *  5. Setel Spreadsheet ke Restricted (jangan "anyone with link").
 *
 * === ATURAN YANG BERUBAH DI TAHAP INI ===
 *  - Tanpa kode/sandi pemilih.
 *  - Sekali kirim saja: submit kedua DITOLAK (bukan lagi menimpa/idempoten).
 *  - Nama yang sudah mengirim dinonaktifkan di daftar pemilih.
 *  - Rekap hanya lewat dashboard bersandi, tidak lagi ditampilkan ke pemilih.
 */

// === SECTION: KONFIGURASI ===============================================

const SPREADSHEET_ID = "https://docs.google.com/spreadsheets/d/1SoERLVWrt5hFXZNrgP7v0NydMhOv-lWQwYUZSRKi2Mg/edit?usp=sharing";
const FOLDER_FOTO_ID = "https://drive.google.com/drive/folders/1hpTR4lAI1-ClvFh5P5Dk1WQ6EdCpJD0G?usp=sharing";
const POIN = { 1: 3, 2: 2, 3: 1 };   // Pilihan 1 / 2 / 3

const SHEET_PEGAWAI   = 'pegawai';
const SHEET_INDIKATOR = 'indikator';
const SHEET_VOTE      = 'vote';
const SHEET_REKAP     = 'rekap';

const CACHE_DETIK = 21600;  // 6 jam
const FOTO_LEBAR  = 400;    // lh3 =w400 → ±95 KB per PNG, alpha bertahan
const KUNCI_MS    = 30000;  // LockService

const DESKRIPSI_FALLBACK = 'Siapa yang paling menunjukkan sikap ini dalam keseharian kerja?';

/**
 * Babak dikelompokkan dari awalan id kelompok. NAMANYA tidak ditulis di sini —
 * diambil dari kolom kelompokNama di sheet, supaya ejaan selalu ikut sheet.
 */
const POLA_BABAK = [
  { pola: /^BO5/i, id: 'BO5' },
  { pola: /^BE4/i, id: 'BE4' },
  { pola: /^CV/i,  id: 'CV'  }
];
const BABAK_LAIN = { id: 'LAIN' };

/**
 * Sandi dashboard. Diperiksa di SERVER — data rekap tidak pernah dikirim
 * sebelum cocok, jadi yang bisa bocor hanyalah sandinya, bukan hasilnya.
 * Untuk menyembunyikannya dari kode, isi Script Property bernama SANDI_ADMIN;
 * nilai di bawah hanya dipakai bila property itu kosong.
 */
function sandiAdmin_() {
  return PropertiesService.getScriptProperties().getProperty('SANDI_ADMIN') || 'admin5108';
}


// === SECTION: ROUTER HTTP ===============================================

function json_(obj, callback) {
  const teks = JSON.stringify(obj);
  if (callback && /^[A-Za-z_$][\w$]*$/.test(callback)) {
    return ContentService.createTextOutput(callback + '(' + teks + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(teks).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Baca. Mendukung JSONP lewat ?callback= sebagai cadangan bila CORS diblokir.
 *   ?aksi=master        daftar pegawai + indikator + siapa yang sudah mengisi
 *   ?aksi=dashboard&sandi=...   seluruh agregasi untuk dashboard
 */
function doGet(e) {
  const p = (e && e.parameter) || {};
  const cb = p.callback;
  try {
    switch (p.aksi) {
      case 'master':    return json_(getMasterData(), cb);
      case 'dashboard': return json_(getDashboard(p.sandi), cb);
      case 'cekSandi':  return json_({ ok: teks_(p.sandi) === sandiAdmin_() }, cb);
      default:
        return json_({ ok: true, layanan: 'Insan Statistik Teladan API',
                       aksi: ['master', 'dashboard', 'cekSandi'] }, cb);
    }
  } catch (err) {
    return json_({ ok: false, message: String(err && err.message ? err.message : err) }, cb);
  }
}

/**
 * Tulis. Klien WAJIB mengirim Content-Type: text/plain agar tidak memicu
 * preflight OPTIONS — Apps Script tidak bisa menjawab preflight.
 */
function doPost(e) {
  try {
    const isi = (e && e.postData && e.postData.contents) || '{}';
    const p = JSON.parse(isi);
    if (p.aksi === 'vote') return json_(submitVote(p));
    return json_({ ok: false, message: 'Aksi tidak dikenali.' });
  } catch (err) {
    return json_({ ok: false, message: String(err && err.message ? err.message : err) });
  }
}


// === SECTION: UTILITAS ==================================================

function idDari_(s) {
  const t = String(s || '').trim();
  const m = t.match(/[-\w]{25,}/);
  return m ? m[0] : t;
}

function ss_() { return SpreadsheetApp.openById(idDari_(SPREADSHEET_ID)); }

function sheet_(nama) {
  const target = String(nama).trim().toLowerCase();
  const semua = ss_().getSheets();
  for (let i = 0; i < semua.length; i++) {
    if (semua[i].getName().trim().toLowerCase() === target) return semua[i];
  }
  throw new Error('Sheet "' + nama + '" tidak ditemukan pada spreadsheet.');
}

function teks_(v) { return String(v === null || v === undefined ? '' : v).trim(); }

function bool_(v) {
  if (v === true) return true;
  if (v === false || v === null || v === undefined) return false;
  const s = String(v).trim().toLowerCase();
  return s === 'true' || s === '1' || s === 'ya' || s === 'yes' || s === 'y';
}

function bacaTabel_(namaSheet) {
  const nilai = sheet_(namaSheet).getDataRange().getValues();
  if (!nilai.length) return { kolom: {}, baris: [] };
  const kolom = {};
  nilai[0].forEach(function (h, i) {
    const k = teks_(h).toLowerCase();
    if (k && !(k in kolom)) kolom[k] = i;
  });
  return { kolom: kolom, baris: nilai.slice(1), header: nilai[0] };
}

function sel_(baris, kolom, nama) {
  const i = kolom[String(nama).toLowerCase()];
  return i === undefined ? '' : baris[i];
}

function normNama_(s) { return String(s || '').trim().toLowerCase().replace(/\s+/g, ' '); }
function namaPendek_(n) { return String(n || '').split(',')[0].trim(); }
function gelarDari_(n) {
  const p = String(n || '').split(',');
  return p.length > 1 ? p.slice(1).join(',').trim() : '';
}
function inisial_(n) {
  const w = String(n || '').replace(/,.*/, '').trim().split(/\s+/).filter(Boolean);
  return (((w[0] || '')[0] || '') + ((w[1] || '')[0] || '')).toUpperCase() || '?';
}
function babakDari_(kelompokId) {
  for (let i = 0; i < POLA_BABAK.length; i++) {
    if (POLA_BABAK[i].pola.test(String(kelompokId))) return POLA_BABAK[i];
  }
  return BABAK_LAIN;
}


// === SECTION: FOTO ======================================================

/** Peta nama berkas → id Drive, di-cache 6 jam. */
function petaFoto_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('petaFoto_v1');
  if (hit) { try { return JSON.parse(hit); } catch (e) {} }

  const peta = {};
  try {
    const berkas = DriveApp.getFolderById(idDari_(FOLDER_FOTO_ID)).getFiles();
    while (berkas.hasNext()) {
      const f = berkas.next();
      const nama = f.getName();
      peta[normNama_(nama)] = f.getId();
      peta[normNama_(nama.replace(/\.[a-z0-9]+$/i, ''))] = f.getId();
    }
  } catch (e) { /* folder tak terbaca → klien pakai avatar inisial */ }

  try { cache.put('petaFoto_v1', JSON.stringify(peta), CACHE_DETIK); } catch (e) {}
  return peta;
}

/**
 * Diuji 2026-08-05: lh3.googleusercontent.com/d/<ID>=w400 mengembalikan PNG
 * RGBA (alpha utuh). Klien WAJIB memakai referrerpolicy="no-referrer" —
 * lh3 menolak permintaan gambar yang membawa header Referer.
 */
function urlFoto_(isiSel, peta, namaPegawai) {
  const v = teks_(isiSel);
  if (!v) return '';
  if (/drive\.google\.com|googleusercontent\.com/i.test(v)) {
    const id = idDari_(v);
    return id ? lh3_(id) : '';
  }
  if (/^[-\w]{25,}$/.test(v)) return lh3_(v);
  let id = peta[normNama_(v)] || peta[normNama_(v.replace(/\.[a-z0-9]+$/i, ''))];
  if (!id && namaPegawai) id = peta[normNama_(namaPendek_(namaPegawai))];
  return id ? lh3_(id) : '';
}

function lh3_(id) { return 'https://lh3.googleusercontent.com/d/' + id + '=w' + FOTO_LEBAR; }


// === SECTION: BACA MASTER ===============================================

function bacaPegawai_() {
  const t = bacaTabel_(SHEET_PEGAWAI);
  const peta = petaFoto_();
  const out = [];
  t.baris.forEach(function (r) {
    const id = teks_(sel_(r, t.kolom, 'id'));
    const nama = teks_(sel_(r, t.kolom, 'nama'));
    if (!id || !nama) return;
    const mentah = teks_(sel_(r, t.kolom, 'fotourl'));
    const url = urlFoto_(mentah, peta, nama);
    out.push({
      id: id, nama: nama, pendek: namaPendek_(nama), gelar: gelarDari_(nama),
      inisial: inisial_(nama),
      bolehMemilih: bool_(sel_(r, t.kolom, 'bolehmemilih')),
      bolehDipilih: bool_(sel_(r, t.kolom, 'bolehdipilih')),
      foto: url, fotoMentah: mentah
    });
  });
  return out;
}

function bacaIndikator_() {
  const t = bacaTabel_(SHEET_INDIKATOR);
  const out = [];
  t.baris.forEach(function (r) {
    const indikatorId = teks_(sel_(r, t.kolom, 'indikatorid'));
    const kelompokId = teks_(sel_(r, t.kolom, 'kelompokid'));
    if (!indikatorId || !kelompokId) return;

    // Kolom rincian pernah bernama `deskripsi`. Keduanya diterima supaya
    // penggantian nama kolom di sheet tidak menghapus 75 rincian secara senyap.
    let rinci = teks_(sel_(r, t.kolom, 'rincian')) || teks_(sel_(r, t.kolom, 'deskripsi'));
    if (rinci === '-' || rinci === '\\-') rinci = '';

    out.push({
      babakId: babakDari_(kelompokId).id,
      kelompokId: kelompokId,
      kelompokNama: teks_(sel_(r, t.kolom, 'kelompoknama')) || kelompokId,
      indikatorId: indikatorId,
      indikatorNama: teks_(sel_(r, t.kolom, 'indikatornama')) || indikatorId,
      rincian: rinci,
      deskripsi: rinci || DESKRIPSI_FALLBACK,   // dipakai UI sebagai teks pertanyaan
      urutan: Number(sel_(r, t.kolom, 'urutan')) || (out.length + 1)
    });
  });
  out.sort(function (a, b) { return a.urutan - b.urutan; });

  // Nama babak diambil dari DATA (kelompokNama baris pertama tiap babak),
  // bukan dari nama yang ditulis di kode — supaya ejaan selalu ikut sheet.
  const namaBabak = {};
  out.forEach(function (i) {
    if (!namaBabak[i.babakId]) namaBabak[i.babakId] = i.kelompokNama;
  });

  // Judul kelompok: pakai indikatorNama bila SERAGAM dalam satu kelompokId
  // (struktur baru — mis. seluruh BO5-1 bernama "Be a Leader, not a Boss"),
  // selain itu pakai kelompokNama. Menjaga tiga lapis tetap terbaca pada
  // kedua bentuk sheet tanpa mengasumsikan salah satunya.
  const kumpul = {};
  out.forEach(function (i) {
    (kumpul[i.kelompokId] = kumpul[i.kelompokId] || []).push(i.indikatorNama);
  });
  const judulKelompok = {};
  Object.keys(kumpul).forEach(function (kid) {
    const semua = kumpul[kid];
    const seragam = semua.every(function (n) { return n === semua[0]; });
    judulKelompok[kid] = seragam ? semua[0] : '';
  });

  out.forEach(function (i) {
    i.babakNama = namaBabak[i.babakId] || i.kelompokNama;
    i.judulKelompok = judulKelompok[i.kelompokId] || i.kelompokNama;
    // Bila judul kelompok sudah memuat indikatorNama, jangan diulang di kartu.
    i.labelIndikator = (i.indikatorNama === i.judulKelompok) ? '' : i.indikatorNama;
  });

  return out;
}

function bacaMaster_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('master_v3');
  if (hit) { try { return JSON.parse(hit); } catch (e) {} }
  const master = { pegawai: bacaPegawai_(), indikator: bacaIndikator_() };
  try { cache.put('master_v3', JSON.stringify(master), CACHE_DETIK); } catch (e) {}
  return master;
}


// === SECTION: BACA VOTE =================================================

function bacaVote_() {
  const t = bacaTabel_(SHEET_VOTE);
  const out = [];
  t.baris.forEach(function (r) {
    const voterId = teks_(sel_(r, t.kolom, 'voterid'));
    const indikatorId = teks_(sel_(r, t.kolom, 'indikatorid'));
    const nomineeId = teks_(sel_(r, t.kolom, 'nomineeid'));
    if (!voterId || !indikatorId || !nomineeId) return;
    out.push({
      voterId: voterId, indikatorId: indikatorId, nomineeId: nomineeId,
      urutan: Number(sel_(r, t.kolom, 'urutan')) || 0,
      poin: Number(sel_(r, t.kolom, 'poin')) || 0
    });
  });
  return out;
}

/** voterId → jumlah indikator berbeda yang sudah terisi. */
function progresPerVoter_(vote) {
  const tampung = {};
  vote.forEach(function (v) {
    if (!tampung[v.voterId]) tampung[v.voterId] = {};
    tampung[v.voterId][v.indikatorId] = 1;
  });
  const out = {};
  Object.keys(tampung).forEach(function (k) { out[k] = Object.keys(tampung[k]).length; });
  return out;
}


// === SECTION: API — master ==============================================

function getMasterData() {
  try {
    const master = bacaMaster_();
    const progres = progresPerVoter_(bacaVote_());

    const pegawai = master.pegawai.map(function (p) {
      return {
        id: p.id, nama: p.nama, pendek: p.pendek, gelar: p.gelar, inisial: p.inisial,
        bolehMemilih: p.bolehMemilih, bolehDipilih: p.bolehDipilih, foto: p.foto,
        // Sekali isi: nama yang sudah mengirim dinonaktifkan di daftar pemilih.
        sudahMengisi: (progres[p.id] || 0) > 0
      };
    });

    const jumlahKandidat = pegawai.filter(function (p) { return p.bolehDipilih; }).length;
    const berhak = pegawai.filter(function (p) { return p.bolehMemilih; });

    return {
      ok: true,
      pegawai: pegawai,
      indikator: master.indikator,
      totalIndikator: master.indikator.length,
      jumlahKandidat: jumlahKandidat,
      galatKonfigurasi: jumlahKandidat < 3
        ? 'Hanya ada ' + jumlahKandidat + ' kandidat dengan bolehDipilih = TRUE. '
          + 'Dibutuhkan minimal 3 agar penilaian dapat dikirim.'
        : '',
      sudahMengisi: berhak.filter(function (p) { return p.sudahMengisi; }).length,
      pemilihBerhak: berhak.length,
      poin: POIN
    };
  } catch (e) {
    return { ok: false, message: 'Gagal membaca data: ' + String(e && e.message ? e.message : e) };
  }
}


// === SECTION: API — submitVote ==========================================

/**
 * payload = { voterId, pilihan: [{indikatorId, nomineeId, urutan}] }
 * SEKALI KIRIM SAJA — bila voterId sudah punya baris, permintaan ditolak.
 */
function submitVote(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(KUNCI_MS);
  } catch (e) {
    return { ok: false, message: 'Sistem sedang menyimpan penilaian lain. Coba lagi beberapa detik.' };
  }

  try {
    const master = bacaMaster_();
    const p = payload || {};
    const voterId = teks_(p.voterId);
    const pilihan = Array.isArray(p.pilihan) ? p.pilihan : [];

    const pemilih = master.pegawai.filter(function (x) { return x.id === voterId; })[0];
    if (!pemilih) return { ok: false, message: 'Pemilih tidak dikenali.' };
    if (!pemilih.bolehMemilih) return { ok: false, message: 'Nama ini tidak berhak memilih.' };

    // --- Sekali kirim: diperiksa DI DALAM kunci agar dua tab tidak lolos bersamaan ---
    const voteLama = bacaVote_();
    if (progresPerVoter_(voteLama)[voterId]) {
      return { ok: false, sudahMengisi: true,
        message: 'Nama ini sudah mengirim penilaian. Setiap orang hanya dapat mengisi satu kali.' };
    }

    const kandidat = {};
    master.pegawai.forEach(function (x) { if (x.bolehDipilih) kandidat[x.id] = x; });
    if (Object.keys(kandidat).length < 3) {
      return { ok: false, message: 'Konfigurasi belum siap: kandidat kurang dari 3 orang.' };
    }

    const indikatorSah = {};
    master.indikator.forEach(function (i) { indikatorSah[i.indikatorId] = i; });
    const total = master.indikator.length;

    const perIndikator = {};
    for (let i = 0; i < pilihan.length; i++) {
      const it = pilihan[i] || {};
      const indId = teks_(it.indikatorId);
      const nomId = teks_(it.nomineeId);
      const urut = Number(it.urutan);

      if (!indikatorSah[indId]) return { ok: false, message: 'Indikator tidak dikenali: ' + indId };
      if (!kandidat[nomId]) return { ok: false, message: 'Kandidat tidak sah pada indikator ' + indId + '.' };
      if (!POIN[urut]) return { ok: false, message: 'Urutan harus 1, 2, atau 3 pada indikator ' + indId + '.' };

      if (!perIndikator[indId]) perIndikator[indId] = {};
      if (perIndikator[indId][urut]) {
        return { ok: false, message: 'Urutan ganda pada indikator ' + indikatorSah[indId].indikatorNama + '.' };
      }
      perIndikator[indId][urut] = nomId;
    }

    const kunci = Object.keys(perIndikator);
    if (kunci.length !== total) {
      return { ok: false,
        message: 'Penilaian belum lengkap: ' + kunci.length + ' dari ' + total + ' indikator terisi.' };
    }
    for (let k = 0; k < kunci.length; k++) {
      const slot = perIndikator[kunci[k]];
      const nama = indikatorSah[kunci[k]].indikatorNama;
      if (!slot[1] || !slot[2] || !slot[3]) {
        return { ok: false, message: 'Indikator "' + nama + '" harus berisi tepat 3 nama.' };
      }
      if (slot[1] === slot[2] || slot[1] === slot[3] || slot[2] === slot[3]) {
        return { ok: false, message: 'Nama ganda dalam indikator "' + nama + '".' };
      }
    }

    // --- Tulis (append; baris lama voterId dijamin tidak ada) ---
    const sh = sheet_(SHEET_VOTE);
    const nilai = sh.getDataRange().getValues();
    const headerDefault = ['timestamp', 'voterId', 'indikatorId', 'nomineeId', 'urutan', 'poin'];
    const header = (nilai.length && teks_(nilai[0][0])) ? nilai[0] : headerDefault;

    const kolom = {};
    header.forEach(function (h, i) { kolom[teks_(h).toLowerCase()] = i; });
    const wajib = ['timestamp', 'voterid', 'indikatorid', 'nomineeid', 'urutan', 'poin'];
    const hilang = wajib.filter(function (k) { return kolom[k] === undefined; });
    if (hilang.length) {
      return { ok: false, message: 'Sheet vote kekurangan kolom: ' + hilang.join(', ')
        + '. Jalankan siapkanSheetVote() lebih dulu.' };
    }

    const stempel = new Date();
    const baru = [];
    master.indikator.forEach(function (ind) {
      const slot = perIndikator[ind.indikatorId];
      [1, 2, 3].forEach(function (u) {
        const baris = new Array(header.length).fill('');
        baris[kolom['timestamp']]   = stempel;
        baris[kolom['voterid']]     = voterId;
        baris[kolom['indikatorid']] = ind.indikatorId;
        baris[kolom['nomineeid']]   = slot[u];
        baris[kolom['urutan']]      = u;
        baris[kolom['poin']]        = POIN[u];
        baru.push(baris);
      });
    });

    const mulai = Math.max(2, nilai.length + (nilai.length ? 1 : 1));
    const barisMulai = (nilai.length ? nilai.length : 1) + 1;
    if (!nilai.length) sh.getRange(1, 1, 1, header.length).setValues([header]);
    sh.getRange(barisMulai, 1, baru.length, header.length).setValues(baru);
    SpreadsheetApp.flush();

    return { ok: true, message: 'Penilaian Anda telah dipersembahkan.',
             jumlahBaris: baru.length, totalIndikator: total };

  } catch (e) {
    return { ok: false, message: 'Gagal menyimpan: ' + String(e && e.message ? e.message : e) };
  } finally {
    lock.releaseLock();
  }
}


// === SECTION: API — dashboard ===========================================

/** Agregasi lengkap untuk dashboard. Digerbang sandi di sisi server. */
function getDashboard(sandi) {
  if (teks_(sandi) !== sandiAdmin_()) {
    return { ok: false, terkunci: true, message: 'Sandi salah.' };
  }
  try {
    return hitungRekap_(bacaMaster_());
  } catch (e) {
    return { ok: false, message: 'Gagal menghitung: ' + String(e && e.message ? e.message : e) };
  }
}

function hitungRekap_(master) {
  const vote = bacaVote_();
  const total = master.indikator.length;
  const progres = progresPerVoter_(vote);

  const petaPeg = {};
  master.pegawai.forEach(function (p) { petaPeg[p.id] = p; });
  const petaInd = {};
  master.indikator.forEach(function (i) { petaInd[i.indikatorId] = i; });

  // Kelompok & babak dalam urutan tampil.
  const kelompok = [], babak = [], lihatK = {}, lihatB = {};
  master.indikator.forEach(function (i) {
    if (!lihatK[i.kelompokId]) {
      lihatK[i.kelompokId] = { id: i.kelompokId, nama: i.kelompokNama,
        babakId: i.babakId, babakNama: i.babakNama, jumlah: 0 };
      kelompok.push(lihatK[i.kelompokId]);
    }
    lihatK[i.kelompokId].jumlah++;
    if (!lihatB[i.babakId]) {
      lihatB[i.babakId] = { id: i.babakId, nama: i.babakNama, jumlah: 0 };
      babak.push(lihatB[i.babakId]);
    }
    lihatB[i.babakId].jumlah++;
  });

  const perOrang = {};
  const skorIndikator = {};
  vote.forEach(function (v) {
    if (!petaPeg[v.nomineeId] || !petaInd[v.indikatorId]) return;
    if (!perOrang[v.nomineeId]) {
      perOrang[v.nomineeId] = { total: 0, suara: 0, p1: 0, p2: 0, p3: 0,
                                perKelompok: {}, perBabak: {}, perIndikator: {} };
    }
    const o = perOrang[v.nomineeId];
    const ind = petaInd[v.indikatorId];
    o.total += v.poin;
    o.suara += 1;
    if (v.urutan === 1) o.p1++; else if (v.urutan === 2) o.p2++; else if (v.urutan === 3) o.p3++;
    o.perIndikator[v.indikatorId] = (o.perIndikator[v.indikatorId] || 0) + v.poin;
    o.perKelompok[ind.kelompokId] = (o.perKelompok[ind.kelompokId] || 0) + v.poin;
    o.perBabak[ind.babakId] = (o.perBabak[ind.babakId] || 0) + v.poin;

    if (!skorIndikator[v.indikatorId]) skorIndikator[v.indikatorId] = {};
    skorIndikator[v.indikatorId][v.nomineeId] =
      (skorIndikator[v.indikatorId][v.nomineeId] || 0) + v.poin;
  });

  // Semua kandidat ikut, termasuk yang nol poin — supaya dashboard jujur.
  const kandidat = master.pegawai.filter(function (p) { return p.bolehDipilih; })
    .map(function (p) {
      const o = perOrang[p.id] || { total: 0, suara: 0, p1: 0, p2: 0, p3: 0,
                                    perKelompok: {}, perBabak: {}, perIndikator: {} };
      return { id: p.id, nama: p.nama, pendek: p.pendek, gelar: p.gelar,
               inisial: p.inisial, foto: p.foto,
               total: o.total, suara: o.suara, p1: o.p1, p2: o.p2, p3: o.p3,
               perKelompok: o.perKelompok, perBabak: o.perBabak, perIndikator: o.perIndikator };
    })
    .sort(function (a, b) { return b.total - a.total || a.nama.localeCompare(b.nama); });

  const partisipasi = master.pegawai.filter(function (p) { return p.bolehMemilih; })
    .map(function (p) {
      return { id: p.id, nama: p.nama, pendek: p.pendek, inisial: p.inisial,
               jumlah: progres[p.id] || 0, tuntas: (progres[p.id] || 0) >= total };
    });

  const tuntas = partisipasi.filter(function (x) { return x.tuntas; }).length;

  return {
    ok: true,
    ringkas: {
      pemilihBerhak: partisipasi.length,
      pemilihTuntas: tuntas,
      totalIndikator: total,
      jumlahKandidat: kandidat.length,
      poinBeredar: kandidat.reduce(function (a, k) { return a + k.total; }, 0),
      poinMaksTeoritis: partisipasi.length * total * 6
    },
    kandidat: kandidat,
    kelompok: kelompok,
    babak: babak,
    indikator: master.indikator.map(function (i) {
      return { indikatorId: i.indikatorId, indikatorNama: i.indikatorNama,
               kelompokId: i.kelompokId, kelompokNama: i.kelompokNama,
               babakId: i.babakId, babakNama: i.babakNama, deskripsi: i.deskripsi };
    }),
    skorIndikator: skorIndikator,
    partisipasi: partisipasi,
    dibuat: new Date().toISOString()
  };
}


// === SECTION: ADMIN (dijalankan dari editor) ============================

function bersihkanCache() {
  CacheService.getScriptCache().removeAll(['master_v3', 'petaFoto_v1']);
  return 'Cache dibersihkan.';
}

function cekFoto() {
  const peg = bacaPegawai_();
  const gagal = peg.filter(function (p) { return !p.foto; })
    .map(function (p) { return p.id + ' — ' + p.nama + ' (fotoUrl: "' + p.fotoMentah + '")'; });
  const laporan = 'Total pegawai: ' + peg.length
    + '\nFoto terpetakan: ' + (peg.length - gagal.length)
    + '\nGagal / kosong:\n  ' + (gagal.length ? gagal.join('\n  ') : '(tidak ada)');
  Logger.log(laporan);
  return laporan;
}

function siapkanSheetVote() {
  const sh = sheet_(SHEET_VOTE);
  if (!teks_(sh.getRange(1, 1).getValue())) {
    sh.getRange(1, 1, 1, 6)
      .setValues([['timestamp', 'voterId', 'indikatorId', 'nomineeId', 'urutan', 'poin']])
      .setFontWeight('bold');
  }
  return 'Sheet vote siap.';
}

/** Cetak hasil ke sheet `rekap`. Angkanya tetap dihitung saat dijalankan. */
function tulisRekap() {
  const master = bacaMaster_();
  const R = hitungRekap_(master);
  if (!R.ok) return R.message || 'Gagal menghitung.';

  const ss = ss_();
  let sh;
  try { sh = sheet_(SHEET_REKAP); } catch (e) { sh = ss.insertSheet(SHEET_REKAP); }
  sh.clear();

  const baris = [];
  baris.push(['REKAP INSAN STATISTIK TELADAN — BPS Kabupaten Buleleng']);
  baris.push(['Dibuat', new Date(), '', 'Pemilih tuntas',
              R.ringkas.pemilihTuntas + ' dari ' + R.ringkas.pemilihBerhak]);
  baris.push([]);
  baris.push(['PERINGKAT KESELURUHAN']);
  baris.push(['#', 'id', 'nama', 'total poin', 'suara', 'pil-1', 'pil-2', 'pil-3']
    .concat(R.kelompok.map(function (k) { return k.nama; })));
  R.kandidat.forEach(function (p, i) {
    baris.push([i + 1, p.id, p.nama, p.total, p.suara, p.p1, p.p2, p.p3]
      .concat(R.kelompok.map(function (k) { return p.perKelompok[k.id] || 0; })));
  });
  baris.push([]);
  baris.push(['LIMA BESAR PER INDIKATOR']);
  baris.push(['babak', 'kelompok', 'indikator', '1', '2', '3', '4', '5']);
  const namaPeg = {};
  R.kandidat.forEach(function (p) { namaPeg[p.id] = p.pendek; });
  R.indikator.forEach(function (ind) {
    const skor = R.skorIndikator[ind.indikatorId] || {};
    const urut = Object.keys(skor)
      .map(function (id) { return { nama: namaPeg[id] || id, poin: skor[id] }; })
      .sort(function (a, b) { return b.poin - a.poin || a.nama.localeCompare(b.nama); })
      .slice(0, 5)
      .map(function (x) { return x.nama + ' (' + x.poin + ')'; });
    while (urut.length < 5) urut.push('');
    baris.push([ind.babakNama, ind.kelompokNama, ind.indikatorNama].concat(urut));
  });
  baris.push([]);
  baris.push(['PARTISIPASI']);
  baris.push(['id', 'nama', 'indikator terisi', 'status']);
  R.partisipasi.forEach(function (x) {
    baris.push([x.id, x.nama, x.jumlah, x.tuntas ? 'TUNTAS' : (x.jumlah ? 'SEBAGIAN' : 'BELUM')]);
  });

  const lebar = baris.reduce(function (a, b) { return Math.max(a, b.length); }, 1);
  const rata = baris.map(function (b) {
    const s = b.slice();
    while (s.length < lebar) s.push('');
    return s;
  });

  sh.getRange(1, 1, rata.length, lebar).setValues(rata);
  sh.getRange(1, 1).setFontWeight('bold').setFontSize(13);
  sh.getRange(5, 1, 1, lebar).setFontWeight('bold');
  SpreadsheetApp.flush();

  return 'Sheet "' + SHEET_REKAP + '" diperbarui — ' + R.kandidat.length + ' kandidat, '
    + R.indikator.length + ' indikator, ' + R.ringkas.pemilihTuntas + ' pemilih tuntas.';
}
