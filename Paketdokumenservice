/**
 * PaketDokumenService.gs
 * ------------------------------------------------------------
 * Tab "Print" di halaman Pengajuan:
 *
 * 1) buatPaketDokumen(idSptjb, token)
 *    Membuat 1 file Google Spreadsheet BARU berisi 4 sheet untuk 1 SPTJB:
 *    Kwitansi, SPBY, SPTJB, Nominatif. Semua angka & nama diambil dari
 *    DB_SPTJB + DB_PENGAJUAN (semua Surat Tugas di bawah SPTJB itu
 *    digabung), REF_AKUN (nama komponen & kode item), dan CONFIG
 *    (satker, DIPA, NPWP, pejabat penandatangan).
 *
 * 2) unduhPdfPaketDokumen(fileId, jenis, token)
 *    Mengekspor SATU sheet dari file paket itu menjadi PDF, lalu
 *    mengirimnya ke browser (base64) supaya langsung terunduh.
 *
 * FORMAT dokumen mengikuti contoh yang dibagikan (versi yang sudah
 * terisi angka). Teks tetap (kop, paragraf resmi, nama bidang, dll)
 * ada di PD_TEKS di bawah -- ubah di sana kalau redaksi resmi berubah.
 *
 * SEKALI SAJA setelah file ini ditambahkan: jalankan
 * otorisasiPaketDokumen() dari editor Apps Script (pilih di dropdown
 * fungsi, klik Run, lalu Allow/Izinkan) karena fitur ini butuh izin
 * Google Drive. Setelah itu buat "New version" di Manage deployments.
 *
 * OPSIONAL: supaya file paket rapi di satu folder, isi Script Property
 * FOLDER_DOKUMEN_ID (Project Settings > Script properties) dengan ID
 * folder Drive tujuan.
 * ------------------------------------------------------------
 */

var PD_PREFIX_NAMA = 'Paket Dokumen SIPERDIN - ';
var PD_KUNCI_METADATA = 'SIPERDIN_PAKET';

// true = file paket otomatis bisa DILIHAT oleh siapa pun yang punya tautannya.
// Perlu true karena pengguna SIPERDIN login manual (belum tentu punya akses Google
// ke Drive pemilik script). Set false kalau mau file hanya bisa dibuka pemilik --
// PDF tetap bisa diunduh dari aplikasi karena dikirim langsung, bukan lewat tautan.
var PD_BAGIKAN_LINK = true;

var PD_FMT_RP = '"Rp "#,##0;"-Rp "#,##0;"Rp -"';

var PD_BULAN = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus',
  'September', 'Oktober', 'November', 'Desember'];

// Urutan sheet di file = urutan array ini. orientasi dipakai saat ekspor PDF.
var PD_LEMBAR = {
  kwitansi:  { nama: 'Kwitansi',  orientasi: 'portrait' },
  spby:      { nama: 'SPBY',      orientasi: 'portrait' },
  sptjb:     { nama: 'SPTJB',     orientasi: 'landscape' },
  nominatif: { nama: 'Nominatif', orientasi: 'landscape' }
};
var PD_URUTAN = ['kwitansi', 'spby', 'sptjb', 'nominatif'];

// Teks tetap dokumen (bukan dari database).
var PD_TEKS = {
  kementerianKop: 'KEMENTERIAN KETENAGAKERJAAN R.I.',
  ditjenKop1: 'DIREKTORAT JENDERAL',
  ditjenKop2: 'PEMBINAAN PENEMPATAN TENAGA KERJA DAN PERLUASAN KESEMPATAN KERJA',
  alamatKop: 'Jalan Jenderal Gatot Subroto Kaveling 51 Lantai 4 Blok A Telp. 5228440 Fax. 5227588 Jakarta Selatan 12950',
  kementerianSpby: 'KEMENTERIAN KETENAGAKERJAAN',
  direktoratSpby: 'DIREKTORAT BINA PERLUASAN KESEMPATAN KERJA',
  bidang: 'Pemberdayaan Tenaga Kerja Mandiri',
  jabatanKetuaTim: 'Ketua Tim Bidang Tenaga Kerja Mandiri',
  jabatanPpk: 'Pejabat Pembuat Komitmen',
  kotaAsal: 'Jakarta',
  penerimaDari: 'Kuasa Pengguna Anggaran/Pejabat Pembuat Komitmen Direktorat Bina Perluasan Kesempatan Kerja',
  paragrafSptjb1: 'Yang bertandatangan di bawah ini atas nama Kuasa Pengguna Anggaran Satuan Kerja Direktorat Bina Perluasan Kesempatan Kerja menyatakan bahwa saya bertanggungjawab secara formal dan material dan kebenaran perhitungan pemungutan pajak atas segala pembayaran tagihan yang telah kami perintahkan dalam SPM ini dengan perincian sebagai berikut :',
  paragrafSptjb2: 'Bukti-bukti pengeluaran anggaran dan asli setoran pajak (SSP/BPN) tersebut di atas disimpan oleh Pengguna Anggaran/Kuasa Pengguna Anggaran untuk kelengkapan administrasi dan pemeriksaan aparat pengawasan fungsional.',
  paragrafSptjb3: 'Demikian Surat Pernyataan ini dibuat dengan sebenarnya.',
  paragrafSpby: 'Saya yang bertanda tangan di bawah ini selaku Pejabat Pembuat Komitmen memerintahkan Bendahara Pengeluaran agar melakukan pembayaran sejumlah :'
};

// Isian CONFIG yang dipakai dokumen; kalau kosong, pengguna diberi peringatan.
var PD_CONFIG_WAJIB = ['kode_satker', 'nama_satker', 'no_dipa', 'tanggal_dipa', 'npwp', 'alamat_satker',
  'nama_ppk', 'nip_ppk', 'nama_ketua_tim_kerja', 'nip_ketua_tim_kerja', 'nama_bendahara', 'nip_bendahara'];

// ============================================================
// FUNGSI YANG DIPANGGIL DARI BROWSER (google.script.run)
// ============================================================

/**
 * Membuat file Spreadsheet paket dokumen untuk 1 SPTJB.
 * @return {Object} {ok:true, fileId, url, nama, jumlahOrang, total, peringatan:[]} atau {ok:false, error}
 */
function buatPaketDokumen(idSptjb, token) {
  let ssBaru = null;
  try {
    if (!getUserByToken_(token)) return { ok: false, error: 'Sesi Anda sudah berakhir. Silakan login kembali.' };

    const D = pdKumpulkanData_(idSptjb);

    const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmm');
    const noAman = pdNamaAman_(D.noSptjb);
    ssBaru = SpreadsheetApp.create(PD_PREFIX_NAMA + noAman + ' - ' + stamp);
    ssBaru.setSpreadsheetLocale('id_ID');
    ssBaru.setSpreadsheetTimeZone(Session.getScriptTimeZone());

    const lembar = {};
    PD_URUTAN.forEach(function (jenis, i) {
      lembar[jenis] = i === 0
        ? ssBaru.getSheets()[0].setName(PD_LEMBAR[jenis].nama)
        : ssBaru.insertSheet(PD_LEMBAR[jenis].nama);
    });

    pdBangunKwitansi_(lembar.kwitansi, D);
    pdBangunSpby_(lembar.spby, D);
    pdBangunSptjb_(lembar.sptjb, D);
    pdBangunNominatif_(lembar.nominatif, D);

    ssBaru.setActiveSheet(lembar.kwitansi);
    // penanda supaya unduhPdfPaketDokumen hanya mau mengekspor file buatan aplikasi ini
    ssBaru.addDeveloperMetadata(PD_KUNCI_METADATA, noAman);
    SpreadsheetApp.flush();

    const fileId = ssBaru.getId();
    const peringatan = D.peringatan.slice();

    try {
      const folderId = PropertiesService.getScriptProperties().getProperty('FOLDER_DOKUMEN_ID');
      if (folderId) DriveApp.getFileById(fileId).moveTo(DriveApp.getFolderById(folderId));
    } catch (e) {
      peringatan.push('File tidak bisa dipindah ke folder tujuan (cek FOLDER_DOKUMEN_ID): ' + e.message);
    }

    if (PD_BAGIKAN_LINK) {
      try {
        DriveApp.getFileById(fileId).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (e) {
        peringatan.push('Tautan spreadsheet tidak bisa dibagikan otomatis (kemungkinan dibatasi kebijakan organisasi). Tombol PDF tetap berfungsi.');
      }
    }

    return {
      ok: true,
      fileId: fileId,
      url: ssBaru.getUrl(),
      nama: ssBaru.getName(),
      jumlahOrang: D.baris.length,
      total: D.jumlah.total,
      peringatan: peringatan
    };
  } catch (err) {
    console.error('buatPaketDokumen gagal: ' + err.message + '\n' + (err.stack || ''));
    if (ssBaru) {
      // file setengah jadi yang baru saja dibuat di panggilan ini -- aman dibuang
      try { DriveApp.getFileById(ssBaru.getId()).setTrashed(true); } catch (e2) { /* abaikan */ }
    }
    return { ok: false, error: 'Paket dokumen gagal dibuat: ' + err.message };
  }
}

/**
 * Ekspor 1 sheet dari file paket menjadi PDF.
 * @param {string} fileId ID spreadsheet hasil buatPaketDokumen
 * @param {string} jenis  'kwitansi' | 'spby' | 'sptjb' | 'nominatif'
 * @return {Object} {ok:true, base64, namaFile} atau {ok:false, error}
 */
function unduhPdfPaketDokumen(fileId, jenis, token) {
  try {
    if (!getUserByToken_(token)) return { ok: false, error: 'Sesi Anda sudah berakhir. Silakan login kembali.' };
    if (!PD_LEMBAR[jenis]) return { ok: false, error: 'Jenis dokumen tidak dikenal.' };
    if (!fileId) return { ok: false, error: 'File paket belum dibuat.' };

    const ss = SpreadsheetApp.openById(fileId);
    const noAman = pdBacaPenandaPaket_(ss);
    if (noAman === null) return { ok: false, error: 'File ini bukan paket dokumen SIPERDIN.' };

    const sh = ss.getSheetByName(PD_LEMBAR[jenis].nama);
    if (!sh) return { ok: false, error: 'Sheet ' + PD_LEMBAR[jenis].nama + ' tidak ada di file paket.' };

    SpreadsheetApp.flush();
    const portrait = PD_LEMBAR[jenis].orientasi === 'portrait';
    const url = 'https://docs.google.com/spreadsheets/d/' + fileId + '/export?' + [
      'format=pdf',
      'gid=' + sh.getSheetId(),
      'size=A4',
      'portrait=' + portrait,
      'fitw=true',
      'gridlines=false',
      'printtitle=false',
      'sheetnames=false',
      'pagenumbers=false',
      'fzr=false',
      'horizontal_alignment=CENTER',
      'top_margin=0.6', 'bottom_margin=0.6', 'left_margin=0.5', 'right_margin=0.5'
    ].join('&');

    const opsi = { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true };
    let resp = UrlFetchApp.fetch(url, opsi);
    if (resp.getResponseCode() === 429 || resp.getResponseCode() >= 500) {
      Utilities.sleep(1500);
      resp = UrlFetchApp.fetch(url, opsi);
    }
    if (resp.getResponseCode() !== 200) {
      console.error('Ekspor PDF gagal, HTTP ' + resp.getResponseCode() + ': ' + resp.getContentText().substring(0, 300));
      return { ok: false, error: 'PDF gagal dibuat (HTTP ' + resp.getResponseCode() + '). Coba lagi sebentar lagi.' };
    }

    return {
      ok: true,
      base64: Utilities.base64Encode(resp.getContent()),
      namaFile: PD_LEMBAR[jenis].nama + ' - ' + (noAman || 'SPTJB') + '.pdf'
    };
  } catch (err) {
    console.error('unduhPdfPaketDokumen gagal: ' + err.message + '\n' + (err.stack || ''));
    return { ok: false, error: 'PDF gagal dibuat: ' + err.message };
  }
}

/**
 * Jalankan SEKALI dari editor Apps Script supaya izin Google Drive
 * disetujui (web app berjalan sebagai pemilik script).
 */
function otorisasiPaketDokumen() {
  DriveApp.getRootFolder();
  ScriptApp.getOAuthToken();
  Logger.log('Otorisasi Drive OK. Sekarang buat "New version" di Deploy > Manage deployments.');
}

// ============================================================
// PENGUMPULAN DATA
// ============================================================

function pdBacaPenandaPaket_(ss) {
  const meta = ss.getDeveloperMetadata();
  for (let i = 0; i < meta.length; i++) {
    if (meta[i].getKey() === PD_KUNCI_METADATA) return meta[i].getValue();
  }
  return null;
}

function pdKumpulkanData_(idSptjb) {
  if (!idSptjb) throw new Error('SPTJB belum dipilih.');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shS = ss.getSheetByName('DB_SPTJB');
  const shP = ss.getSheetByName('DB_PENGAJUAN');
  if (!shS || !shP) throw new Error('Sheet DB_SPTJB atau DB_PENGAJUAN tidak ditemukan.');

  // ---- SPTJB (kolom: A=ID, B=No_SPTJB, D=Kegiatan, J=No_Akun, L=Nama_Akun, S=Klasifikasi) ----
  const dataS = shS.getDataRange().getValues();
  let sptjb = null;
  for (let r = 1; r < dataS.length; r++) {
    if (dataS[r][0] === idSptjb) {
      sptjb = {
        noSptjb: String(dataS[r][1] || ''),
        kegiatan: String(dataS[r][3] || ''),
        noAkun: String(dataS[r][9] || '').trim(),
        namaAkun: String(dataS[r][11] || ''),
        klasifikasi: String(dataS[r][18] || '').trim()
      };
      break;
    }
  }
  if (!sptjb) throw new Error('SPTJB tidak ditemukan.');

  // ---- Baris personil semua Surat Tugas di bawah SPTJB ini ----
  const dataP = shP.getDataRange().getValues();
  const baris = [];
  const noStList = [];
  let tglMulai = null, tglSelesai = null;
  for (let r = 1; r < dataP.length; r++) {
    const row = dataP[r];
    if (row[3] !== idSptjb) continue;
    const tglB = pdParseTanggal_(row[17]);
    const tglK = pdParseTanggal_(row[18]);
    let lama = Number(row[19]) || 0;
    if (!lama && tglB && tglK) lama = Math.max(Math.round((tglK - tglB) / 86400000) + 1, 0);
    const b = {
      nama: String(row[8] || ''),
      tujuan: String(row[14] || ''),
      provinsi: String(row[15] || ''),
      daerah: String(row[16] || ''),
      tglB: tglB, tglK: tglK, lama: lama,
      noSt: String(row[29] || '').trim(),
      tiket: pdN_(row[30]), kedudukan: pdN_(row[31]), tujuanTaksi: pdN_(row[32]),
      uangHarian: pdN_(row[33]), penginapan: pdN_(row[34]), representatif: pdN_(row[35]),
      lainnya: pdN_(row[37])
    };
    // total dihitung dari komponen supaya selalu cocok dengan kolom yang tampil
    b.total = b.tiket + b.kedudukan + b.tujuanTaksi + b.lainnya + b.uangHarian + b.penginapan + b.representatif;
    baris.push(b);
    if (b.noSt && noStList.indexOf(b.noSt) === -1) noStList.push(b.noSt);
    if (tglB && (!tglMulai || tglB < tglMulai)) tglMulai = tglB;
    if (tglK && (!tglSelesai || tglK > tglSelesai)) tglSelesai = tglK;
  }
  if (baris.length === 0) {
    throw new Error('SPTJB ' + sptjb.noSptjb + ' belum punya Surat Tugas/personil. Tambahkan dulu di tab Surat Tugas.');
  }

  const jumlah = { tiket: 0, kedudukan: 0, tujuanTaksi: 0, lainnya: 0, uangHarian: 0, penginapan: 0, representatif: 0, total: 0 };
  baris.forEach(function (b) {
    Object.keys(jumlah).forEach(function (k) { jumlah[k] += b[k]; });
  });

  // ---- CONFIG ----
  const cfg = {};
  const cfgLabel = {};
  getConfig_().forEach(function (f) {
    cfg[f.key] = String(f.value || '').trim();
    cfgLabel[f.key] = f.label;
  });
  const peringatan = [];
  const kosong = PD_CONFIG_WAJIB.filter(function (k) { return !cfg[k]; }).map(function (k) { return cfgLabel[k] || k; });
  if (kosong.length) {
    peringatan.push('Konfigurasi belum lengkap, bagian ini tampil kosong/placeholder di dokumen: ' + kosong.join(', ') + '. Lengkapi di menu Konfigurasi lalu buat paket ulang.');
  }

  const akun = pdCariInfoAkun_(sptjb.klasifikasi, sptjb.noAkun);

  return {
    idSptjb: idSptjb,
    noSptjb: sptjb.noSptjb,
    kegiatan: sptjb.kegiatan,
    klasifikasi: sptjb.klasifikasi || sptjb.noAkun,
    noAkun: sptjb.noAkun,
    namaAkun: sptjb.namaAkun || akun.namaAkun,
    akun: akun,
    cfg: cfg,
    baris: baris,
    jumlah: jumlah,
    uraian: pdBuatUraian_(sptjb.kegiatan, baris, tglMulai, tglSelesai),
    noStGabung: noStList.length ? noStList.join(', ') : sptjb.noSptjb,
    penerima: baris[0].nama + (baris.length > 1 ? ',dkk' : ''),
    noDipaTeks: (cfg.no_dipa || '-') + (cfg.tanggal_dipa ? '  Tanggal ' + cfg.tanggal_dipa : ''),
    peringatan: peringatan
  };
}

/**
 * Nama komponen/sub komponen/akun + kode item (Tiket, Transport, dst) dari REF_AKUN.
 * Klasifikasi dibaca dari belakang karena kode kegiatan sendiri berisi titik:
 * "2175.BDC.003.052.0A.524111" -> program 2175, kegiatan BDC.003, output 052, sub 0A, akun 524111.
 * Kalau tidak ketemu, hasilnya kosong (dokumen tetap jadi, bagian itu tampil "-").
 */
function pdCariInfoAkun_(klasifikasi, noAkun) {
  const info = { kodeKomponen: '', namaKomponen: '', kodeSub: '', namaSub: '', namaAkun: '',
    item: { tiket: '', transport: '', uangHarian: '', uangHotel: '', transportBandara: '', representatif: '' } };
  const parts = String(klasifikasi || '').split('.');
  if (parts.length < 6) return info;

  const n = parts.length;
  const akun = parts[n - 1];
  const suffix = parts[n - 2];
  const output = parts[n - 3];
  const kegiatan = parts.slice(1, n - 3).join('.');
  const subKode = output + '.' + suffix;
  info.kodeKomponen = parts[0] + '.' + kegiatan;
  info.kodeSub = subKode;

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('REF_AKUN');
  if (!sh) return info;
  const data = sh.getDataRange().getValues();
  for (let r = 1; r < data.length; r++) {
    const level = String(data[r][0]).trim();
    const kode = String(data[r][1]).trim();
    const nama = String(data[r][2]).trim();
    const kKeg = String(data[r][3]).trim();
    const kOut = String(data[r][4]).trim();
    const kSub = String(data[r][5]).trim();
    const kAkun = String(data[r][6]).trim();

    if (level === 'Kegiatan' && kode === kegiatan) {
      info.namaKomponen = nama;
    } else if (level === 'Output' && kode === output && kKeg === kegiatan) {
      info.namaSub = nama;
    } else if (level === 'Akun' && kode === akun && kKeg === kegiatan && kSub === subKode) {
      info.namaAkun = nama;
    } else if (level === 'Item' && kAkun === akun && kKeg === kegiatan && kSub === subKode) {
      const nm = nama.toLowerCase();
      if (nm.indexOf('tiket') !== -1) info.item.tiket = kode;
      else if (nm.indexOf('transport bandara') !== -1) info.item.transportBandara = kode;
      else if (nm === 'transport') info.item.transport = kode;
      else if (nm.indexOf('uang harian') !== -1) info.item.uangHarian = kode;
      else if (nm.indexOf('hotel') !== -1) info.item.uangHotel = kode;
      else if (nm.indexOf('representatif') !== -1) info.item.representatif = kode;
    }
  }
  return info;
}

// Baris rincian di SPTJB: hanya komponen yang nilainya > 0
function pdBarisRincianSptjb_(D) {
  const it = D.akun.item;
  const j = D.jumlah;
  const daftar = [
    { label: 'Tiket', kode: it.tiket, jumlah: j.tiket },
    { label: 'Transport', kode: it.transport, jumlah: j.lainnya },
    { label: 'Uang Harian', kode: it.uangHarian, jumlah: j.uangHarian },
    { label: 'Uang Hotel', kode: it.uangHotel, jumlah: j.penginapan },
    { label: 'Transport Bandara', kode: it.transportBandara, jumlah: j.kedudukan + j.tujuanTaksi },
    { label: 'Uang Representatif', kode: it.representatif, jumlah: j.representatif }
  ];
  return daftar.filter(function (x) { return x.jumlah > 0; }).map(function (x) {
    return { teks: x.label + (x.kode ? ' (' + x.kode + ')' : ''), jumlah: x.jumlah };
  });
}

// ============================================================
// PEMBANGUN SHEET
// ============================================================

function pdBangunKwitansi_(sh, D) {
  const L = pdLembar_(sh, [140, 14, 105, 105, 105, 105, 105], 10);
  const T = PD_TEKS;
  const cfg = D.cfg;

  pdIsi_(L, 1, 1, 1, 7, T.kementerianKop, { b: true, fs: 11, h: 'center' });
  pdIsi_(L, 2, 1, 2, 7, T.ditjenKop1, { b: true, fs: 11, h: 'center' });
  pdIsi_(L, 3, 1, 3, 7, T.ditjenKop2, { b: true, fs: 10, h: 'center' });
  pdIsiTeks_(L, 4, 1, 7, T.alamatKop, { fs: 8, h: 'center', bb: true });
  pdTinggi_(L, 5, 1, 10);
  pdIsi_(L, 6, 1, 6, 7, 'KWITANSI/BUKTI PEMBAYARAN', { b: true, fs: 12, h: 'center' });
  pdTinggi_(L, 6, 1, 28);
  pdIsi_(L, 7, 5, 7, 7, 'Tahun Anggaran : ' + new Date().getFullYear(), { h: 'right' });
  pdBarisIsian_(L, 8, 'Mak', D.klasifikasi, 3, 7, { b: true });
  pdTinggi_(L, 9, 1, 10);
  pdBarisIsian_(L, 10, 'Telah Terima dari', T.penerimaDari, 3, 7);
  pdBarisIsian_(L, 11, 'Jumlah Uang', 'Rp' + pdAngka_(D.jumlah.total), 3, 7, { b: true, fs: 12 });
  pdBarisIsian_(L, 12, 'Terbilang', pdTerbilang_(D.jumlah.total), 3, 7, { b: true, i: true });
  pdBarisIsian_(L, 13, 'Untuk Pembayaran', D.uraian, 3, 7);
  pdTinggi_(L, 14, 1, 14);

  const tgl = pdKotaTanggal_(cfg.kota_penandatanganan);
  pdTandaTangan_(L, 15, 60, [
    { c1: 1, c2: 3, baris: ['Setuju dibayar a.n.', 'Kuasa Pengguna Anggaran', T.jabatanPpk],
      nama: pdIsian_(cfg.nama_ppk, 'Nama PPK'), nip: cfg.nip_ppk },
    { c1: 4, c2: 5, baris: ['Lunas Tgl.' + pdTitikTanggal_(), 'Bendahara Pengeluaran'],
      nama: pdIsian_(cfg.nama_bendahara, 'Nama Bendahara'), nip: cfg.nip_bendahara },
    { c1: 6, c2: 7, baris: [tgl, 'Yang Menerima'],
      nama: pdIsian_(cfg.nama_ketua_tim_kerja, 'Nama Ketua Tim Kerja'), nip: cfg.nip_ketua_tim_kerja }
  ]);
  pdTerapkan_(L);
}

function pdBangunSpby_(sh, D) {
  const L = pdLembar_(sh, [215, 14, 110, 110, 110, 110], 10);
  const T = PD_TEKS;
  const cfg = D.cfg;

  pdIsi_(L, 1, 1, 1, 6, T.kementerianSpby, { b: true, fs: 11, h: 'center' });
  pdIsi_(L, 2, 1, 2, 6, T.direktoratSpby, { b: true, fs: 11, h: 'center', bb: true });
  pdTinggi_(L, 3, 1, 14);
  pdIsi_(L, 4, 1, 4, 6, 'SURAT PERINTAH BAYAR', { b: true, fs: 12, h: 'center' });
  pdTinggi_(L, 4, 1, 28);
  pdIsi_(L, 5, 1, 5, 3, 'Tanggal :' + pdTitikTanggal_(), { txt: true });
  pdIsi_(L, 5, 4, 5, 6, 'Nomor :', {});
  pdTinggi_(L, 6, 1, 10);
  pdIsiTeks_(L, 7, 1, 6, T.paragrafSpby);
  pdIsi_(L, 8, 1, 8, 6, 'Rp ' + pdAngka_(D.jumlah.total), { b: true, fs: 12 });
  pdTinggi_(L, 8, 1, 26);
  pdIsiTeks_(L, 9, 1, 6, pdTerbilang_(D.jumlah.total), { b: true, i: true });
  pdTinggi_(L, 10, 1, 10);
  pdBarisIsian_(L, 11, 'Kepada', D.penerima, 3, 6);
  pdBarisIsian_(L, 12, 'Untuk pembayaran', D.uraian, 3, 6);
  pdTinggi_(L, 13, 1, 10);
  pdIsi_(L, 14, 1, 14, 6, 'Atas dasar :', {});
  pdBarisIsian_(L, 15, '1. No. Surat Tugas', D.noStGabung, 3, 6);
  pdBarisIsian_(L, 16, '2. Nota/bukti penerimaan barang/jasa/ Bukti lainnya', '', 3, 6);
  pdTinggi_(L, 17, 1, 10);
  pdIsi_(L, 18, 1, 18, 6, 'Dibebankan pada :', {});
  pdBarisIsian_(L, 19, 'Kegiatan, Output, MAK', D.klasifikasi, 3, 6);
  pdBarisIsian_(L, 20, 'Kode Akun', D.noAkun, 3, 6, { txt: true });
  pdTinggi_(L, 21, 1, 14);

  pdTandaTangan_(L, 22, 60, [
    { c1: 1, c2: 3, baris: ['Setuju/lunas dibayar, Tgl' + pdTitikTanggal_(), 'Bendahara Pengeluaran', 'Direktorat Bina PKK'],
      nama: pdIsian_(cfg.nama_bendahara, 'Nama Bendahara'), nip: cfg.nip_bendahara },
    { c1: 4, c2: 6, baris: [pdKotaTanggal_(cfg.kota_penandatanganan), T.jabatanPpk, ''],
      nama: pdIsian_(cfg.nama_ppk, 'Nama PPK'), nip: cfg.nip_ppk }
  ]);
  pdTerapkan_(L);
}

function pdBangunSptjb_(sh, D) {
  const L = pdLembar_(sh, [35, 150, 120, 115, 150, 240, 110, 85, 110], 10);
  const T = PD_TEKS;
  const cfg = D.cfg;

  pdIsi_(L, 1, 1, 1, 9, 'SURAT PERNYATAAN TANGGUNG JAWAB BELANJA', { b: true, fs: 12, h: 'center' });
  pdTinggi_(L, 1, 1, 26);
  pdIsi_(L, 2, 1, 2, 9, 'Nomor  : ' + D.noSptjb, { h: 'center', b: true });
  pdTinggi_(L, 3, 1, 10);

  const info = [
    ['1', 'Kode Satuan Kerja', ': ' + (cfg.kode_satker || '-')],
    ['2', 'Nama Satuan Kerja', ': ' + (cfg.nama_satker || '-')],
    ['3', 'Tanggal/No. DIPA', ': ' + D.noDipaTeks],
    ['4', 'Klasifikasi Anggaran', ': ' + D.klasifikasi],
    ['5', 'Bidang', ': ' + T.bidang]
  ];
  info.forEach(function (x, i) {
    const r = 4 + i;
    pdIsi_(L, r, 1, r, 1, x[0], { h: 'center' });
    pdIsi_(L, r, 2, r, 3, x[1], {});
    pdIsi_(L, r, 4, r, 9, x[2], { txt: true });
  });
  pdTinggi_(L, 9, 1, 10);
  pdIsiTeks_(L, 10, 1, 9, T.paragrafSptjb1);
  pdTinggi_(L, 11, 1, 8);

  // ---- tabel ----
  const rHead = 12;
  const kepala = [[1, 1, 'No.'], [2, 2, 'Akun'], [3, 3, 'Penerima'], [4, 4, 'Npwp'], [5, 5, 'Alamat'],
    [6, 6, 'Uraian'], [9, 9, 'Jumlah']];
  kepala.forEach(function (h) {
    pdIsi_(L, rHead, h[0], rHead + 1, h[1], h[2], { b: true, h: 'center', bg: '#E8EAF0' });
  });
  pdIsi_(L, rHead, 7, rHead, 8, 'Bukti', { b: true, h: 'center', bg: '#E8EAF0' });
  pdIsi_(L, rHead + 1, 7, rHead + 1, 7, 'Nomor', { b: true, h: 'center', bg: '#E8EAF0' });
  pdIsi_(L, rHead + 1, 8, rHead + 1, 8, 'Tanggal', { b: true, h: 'center', bg: '#E8EAF0' });
  pdTinggi_(L, rHead, 2, 22);

  const rincian = pdBarisRincianSptjb_(D);
  const r1 = rHead + 2;
  const rAkhir = r1 + rincian.length;
  const vert = { h: 'left', v: 'top', w: true };
  pdIsi_(L, r1, 1, rAkhir, 1, 1, { h: 'center', v: 'top' });
  pdIsi_(L, r1, 2, rAkhir, 2, D.klasifikasi, vert);
  pdIsi_(L, r1, 3, rAkhir, 3, D.penerima, vert);
  pdIsi_(L, r1, 4, rAkhir, 4, cfg.npwp || '-', { h: 'left', v: 'top', w: true, txt: true });
  pdIsi_(L, r1, 5, rAkhir, 5, cfg.alamat_satker || '-', vert);
  pdIsi_(L, r1, 6, r1, 6, D.uraian, vert);
  pdTinggi_(L, r1, 1, pdEstimasiTinggi_(D.uraian, 240, 10));
  pdIsi_(L, r1, 7, rAkhir, 7, D.noStGabung, { h: 'left', v: 'top', w: true, txt: true });
  pdIsi_(L, r1, 8, rAkhir, 8, pdTitikTanggal_().replace(/^\s+/, ''), { h: 'left', v: 'top', w: true, txt: true });
  rincian.forEach(function (x, i) {
    pdIsi_(L, r1 + 1 + i, 6, r1 + 1 + i, 6, x.teks, {});
    pdIsi_(L, r1 + 1 + i, 9, r1 + 1 + i, 9, x.jumlah, { fmt: PD_FMT_RP, h: 'right' });
  });
  const rTotal = rAkhir + 1;
  pdIsi_(L, rTotal, 1, rTotal, 8, 'J u m l a h', { b: true, h: 'center' });
  pdIsi_(L, rTotal, 9, rTotal, 9, D.jumlah.total, { b: true, fmt: PD_FMT_RP, h: 'right' });
  pdIsi_(L, rHead, 1, rTotal, 9, undefined, { brd: true, m: false });

  // ---- penutup ----
  let r = rTotal + 2;
  pdIsiTeks_(L, r, 1, 9, T.paragrafSptjb2);
  r++;
  pdIsi_(L, r, 1, r, 9, T.paragrafSptjb3, {});
  r++;
  pdTinggi_(L, r, 1, 10);
  r++;
  pdTandaTangan_(L, r, 60, [
    { c1: 2, c2: 4, baris: ['', 'Yang Mengusulkan atas nama', T.jabatanKetuaTim],
      nama: pdIsian_(cfg.nama_ketua_tim_kerja, 'Nama Ketua Tim Kerja'), nip: cfg.nip_ketua_tim_kerja },
    { c1: 7, c2: 9, baris: [pdKotaTanggal_(cfg.kota_penandatanganan), 'Mengetahui', T.jabatanPpk],
      nama: pdIsian_(cfg.nama_ppk, 'Nama PPK'), nip: cfg.nip_ppk }
  ]);
  pdTerapkan_(L);
}

function pdBangunNominatif_(sh, D) {
  const L = pdLembar_(sh, [30, 160, 60, 150, 115, 55, 90, 90, 90, 80, 80, 45, 95, 80, 45, 95, 85, 100], 9);
  const T = PD_TEKS;
  const cfg = D.cfg;

  pdIsi_(L, 1, 1, 1, 3, 'Daftar Nominatif :', { b: true });
  pdIsiTeks_(L, 1, 4, 18, D.uraian, { b: true });
  pdIsi_(L, 2, 1, 2, 3, 'Dengan nomor :', { b: true });
  pdIsiTeks_(L, 2, 4, 18, D.noStGabung, { b: true, txt: true });
  pdTinggi_(L, 3, 1, 8);

  const info = [
    ['1', 'Tanggal /No.DIPA', ': ' + D.noDipaTeks, ''],
    ['2', 'Komponen', ': ' + (D.akun.namaKomponen || '-'), D.akun.kodeKomponen],
    ['3', 'SubKomponen', ': ' + (D.akun.namaSub || '-'), D.akun.kodeSub],
    ['4', 'Akun', ': ' + (D.namaAkun || '-'), D.noAkun]
  ];
  info.forEach(function (x, i) {
    const r = 4 + i;
    pdIsi_(L, r, 1, r, 1, x[0], { h: 'center' });
    pdIsi_(L, r, 2, r, 3, x[1], {});
    pdIsiTeks_(L, r, 4, 16, x[2], { txt: true });
    pdIsi_(L, r, 17, r, 18, x[3], { h: 'right', b: true, txt: true });
  });
  pdTinggi_(L, 8, 1, 10);

  // ---- kepala tabel (2 baris digabung) ----
  const kepala = [[1, 1, 'No'], [2, 2, 'N a m a'], [3, 3, 'Asal'], [4, 4, 'Tujuan'],
    [5, 5, 'Tgl Keberangkatan - Tgl Kembali'], [6, 6, 'Lama Perjalanan'], [7, 7, 'Tiket PP'],
    [8, 8, 'Taksi Kedudukan (Bandara/Stasiun/Terminal) PP'], [9, 9, 'Taksi Tujuan (Bandara/Stasiun/Terminal) PP'],
    [10, 10, 'Transport'], [11, 13, 'Uang Harian'], [14, 16, 'Uang Hotel'], [17, 17, 'Uang Representatif'], [18, 18, 'Total']];
  kepala.forEach(function (h) {
    pdIsi_(L, 9, h[0], 10, h[1], h[2], { b: true, h: 'center', w: true, bg: '#E8EAF0', fs: 9 });
  });
  pdTinggi_(L, 9, 2, 36);

  // ---- isi tabel ----
  const r1 = 11;
  const rows = D.baris.map(function (b, i) {
    const malam = Math.max(b.lama - 1, 0);
    // tarif "x hari" ditampilkan hanya kalau totalnya habis dibagi (angka belum diubah manual)
    const uhOk = b.lama > 0 && b.uangHarian > 0 && b.uangHarian % b.lama === 0;
    const hotelOk = malam > 0 && b.penginapan > 0 && b.penginapan % malam === 0;
    return [
      i + 1, b.nama, T.kotaAsal, b.tujuan || b.daerah || '-', pdRentangSingkat_(b.tglB, b.tglK),
      b.lama ? b.lama + ' Hari' : '',
      b.tiket, b.kedudukan, b.tujuanTaksi, b.lainnya,
      uhOk ? b.uangHarian / b.lama : '', uhOk ? 'x ' + b.lama + ' =' : '', b.uangHarian,
      hotelOk ? b.penginapan / malam : '', hotelOk ? 'x ' + malam + ' =' : '', b.penginapan,
      b.representatif, b.total
    ];
  });
  pdBlok_(L, r1, 1, rows);
  const rEnd = r1 + rows.length - 1;
  const kol = function (c1, c2, o) { pdIsi_(L, r1, c1, rEnd, c2, undefined, Object.assign({ m: false }, o)); };
  kol(1, 1, { h: 'center' });
  kol(2, 2, { h: 'left', w: true });
  kol(3, 3, { h: 'center' });
  kol(4, 4, { h: 'left', w: true });
  kol(5, 5, { h: 'center', w: true, txt: true });
  kol(6, 6, { h: 'center' });
  kol(7, 11, { h: 'right', fmt: PD_FMT_RP });
  kol(12, 12, { h: 'center' });
  kol(13, 14, { h: 'right', fmt: PD_FMT_RP });
  kol(15, 15, { h: 'center' });
  kol(16, 18, { h: 'right', fmt: PD_FMT_RP });
  rows.forEach(function (row, i) {
    const px = Math.max(pdEstimasiTinggi_(row[1], 160, 9), pdEstimasiTinggi_(row[3], 150, 9), pdEstimasiTinggi_(row[4], 115, 9), 22);
    pdTinggi_(L, r1 + i, 1, px);
  });

  // ---- baris jumlah ----
  const rT = rEnd + 1;
  const j = D.jumlah;
  const tot = { b: true, h: 'right', fmt: PD_FMT_RP, bg: '#F2F3F7' };
  pdIsi_(L, rT, 1, rT, 6, 'JUMLAH SELURUHNYA', { b: true, h: 'center', bg: '#F2F3F7' });
  pdIsi_(L, rT, 7, rT, 7, j.tiket, tot);
  pdIsi_(L, rT, 8, rT, 8, j.kedudukan, tot);
  pdIsi_(L, rT, 9, rT, 9, j.tujuanTaksi, tot);
  pdIsi_(L, rT, 10, rT, 10, j.lainnya, tot);
  pdIsi_(L, rT, 11, rT, 12, undefined, { bg: '#F2F3F7' });
  pdIsi_(L, rT, 13, rT, 13, j.uangHarian, tot);
  pdIsi_(L, rT, 14, rT, 15, undefined, { bg: '#F2F3F7' });
  pdIsi_(L, rT, 16, rT, 16, j.penginapan, tot);
  pdIsi_(L, rT, 17, rT, 17, j.representatif, tot);
  pdIsi_(L, rT, 18, rT, 18, j.total, tot);
  pdTinggi_(L, rT, 1, 24);
  pdIsi_(L, 9, 1, rT, 18, undefined, { brd: true, m: false });

  // ---- tanda tangan ----
  pdTinggi_(L, rT + 1, 1, 12);
  pdTandaTangan_(L, rT + 2, 60, [
    { c1: 2, c2: 4, baris: ['Setuju Dibayarkan,', T.jabatanPpk],
      nama: pdIsian_(cfg.nama_ppk, 'Nama PPK'), nip: cfg.nip_ppk },
    { c1: 6, c2: 10, baris: ['Yang Bertanggungjawab', T.jabatanKetuaTim],
      nama: pdIsian_(cfg.nama_ketua_tim_kerja, 'Nama Ketua Tim Kerja'), nip: cfg.nip_ketua_tim_kerja },
    { c1: 14, c2: 18, baris: [pdKotaTanggal_(cfg.kota_penandatanganan), 'Bendahara Pengeluaran'],
      nama: pdIsian_(cfg.nama_bendahara, 'Nama Bendahara'), nip: cfg.nip_bendahara }
  ]);
  pdTerapkan_(L);
}

// ============================================================
// PEMBANTU TATA LETAK
// Lembar dibangun dulu di memori (nilai + daftar format), lalu
// ditulis ke sheet sekaligus di pdTerapkan_ -- jauh lebih cepat
// daripada memanggil API spreadsheet cell demi cell.
// ============================================================

function pdLembar_(sh, lebar, fsDasar) {
  return { sh: sh, lebar: lebar, fs: fsDasar || 10, nilai: [], op: [], tinggi: [], maxRow: 0 };
}

/**
 * Catat nilai + format untuk blok sel (r1,c1)-(r2,c2).
 * Blok lebih dari 1 sel otomatis digabung (merge) kecuali opsi.m === false.
 * opsi: b tebal, i miring, u garis bawah, fs ukuran font, h/v perataan, w bungkus teks,
 *       fmt format angka, txt paksa format teks, bg warna latar, brd garis kotak, bb garis bawah tebal.
 */
function pdIsi_(L, r1, c1, r2, c2, nilai, opsi) {
  L.op.push({ r1: r1, c1: c1, r2: r2, c2: c2, o: opsi || {} });
  if (nilai !== undefined && nilai !== null && nilai !== '') L.nilai.push([r1, c1, nilai]);
  if (r2 > L.maxRow) L.maxRow = r2;
}

// Teks 1 baris yang bisa panjang: dibungkus, tinggi baris diperkirakan otomatis
function pdIsiTeks_(L, r, c1, c2, teks, opsi) {
  const o = Object.assign({ w: true }, opsi || {});
  pdIsi_(L, r, c1, r, c2, teks, o);
  let lebar = 0;
  for (let c = c1; c <= c2; c++) lebar += L.lebar[c - 1];
  pdTinggi_(L, r, 1, pdEstimasiTinggi_(teks, lebar, o.fs || L.fs));
}

// Baris "Label : isi" -- label di kolom A, titik dua di kolom B, isi digabung c1..c2
function pdBarisIsian_(L, r, label, isi, c1, c2, opsi) {
  const o = opsi || {};
  pdIsiTeks_(L, r, 1, 1, label, { fs: o.fs });
  pdIsi_(L, r, 2, r, 2, ':', { h: 'center', v: 'top' });
  pdIsiTeks_(L, r, c1, c2, isi, o);
  // label bisa lebih panjang dari 1 baris; ambil tinggi terbesar dari label & isi
  const tLabel = pdEstimasiTinggi_(label, L.lebar[0], o.fs || L.fs);
  const tIsi = pdEstimasiTinggi_(isi, L.lebar.slice(c1 - 1, c2).reduce(function (a, b) { return a + b; }, 0), o.fs || L.fs);
  pdTinggi_(L, r, 1, Math.max(tLabel, tIsi));
  L.op[L.op.length - 3].o.v = 'top';
  L.op[L.op.length - 1].o.v = 'top';
}

function pdBlok_(L, rAwal, cAwal, matriks) {
  matriks.forEach(function (row, i) {
    row.forEach(function (v, j) {
      if (v !== '' && v !== null && v !== undefined) L.nilai.push([rAwal + i, cAwal + j, v]);
    });
  });
  const rAkhir = rAwal + matriks.length - 1;
  if (rAkhir > L.maxRow) L.maxRow = rAkhir;
}

function pdTinggi_(L, rAwal, jumlahBaris, px) {
  L.tinggi.push({ r: rAwal, n: jumlahBaris, px: px });
  if (rAwal + jumlahBaris - 1 > L.maxRow) L.maxRow = rAwal + jumlahBaris - 1;
}

/**
 * Blok tanda tangan. Tiap blok: {c1, c2, baris:[teks...], nama, nip}.
 * Semua blok dijajarkan; ruang tanda tangan (tinggiRuang px) di bawah baris teks,
 * lalu nama (tebal, bergaris bawah) dan NIP.
 * @return {number} nomor baris terakhir yang dipakai
 */
function pdTandaTangan_(L, rAwal, tinggiRuang, bloks) {
  let maks = 0;
  bloks.forEach(function (b) { maks = Math.max(maks, b.baris.length); });
  for (let i = 0; i < maks; i++) {
    bloks.forEach(function (b) {
      pdIsi_(L, rAwal + i, b.c1, rAwal + i, b.c2, b.baris[i] || '', { h: 'center', txt: true });
    });
  }
  const rRuang = rAwal + maks;
  pdTinggi_(L, rRuang, 1, tinggiRuang);
  pdIsi_(L, rRuang, 1, rRuang, 1, undefined, { m: false }); // pastikan baris tercatat
  bloks.forEach(function (b) {
    pdIsi_(L, rRuang + 1, b.c1, rRuang + 1, b.c2, b.nama, { h: 'center', b: true, u: true });
    pdIsi_(L, rRuang + 2, b.c1, rRuang + 2, b.c2, 'NIP. ' + (b.nip || '-'), { h: 'center', txt: true });
  });
  return rRuang + 2;
}

function pdTerapkan_(L) {
  const sh = L.sh;
  const nk = L.lebar.length;
  const maxRow = L.maxRow;

  L.lebar.forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.setHiddenGridlines(true);
  sh.getRange(1, 1, maxRow, nk).setFontFamily('Arial').setFontSize(L.fs).setVerticalAlignment('middle');

  // 1) sel yang harus berformat teks (cegah "September 2026" dibaca sebagai tanggal)
  L.op.forEach(function (op) {
    if (op.o.txt) sh.getRange(op.r1, op.c1, op.r2 - op.r1 + 1, op.c2 - op.c1 + 1).setNumberFormat('@');
  });

  // 2) semua nilai sekaligus
  const arr = [];
  for (let r = 0; r < maxRow; r++) {
    const baris = [];
    for (let c = 0; c < nk; c++) baris.push('');
    arr.push(baris);
  }
  L.nilai.forEach(function (v) { arr[v[0] - 1][v[1] - 1] = v[2]; });
  sh.getRange(1, 1, maxRow, nk).setValues(arr);

  // 3) gabung sel & format
  L.op.forEach(function (op) {
    const o = op.o;
    const rng = sh.getRange(op.r1, op.c1, op.r2 - op.r1 + 1, op.c2 - op.c1 + 1);
    if (o.m !== false && (op.r2 > op.r1 || op.c2 > op.c1)) rng.merge();
    if (o.b) rng.setFontWeight('bold');
    if (o.i) rng.setFontStyle('italic');
    if (o.u) rng.setFontLine('underline');
    if (o.fs) rng.setFontSize(o.fs);
    if (o.h) rng.setHorizontalAlignment(o.h);
    if (o.v) rng.setVerticalAlignment(o.v);
    if (o.w) rng.setWrap(true);
    if (o.fmt) rng.setNumberFormat(o.fmt);
    if (o.bg) rng.setBackground(o.bg);
    if (o.brd) rng.setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
    if (o.bb) rng.setBorder(null, null, true, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  });

  // 4) tinggi baris
  L.tinggi.forEach(function (t) { sh.setRowHeights(t.r, t.n, t.px); });
}

// Perkiraan tinggi baris (px) untuk teks yang dibungkus di kolom selebar lebarPx
function pdEstimasiTinggi_(teks, lebarPx, fs) {
  const perKarakter = fs * 0.68;
  const kpb = Math.max(Math.floor((lebarPx - 10) / perKarakter), 8);
  let baris = 0;
  String(teks === undefined || teks === null ? '' : teks).split('\n').forEach(function (p) {
    baris += Math.max(1, Math.ceil(p.length / kpb));
  });
  return Math.max(21, Math.round(baris * fs * 1.55 + 8));
}

// ============================================================
// PEMBANTU TEKS / ANGKA / TANGGAL
// ============================================================

function pdN_(v) { return Number(v) || 0; }

function pdNamaAman_(s) {
  return String(s || '').replace(/[\\\/:*?"<>|]+/g, '-').replace(/\s+/g, ' ').trim();
}

// Nama/NIP kosong -> penanda jelas di dokumen (bukan dibiarkan kosong diam-diam)
function pdIsian_(v, label) {
  return v ? v : '(' + label + ' - isi di Konfigurasi)';
}

// 95812000 -> "95.812.000" (tanpa toLocaleString: di Apps Script hasilnya tidak selalu berlokal Indonesia)
function pdAngka_(n) {
  return String(Math.round(Number(n) || 0)).replace(/\B(?=(\d{3})+(?!\d))/g, '.');
}

function pdTerbilangRec_(n) {
  const s = ['', 'Satu', 'Dua', 'Tiga', 'Empat', 'Lima', 'Enam', 'Tujuh', 'Delapan', 'Sembilan', 'Sepuluh', 'Sebelas'];
  if (n < 12) return s[n];
  if (n < 20) return pdTerbilangRec_(n - 10) + ' Belas';
  if (n < 100) return pdTerbilangRec_(Math.floor(n / 10)) + ' Puluh' + (n % 10 ? ' ' + pdTerbilangRec_(n % 10) : '');
  if (n < 200) return 'Seratus' + (n - 100 ? ' ' + pdTerbilangRec_(n - 100) : '');
  if (n < 1000) return pdTerbilangRec_(Math.floor(n / 100)) + ' Ratus' + (n % 100 ? ' ' + pdTerbilangRec_(n % 100) : '');
  if (n < 2000) return 'Seribu' + (n - 1000 ? ' ' + pdTerbilangRec_(n - 1000) : '');
  if (n < 1000000) return pdTerbilangRec_(Math.floor(n / 1000)) + ' Ribu' + (n % 1000 ? ' ' + pdTerbilangRec_(n % 1000) : '');
  if (n < 1000000000) return pdTerbilangRec_(Math.floor(n / 1000000)) + ' Juta' + (n % 1000000 ? ' ' + pdTerbilangRec_(n % 1000000) : '');
  if (n < 1000000000000) return pdTerbilangRec_(Math.floor(n / 1000000000)) + ' Miliar' + (n % 1000000000 ? ' ' + pdTerbilangRec_(n % 1000000000) : '');
  return pdTerbilangRec_(Math.floor(n / 1000000000000)) + ' Triliun' + (n % 1000000000000 ? ' ' + pdTerbilangRec_(n % 1000000000000) : '');
}

// 95812000 -> "Sembilan Puluh Lima Juta Delapan Ratus Dua Belas Ribu Rupiah"
function pdTerbilang_(n) {
  const v = Math.floor(Math.abs(Number(n) || 0));
  return (v === 0 ? 'Nol' : pdTerbilangRec_(v)) + ' Rupiah';
}

function pdParseTanggal_(v) {
  if (!v) return null;
  if (Object.prototype.toString.call(v) === '[object Date]') return isNaN(v.getTime()) ? null : v;
  const m = String(v).match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  const d = new Date(v);
  return isNaN(d.getTime()) ? null : d;
}

function pdTanggalPanjang_(d) {
  return d.getDate() + ' ' + PD_BULAN[d.getMonth()] + ' ' + d.getFullYear();
}

function pdSamaHari_(a, b) {
  return a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
}

// Untuk uraian: "17 September - 25 September 2026"
function pdRentangTanggal_(d1, d2) {
  if (!d1 && !d2) return '';
  if (!d1 || !d2) return pdTanggalPanjang_(d1 || d2);
  if (pdSamaHari_(d1, d2)) return pdTanggalPanjang_(d1);
  if (d1.getFullYear() === d2.getFullYear()) {
    return d1.getDate() + ' ' + PD_BULAN[d1.getMonth()] + ' - ' + d2.getDate() + ' ' + PD_BULAN[d2.getMonth()] + ' ' + d1.getFullYear();
  }
  return pdTanggalPanjang_(d1) + ' - ' + pdTanggalPanjang_(d2);
}

// Untuk tabel: "22 - 25 September 2026"
function pdRentangSingkat_(d1, d2) {
  if (!d1 && !d2) return '';
  if (!d1 || !d2) return pdTanggalPanjang_(d1 || d2);
  if (pdSamaHari_(d1, d2)) return pdTanggalPanjang_(d1);
  if (d1.getFullYear() === d2.getFullYear() && d1.getMonth() === d2.getMonth()) {
    return d1.getDate() + ' - ' + d2.getDate() + ' ' + PD_BULAN[d1.getMonth()] + ' ' + d1.getFullYear();
  }
  return pdRentangTanggal_(d1, d2);
}

// "        September 2026" -- tanggalnya sengaja dikosongkan untuk ditulis tangan, seperti contoh
function pdTitikTanggal_() {
  const n = new Date();
  return '      ' + PD_BULAN[n.getMonth()] + ' ' + n.getFullYear();
}

function pdKotaTanggal_(kota) {
  return (kota || 'Jakarta') + ',' + pdTitikTanggal_();
}

// "R I A U" -> "Riau", "B A L I" -> "Bali", "D.K.I. Jakarta" -> "DKI Jakarta"
// (ejaan provinsi di REF_TARIF sengaja dibuat begitu supaya cocok antar blok tarif)
function pdRapikanProvinsi_(p) {
  let s = String(p || '').trim();
  if (/^(?:[A-Za-z]\s){2,}[A-Za-z]$/.test(s)) {
    s = s.replace(/\s+/g, '');
    s = s.charAt(0).toUpperCase() + s.slice(1).toLowerCase();
  }
  return s.replace(/^D\.K\.I\.\s*/i, 'DKI ');
}

/**
 * "Perjalanan Dinas dalam rangka <kegiatan> di <daerah>, Provinsi <X> Pada Tanggal <rentang>".
 * Kalau provinsinya lebih dari satu: "... di beberapa Kabupaten/Kota di Provinsi A, B dan C ..."
 * (persis pola di contoh dokumen).
 */
function pdBuatUraian_(kegiatan, baris, d1, d2) {
  const prov = [];
  const daerah = [];
  baris.forEach(function (b) {
    const p = pdRapikanProvinsi_(b.provinsi);
    if (p && prov.indexOf(p) === -1) prov.push(p);
    pecahDaftarDaerah_(b.daerah || b.tujuan).forEach(function (d) {
      if (daerah.indexOf(d) === -1) daerah.push(d);
    });
  });

  let teks = 'Perjalanan Dinas dalam rangka ' + kegiatan;
  if (prov.length > 1) {
    teks += ' di beberapa Kabupaten/Kota di Provinsi ' + gabungDaftarDaerah_(prov);
  } else if (prov.length === 1) {
    teks += (daerah.length ? ' di ' + gabungDaftarDaerah_(daerah) + ',' : '') + ' Provinsi ' + prov[0];
  } else if (daerah.length) {
    teks += ' di ' + gabungDaftarDaerah_(daerah);
  }
  const rentang = pdRentangTanggal_(d1, d2);
  if (rentang) teks += ' Pada Tanggal ' + rentang;
  return teks;
}
