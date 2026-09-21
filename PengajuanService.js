/**
 * PengajuanService.gs
 * ------------------------------------------------------------
 * Logic server untuk halaman Form Pengajuan.
 *
 * CATATAN FITUR "BAGI PER KELOMPOK DAERAH" (multi-destinasi personil
 * dalam 1 Surat Tugas):
 * - TIDAK ADA PERUBAHAN SKEMA. Kolom Tujuan/Provinsi/Daerah di
 *   DB_PENGAJUAN sudah per-baris (per-orang) sejak awal -- yang baru
 *   adalah tulisSatuTrip_() sekarang membolehkan tiap orang punya
 *   tujuan/provinsi/daerah SENDIRI (override), fallback ke nilai trip
 *   kalau tidak diisi (mode "Satu Daerah Bersama" -- 100% backward
 *   compatible dengan data lama).
 * - getPengajuanUntukEdit() mendeteksi otomatis apakah sebuah Surat
 *   Tugas punya >1 kombinasi Provinsi+Daerah berbeda di antara
 *   baris-barisnya -> mode 'kelompok', lalu direkonstruksi jadi
 *   kelompokList untuk diedit. Kalau cuma 1 kombinasi -> mode
 *   'standar' seperti sebelumnya.
 * - cetakDaftarNominatif (DokumenService.gs), getDashboardData_
 *   (Code.gs), getPengajuanListUntukRealisasi_ (RealisasiService.gs)
 *   SUDAH membaca Tujuan/Provinsi per BARIS -- jadi otomatis kompatibel
 *   dengan data kelompok tanpa perlu diubah (sudah dicek).
 *
 * CATATAN FITUR "BEBERAPA KAB/KOTA DALAM 1 PROVINSI":
 * - Satu Surat Tugas (mode standar) atau satu kelompok daerah boleh
 *   punya lebih dari 1 kab/kota tujuan (satu provinsi yang sama).
 * - Disimpan di SATU kolom (Daerah, kolom Q, dan Tujuan, kolom O bila
 *   Tujuan tidak diisi manual) dengan format: "A, B dan C"
 *   (dua item: "A dan B"). Lihat gabungDaftarDaerah_() /
 *   pecahDaftarDaerah_(). Data lama (1 daerah saja) tetap terbaca.
 * ------------------------------------------------------------
 */

/**
 * Daftar provinsi unik untuk dropdown Provinsi (section SPTJB & Trip),
 * diambil dari kolom "provinsi" di blok REF_TARIF (UH_PEGAWAI, fallback
 * ke HOTEL_PEJABAT kalau blok itu tidak ketemu).
 */
function getProvinsiList() {
  var blok = getBlokRefTarif_();
  var b = blok.UH_PEGAWAI || blok.HOTEL_PEJABAT;
  if (!b) return [];
  var colProv = findColIndex_(b.headers, ['provinsi']);
  if (colProv === -1) return [];
  var set = {};
  b.data.forEach(function (row) {
    var v = String(row[colProv] || '').trim();
    if (v) set[v] = true;
  });
  return Object.keys(set).sort();
}

/**
 * Daftar kab/kota per provinsi -- lihat getDaerahByProvinsi() di
 * WilayahService.gs (sumbernya API resmi wilayah Indonesia, bukan
 * REF_TARIF lagi, supaya cakupannya lengkap & bisa difilter per
 * provinsi yang dipilih).
 */


function getPegawaiList_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('REF_PEGAWAI');
  if (!sh) return [];
  const data = sh.getDataRange().getValues();
  const out = [];
  for (let r = 1; r < data.length; r++) {
    if (!data[r][0]) continue;
    out.push({ nama: data[r][0], nip: data[r][1], jabatan: data[r][2] });
  }
  return out;
}

// Hierarki Kegiatan > Output > SubOutput > Akun untuk dropdown bertingkat
// "Klasifikasi Anggaran" di Form Pengajuan.
function getRefAkunTree_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('REF_AKUN');
  if (!sh) return { kegiatan: [], output: [], suboutput: [], akun: [] };
  const data = sh.getDataRange().getValues();
  const tree = { kegiatan: [], output: [], suboutput: [], akun: [] };
  const seen = {};

  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const level = row[0], kode = row[1], nama = row[2];
    const kegiatanKode = row[3], outputKode = row[4], suboutputKode = row[5];
    if (!kode) continue;
    const dedupKey = level + '|' + kode + '|' + kegiatanKode + '|' + outputKode + '|' + suboutputKode;
    if (seen[dedupKey]) continue;
    seen[dedupKey] = true;

    if (level === 'Kegiatan') {
      tree.kegiatan.push({ kode: kode, nama: nama });
    } else if (level === 'Output') {
      tree.output.push({ kode: kode, nama: nama, kegiatanKode: kegiatanKode });
    } else if (level === 'SubOutput') {
      tree.suboutput.push({ kode: kode, nama: nama, kegiatanKode: kegiatanKode, outputKode: outputKode });
    } else if (level === 'Akun') {
      tree.akun.push({ kode: kode, nama: nama, kegiatanKode: kegiatanKode, outputKode: outputKode, suboutputKode: suboutputKode });
    }
  }
  return tree;
}


// Daftar SPTJB ringkas, untuk dropdown "pilih kegiatan yang sudah ada / buat baru"
function getSptjbListRingkas_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DB_SPTJB');
  if (!sh) return [];
  const data = sh.getDataRange().getValues();
  const out = [];
  for (let r = data.length - 1; r >= 1; r--) {
    if (!data[r][0]) continue;
    out.push({ idSptjb: data[r][0], noSptjb: data[r][1], kegiatan: data[r][3], noAkun: data[r][9] });
  }
  return out;
}

/**
 * Submit form SPTJB (BARU saja, TANPA Surat Tugas) -- dipanggil dari
 * tab "SPTJB" di Form Pengajuan. Hanya membuat 1 baris baru di
 * DB_SPTJB. Untuk menambahkan Surat Tugas ke SPTJB ini (atau ke SPTJB
 * lain yang sudah ada), gunakan submitSuratTugas().
 *
 * formData = {
 *   kegiatan, tglMulai, tglSelesai, noAkun, namaAkun,
 *   klasifikasiAnggaran, jumlahPengajuan, token
 * }
 * @return {Object} {ok:true, idSptjb, noSptjb, kegiatan} atau {ok:false, error}
 */
function submitSptjb(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  const pengirim = getUserByToken_(formData.token);
  const emailPengirim = pengirim ? pengirim.email : '(tidak diketahui)';

  try {
    if (!formData.kegiatan) return { ok: false, error: 'Nama Kegiatan wajib diisi.' };

    const shSptjb = ss.getSheetByName('DB_SPTJB');
    const idSptjb = 'SPTJB-2026-' + nextSequence_('SPTJB');
    const noSptjbAuto = generateNoSptjb_();
    const rowIdx = shSptjb.getLastRow() + 1;
    shSptjb.getRange(rowIdx, 1, 1, 19).setValues([[
      idSptjb, noSptjbAuto, '', formData.kegiatan || '', '',
      '', '', formData.tglMulai || '', formData.tglSelesai || '',
      formData.noAkun || '', '', formData.namaAkun || '', '', Number(formData.jumlahPengajuan) || 0,
      'Aktif', emailPengirim, new Date(), '',
      formData.klasifikasiAnggaran || ''
    ]]);

    return { ok: true, idSptjb: idSptjb, noSptjb: noSptjbAuto, kegiatan: formData.kegiatan };
  } catch (err) {
    return { ok: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Submit Surat Tugas (satu atau lebih sekaligus) untuk SPTJB yang
 * SUDAH ADA -- dipanggil dari tab "Surat Tugas". idSptjbExisting bisa
 * SPTJB MANAPUN (baru maupun lama).
 *
 * No Akun tiap Surat Tugas SELALU ikut No Akun SPTJB induknya.
 *
 * formData = {
 *   idSptjbExisting: '...',
 *   tripList: [
 *     { trip: { tujuan, provinsi, daerah, tglBerangkat, tglKembali, hotelVendor, noSpm, tahap, kegiatanTrip, noSt, noSk },
 *       orangList: [ { nama, nip, tipeId, namaKelompok,
 *                      tujuan, provinsi, daerah, // opsional -- override per-orang untuk mode "kelompok"
 *                      transportPP, transportKedudukan, transportTujuan, transportLainnya,
 *                      uangHarian, penginapan, uangRepresentatif, totalNominatif }, ... ] },
 *     ...
 *   ],
 *   token: '...'
 * }
 * CATATAN: "daerah" (di trip maupun di orang) boleh berupa STRING ("Kab. A")
 * atau ARRAY (["Kab. A", "Kota B"]) untuk beberapa kab/kota dalam 1 provinsi.
 * @return {Object} {ok:true, isEdit:false, status, jumlahTrip, jumlahOrang} atau {ok:false, error}
 */
function submitSuratTugas(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  const pengirim = getUserByToken_(formData.token);
  const emailPengirim = pengirim ? pengirim.email : '(tidak diketahui)';

  try {
    const shPengajuan = ss.getSheetByName('DB_PENGAJUAN');
    const idSptjb = formData.idSptjbExisting;
    if (!idSptjb) return { ok: false, error: 'Pilih SPTJB terlebih dahulu.' };
    if (!formData.tripList || formData.tripList.length === 0) {
      return { ok: false, error: 'Isi minimal 1 Surat Tugas dengan minimal 1 orang.' };
    }

    const noAkunTrip = lookupNoAkunSptjb_(idSptjb) || '';

    let jumlahOrangTotal = 0;
    let statusTerakhir = '';
    formData.tripList.forEach(function (item) {
      const idTrip = 'TRIP-' + nextSequence_('TRIP');
      const idPengajuan = 'PENG-2026-' + nextSequence_('PENG');
      const hasil = tulisSatuTrip_(shPengajuan, item, idSptjb, noAkunTrip, emailPengirim, idTrip, idPengajuan);
      jumlahOrangTotal += hasil.jumlahOrang;
      statusTerakhir = hasil.status;
    });

    return {
      ok: true, isEdit: false, status: statusTerakhir,
      jumlahTrip: formData.tripList.length, jumlahOrang: jumlahOrangTotal
    };
  } catch (err) {
    return { ok: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Submit EDIT untuk Surat Tugas yang SUDAH ADA (selalu 1 Surat Tugas
 * per panggilan) -- dipanggil dari tab "Cari & Edit".
 *
 * formData = {
 *   idPengajuanEdit: '...',
 *   tripList: [ { trip: {...}, orangList: [...] } ], // hanya elemen pertama dipakai
 *   token: '...'
 * }
 *
 * CATATAN KETERBATASAN: hapus+tulis ulang ID_Baris BARU untuk semua
 * orang (bukan update di tempat) -- kalau pengajuan yang diedit SUDAH
 * punya data di DB_REALISASI yang tertaut ke ID_Baris lama, tautannya
 * akan putus. Aman untuk edit pengajuan yang belum direalisasikan.
 */
function submitPengajuan(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  const pengirim = getUserByToken_(formData.token);
  const emailPengirim = pengirim ? pengirim.email : '(tidak diketahui)';

  try {
    const shPengajuan = ss.getSheetByName('DB_PENGAJUAN');

    if (!formData.idPengajuanEdit) {
      return { ok: false, error: 'submitPengajuan() sekarang hanya untuk mode edit Surat Tugas -- gunakan submitSptjb() atau submitSuratTugas() untuk data baru.' };
    }
    if (!formData.tripList || !formData.tripList[0]) {
      return { ok: false, error: 'Data Surat Tugas tidak lengkap.' };
    }

    const existing = getPengajuanUntukEdit(formData.idPengajuanEdit);
    if (!existing) return { ok: false, error: 'Data pengajuan yang mau diedit tidak ditemukan.' };
    const idSptjb = existing.idSptjb;
    hapusBarisPengajuan_(existing.idPengajuan);

    const noAkunTrip = lookupNoAkunSptjb_(idSptjb);
    const hasil = tulisSatuTrip_(shPengajuan, formData.tripList[0], idSptjb, noAkunTrip, emailPengirim,
      existing.idTrip, existing.idPengajuan);

    return {
      ok: true, isEdit: true, noSptjb: '', status: hasil.status,
      jumlahTrip: 1, jumlahOrang: hasil.jumlahOrang
    };
  } catch (err) {
    return { ok: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Terima array ATAU string ("A, B dan C") -> array nama bersih tanpa
 * duplikat. Dipakai untuk fitur beberapa kab/kota dalam 1 provinsi.
 */
function pecahDaftarDaerah_(v) {
  const arr = Array.isArray(v) ? v : String(v || '').split(/\s*,\s*|\s+dan\s+/i);
  return arr
    .map(function (s) { return String(s).trim(); })
    .filter(function (s, i, a) { return s && a.indexOf(s) === i; });
}

/**
 * ["A","B","C"] -> "A, B dan C"; ["A","B"] -> "A dan B"; ["A"] -> "A".
 * Format ini sama dengan kolom Tujuan di Daftar Nominatif.
 */
function gabungDaftarDaerah_(v) {
  const a = pecahDaftarDaerah_(v);
  if (a.length <= 1) return a[0] || '';
  return a.slice(0, -1).join(', ') + ' dan ' + a[a.length - 1];
}

/**
 * Tulis baris DB_PENGAJUAN untuk SATU Surat Tugas (1 idTrip/idPengajuan)
 * beserta semua orangnya.
 *
 * Fitur kelompok daerah: setiap orang BOLEH punya tujuan/provinsi/daerah
 * SENDIRI (o.tujuan/o.provinsi/o.daerah) yang meng-override nilai
 * trip-level. Kalau tidak diisi (mode "Satu Daerah Bersama"), fallback
 * ke nilai trip -- 100% backward compatible.
 *
 * Fitur multi kab/kota: daerah (di orang maupun trip) boleh array;
 * disimpan sebagai satu string "A, B dan C". Kolom Tujuan: isian manual
 * menang; kalau kosong, otomatis sama dengan daftar daerah.
 */
function tulisSatuTrip_(shPengajuan, item, idSptjb, noAkunTrip, emailPengirim, idTrip, idPengajuan) {
  const t = item.trip;
  const status = t.noSpm ? 'Final' : 'Draft';

  const rows = item.orangList.map(function (o) {
    const idBaris = 'BR-' + nextSequence_('BARIS');
    const daerahPilihan = pecahDaftarDaerah_(o.daerah);
    const adaDaerahOrang = daerahPilihan.length > 0;
    const daerahOrang = gabungDaftarDaerah_(adaDaerahOrang ? daerahPilihan : t.daerah);
    const provinsiOrang = o.provinsi || t.provinsi || '';
    const tujuanOrang = o.tujuan || (adaDaerahOrang ? daerahOrang : (t.tujuan || daerahOrang));
    return [
      idBaris, idTrip, idPengajuan, idSptjb,
      noAkunTrip,
      t.noSpm || '', t.tahap || '', t.kegiatanTrip || '',
      o.nama || '', o.nip || '', o.tipeId || '', t.noSk || '', o.namaKelompok || '',
      '', tujuanOrang, provinsiOrang, daerahOrang,
      t.tglBerangkat || '', t.tglKembali || '', hitungLamaHari_(t.tglBerangkat, t.tglKembali),
      t.hotelVendor || '', item.orangList.length,
      status, '', '', '',
      emailPengirim, new Date(), '',
      t.noSt || '',
      Number(o.transportPP) || 0,
      Number(o.transportKedudukan) || 0,
      Number(o.transportTujuan) || 0,
      Number(o.uangHarian) || 0,
      Number(o.penginapan) || 0,
      Number(o.uangRepresentatif) || 0,
      Number(o.totalNominatif) || 0,
      Number(o.transportLainnya) || 0 // kolom baru, ditambah di UJUNG (lihat AddFieldsV4.gs)
    ];
  });

  const startRow = shPengajuan.getLastRow() + 1;
  shPengajuan.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);

  return { status: status, jumlahOrang: rows.length };
}

/**
 * Hapus semua baris DB_PENGAJUAN milik satu ID_Pengajuan (dipakai
 * saat mode EDIT, sebelum menulis ulang baris terbaru).
 */
function hapusBarisPengajuan_(idPengajuan) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DB_PENGAJUAN');
  const data = sh.getDataRange().getValues();
  for (let r = data.length - 1; r >= 1; r--) {
    if (data[r][2] === idPengajuan) sh.deleteRow(r + 1);
  }
}

/**
 * Dipanggil dari client untuk memuat ulang satu Pengajuan (Surat
 * Tugas) ke form, supaya bisa diedit.
 *
 * Fungsi ini MENDETEKSI OTOMATIS apakah Surat Tugas ini punya lebih
 * dari 1 kombinasi Provinsi+Daerah berbeda di antara baris-barisnya:
 * - Kalau HANYA 1 kombinasi -> mode: 'standar' (orangList diisi,
 *   kelompokList kosong).
 * - Kalau LEBIH dari 1 kombinasi -> mode: 'kelompok', orangList
 *   dikosongkan dan kelompokList diisi per kombinasi
 *   (provinsi, daerah, daerahList, tujuan, namaKelompok, orangList).
 *
 * daerahList = daftar kab/kota hasil pecahan string Daerah ("A, B dan C").
 * trip.tujuan dikembalikan KOSONG kalau isinya sama dengan Daerah
 * (artinya otomatis), supaya perubahan pilihan daerah saat edit ikut
 * memperbarui Tujuan dan tidak "terkunci" sebagai isian manual.
 */
function getPengajuanUntukEdit(idPengajuan) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DB_PENGAJUAN');
  if (!sh) return null;
  const data = sh.getDataRange().getValues();
  const rows = [];
  for (let r = 1; r < data.length; r++) {
    if (data[r][2] === idPengajuan) rows.push(data[r]);
  }
  if (rows.length === 0) return null;

  const first = rows[0];

  const comboOrder = [];
  const comboMap = {};
  rows.forEach(function (row) {
    const provinsi = row[15] || '';
    const daerah = row[16] || '';
    const key = provinsi + '||' + daerah;
    if (!comboMap[key]) {
      comboMap[key] = {
        provinsi: provinsi, daerah: daerah, daerahList: pecahDaftarDaerah_(daerah), tujuan: row[14] || '',
        namaKelompok: row[12] || '', orangList: []
      };
      comboOrder.push(key);
    }
    comboMap[key].orangList.push({
      nama: row[8], nip: row[9], tipeId: row[10], namaKelompok: row[12],
      transportPP: row[30], transportKedudukan: row[31], transportTujuan: row[32],
      uangHarian: row[33], penginapan: row[34], uangRepresentatif: row[35], totalNominatif: row[36],
      transportLainnya: row[37] || 0
    });
  });

  const isKelompok = comboOrder.length > 1;

  return {
    idPengajuan: idPengajuan,
    idTrip: first[1],
    idSptjb: first[3],
    mode: isKelompok ? 'kelompok' : 'standar',
    trip: {
      tujuan: first[14] === first[16] ? '' : first[14],
      provinsi: first[15], daerah: first[16], daerahList: pecahDaftarDaerah_(first[16]),
      tglBerangkat: formatTanggalInput_(first[17]), tglKembali: formatTanggalInput_(first[18]),
      hotelVendor: first[20], noSpm: first[5], noSt: first[29], noSk: first[11],
      tahap: first[6], kegiatanTrip: first[7]
    },
    orangList: isKelompok ? [] : comboMap[comboOrder[0]].orangList,
    kelompokList: isKelompok ? comboOrder.map(function (k) { return comboMap[k]; }) : []
  };
}

/**
 * Daftar ringkas semua Pengajuan (dikelompokkan per ID_Pengajuan)
 * untuk ditampilkan di form supaya bisa dipilih untuk di-Edit.
 * Dipanggil dari server (Code.gs doGet), bukan dari client.
 *
 * Kalau satu Surat Tugas ternyata punya lebih dari 1 tujuan unik
 * (mode kelompok), label tujuan yang ditampilkan jadi
 * "<tujuan pertama> + N daerah lain" supaya tidak menyesatkan.
 */
function getDaftarPengajuanUntukEdit_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DB_PENGAJUAN');
  if (!sh) return [];
  const data = sh.getDataRange().getValues();
  const map = {};
  const order = [];
  for (let r = 1; r < data.length; r++) {
    const idPengajuan = data[r][2];
    if (!idPengajuan) continue;
    if (!map[idPengajuan]) {
      map[idPengajuan] = {
        idPengajuan: idPengajuan,
        tujuan: data[r][14] || '-',
        tglBerangkat: formatTanggal_(data[r][17]),
        tglKembali: formatTanggal_(data[r][18]),
        jumlahOrang: 0,
        status: data[r][22] || '-',
        noSpm: data[r][5] || '',
        _tujuanSet: {}
      };
      order.push(idPengajuan);
    }
    map[idPengajuan].jumlahOrang++;
    if (data[r][14]) map[idPengajuan]._tujuanSet[data[r][14]] = true;
  }
  return order.map(function (id) {
    const item = map[id];
    const tujuanUnik = Object.keys(item._tujuanSet);
    delete item._tujuanSet;
    if (tujuanUnik.length > 1) {
      item.tujuan = tujuanUnik[0] + ' + ' + (tujuanUnik.length - 1) + ' daerah lain';
    }
    return item;
  }).reverse();
}

function formatTanggalInput_(v) {
  if (!v) return '';
  try {
    if (Object.prototype.toString.call(v) === '[object Date]') {
      return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
    return String(v);
  } catch (e) {
    return '';
  }
}

function lookupNoAkunSptjb_(idSptjb) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DB_SPTJB');
  const data = sh.getDataRange().getValues();
  for (let r = 1; r < data.length; r++) {
    if (data[r][0] === idSptjb) return data[r][9];
  }
  return '';
}

function hitungLamaHari_(tglMulai, tglSelesai) {
  try {
    const d1 = new Date(tglMulai);
    const d2 = new Date(tglSelesai);
    const diff = Math.round((d2 - d1) / (1000 * 60 * 60 * 24)) + 1;
    return diff > 0 ? diff : '';
  } catch (err) {
    return '';
  }
}

// Penomoran berurutan sederhana, disimpan di Properties supaya tidak
// perlu scan seluruh sheet tiap kali submit.
function toRoman_(num) {
  const map = [
    [1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'],
    [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'],
    [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']
  ];
  let result = '';
  map.forEach(function (pair) {
    while (num >= pair[0]) {
      result += pair[1];
      num -= pair[0];
    }
  });
  return result;
}

/**
 * Generate No SPTJB otomatis dengan format resmi:
 *   3/{nomor urut per tahun}/PK.03.03/{bulan romawi}/{tahun}
 */
function generateNoSptjb_() {
  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth() + 1;
  const props = PropertiesService.getScriptProperties();
  const propKey = 'SPTJB_COUNTER_' + year;
  const next = Number(props.getProperty(propKey) || '0') + 1;
  props.setProperty(propKey, String(next));
  return '3/' + next + '/PK.03.03/' + toRoman_(month) + '/' + year;
}

/**
 * Preview No SPTJB berikutnya TANPA mengunci counter -- aman dipanggil
 * berkali-kali (refresh halaman, ganti tab, dst).
 */
function previewNoSptjb() {
  const now = new Date();
  const year = now.getFullYear();
  const month = now.getMonth() + 1;
  const props = PropertiesService.getScriptProperties();
  const propKey = 'SPTJB_COUNTER_' + year;
  const next = Number(props.getProperty(propKey) || '0') + 1;
  return '3/' + next + '/PK.03.03/' + toRoman_(month) + '/' + year;
}

function nextSequence_(key) {
  const props = PropertiesService.getScriptProperties();
  const propKey = 'SEQ_' + key;
  let n = Number(props.getProperty(propKey) || '0') + 1;
  props.setProperty(propKey, String(n));
  return String(n).padStart(5, '0');
}

/**
 * PENTING: jalankan fungsi ini SEKALI SAJA sebelum pertama kali pakai
 * form Pengajuan/Realisasi.
 */
function initSequencesV1() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();

  function maxSuffix(sheetName, colIndex, prefix) {
    const sh = ss.getSheetByName(sheetName);
    if (!sh) return 0;
    const data = sh.getDataRange().getValues();
    let max = 0;
    for (let r = 1; r < data.length; r++) {
      const v = String(data[r][colIndex] || '');
      if (v.indexOf(prefix) === 0) {
        const n = parseInt(v.slice(prefix.length), 10);
        if (!isNaN(n) && n > max) max = n;
      }
    }
    return max;
  }

  props.setProperty('SEQ_SPTJB', String(maxSuffix('DB_SPTJB', 0, 'SPTJB-2026-')));
  props.setProperty('SEQ_PENG', String(maxSuffix('DB_PENGAJUAN', 2, 'PENG-2026-')));
  props.setProperty('SEQ_TRIP', '0');
  props.setProperty('SEQ_BARIS', '0');
  props.setProperty('SEQ_REAL', String(maxSuffix('DB_REALISASI', 0, 'REAL-')));

  seedNoSptjbCounters_(ss, props);

  Logger.log('Sequence diinisialisasi: ' + JSON.stringify(props.getProperties()));
}

function seedNoSptjbCounters_(ss, props) {
  const sh = ss.getSheetByName('DB_SPTJB');
  if (!sh) return;
  const data = sh.getDataRange().getValues();
  const maxByYear = {};
  const pattern = /^3\/(\d+)\/PK\.03\.03\/[IVXLCDM]+\/(\d{4})$/;
  for (let r = 1; r < data.length; r++) {
    const v = String(data[r][1] || '');
    const m = v.match(pattern);
    if (m) {
      const n = parseInt(m[1], 10);
      const year = m[2];
      if (!maxByYear[year] || n > maxByYear[year]) maxByYear[year] = n;
    }
  }
  Object.keys(maxByYear).forEach(function (year) {
    props.setProperty('SPTJB_COUNTER_' + year, String(maxByYear[year]));
  });
}

/**
 * ============================================================================
 * MODUL SARAN NOMINATIF OTOMATIS — form Pengajuan SIPERDIN
 * ----------------------------------------------------------------------------
 * Menarik batas MAKSIMAL tarif dari sheet REF_TARIF (sesuai PMK 32/2025)
 * untuk tiap komponen biaya. Hasilnya hanya SARAN -- field di form tetap
 * bisa diubah manual.
 *
 * PERBAIKAN TERBARU:
 * - Tarif hotel: kolom dipilih dari nama header berdasarkan Golongan_Hotel
 *   (1-4) di REF_PEGAWAI (dulu mencari angka di nama header, tidak pernah cocok).
 * - Pencocokan provinsi memakai kunciProvinsi_() supaya tahan terhadap
 *   ejaan berbeda antar blok ("R I A U", "A C E H", Sumatra/Sumatera).
 * - Beberapa kab/kota dalam 1 provinsi: tiket & transport pakai tarif
 *   TERTINGGI dari semua kab/kota terpilih (SBM = batas atas).
 * - Hasil sekarang menyertakan "catatan" bila ada komponen yang tidak
 *   ditemukan tarifnya, supaya tidak gagal diam-diam.
 * - Transport Kedudukan & Transport Tujuan (taksi/transport bandara) dikali 2
 *   (pergi-pulang), lihat KALI_TRANSPORT_BANDARA.
 * - Tiket pesawat: pencocokan kota tidak lagi bergantung awalan Kab./Kota, dan
 *   bila tidak ada rute langsung ke kab/kota dipakai rute ke kota bandara provinsi
 *   (KOTA_BANDARA_PROVINSI) dengan catatan yang jelas.
 * ============================================================================
 */

var REF_TARIF_SHEET_NAME = 'REF_TARIF';

var TITLE_KEYWORDS = {
  UH_PEGAWAI: 'uh pegawai',
  HOTEL_PEJABAT: 'hotel pejabat',
  FULLBOARD_PESERTA: 'fullboard peserta',
  SEWA_MOBIL: 'sewa mobil',
  TRANSPORT_PESERTA: 'transport peserta',
  TAKSI: 'taksi',
  TIKET_PESAWAT: 'tiket pesawat'
};

// Golongan_Hotel di REF_PEGAWAI -> nama kolom di blok HOTEL PEJABAT
// (sudah huruf kecil, sesuai hasil closeBlock_)
var KOLOM_HOTEL_PER_GOLONGAN = {
  '1': 'tarif_pejabat_negara_wamen_eselon_i',
  '2': 'tarif_eselon_ii',
  '3': 'tarif_eselon_iii_gol_iv',
  '4': 'tarif_eselon_iv_gol_iii_ii_i'
};

// Transport Kedudukan & Transport Tujuan dihitung PERGI-PULANG: tarif SBM taksi/transport
// bandara di REF_TARIF adalah per sekali jalan, jadi dikali 2.
var KALI_TRANSPORT_BANDARA = 2;

// Kota bandara utama per provinsi (kunci = kunciProvinsi_). Dipakai HANYA sebagai cadangan
// saat tidak ada rute Jakarta-<kab/kota tujuan> di blok TIKET PESAWAT (mis. kabupaten tanpa
// bandara). Provinsi tanpa rute pesawat dari Jakarta (Jawa Barat, Banten, DKI) sengaja tidak ada.
var KOTA_BANDARA_PROVINSI = {
  'aceh': 'Banda Aceh', 'sumatrautara': 'Medan', 'riau': 'Pekanbaru', 'kepulauanriau': 'Batam',
  'jambi': 'Jambi', 'sumatrabarat': 'Padang', 'sumatraselatan': 'Palembang', 'lampung': 'Bandar Lampung',
  'bengkulu': 'Bengkulu', 'bangkabelitung': 'Pangkal Pinang',
  'jawatengah': 'Semarang', 'diyogyakarta': 'Yogyakarta', 'jawatimur': 'Surabaya',
  'bali': 'Denpasar', 'nusatenggarabarat': 'Mataram', 'nusatenggaratimur': 'Kupang',
  'kalimantanbarat': 'Pontianak', 'kalimantantengah': 'Palangkaraya', 'kalimantanselatan': 'Banjarmasin',
  'kalimantantimur': 'Balikpapan', 'kalimantanutara': 'Tanjung Selor',
  'sulawesiutara': 'Manado', 'gorontalo': 'Gorontalo', 'sulawesibarat': 'Mamuju',
  'sulawesiselatan': 'Makassar', 'sulawesitengah': 'Palu', 'sulawesitenggara': 'Kendari',
  'maluku': 'Ambon', 'malukuutara': 'Ternate', 'papua': 'Jayapura', 'papuabarat': 'Manokwari',
  'papuatengah': 'Timika'
};

// Beda penulisan nama kota antara data wilayah dan blok TIKET PESAWAT (kunci & nilai = kunciKota_)
var KOTA_ALIAS = { 'surakarta': 'solo' };

/**
 * Fungsi utama — dipanggil dari form lewat google.script.run.
 * @param {Object} p {namaPegawai, provinsiTujuan, kabKotaList (array) ATAU kabKotaTujuan (string), lamaHari}
 * @return {Object} breakdown nominatif per komponen + total + catatan
 */
function hitungNominatifSaran(p) {
  var pegawai = getPegawaiByNama_(p.namaPegawai);
  var blok = getBlokRefTarif_();

  var daftarKabKota = pecahDaftarDaerah_(p.kabKotaList && p.kabKotaList.length ? p.kabKotaList : p.kabKotaTujuan);

  // Tiket: untuk tiap kab/kota cari rute Jakarta-tujuan; kalau tidak ada, pakai kota bandara
  // provinsi itu. Bila banyak kab/kota, ambil tarif TERTINGGI (SBM = batas atas).
  var catatanTiket = [];
  var transportPP = tarifTertinggi_(daftarKabKota, function (kk) {
    var hasil = cariTiketDenganFallback_(blok.TIKET_PESAWAT, kk, p.provinsiTujuan, pegawai.golongan);
    if (hasil.via) catatanTiket.push('Tiket ke ' + kk + ' memakai rute Jakarta-' + hasil.via + ' (kota bandara provinsi; tidak ada rute langsung).');
    return hasil.tarif;
  });

  // TAKSI di REF_TARIF diindeks per PROVINSI (bukan per kota) -- jadi
  // kedudukan (asal) pakai provinsi DKI Jakarta, tujuan pakai provinsi
  // tujuan trip, BUKAN nama kab/kota-nya. Keduanya dikali 2 (pergi-pulang).
  var transportKedudukan = cariTaksi_(blok.TAKSI, 'D.K.I. Jakarta') * KALI_TRANSPORT_BANDARA;
  var taksiTujuan = cariTaksi_(blok.TAKSI, p.provinsiTujuan);
  var transportTujuan = taksiTujuan
    ? taksiTujuan * KALI_TRANSPORT_BANDARA
    : tarifTertinggi_(daftarKabKota, function (kk) {
        return cariTransportPeserta_(blok.TRANSPORT_PESERTA, kk); // tarif per kab/kota, tidak dikali 2
      });
  var uhHarian = cariUhPegawai_(blok.UH_PEGAWAI, p.provinsiTujuan, 'luar_kota');
  var uangHarian = uhHarian * p.lamaHari;
  var tarifHotel = cariHotelPejabat_(blok.HOTEL_PEJABAT, p.provinsiTujuan, pegawai.golonganHotel);
  var penginapan = tarifHotel * Math.max(p.lamaHari - 1, 0);
  var uangRepresentatif = 0; // TODO: belum ada blok referensi, isi manual dulu

  var total = transportPP + transportKedudukan + transportTujuan
            + uangHarian + penginapan + uangRepresentatif;

  var catatan = [];
  if (!uhHarian) catatan.push('Uang harian provinsi ini tidak ditemukan di REF_TARIF.');
  if (!tarifHotel && p.lamaHari > 1) catatan.push('Tarif hotel provinsi ini tidak ditemukan di REF_TARIF.');
  if (!transportPP) catatan.push('Tiket pesawat Jakarta-tujuan tidak ada di REF_TARIF (Rp0; isi manual bila naik pesawat).');
  else catatan = catatan.concat(catatanTiket);
  if (!transportTujuan) catatan.push('Tarif transport tujuan tidak ditemukan di REF_TARIF.');

  return {
    transportPP: transportPP,
    transportKedudukan: transportKedudukan,
    transportTujuan: transportTujuan,
    uangHarian: uangHarian,
    penginapan: penginapan,
    uangRepresentatif: uangRepresentatif,
    total: total,
    catatan: catatan.join(' ')
  };
}

// Nilai tertinggi dari fn(item) untuk semua item di daftar (0 kalau kosong)
function tarifTertinggi_(daftar, fn) {
  return daftar.reduce(function (maks, item) {
    return Math.max(maks, Number(fn(item)) || 0);
  }, 0);
}

// Kunci pembanding provinsi: tahan terhadap "R I A U", "A C E H",
// "D.K.I. Jakarta", dan beda ejaan Sumatra/Sumatera antar blok REF_TARIF.
function kunciProvinsi_(s) {
  return String(s || '').toLowerCase().replace(/[^a-z]/g, '').replace('sumatera', 'sumatra');
}

// Kunci pembanding nama kota: tanpa awalan Kab./Kota/Kabupaten, tanpa spasi & tanda baca
// ("Kota Medan" -> "medan", "Kota Surakarta" -> "solo")
function kunciKota_(s) {
  var k = String(s || '').toLowerCase()
    .replace(/^\s*(kabupaten|kab\.?|kota)\s+/, '')
    .replace(/[^a-z]/g, '');
  return KOTA_ALIAS[k] || k;
}

// Tiket pesawat dengan cadangan: rute langsung ke kab/kota; kalau tidak ada, rute ke kota
// bandara provinsi. via = nama kota bandara bila cadangan dipakai, '' bila rute langsung.
function cariTiketDenganFallback_(blokData, kabKota, provinsi, golongan) {
  var langsung = cariTiketPesawat_(blokData, 'Jakarta', kabKota, golongan);
  if (langsung) return { tarif: langsung, via: '' };
  var kotaBandara = KOTA_BANDARA_PROVINSI[kunciProvinsi_(provinsi)];
  if (kotaBandara) {
    var viaBandara = cariTiketPesawat_(blokData, 'Jakarta', kotaBandara, golongan);
    if (viaBandara) return { tarif: viaBandara, via: kotaBandara };
  }
  return { tarif: 0, via: '' };
}

function getBlokRefTarif_(skipCache) {
  var cache = CacheService.getScriptCache();
  if (!skipCache) {
    var cached = cache.get('BLOK_REF_TARIF');
    if (cached) return JSON.parse(cached);
  }

  var sheet = SpreadsheetApp.getActive().getSheetByName(REF_TARIF_SHEET_NAME);
  if (!sheet) throw new Error('Sheet "' + REF_TARIF_SHEET_NAME + '" tidak ditemukan di spreadsheet ini.');
  var lastCol = sheet.getLastColumn();
  var titleRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var headerRow = sheet.getRange(2, 1, 1, lastCol).getValues()[0];

  var blocks = {};
  var currentKey = null;
  var currentStart = null;

  for (var c = 0; c < lastCol; c++) {
    // pakai normalisasi_ (bukan toLowerCase/trim biasa) supaya spasi ganda
    // atau spasi tak biasa di judul sheet tidak bikin deteksi blok gagal
    var titleCell = normalisasi_(titleRow[c]);
    if (titleCell) {
      if (currentKey) {
        blocks[currentKey] = closeBlock_(sheet, currentStart, c, headerRow);
      }
      currentKey = matchTitleKeyword_(titleCell);
      currentStart = c;
    }
  }
  if (currentKey) {
    blocks[currentKey] = closeBlock_(sheet, currentStart, lastCol, headerRow);
  }

  // CacheService membatasi 100 KB per key -- kalau kelewat, jangan sampai
  // seluruh saran SBM ikut gagal; cukup lewati cache.
  try {
    cache.put('BLOK_REF_TARIF', JSON.stringify(blocks), 300); // cache 5 menit
  } catch (e) {
    console.warn('Cache BLOK_REF_TARIF gagal disimpan: ' + e.message);
  }
  return blocks;
}

function closeBlock_(sheet, startCol, endCol, headerRow) {
  var width = endCol - startCol;
  var headers = headerRow.slice(startCol, endCol).map(function (h) {
    return String(h || '').toLowerCase().trim();
  });
  var lastRow = sheet.getLastRow();
  // baris yang seluruh selnya kosong dibuang: memperkecil ukuran cache
  var data = sheet.getRange(3, startCol + 1, lastRow - 2, width).getValues()
    .filter(function (row) {
      return row.some(function (c) { return c !== ''; });
    });
  return { startCol: startCol, headers: headers, data: data };
}

function matchTitleKeyword_(titleLower) {
  for (var key in TITLE_KEYWORDS) {
    if (titleLower.indexOf(TITLE_KEYWORDS[key]) !== -1) return key;
  }
  return null;
}

function findColIndex_(headers, candidates) {
  for (var i = 0; i < headers.length; i++) {
    for (var j = 0; j < candidates.length; j++) {
      if (headers[i].indexOf(candidates[j]) !== -1) return i;
    }
  }
  return -1;
}

function normalisasi_(s) {
  return String(s || '').toLowerCase().replace(/\s+/g, ' ').trim();
}

// --- Pencari per blok ---------------------------------------------------

function cariUhPegawai_(blokData, provinsi, jenis) {
  if (!blokData) return 0;
  var colProv = findColIndex_(blokData.headers, ['provinsi']);
  var colLuar = findColIndex_(blokData.headers, ['luar kota', 'luar_kota']);
  var colDalam = findColIndex_(blokData.headers, ['dalam kota', 'dalam_kota']);
  var colTarget = jenis === 'dalam_kota' ? colDalam : colLuar;
  if (colProv === -1 || colTarget === -1) return 0;

  var target = kunciProvinsi_(provinsi);
  for (var i = 0; i < blokData.data.length; i++) {
    if (kunciProvinsi_(blokData.data[i][colProv]) === target) {
      return Number(blokData.data[i][colTarget]) || 0;
    }
  }
  return 0;
}

function cariHotelPejabat_(blokData, provinsi, golonganHotel) {
  if (!blokData) return 0;
  var colProv = findColIndex_(blokData.headers, ['provinsi']);
  // golongan hotel kosong/tidak dikenal -> pakai golongan 4 (tarif terendah)
  var namaKolom = KOLOM_HOTEL_PER_GOLONGAN[String(golonganHotel).trim()] || KOLOM_HOTEL_PER_GOLONGAN['4'];
  var colGol = blokData.headers.indexOf(namaKolom);
  if (colProv === -1 || colGol === -1) return 0;

  var target = kunciProvinsi_(provinsi);
  for (var i = 0; i < blokData.data.length; i++) {
    if (kunciProvinsi_(blokData.data[i][colProv]) === target) {
      return Number(blokData.data[i][colGol]) || 0;
    }
  }
  return 0;
}

function cariTaksi_(blokData, provinsi) {
  if (!blokData) return 0;
  var colProv = findColIndex_(blokData.headers, ['provinsi']);
  var colTarif = findColIndex_(blokData.headers, ['tarif', 'taksi', 'nilai']);
  if (colProv === -1 || colTarif === -1) return 0;

  var target = kunciProvinsi_(provinsi);
  for (var i = 0; i < blokData.data.length; i++) {
    if (kunciProvinsi_(blokData.data[i][colProv]) === target) {
      return Number(blokData.data[i][colTarif]) || 0;
    }
  }
  return 0;
}

function cariTransportPeserta_(blokData, kabKota) {
  if (!blokData) return 0;
  var colTujuan = findColIndex_(blokData.headers, ['kab', 'kota', 'tujuan']);
  var colTarif = findColIndex_(blokData.headers, ['tarif', 'nilai']);
  if (colTujuan === -1 || colTarif === -1) return 0;

  var target = normalisasi_(kabKota);
  for (var i = 0; i < blokData.data.length; i++) {
    var cell = normalisasi_(blokData.data[i][colTujuan]);
    if (cell.indexOf(target) !== -1 || target.indexOf(cell) !== -1) {
      return Number(blokData.data[i][colTarif]) || 0;
    }
  }
  return 0;
}

function cariTiketPesawat_(blokData, kotaAsal, kotaTujuan, golongan) {
  if (!blokData) return 0;
  var colAsal = findColIndex_(blokData.headers, ['asal']);
  var colTujuan = findColIndex_(blokData.headers, ['tujuan']);
  var kelas = golonganTermasukBisnis_(golongan) ? 'bisnis' : 'ekonomi';
  var colTarif = findColIndex_(blokData.headers, [kelas]);
  if (colAsal === -1 || colTujuan === -1 || colTarif === -1) return 0;

  var kAsal = kunciKota_(kotaAsal);
  var kTujuan = kunciKota_(kotaTujuan);
  if (!kAsal || !kTujuan) return 0;

  for (var i = 0; i < blokData.data.length; i++) {
    if (kunciKota_(blokData.data[i][colAsal]) !== kAsal) continue;
    if (kunciKota_(blokData.data[i][colTujuan]) === kTujuan) {
      return Number(blokData.data[i][colTarif]) || 0;
    }
  }
  return 0;
}

/**
 * Aturan kelas tiket bisnis biasanya untuk pejabat eselon I/II (golongan IV/d ke atas).
 * TODO: sesuaikan ambang golongan ini dengan aturan resmi di PMK 32/2025 jika beda.
 */
function golonganTermasukBisnis_(golongan) {
  if (!golongan) return false;
  var g = String(golongan).toUpperCase().replace(/\s+/g, '');
  return g.indexOf('IV/D') !== -1 || g.indexOf('IV/E') !== -1 || g.indexOf('IVD') !== -1 || g.indexOf('IVE') !== -1;
}

/**
 * Ambil data pegawai (golongan, golongan_hotel) dari REF_PEGAWAI berdasar nama.
 */
function getPegawaiByNama_(nama) {
  var sheet = SpreadsheetApp.getActive().getSheetByName('REF_PEGAWAI');
  var data = sheet.getDataRange().getValues();
  var headers = data[0].map(function (h) { return normalisasi_(h); });
  var colNama = findColIndex_(headers, ['nama']);
  var colGol = findColIndex_(headers, ['golongan']);
  var colGolHotel = findColIndex_(headers, ['golongan_hotel', 'golongan hotel']);

  var target = normalisasi_(nama);
  for (var i = 1; i < data.length; i++) {
    if (normalisasi_(data[i][colNama]) === target) {
      return {
        nama: data[i][colNama],
        golongan: colGol !== -1 ? data[i][colGol] : '',
        golonganHotel: colGolHotel !== -1 ? data[i][colGolHotel] : ''
      };
    }
  }
  return { nama: nama, golongan: '', golonganHotel: '' };
}

/**
 * DIAGNOSTIK — jalankan fungsi ini manual dari editor Apps Script
 * (pilih debugRefTarif di dropdown atas, klik Run), lalu buka
 * Executions / View > Logs untuk lihat hasilnya.
 * Sekaligus menyegarkan cache BLOK_REF_TARIF (skipCache = true).
 */
function debugRefTarif() {
  var sheet = SpreadsheetApp.getActive().getSheetByName(REF_TARIF_SHEET_NAME);
  if (!sheet) {
    Logger.log('Sheet "' + REF_TARIF_SHEET_NAME + '" TIDAK DITEMUKAN. Cek nama sheet persis.');
    return;
  }

  var lastCol = sheet.getLastColumn();
  var titleRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  Logger.log('=== Baris judul (row 1), mentah ===');
  Logger.log(JSON.stringify(titleRow));

  var blok = getBlokRefTarif_(true); // skip cache, baca ulang dari sheet
  var keys = Object.keys(blok);
  Logger.log('=== Blok yang TERDETEKSI: ' + (keys.length ? keys.join(', ') : '(tidak ada satupun)') + ' ===');

  Object.keys(TITLE_KEYWORDS).forEach(function (key) {
    if (blok[key]) {
      Logger.log(key + ' -> headers: ' + JSON.stringify(blok[key].headers) + ' | jumlah baris data: ' + blok[key].data.length);
    } else {
      Logger.log(key + ' TIDAK terdeteksi (cari kata kunci "' + TITLE_KEYWORDS[key] + '" di baris judul)');
    }
  });

  Logger.log('=== Tes getProvinsiList() ===');
  Logger.log(JSON.stringify(getProvinsiList()));
  Logger.log('=== Tes getDaerahByProvinsi("Jawa Barat") ===');
  Logger.log(JSON.stringify(getDaerahByProvinsi('Jawa Barat')));
}

/**
 * DIAGNOSTIK Saran SBM lengkap — jalankan manual dari editor, ubah
 * parameter di bagian atas fungsi sesuai kasus yang bermasalah, lalu cek Logs.
 */
function debugSaranSbm() {
  var nama = 'Nidhomul Haq S.Psi.';       // GANTI: nama persis di REF_PEGAWAI
  var provinsi = 'Sumatra Utara';          // GANTI: provinsi seperti di dropdown
  var kabKota = ['Kota Medan', 'Kab. Deli Serdang']; // GANTI: satu atau beberapa kab/kota
  var lamaHari = 3;

  var hasil = hitungNominatifSaran({ namaPegawai: nama, provinsiTujuan: provinsi, kabKotaList: kabKota, lamaHari: lamaHari });
  Logger.log(JSON.stringify(hasil, null, 2));
  Logger.log('Gabungan daerah yang akan tersimpan: ' + gabungDaftarDaerah_(kabKota));
}

/**
 * DIAGNOSTIK Transport Tiket -- jalankan manual dari editor, isi
 * parameter di bagian atas sesuai kasus yang bermasalah, lalu cek Logs.
 */
function debugTiketPesawat() {
  var kotaTujuan = 'Kota Medan'; // GANTI sesuai kab/kota yang dicoba di form
  var golongan = 'III/a'; // GANTI sesuai golongan pegawai yang dicoba

  var blok = getBlokRefTarif_(true);
  var b = blok.TIKET_PESAWAT;
  if (!b) { Logger.log('Blok TIKET_PESAWAT tidak terdeteksi sama sekali.'); return; }

  Logger.log('=== Header blok TIKET_PESAWAT ===');
  Logger.log(JSON.stringify(b.headers));
  Logger.log('Jumlah baris data: ' + b.data.length);
  Logger.log('=== 10 baris pertama (mentah) ===');
  Logger.log(JSON.stringify(b.data.slice(0, 10)));

  var kelas = golonganTermasukBisnis_(golongan) ? 'bisnis' : 'ekonomi';
  Logger.log('Golongan "' + golongan + '" -> kelas dipakai: ' + kelas);

  var hasil = cariTiketPesawat_(b, 'Jakarta', kotaTujuan, golongan);
  Logger.log('=== Hasil cariTiketPesawat_("Jakarta", "' + kotaTujuan + '", "' + golongan + '") ===');
  Logger.log('Tarif ditemukan: ' + hasil);

  var kw = normalisasi_(kotaTujuan).split(' ').pop();
  Logger.log('=== Baris yang mengandung kata "' + kw + '" di kolom manapun ===');
  b.data.forEach(function (row) {
    var gabung = row.join(' | ').toLowerCase();
    if (gabung.indexOf(kw) !== -1) Logger.log(JSON.stringify(row));
  });
}
