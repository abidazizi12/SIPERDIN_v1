/**
 * TemplateCetakService.gs
 * ------------------------------------------------------------
 * Mengisi 4 sheet cetak yang SUDAH ADA di spreadsheet database
 * (Kwitansi, SPBY, SPTJB, Nominatif -- format persis punya Abid)
 * berdasarkan SATU sel kontrol "No SPTJB". Ganti pilihan di
 * dropdown itu -> keempat sheet otomatis terisi ulang, seperti
 * kebiasaan "ubah 1 kode semua berubah" di Excel.
 *
 * KENAPA BUKAN RUMUS SPREADSHEET SEPERTI TEMPLATE ASLI:
 * Rumus-rumus asli di 4 sheet itu menunjuk ke sheet "Pengajuan",
 * "Form Rincian", "Realisasi All", "Ref" -- sheet-sheet itu ada di
 * file template lama Abid, TAPI TIDAK IKUT TERTEMPEL ke database
 * SIPERDIN, makanya jadi #REF!. Daripada menulis ulang puluhan
 * rumus array yang rapuh (gampang rusak kalau ada yang salah
 * geser baris), lebih aman fungsi Apps Script ini yang menghitung
 * lalu MENULIS NILAI JADI ke sel yang sama persis seperti template
 * asli -- tampilannya tetap sama, tapi sumber datanya DB_SPTJB /
 * DB_PENGAJUAN / REF_AKUN / CONFIG yang sudah pasti ada.
 *
 * KETERGANTUNGAN: file ini memakai ulang fungsi-fungsi dari
 * PaketDokumenService.gs (pdKumpulkanData_, pdTerbilang_,
 * pdKotaTanggal_, pdTitikTanggal_, pdRentangSingkat_,
 * pdBarisRincianSptjb_, pdN_) -- pastikan file itu ADA di project
 * yang sama (satu project Apps Script = satu namespace, jadi boleh
 * beda file .gs, tidak perlu import apa pun).
 *
 * SETUP (SEKALI SAJA):
 * 1. Tambahkan file ini + pastikan PaketDokumenService.gs juga ada.
 * 2. Di editor Apps Script, jalankan fungsi pasangKontrolCetak()
 *    sekali (pilih di dropdown fungsi, klik Run, Allow kalau diminta
 *    izin). Ini menambahkan label + dropdown "Pilih No SPTJB" di
 *    sheet Nominatif, kolom V-W baris 1 (area kosong, tidak
 *    menabrak layout cetak yang sudah ada di A1:T46).
 * 3. Reload spreadsheet-nya (supaya menu custom "SIPERDIN" muncul).
 * 4. Pilih No SPTJB dari dropdown -> 4 sheet terisi otomatis.
 *    Kalau karena suatu sebab dropdown tidak memicu (jarang, tapi
 *    Google Sheets kadang begitu untuk paste massal), pakai menu
 *    SIPERDIN > Muat Ulang Dokumen Cetak.
 *
 * KETERBATASAN (batas wajar untuk versi pertama):
 * - Sheet Nominatif tabelnya bisa menampung maksimal 18 orang
 *   (baris 13-30, sama seperti template asli). Kalau SPTJB itu
 *   personilnya lebih dari 18, sisanya TIDAK tercetak dan akan
 *   muncul peringatan -- beri tahu saya kalau ini sering kejadian,
 *   nanti saya buatkan agar barisnya menyisip otomatis.
 * - Kolom "Akun" di Nominatif (label REF_AKUN tingkat SubOutput)
 *   dan kode item (000422 dst di SPTJB) diambil dari REF_AKUN --
 *   kalau ada Kegiatan/Output/SubOutput/Akun baru yang belum ada di
 *   REF_AKUN, kolom itu akan kosong (bukan error), tinggal
 *   dilengkapi di sheet REF_AKUN.
 * - SPJ RAMPUNG & Sppd Depan BELUM disentuh (sesuai kesepakatan --
 *   fokus 4 dokumen dulu).
 * ------------------------------------------------------------
 */

var TC_SHEET_KONTROL = 'Nominatif';
var TC_SEL_DROPDOWN = 'W1';
var TC_SEL_LABEL = 'V1';

// ============================================================
// SETUP SEKALI SAJA
// ============================================================

function pasangKontrolCetak() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(TC_SHEET_KONTROL);
  if (!sh) throw new Error('Sheet "' + TC_SHEET_KONTROL + '" tidak ditemukan.');

  sh.getRange(TC_SEL_LABEL).setValue('PILIH NO SPTJB:').setFontWeight('bold').setBackground('#FFF3CD');
  const selDropdown = sh.getRange(TC_SEL_DROPDOWN);
  selDropdown.setBackground('#FFF3CD').setFontWeight('bold');

  const shSptjb = ss.getSheetByName('DB_SPTJB');
  const lastRow = Math.max(shSptjb.getLastRow(), 2);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(shSptjb.getRange('B2:B' + lastRow), true)
    .setAllowInvalid(false)
    .build();
  selDropdown.setDataValidation(rule);

  Logger.log('Kontrol cetak terpasang di ' + TC_SHEET_KONTROL + '!' + TC_SEL_DROPDOWN + '. Reload spreadsheet untuk melihat menu SIPERDIN.');
}

/**
 * Menu custom -- cadangan manual kalau dropdown tidak memicu onEdit,
 * dan supaya jelas ke mana harus klik untuk "refresh".
 * CATATAN: kalau project ini SUDAH punya fungsi onOpen() di file
 * lain, gabungkan isi fungsi ini ke situ (Apps Script cuma memakai
 * SATU fungsi bernama onOpen per project).
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('SIPERDIN')
    .addItem('🔄 Muat Ulang Dokumen Cetak', 'muatUlangDariMenu')
    .addItem('⚙️ Pasang/Reset Kontrol Cetak', 'pasangKontrolCetak')
    .addToUi();
}

function muatUlangDariMenu() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TC_SHEET_KONTROL);
  const noSptjb = sh.getRange(TC_SEL_DROPDOWN).getValue();
  if (!noSptjb) { SpreadsheetApp.getUi().alert('Pilih No SPTJB dulu di ' + TC_SHEET_KONTROL + '!' + TC_SEL_DROPDOWN + '.'); return; }
  muatDokumenCetak_(noSptjb);
}

// ============================================================
// PEMICU OTOMATIS
// CATATAN: kalau project ini SUDAH punya fungsi onEdit() di file
// lain, gabungkan isi fungsi ini ke situ (sama seperti onOpen,
// Apps Script cuma memakai SATU fungsi bernama onEdit per project).
// ============================================================

function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet();
    if (sh.getName() !== TC_SHEET_KONTROL) return;
    if (e.range.getA1Notation() !== TC_SEL_DROPDOWN) return;

    const noSptjb = e.value;
    if (!noSptjb) return;
    muatDokumenCetak_(noSptjb);
  } catch (err) {
    SpreadsheetApp.getUi().alert('Gagal memuat dokumen: ' + err.message);
  }
}

// ============================================================
// INTI: AMBIL DATA (pakai ulang PaketDokumenService.gs) LALU TULIS
// ============================================================

function muatDokumenCetak_(noSptjbTeks) {
  const idSptjb = tcCariIdSptjbDariNoSptjb_(noSptjbTeks);
  if (!idSptjb) { SpreadsheetApp.getUi().alert('No SPTJB "' + noSptjbTeks + '" tidak ditemukan di DB_SPTJB.'); return; }

  let D;
  try {
    D = pdKumpulkanData_(idSptjb); // fungsi ini ada di PaketDokumenService.gs
  } catch (err) {
    SpreadsheetApp.getUi().alert('Gagal memuat data: ' + err.message);
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  tcTulisKwitansi_(ss.getSheetByName('Kwitansi'), D);
  tcTulisSpby_(ss.getSheetByName('SPBY'), D);
  tcTulisSptjb_(ss.getSheetByName('SPTJB'), D);
  tcTulisNominatif_(ss.getSheetByName('Nominatif'), D);

  if (D.peringatan.length) {
    SpreadsheetApp.getUi().alert('Dokumen dimuat, tapi ada yang perlu dicek:\n\n' + D.peringatan.join('\n'));
  }
}

function tcCariIdSptjbDariNoSptjb_(noSptjbTeks) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DB_SPTJB');
  const data = sh.getDataRange().getValues();
  const target = String(noSptjbTeks || '').trim();
  for (let r = 1; r < data.length; r++) {
    if (String(data[r][1]).trim() === target) return data[r][0];
  }
  return '';
}

// ============================================================
// KWITANSI
// ============================================================

function tcTulisKwitansi_(sh, D) {
  if (!sh) return;
  const cfg = D.cfg;
  const tahun = (D.baris[0].tglB ? D.baris[0].tglB.getFullYear() : new Date().getFullYear());

  sh.getRange('M10').setValue(tahun);
  sh.getRange('M12').setValue(D.klasifikasi);
  sh.getRange('F16').setValue(D.jumlah.total);
  sh.getRange('F18').setValue(pdTerbilang_(D.jumlah.total));
  sh.getRange('F20').setValue(D.uraian);
  sh.getRange('F24').setValue('Lunas Tgl. ' + pdTitikTanggal_());
  sh.getRange('L24').setValue(pdKotaTanggal_(cfg.kota_penandatanganan));

  sh.getRange('A32').setValue(pdIsian_(cfg.nama_ppk, 'Nama PPK'));
  sh.getRange('A33').setValue('NIP. ' + (cfg.nip_ppk || '-'));
  sh.getRange('F32').setValue(pdIsian_(cfg.nama_bendahara, 'Nama Bendahara'));
  sh.getRange('F33').setValue('NIP. ' + (cfg.nip_bendahara || '-'));
  sh.getRange('L32').setValue(pdIsian_(cfg.nama_ketua_tim_kerja, 'Nama Ketua Tim Kerja'));
  sh.getRange('L33').setValue('NIP. ' + (cfg.nip_ketua_tim_kerja || '-'));
}

// ============================================================
// SPBY
// ============================================================

function tcTulisSpby_(sh, D) {
  if (!sh) return;
  const cfg = D.cfg;

  sh.getRange('C6').setValue('Tanggal :' + pdTitikTanggal_() + '     Nomor');
  sh.getRange('A9').setValue(D.jumlah.total);
  sh.getRange('A10').setValue(pdTerbilang_(D.jumlah.total));
  sh.getRange('C12').setValue(D.penerima);
  sh.getRange('C13').setValue(D.uraian);
  sh.getRange('D17').setValue(': ' + D.noStGabung);
  sh.getRange('C22').setValue(D.klasifikasi);
  sh.getRange('C23').setValue(D.noAkun);
  sh.getRange('C23').setNumberFormat('@'); // cegah kode akun dibaca sebagai angka/tanggal

  sh.getRange('A27').setValue('Setuju/lunas dibayar, Tgl' + pdTitikTanggal_());
  sh.getRange('E27').setValue(pdKotaTanggal_(cfg.kota_penandatanganan));
  sh.getRange('A33').setValue(pdIsian_(cfg.nama_bendahara, 'Nama Bendahara'));
  sh.getRange('A34').setValue('NIP  ' + (cfg.nip_bendahara || '-'));
  sh.getRange('E33').setValue(pdIsian_(cfg.nama_ppk, 'Nama PPK'));
  sh.getRange('E34').setValue('NIP  ' + (cfg.nip_ppk || '-'));
}

// ============================================================
// SPTJB
// ============================================================

function tcTulisSptjb_(sh, D) {
  if (!sh) return;
  const cfg = D.cfg;

  sh.getRange('A3').setValue('Nomor  : ' + D.noSptjb);
  sh.getRange('C5').setValue(': ' + (cfg.kode_satker || '-'));
  sh.getRange('C6').setValue(': ' + (cfg.nama_satker || '-'));
  sh.getRange('C7').setValue(': ' + D.noDipaTeks);
  sh.getRange('C8').setValue(': ' + D.klasifikasi);
  sh.getRange('C9').setValue(': ' + (cfg.bidang || 'Pemberdayaan Tenaga Kerja Mandiri'));

  sh.getRange('B16').setValue(D.klasifikasi);
  sh.getRange('C16').setValue(D.penerima);
  sh.getRange('D16').setValue(cfg.npwp || '-');
  sh.getRange('D16').setNumberFormat('@');
  sh.getRange('E16').setValue(cfg.alamat_satker || '-');
  sh.getRange('F16').setValue(D.uraian);
  sh.getRange('G16').setValue(D.noStGabung);
  sh.getRange('G16').setNumberFormat('@');
  sh.getRange('H16').setValue(new Date());

  // rincian (Tiket, Transport, Uang Harian, Uang Hotel, Transport Bandara,
  // Uang Representatif) -- maksimal 5 baris seperti template (baris 17-21)
  const rincian = pdBarisRincianSptjb_(D);
  sh.getRange('F17:I21').clearContent();
  rincian.slice(0, 5).forEach(function (x, i) {
    sh.getRange(17 + i, 6).setValue(x.teks);
    sh.getRange(17 + i, 9).setValue(x.jumlah);
  });
  if (rincian.length > 5) {
    SpreadsheetApp.getUi().alert('SPTJB: ada ' + rincian.length + ' komponen biaya, tapi tabel cuma muat 5 baris -- komponen terakhir tidak tercetak. Beri tahu saya kalau ini sering terjadi.');
  }

  sh.getRange('G29').setValue(pdKotaTanggal_(cfg.kota_penandatanganan));
  sh.getRange('C37').setValue(pdIsian_(cfg.nama_ketua_tim_kerja, 'Nama Ketua Tim Kerja'));
  sh.getRange('C38').setValue('NIP ' + (cfg.nip_ketua_tim_kerja || '-'));
  sh.getRange('G37').setValue(pdIsian_(cfg.nama_ppk, 'Nama PPK'));
  sh.getRange('G38').setValue('NIP ' + (cfg.nip_ppk || '-'));
}

// ============================================================
// NOMINATIF
// ============================================================

var TC_NOMINATIF_MAKS_ORANG = 18; // baris 13-30, sama seperti template asli

function tcTulisNominatif_(sh, D) {
  if (!sh) return;
  const cfg = D.cfg;
  const akun = D.akun;

  sh.getRange('N1').setValue(D.uraian);
  sh.getRange('N3').setValue(D.noStGabung);
  sh.getRange('N3').setNumberFormat('@');

  sh.getRange('C5').setValue(': ' + D.noDipaTeks);
  sh.getRange('C6').setValue(': ' + (akun.namaKomponen || '-'));
  sh.getRange('S6').setValue(akun.kodeKomponen || '');
  sh.getRange('C7').setValue(': ' + (akun.namaSub || '-'));

  const subParts = String(akun.kodeSub || '').split('.');
  sh.getRange('S7').setValue(subParts[0] || '');
  sh.getRange('S8').setValue((subParts[1] || '').replace(/^0/, ''));

  // Label "4. Akun" di template ini memuat nama tingkat SubOutput
  // (bukan nama akun 524111-nya) -- cari langsung dari REF_AKUN.
  const namaSubOutput = tcCariNamaSubOutput_(akun.kodeKomponen, subParts[0], akun.kodeSub);
  sh.getRange('C9').setValue(': ' + (namaSubOutput || akun.namaAkun || '-'));
  sh.getRange('S9').setValue(D.noAkun);
  sh.getRange('S9').setNumberFormat('@');

  const baris = D.baris.slice(0, TC_NOMINATIF_MAKS_ORANG);
  sh.getRange('A13:S30').clearContent();
  baris.forEach(function (b, i) {
    const r = 13 + i;
    const malam = Math.max(b.lama - 1, 0);
    sh.getRange(r, 1).setValue(i + 1);
    sh.getRange(r, 2).setValue(b.nama);
    sh.getRange(r, 3).setValue('Jakarta');
    sh.getRange(r, 4).setValue(b.tujuan || b.daerah || '-');
    sh.getRange(r, 5).setValue(pdRentangSingkat_(b.tglB, b.tglK));
    sh.getRange(r, 6).setValue(b.lama || '');
    sh.getRange(r, 7).setValue(b.lama ? 'Hari' : '');
    sh.getRange(r, 8).setValue(pdN_(b.tiket));
    sh.getRange(r, 9).setValue(pdN_(b.kedudukan));
    sh.getRange(r, 10).setValue(pdN_(b.tujuanTaksi));
    sh.getRange(r, 11).setValue(pdN_(b.lainnya));
    sh.getRange(r, 12).setValue(b.lama > 0 ? b.uangHarian / b.lama : '');
    sh.getRange(r, 13).setValue(b.lama ? 'x ' + b.lama + ' =' : '');
    sh.getRange(r, 14).setValue(pdN_(b.uangHarian));
    sh.getRange(r, 15).setValue(malam > 0 ? b.penginapan / malam : '');
    sh.getRange(r, 16).setValue(malam > 0 ? 'x ' + malam + ' =' : '');
    sh.getRange(r, 17).setValue(pdN_(b.penginapan));
    sh.getRange(r, 18).setValue(pdN_(b.representatif));
    sh.getRange(r, 19).setValue(pdN_(b.total));
  });
  if (D.baris.length > TC_NOMINATIF_MAKS_ORANG) {
    SpreadsheetApp.getUi().alert('Nominatif: SPTJB ini punya ' + D.baris.length + ' orang, tapi tabel cetak cuma muat ' +
      TC_NOMINATIF_MAKS_ORANG + ' baris -- ' + (D.baris.length - TC_NOMINATIF_MAKS_ORANG) + ' orang terakhir TIDAK tercetak. Beri tahu saya untuk saya perbesar tabelnya.');
  }
  // baris 31 (JUMLAH SELURUHNYA) TIDAK disentuh -- rumus SUM yang
  // sudah ada di sheet (=SUM(H13:H30) dst) otomatis mengikuti.

  sh.getRange('O35').setValue(pdKotaTanggal_(cfg.kota_penandatanganan));
  sh.getRange('A41').setValue(pdIsian_(cfg.nama_ppk, 'Nama PPK'));
  sh.getRange('A42').setValue('NIP. ' + (cfg.nip_ppk || '-'));
  sh.getRange('G41').setValue(pdIsian_(cfg.nama_ketua_tim_kerja, 'Nama Ketua Tim Kerja'));
  sh.getRange('G42').setValue('NIP. ' + (cfg.nip_ketua_tim_kerja || '-'));
  sh.getRange('O41').setValue(pdIsian_(cfg.nama_bendahara, 'Nama Bendahara'));
  sh.getRange('O42').setValue('NIP. ' + (cfg.nip_bendahara || '-'));
}

/**
 * Nama level SubOutput (mis. "052.0A" -> "Pendampingan Penyusunan LPJ
 * Bantuan TKM Pemula") -- pdCariInfoAkun_ di PaketDokumenService.gs
 * tidak mengambil level ini (dia berhenti di level Output & Akun),
 * jadi dicari terpisah di sini.
 */
function tcCariNamaSubOutput_(kodeKegiatanPenuh, outputKode, subKode) {
  // kodeKegiatanPenuh contoh "2175.BDC.003" -> kegiatan-nya "BDC.003"
  const parts = String(kodeKegiatanPenuh || '').split('.');
  const kegiatan = parts.slice(1).join('.');
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('REF_AKUN');
  if (!sh || !kegiatan || !subKode) return '';
  const data = sh.getDataRange().getValues();
  for (let r = 1; r < data.length; r++) {
    if (String(data[r][0]).trim() === 'SubOutput' &&
      String(data[r][3]).trim() === kegiatan &&
      String(data[r][4]).trim() === outputKode &&
      String(data[r][5]).trim() === subKode) {
      return String(data[r][2]).trim();
    }
  }
  return '';
}
