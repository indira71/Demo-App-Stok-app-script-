/**DelCab
 * SISTEM PENJUALAN & STOK — ES TEH NUSANTARA (v5: CABANG DINAMIS + UBAH NAMA/HAPUS CABANG)
 *
 * - Stok GUDANG satu (bersama). Stok AREA SALES terpisah per cabang.
 * - Penjualan, pengeluaran, transfer (gudang -> area cabang), penyesuaian kas/area, tutup buku: per cabang.
 * - Belanja bahan baku (stok masuk) masuk gudang & dicatat di cabang "Pusat".
 * - Laporan/dashboard: pilih satu cabang atau "Semua" (gabungan + rincian per cabang).
 * - Data lama (tanpa kolom Cabang) otomatis dianggap milik CABANG[0].
 * - Cabang bisa ditambah dari aplikasi (tombol "+ Cabang"), disimpan di sheet CABANG_LIST.
 * - BARU: cabang bisa diubah namanya (semua data ikut berganti) dan dihapus (dengan konfirmasi bila masih punya data).
 * - Transaksi penjualan bisa dihapus dari menu Kasir (stok otomatis kembali).
 *
 * Upgrade dari v4: tempel kode ini, lalu Deploy > Manage deployments > Edit > New version.
 * Spreadsheet baru: SETUP_DATABASE() -> (opsional) SEED_CONTOH_DATA().
 */

var TZ = "Asia/Makassar";
var HTML_FILE = "index";
var METODE = ["Tunai", "Qris"];
var UKURAN = ["18oz", "22oz"];
var JENIS_DEFAULT = ["Operasional", "Bahan Baku"];
var CABANG_DEFAULT = ["Cabang 1", "Cabang 2"];
var CABANG = loadCabang_();
var PUSAT = "Pusat";

/** Daftar cabang dibaca dari sheet CABANG_LIST; jika belum ada, pakai CABANG_DEFAULT. */
function loadCabang_() {
  try {
    var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CABANG_LIST");
    if (!s || s.getLastRow() < 2) return CABANG_DEFAULT.slice();
    var list = s.getRange(2, 1, s.getLastRow() - 1, 1).getValues()
      .map(function (r) { return String(r[0]).trim(); })
      .filter(function (x) { return x; });
    return list.length ? list : CABANG_DEFAULT.slice();
  } catch (e) { return CABANG_DEFAULT.slice(); }
}

var SCHEMA = {
  MENU:           ["ID Menu", "Nama Menu", "Tipe", "Harga 18oz", "Harga 22oz", "HPP 18oz", "HPP 22oz"],
  RESEP:          ["ID Menu", "ID Bahan", "Jumlah per Porsi", "Ukuran"],
  STOK:           ["ID Barang", "Nama Barang", "Tipe", "Satuan", "Stok Awal Gudang", "Stok Awal Area Sales", "Stok Awal Area Sales 2"],
  PENJUALAN:      ["ID Transaksi", "Tanggal", "ID Menu", "Nama Menu", "Ukuran", "Jumlah", "Harga Jual",
                   "Metode Bayar", "Total Penjualan", "Total HPP", "Profit", "Cabang"],
  PEMAKAIAN_STOK: ["ID", "Tanggal", "ID Transaksi", "ID Barang", "Nama Barang", "Jumlah Terpakai", "Cabang"],
  STOK_MASUK:     ["ID Masuk", "Tanggal", "ID Barang", "Nama Barang", "Jumlah Masuk", "Harga Beli Total", "Keterangan"],
  TRANSFER_STOK:  ["ID", "Tanggal", "ID Barang", "Nama Barang", "Jumlah", "Keterangan", "Cabang"],
  PENYESUAIAN:    ["ID", "Tanggal", "Jenis", "ID Barang", "Nama Barang", "Jumlah/Nominal", "Metode Bayar", "Keterangan", "Lokasi", "Cabang"],
  PENGELUARAN:    ["ID Pengeluaran", "Tanggal", "Kategori", "Jenis", "Keterangan", "Jumlah Biaya", "Metode Bayar", "Cabang"],
  JENIS_PENGELUARAN: ["Jenis Pengeluaran"],
  TUTUP_BUKU:     ["Tanggal", "Ditutup Pada", "Omset", "HPP", "Profit", "Biaya Operasional", "Laba Bersih", "Saldo Tunai", "Saldo Qris", "Cabang"],
  STOK_BULANAN:   ["Bulan", "ID Barang", "Nama Barang", "Lokasi", "Stok Awal", "Masuk", "Keluar/Terpakai", "Penyesuaian", "Stok Sisa"]
};

var TEXT_COLS = {
  PENJUALAN: [2], PEMAKAIAN_STOK: [2], STOK_MASUK: [2], TRANSFER_STOK: [2],
  PENYESUAIAN: [2], PENGELUARAN: [2], STOK_BULANAN: [1], TUTUP_BUKU: [1, 2]
};

// indeks kolom Cabang (0-based) per sheet
var CI = { PENJUALAN: 11, PEMAKAIAN: 6, TRANSFER: 6, ADJ: 9, KELUAR: 7, TUTUP: 9 };

// [nama sheet, indeks kolom Cabang (0-based)] — dipakai untuk ubah nama / hapus cabang
var CAB_SHEETS = [
  ["PENJUALAN", CI.PENJUALAN], ["PEMAKAIAN_STOK", CI.PEMAKAIAN],
  ["TRANSFER_STOK", CI.TRANSFER], ["PENYESUAIAN", CI.ADJ],
  ["PENGELUARAN", CI.KELUAR], ["TUTUP_BUKU", CI.TUTUP]
];

var SEQ_ = 0;

// ================= SETUP =================

function SETUP_DATABASE() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var stamp = Utilities.formatDate(new Date(), TZ, "yyMMdd");

  Object.keys(SCHEMA).forEach(function (name) {
    var headers = SCHEMA[name];
    var eq = function (h, i) { return h === headers[i]; };
    var sheet = ss.getSheetByName(name);

    if (sheet) {
      var lastCol = sheet.getLastColumn();
      var existing = lastCol ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
      while (existing.length && existing[existing.length - 1] === "") existing.pop();

      // Header sudah sama, atau lebih panjang (kolom cabang tambahan) -> aman
      if (existing.length >= headers.length && headers.every(function (h, i) { return existing[i] === h; })) return;

      // Header lama = awalan header baru -> cukup tambah kolom baru (data lama aman)
      if (existing.length > 0 && existing.length < headers.length && existing.every(eq)) {
        if (sheet.getMaxColumns() < headers.length) {
          sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
        }
        sheet.getRange(1, 1, 1, headers.length).setValues([headers])
          .setFontWeight("bold").setBackground("#e0e0e0");
        return;
      }
      sheet.setName(name + "_LAMA_" + stamp);
    }

    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight("bold").setBackground("#e0e0e0");
    sheet.setFrozenRows(1);
    (TEXT_COLS[name] || []).forEach(function (c) {
      sheet.getRange(2, c, sheet.getMaxRows() - 1, 1).setNumberFormat("@");
    });
  });

  var def = ss.getSheetByName("Sheet1");
  if (def && ss.getSheets().length > 1) ss.deleteSheet(def);
  getJenisPengeluaran();
}

/** Migrasi dari v1 -> v2/v3 (hanya jika masih punya sheet STOK v1 dengan kolom "Harga Beli"). */
function MIGRASI_V2() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var old = ss.getSheetByName("STOK");
  var transfer = [];

  if (old && old.getLastColumn() >= 5 && old.getRange(1, 5).getValue() === "Harga Beli") {
    var stamp = Utilities.formatDate(new Date(), TZ, "yyMMdd_HHmm");
    var last = old.getLastRow();
    var rows = last >= 2 ? old.getRange(2, 1, last - 1, 10).getValues().filter(function (r) { return r[0] !== ""; }) : [];
    old.copyTo(ss).setName("STOK_LAMA_" + stamp);

    var baru = rows.map(function (r) { return [r[0], r[1], r[2], r[3], 0, num_(r[5]), 0]; });
    old.clear();
    old.getRange(1, 1, 1, SCHEMA.STOK.length).setValues([SCHEMA.STOK])
      .setFontWeight("bold").setBackground("#e0e0e0");
    old.setFrozenRows(1);
    if (baru.length) old.getRange(2, 1, baru.length, 7).setValues(baru);

    var masuk = ss.getSheetByName("STOK_MASUK");
    if (masuk && masuk.getLastRow() >= 2) {
      masuk.getRange(2, 1, masuk.getLastRow() - 1, 7).getValues()
        .filter(function (r) { return r[0] !== ""; })
        .forEach(function (r) {
          transfer.push([id_("TRF"), normDate_(r[1]), r[2], r[3], num_(r[4]), "Migrasi: stok masuk lama", CABANG[0]]);
        });
    }
  }

  SETUP_DATABASE();
  appendRows_("TRANSFER_STOK", transfer);
  SpreadsheetApp.getUi().alert("Migrasi selesai. Sheet MENU & RESEP versi baru sudah dibuat (yang lama dicadangkan *_LAMA_*). Silakan input ulang menu lewat aplikasi.");
}

function SEED_ULANG() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var stamp = Utilities.formatDate(new Date(), TZ, "yyMMdd_HHmm");
  Object.keys(SCHEMA).forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (sh) sh.setName(name + "_BACKUP_" + stamp);
  });
  SETUP_DATABASE();
  SEED_CONTOH_DATA();
}

function tglOffset_(n) {
  var d = new Date();
  d.setDate(d.getDate() - n);
  return Utilities.formatDate(d, TZ, "yyyy-MM-dd");
}

function SEED_CONTOH_DATA() {
  if (readRows_("STOK").length || readRows_("MENU").length) {
    throw new Error("STOK/MENU sudah berisi data. Jalankan SEED_ULANG() untuk mencadangkan data lama lalu mengisi ulang.");
  }
  var C1 = CABANG[0], C2 = CABANG[1];

  // 1. STOK: [nama, kategori, gudang, area cabang 1] ; area cabang 2 = 70% dari cabang 1
  var stok = [
    ["Cup Ori Polos", "Bahan", 100, 70], ["Cup Teh Nusantara 18oz ETN", "Bahan", 100, 45],
    ["Cup Ori Polos (Jumbo)", "Bahan", 60, 45], ["Cup Teh Nusantara 22oz ETN (JUMBO)", "Bahan", 60, 28],
    ["Cup Hot", "Bahan", 50, 25], ["Cup ETN (ramadhan)", "Bahan", 0, 0], ["Cream Machiatto/100gr", "Bahan", 0, 0],
    ["Chocolate 0,5 (Choco)", "Bahan Baku", 20, 10], ["Peach Tea", "Bahan Baku", 10, -1],
    ["Lychee Tea", "Bahan Baku", 15, 9], ["Lemon Tea", "Bahan Baku", 15, 14], ["GreenTea", "Bahan Baku", 10, 4],
    ["Taro", "Bahan Baku", 10, 2], ["Teh Jahe Jeruk", "Bahan Baku", 15, 11], ["Teh Sereh Wangi", "Bahan Baku", 15, 16],
    ["Melon", "Bahan Baku", 15, 13], ["Honey Lime", "Bahan Baku", 10, 8], ["CocoPandan", "Bahan Baku", 15, 12],
    ["Red Velvet", "Bahan Baku", 10, 6], ["CaffeLatte", "Bahan Baku", 15, 9], ["Mango", "Bahan Baku", 15, 10]
  ];
  var idByName = {};
  appendRows_("STOK", stok.map(function (s, i) {
    var id = "BRG-" + ("00" + (i + 1)).slice(-3);
    idByName[s[0]] = id;
    return [id, s[0], s[1], "pcs", s[2], s[3], Math.floor(s[3] * 0.7)];
  }));

  // 2. MENU & RESEP
  var CUP18 = "Cup Teh Nusantara 18oz ETN", CUP22 = "Cup Teh Nusantara 22oz ETN (JUMBO)";
  var menu = [["Teh", "Teh Nusantara", 6000, 8000, 450, 600,
    [["Cup Ori Polos", "18oz", 1], ["Cup Ori Polos (Jumbo)", "22oz", 1]]]];
  [["Milktea", "Teh Nusantara", 7000, 10000], ["Milo", "Susu Bubuk", 9000, 12000], ["Matcha", "Teh Poci", 10000, 13000]]
    .forEach(function (m) { menu.push([m[0], m[1], m[2], m[3], 600, 800, [[CUP18, "18oz", 1], [CUP22, "22oz", 1]]]); });
  [
    ["Peach Tea", "Peach Tea", "Tea Series", 8000, 12000], ["Lychee Tea", "Lychee Tea", "Tea Series", 8000, 12000],
    ["Lemon Tea", "Lemon Tea", "Tea Series", 8000, 12000], ["Teh Jahe Jeruk", "Teh Jahe Jeruk", "Tea Series", 8000, 12000],
    ["Teh Sereh Wangi", "Teh Sereh Wangi", "Tea Series", 8000, 12000], ["Honey Lime", "Honey Lime", "Tea Series", 8000, 13000],
    ["GreenTea", "GreenTea", "Milk Series", 8000, 12000], ["Taro", "Taro", "Milk Series", 8000, 12000],
    ["Red Velvet", "Red Velvet", "Milk Series", 9000, 13000], ["Choco", "Chocolate 0,5 (Choco)", "Milk Series", 9000, 13000],
    ["CaffeLatte", "CaffeLatte", "Coffee", 8000, 12000], ["Melon", "Melon", "Fruit Series", 8000, 12000],
    ["CocoPandan", "CocoPandan", "Fruit Series", 8000, 12000], ["Mango", "Mango", "Fruit Series", 9000, 13000]
  ].forEach(function (r) {
    menu.push([r[0], r[2], r[3], r[4], 3100, 3300, [[CUP18, "18oz", 1], [CUP22, "22oz", 1], [r[1], "Semua", 1]]]);
  });
  var menuRows = [], resepRows = [], idMenu = {};
  menu.forEach(function (m, i) {
    var id = "MNU-" + ("00" + (i + 1)).slice(-3);
    idMenu[m[0]] = id;
    menuRows.push([id, m[0], m[1], m[2], m[3], m[4], m[5]]);
    m[6].forEach(function (b) { resepRows.push([id, idByName[b[0]], b[2], b[1]]); });
  });
  appendRows_("MENU", menuRows);
  appendRows_("RESEP", resepRows);

  // 3. Tanggal: maks 7 hari terakhir di bulan berjalan
  var bulan = today_().slice(0, 7), dates = [];
  for (var k = 6; k >= 0; k--) { var t = tglOffset_(k); if (t.slice(0, 7) === bulan) dates.push(t); }
  var awal = dates[0], akhir = dates[dates.length - 1];

  // 4. Kas awal tiap cabang
  CABANG.forEach(function (c) {
    addPenyesuaian({ tanggal: awal, cabang: c, jenis: "Kas", jumlah: 250000, metodeBayar: "Tunai", keterangan: "Modal awal kas" });
  });

  // 5. Stok masuk (ke GUDANG bersama)
  [["Cup Teh Nusantara 18oz ETN", 50, 30000, "Beli cup 18oz"], ["Cup Teh Nusantara 22oz ETN (JUMBO)", 20, 16000, "Beli cup 22oz"],
   ["Lychee Tea", 12, 30000, "Restok Lychee"], ["GreenTea", 10, 27000, "Restok GreenTea"], ["Taro", 10, 30000, "Restok Taro"]
  ].forEach(function (m) {
    addStokMasuk({ tanggal: awal, idBarang: idByName[m[0]], jumlah: m[1], hargaBeliTotal: m[2],
                   keterangan: m[3], metodeBayar: "Tunai", catatPengeluaran: true });
  });

  // 5b. Transfer gudang -> area tiap cabang
  [["Cup Teh Nusantara 18oz ETN", 15], ["Cup Teh Nusantara 22oz ETN (JUMBO)", 8], ["GreenTea", 3], ["Taro", 3]]
    .forEach(function (m) {
      CABANG.forEach(function (c) {
        addTransfer({ tanggal: awal, cabang: c, idBarang: idByName[m[0]], jumlah: m[1], keterangan: "Isi area " + c });
      });
    });

  // 6. Penjualan: 3 transaksi per hari, cabang & metode bayar bergantian
  var paket = [
    [["Lychee Tea", "18oz", 2], ["Teh", "18oz", 1]], [["Taro", "18oz", 1], ["Honey Lime", "22oz", 1]],
    [["GreenTea", "18oz", 2]], [["Mango", "22oz", 1], ["Milktea", "18oz", 2]],
    [["CaffeLatte", "18oz", 1], ["Lemon Tea", "18oz", 1], ["Teh", "22oz", 1]],
    [["Red Velvet", "18oz", 1], ["Choco", "22oz", 1]], [["Peach Tea", "18oz", 1], ["Melon", "18oz", 2]]
  ];
  dates.forEach(function (tgl, d) {
    for (var j = 0; j < 3; j++) {
      var p = paket[(d * 2 + j) % paket.length];
      addPenjualan({
        tanggal: tgl, cabang: CABANG[(d + j) % 2], metodeBayar: (d + j) % 2 === 0 ? "Tunai" : "Qris",
        items: p.map(function (x) { return { idMenu: idMenu[x[0]], ukuran: x[1], jumlah: x[2] }; })
      });
    }
  });

  // 7. Pengeluaran per cabang
  addJenisPengeluaran({ nama: "Operasional Lainnya" });
  addPengeluaran({ tanggal: awal, cabang: C1, kategori: "Listrik", jenis: "Operasional", keterangan: "Token listrik", jumlah: 50000, metodeBayar: "Tunai" });
  addPengeluaran({ tanggal: awal, cabang: C2, kategori: "Listrik", jenis: "Operasional", keterangan: "Token listrik", jumlah: 40000, metodeBayar: "Tunai" });
  dates.forEach(function (tgl) {
    CABANG.forEach(function (c) {
      addPengeluaran({ tanggal: tgl, cabang: c, kategori: "Es Batu", jenis: "Operasional", keterangan: "Es batu harian", jumlah: 20000, metodeBayar: "Tunai" });
    });
  });
  addPengeluaran({ tanggal: akhir, cabang: C1, kategori: "Lainnya", jenis: "Operasional", keterangan: "Plastik & sedotan", jumlah: 15000, metodeBayar: "Tunai" });
  addPengeluaran({ tanggal: akhir, cabang: C2, kategori: "Parkir & Kebersihan", jenis: "Operasional Lainnya", keterangan: "Iuran kebersihan", jumlah: 10000, metodeBayar: "Qris" });

  // 8. Penyesuaian
  addPenyesuaian({ tanggal: akhir, jenis: "Stok", lokasi: "Area Sales", cabang: C1, idBarang: idByName["Cup Hot"], jumlah: -2, keterangan: "Cup rusak" });
  addPenyesuaian({ tanggal: akhir, jenis: "Stok", lokasi: "Area Sales", cabang: C2, idBarang: idByName["Lemon Tea"], jumlah: -1, keterangan: "Tumpah" });
  addPenyesuaian({ tanggal: akhir, jenis: "Kas", cabang: C1, jumlah: -2000, metodeBayar: "Tunai", keterangan: "Selisih kas (uang kurang)" });

  // 9. Contoh tutup buku: hanya tanggal pertama untuk cabang pertama
  if (dates.length > 1) tutupBuku(awal, C1);
}

function HAPUS_SEMUA_DATA() {
  var ui = SpreadsheetApp.getUi();
  var jawab = ui.alert("Hapus SEMUA data?",
    "Isi semua sheet data akan dikosongkan. Header tetap ada. Tindakan ini tidak bisa dibatalkan.", ui.ButtonSet.YES_NO);
  if (jawab !== ui.Button.YES) return;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SCHEMA).forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (!sh || sh.getMaxRows() < 2) return;
    sh.getRange(2, 1, sh.getMaxRows() - 1, sh.getMaxColumns()).clearContent();
  });
  SpreadsheetApp.flush();
  ui.alert("Selesai. Semua data sudah dikosongkan. (Daftar cabang di sheet CABANG_LIST tidak dihapus.)");
}

function HAPUS_SHEET_BACKUP() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var target = ss.getSheets().filter(function (sh) { return /_BACKUP_|_LAMA_/.test(sh.getName()); });
  if (!target.length) { ui.alert("Tidak ada sheet cadangan."); return; }
  var jawab = ui.alert("Hapus " + target.length + " sheet cadangan?",
    target.map(function (s) { return s.getName(); }).join("\n") + "\n\nTindakan ini tidak bisa dibatalkan.", ui.ButtonSet.YES_NO);
  if (jawab !== ui.Button.YES) return;
  target.forEach(function (sh) { ss.deleteSheet(sh); });
  ui.alert("Selesai. Sheet cadangan dihapus.");
}

// ================= WEB APP =================

function doGet() {
  var tpl;
  try { tpl = HtmlService.createTemplateFromFile(HTML_FILE); }
  catch (e) { tpl = HtmlService.createTemplateFromFile(HTML_FILE.charAt(0).toUpperCase() + HTML_FILE.slice(1)); }
  return tpl.evaluate()
    .setTitle("Aplikasi Penjualan & Stok")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Konfigurasi untuk tampilan: daftar cabang & jenis pengeluaran. */
function getConfig() {
  return { cabang: CABANG, jenis: getJenisPengeluaran() };
}

// ================= CABANG =================

/** data: {nama} -> mengembalikan daftar cabang terbaru. */
function addCabang(data) {
  return withLock_(function () {
    var nama = String((data && data.nama) || "").trim();
    if (!nama) throw new Error("Nama cabang wajib diisi");
    if (nama.indexOf("|") >= 0) throw new Error('Nama cabang tidak boleh memakai karakter "|"');
    var low = nama.toLowerCase();
    if (low === PUSAT.toLowerCase() || CABANG.some(function (c) { return c.toLowerCase() === low; })) {
      throw new Error('Cabang "' + nama + '" sudah ada');
    }

    var ss = ss_();
    var sh = ss.getSheetByName("CABANG_LIST");
    if (!sh) {
      sh = ss.insertSheet("CABANG_LIST");
      sh.getRange(1, 1).setValue("Nama Cabang").setFontWeight("bold").setBackground("#e0e0e0");
      sh.setFrozenRows(1);
      sh.getRange(2, 1, CABANG.length, 1).setValues(CABANG.map(function (c) { return [c]; }));
    }
    sh.appendRow([nama]);

    // kolom stok awal area untuk cabang baru
    var st = sheet_("STOK");
    var col = 6 + CABANG.length;
    if (st.getMaxColumns() < col) st.insertColumnsAfter(st.getMaxColumns(), col - st.getMaxColumns());
    st.getRange(1, col).setValue("Stok Awal Area Sales " + (CABANG.length + 1))
      .setFontWeight("bold").setBackground("#e0e0e0");

    return CABANG.concat([nama]);
  });
}

/** Pastikan sheet CABANG_LIST ada & terisi (dibuat dari daftar saat ini bila belum). */
function ensureCabangSheet_() {
  var ss = ss_();
  var sh = ss.getSheetByName("CABANG_LIST");
  if (!sh) {
    sh = ss.insertSheet("CABANG_LIST");
    sh.getRange(1, 1).setValue("Nama Cabang").setFontWeight("bold").setBackground("#e0e0e0");
    sh.setFrozenRows(1);
  }
  if (sh.getLastRow() < 2) {
    sh.getRange(2, 1, CABANG.length, 1).setValues(CABANG.map(function (c) { return [c]; }));
  }
  return sh;
}

/** Ganti nilai kolom Cabang (col = nomor kolom 1-based) dari 'lama' ke 'baru'. */
function gantiCab_(name, col, lama, baru) {
  var s = sheet_(name), last = s.getLastRow();
  if (last < 2 || s.getMaxColumns() < col) return;
  var rg = s.getRange(2, col, last - 1, 1);
  rg.setValues(rg.getValues().map(function (r) {
    return [String(r[0]) === lama ? baru : r[0]];
  }));
}

/** data: {lama, baru} -> daftar cabang terbaru. Semua data lama ikut berganti nama. */
function renameCabang(data) {
  return withLock_(function () {
    var lama = String((data && data.lama) || "").trim();
    var baru = String((data && data.baru) || "").trim();
    var idx = CABANG.indexOf(lama);
    if (idx < 0) throw new Error("Cabang tidak ditemukan");
    if (!baru) throw new Error("Nama cabang wajib diisi");
    if (baru.indexOf("|") >= 0) throw new Error('Nama cabang tidak boleh memakai karakter "|"');
    var low = baru.toLowerCase();
    if (low === PUSAT.toLowerCase() ||
        CABANG.some(function (c, i) { return i !== idx && c.toLowerCase() === low; })) {
      throw new Error('Cabang "' + baru + '" sudah ada');
    }
    var hasil = CABANG.map(function (c) { return c === lama ? baru : c; });
    if (baru === lama) return hasil;

    ensureCabangSheet_().getRange(idx + 2, 1).setValue(baru);
    CAB_SHEETS.forEach(function (x) { gantiCab_(x[0], x[1] + 1, lama, baru); });

    // arsip stok bulanan: kolom Lokasi "Area Sales - <cabang>"
    var sb = sheet_("STOK_BULANAN"), lb = sb.getLastRow();
    if (lb >= 2) {
      var rg = sb.getRange(2, 4, lb - 1, 1);
      rg.setValues(rg.getValues().map(function (r) {
        return [r[0] === "Area Sales - " + lama ? "Area Sales - " + baru : r[0]];
      }));
    }
    return hasil;
  });
}

/** Hitung (hanyaHitung=true) atau hapus semua baris transaksi milik cabang. Mengembalikan jumlah baris. */
function hapusDataCab_(nama, hanyaHitung) {
  var n = 0;
  CAB_SHEETS.forEach(function (x) {
    var s = sheet_(x[0]), last = s.getLastRow();
    if (last < 2) return;
    var vals = s.getRange(2, 1, last - 1, SCHEMA[x[0]].length).getValues();
    for (var i = vals.length - 1; i >= 0; i--) {
      var r = vals[i];
      if (r[0] === "" || r[0] === null) continue;
      if (x[0] === "PENYESUAIAN" && r[8] === "Gudang") continue; // penyesuaian gudang milik bersama
      if (cabOf_(r, x[1]) !== nama) continue;
      n++;
      if (!hanyaHitung) s.deleteRow(i + 2);
    }
  });
  return n;
}

/**
 * data: {nama, paksa}
 * Jika cabang masih punya transaksi dan paksa != true -> {ok:false, jumlah}.
 * Jika paksa == true -> cabang + semua datanya dihapus.
 */
function hapusCabang(data) {
  return withLock_(function () {
    var nama = String((data && data.nama) || "").trim();
    var idx = CABANG.indexOf(nama);
    if (idx < 0) throw new Error("Cabang tidak ditemukan");
    if (CABANG.length <= 1) throw new Error("Minimal harus ada 1 cabang");

    var jml = hapusDataCab_(nama, true);
    if (jml > 0 && !(data && data.paksa)) return { ok: false, jumlah: jml };
    if (jml > 0) hapusDataCab_(nama, false);

    // hapus arsip stok bulanan cabang ini
    var sb = sheet_("STOK_BULANAN"), lb = sb.getLastRow();
    if (lb >= 2) {
      var v = sb.getRange(2, 1, lb - 1, 9).getValues();
      for (var i = v.length - 1; i >= 0; i--) {
        if (v[i][3] === "Area Sales - " + nama) sb.deleteRow(i + 2);
      }
    }

    // hapus dari daftar cabang
    ensureCabangSheet_().deleteRow(idx + 2);

    // hapus kolom "Stok Awal Area Sales" milik cabang ini lalu rapikan header
    var st = sheet_("STOK");
    st.deleteColumn(6 + idx);
    if (st.getMaxColumns() < 7) st.insertColumnsAfter(st.getMaxColumns(), 7 - st.getMaxColumns());
    var sisa = CABANG.length - 1, tot = Math.max(2, sisa), hdr = [];
    for (var k = 0; k < tot; k++) hdr.push("Stok Awal Area Sales" + (k ? " " + (k + 1) : ""));
    st.getRange(1, 6, 1, tot).setValues([hdr]).setFontWeight("bold").setBackground("#e0e0e0");

    return { ok: true, cabang: CABANG.filter(function (c) { return c !== nama; }) };
  });
}

// ================= HELPER =================

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function sheet_(name) {
  var s = ss_().getSheetByName(name);
  if (!s) throw new Error('Sheet "' + name + '" tidak ditemukan. Jalankan SETUP_DATABASE terlebih dahulu.');
  return s;
}

function readRows_(name) {
  var s = sheet_(name);
  var last = s.getLastRow();
  if (last < 2) return [];
  var n = name === "STOK" ? Math.max(SCHEMA.STOK.length, 5 + CABANG.length) : SCHEMA[name].length;
  if (s.getMaxColumns() < n) throw new Error("Sheet " + name + " belum diperbarui. Jalankan SETUP_DATABASE sekali.");
  return s.getRange(2, 1, last - 1, n).getValues()
    .filter(function (r) { return r[0] !== "" && r[0] !== null; });
}

function appendRows_(name, rows) {
  if (!rows.length) return;
  var s = sheet_(name);
  var start = s.getLastRow() + 1;
  var need = start + rows.length - 1;
  if (need > s.getMaxRows()) s.insertRowsAfter(s.getMaxRows(), need - s.getMaxRows());
  s.getRange(start, 1, rows.length, rows[0].length).setValues(rows);
}

function num_(v) { return Number(v) || 0; }

function normDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, "yyyy-MM-dd");
  return String(v || "").slice(0, 10);
}

function today_() { return Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd"); }

function tgl_(v) {
  var t = v ? normDate_(v) : today_();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(t)) throw new Error("Format tanggal harus yyyy-MM-dd");
  return t;
}

function metode_(v) {
  var m = v || "Tunai";
  if (METODE.indexOf(m) === -1) throw new Error("Metode bayar harus Tunai atau Qris");
  return m;
}

function lokasi_(v) { return v === "Gudang" ? "Gudang" : "Area Sales"; }

/** Cabang wajib untuk transaksi tulis. */
function cabReq_(v) {
  if (CABANG.indexOf(v) === -1) throw new Error("Pilih cabang terlebih dahulu");
  return v;
}

/** Nilai cabang pada baris; kosong (data lama) = cabang pertama. */
function cabOf_(r, idx) { return r[idx] === "" || r[idx] == null ? CABANG[0] : r[idx]; }

/** Filter baris per cabang; cab kosong = semua cabang. */
function filterCab_(rows, idx, cab) {
  if (!cab) return rows;
  return rows.filter(function (r) { return cabOf_(r, idx) === cab; });
}

function id_(prefix) {
  SEQ_++;
  return prefix + "-" + Utilities.formatDate(new Date(), TZ, "yyMMddHHmmss") + "-" +
    Utilities.getUuid().slice(0, 4) + SEQ_;
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try { var res = fn(); SpreadsheetApp.flush(); return res; }
  finally { lock.releaseLock(); }
}

function bulanRange_(bulanTahun) {
  if (!/^\d{4}-\d{2}$/.test(bulanTahun || "")) throw new Error("Format bulan harus yyyy-MM");
  var y = Number(bulanTahun.slice(0, 4)), m = Number(bulanTahun.slice(5, 7));
  var next = m === 12 ? (y + 1) + "-01-01" : y + "-" + ("0" + (m + 1)).slice(-2) + "-01";
  return { from: bulanTahun + "-01", to: next };
}

function fmtRow_(row, dateIdx) {
  var r = row.slice();
  r[dateIdx] = normDate_(r[dateIdx]);
  return r;
}

// ---------- Tutup buku (per tanggal + cabang) ----------

/** { "yyyy-MM-dd|Cabang": "waktu ditutup" } */
function tutupMap_() {
  var m = {};
  readRows_("TUTUP_BUKU").forEach(function (r) {
    var w = r[1] instanceof Date ? Utilities.formatDate(r[1], TZ, "dd/MM/yyyy HH:mm") : String(r[1] || "");
    m[normDate_(r[0]) + "|" + cabOf_(r, CI.TUTUP)] = w;
  });
  return m;
}

/** Status tutup: cab diisi = cabang itu; kosong = true jika SEMUA cabang sudah ditutup. */
function tutup_(t, cab) {
  var m = tutupMap_();
  if (cab) return m[t + "|" + cab] !== undefined;
  return CABANG.every(function (c) { return m[t + "|" + c] !== undefined; });
}

function cekBuka_(tgl, cab) {
  if (tutup_(tgl, cab)) throw new Error("Buku " + cab + " tanggal " + tgl + " sudah ditutup. Transaksi tidak bisa ditambahkan.");
}

/** Untuk transaksi gudang (milik bersama): ditolak hanya jika semua cabang sudah tutup buku. */
function cekBukaGudang_(tgl) {
  if (tutup_(tgl, "")) throw new Error("Buku tanggal " + tgl + " sudah ditutup di semua cabang.");
}

// ---------- Stok (gudang bersama + area per cabang) ----------

/**
 * Gudang : awal + masuk - transfer (semua cabang) + penyesuaian(Gudang)
 * Area[c]: awal + transfer ke c - terpakai (penjualan c) + penyesuaian(Area, c)
 * Periode [from, to) opsional: log sebelum 'from' masuk ke stok awal periode.
 */
function hitungStok_(from, to) {
  var map = {}, order = [];
  readRows_("STOK").forEach(function (r) {
    var it = { id: r[0], nama: r[1], tipe: r[2], satuan: r[3],
               gAwal: num_(r[4]), gMasuk: 0, gKeluar: 0, gAdj: 0, area: {} };
    CABANG.forEach(function (c, i) { it.area[c] = { awal: num_(r[5 + i]), masuk: 0, terpakai: 0, adj: 0 }; });
    map[r[0]] = it;
    order.push(r[0]);
  });

  // pick(r, item) -> {o: objek, awal: kunci, per: kunci, sign} atau null
  function proses(rows, idIdx, qtyIdx, pick) {
    rows.forEach(function (r) {
      var it = map[r[idIdx]];
      if (!it) return;
      var p = pick(r, it);
      if (!p || !p.o) return;
      var t = normDate_(r[1]), q = num_(r[qtyIdx]);
      if (to && t >= to) return;
      if (from && t < from) p.o[p.awal] += p.sign * q;
      else p.o[p.per] += q;
    });
  }
  proses(readRows_("STOK_MASUK"), 2, 4, function (r, it) { return { o: it, awal: "gAwal", per: "gMasuk", sign: 1 }; });
  var trf = readRows_("TRANSFER_STOK");
  proses(trf, 2, 4, function (r, it) { return { o: it, awal: "gAwal", per: "gKeluar", sign: -1 }; });
  proses(trf, 2, 4, function (r, it) { return { o: it.area[cabOf_(r, CI.TRANSFER)], awal: "awal", per: "masuk", sign: 1 }; });
  proses(readRows_("PEMAKAIAN_STOK"), 3, 5, function (r, it) {
    return { o: it.area[cabOf_(r, CI.PEMAKAIAN)], awal: "awal", per: "terpakai", sign: -1 };
  });
  proses(readRows_("PENYESUAIAN"), 3, 5, function (r, it) {
    if (r[2] !== "Stok") return null;
    if (r[8] === "Gudang") return { o: it, awal: "gAwal", per: "gAdj", sign: 1 };
    return { o: it.area[cabOf_(r, CI.ADJ)], awal: "awal", per: "adj", sign: 1 };
  });

  order.forEach(function (id) {
    var it = map[id];
    it.gSisa = it.gAwal + it.gMasuk - it.gKeluar + it.gAdj;
    CABANG.forEach(function (c) {
      var a = it.area[c];
      a.sisa = a.awal + a.masuk - a.terpakai + a.adj;
    });
  });
  return { map: map, order: order };
}

function resepMap_() {
  var m = {};
  readRows_("RESEP").forEach(function (r) {
    (m[r[0]] = m[r[0]] || []).push({ idBahan: r[1], jumlah: num_(r[2]), ukuran: r[3] || "Semua" });
  });
  return m;
}

function resepRows_(idMenu, items) {
  return (items || []).map(function (b) {
    var uk = UKURAN.indexOf(b.ukuran) >= 0 ? b.ukuran : "Semua";
    return [idMenu, b.idBahan, Number(b.jumlah) || 0, uk];
  });
}

// ================= MASTER DATA =================

function getMenuData() {
  var resep = resepMap_();
  var stok = hitungStok_().map;
  return readRows_("MENU").map(function (r) {
    return {
      id: r[0], nama: r[1], tipe: r[2],
      harga18: num_(r[3]), harga22: num_(r[4]), hpp18: num_(r[5]), hpp22: num_(r[6]),
      resep: (resep[r[0]] || []).map(function (b) {
        return { idBahan: b.idBahan, namaBahan: stok[b.idBahan] ? stok[b.idBahan].nama : "(dihapus)",
                 jumlah: b.jumlah, ukuran: b.ukuran };
      })
    };
  });
}

function cekHargaMenu_(data) {
  if (!data.nama || !String(data.nama).trim()) throw new Error("Nama menu wajib diisi");
  if (!(Number(data.harga18) > 0) || !(Number(data.harga22) > 0)) throw new Error("Harga ukuran 18oz dan 22oz wajib diisi");
}

function addMenu(data) {
  return withLock_(function () {
    cekHargaMenu_(data);
    var id = id_("MNU");
    appendRows_("MENU", [[id, String(data.nama).trim(), data.tipe || "",
      Number(data.harga18), Number(data.harga22), num_(data.hpp18), num_(data.hpp22)]]);
    appendRows_("RESEP", resepRows_(id, data.resep));
    return id;
  });
}

function saveResep(idMenu, items) {
  return withLock_(function () {
    hapusResepRows_(idMenu);
    appendRows_("RESEP", resepRows_(idMenu, items));
    return true;
  });
}

/** Daftar stok saat ini: gudang (gSisa...) dan area[cabang] (awal, masuk, terpakai, adj, sisa). */
function getStokData() {
  var h = hitungStok_();
  return h.order.map(function (id) { return h.map[id]; });
}

/** data: {nama, tipe, satuan, stokAwalGudang, stokAwalArea:[cabang1, cabang2, ...]} */
function addStokBarang(data) {
  return withLock_(function () {
    if (!data.nama) throw new Error("Nama barang wajib diisi");
    var id = id_("BRG"), aw = data.stokAwalArea || [];
    var row = [id, data.nama, data.tipe || "Bahan Baku", data.satuan || "pcs", Number(data.stokAwalGudang) || 0];
    CABANG.forEach(function (c, i) { row.push(Number(aw[i]) || 0); });
    appendRows_("STOK", [row]);
    return id;
  });
}

function cariBarisStok_(s, id) {
  var last = s.getLastRow();
  if (last >= 2) {
    var ids = s.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) if (ids[i][0] === id) return i + 2;
  }
  throw new Error("Barang tidak ditemukan");
}

/** Edit barang. data: {id, nama, tipe, satuan, stokAwalGudang, stokAwalArea:[...]} */
function updateStokBarang(data) {
  return withLock_(function () {
    if (!data.nama || !String(data.nama).trim()) throw new Error("Nama barang wajib diisi");
    var s = sheet_("STOK");
    var row = cariBarisStok_(s, data.id);
    var cur = s.getRange(row, 1, 1, Math.max(7, 5 + CABANG.length)).getValues()[0], aw = data.stokAwalArea || [];
    var kosong = function (x) { return x === "" || x === null || x === undefined; };
    var vals = [String(data.nama).trim(), data.tipe || cur[2], data.satuan || cur[3],
      kosong(data.stokAwalGudang) ? num_(cur[4]) : (Number(data.stokAwalGudang) || 0)];
    CABANG.forEach(function (c, i) { vals.push(kosong(aw[i]) ? num_(cur[5 + i]) : (Number(aw[i]) || 0)); });
    s.getRange(row, 2, 1, vals.length).setValues([vals]);
    return true;
  });
}

function hapusStokBarang(id) {
  return withLock_(function () {
    var s = sheet_("STOK");
    var row = cariBarisStok_(s, id);
    var dipakai = readRows_("RESEP").filter(function (r) { return r[1] === id; }).map(function (r) { return r[0]; });
    if (dipakai.length) {
      var nm = {};
      readRows_("MENU").forEach(function (m) { nm[m[0]] = m[1]; });
      var daftar = dipakai.filter(function (x, i) { return dipakai.indexOf(x) === i; }).map(function (x) { return nm[x] || x; });
      throw new Error("Barang masih dipakai di resep: " + daftar.join(", ") + ". Hapus dari resep dulu.");
    }
    s.deleteRow(row);
    return true;
  });
}

// ================= JENIS PENGELUARAN =================

function getJenisPengeluaran() {
  var list = readRows_("JENIS_PENGELUARAN").map(function (r) { return String(r[0]).trim(); }).filter(function (x) { return x; });
  if (!list.length) {
    appendRows_("JENIS_PENGELUARAN", JENIS_DEFAULT.map(function (j) { return [j]; }));
    list = JENIS_DEFAULT.slice();
  }
  return list;
}

function addJenisPengeluaran(data) {
  return withLock_(function () {
    var nama = String((data && data.nama) || "").trim();
    if (!nama) throw new Error("Nama jenis wajib diisi");
    var list = getJenisPengeluaran();
    if (list.some(function (x) { return x.toLowerCase() === nama.toLowerCase(); })) throw new Error('Jenis "' + nama + '" sudah ada');
    appendRows_("JENIS_PENGELUARAN", [[nama]]);
    return list.concat([nama]);
  });
}

// ================= TRANSAKSI =================

/**
 * Stok masuk (belanja) -> menambah stok GUDANG (bersama). Pengeluaran dicatat di cabang "Pusat".
 * data: {tanggal, idBarang, jumlah, hargaBeliTotal, keterangan, metodeBayar, catatPengeluaran}
 */
function addStokMasuk(data) {
  return withLock_(function () {
    var stok = hitungStok_().map[data.idBarang];
    if (!stok) throw new Error("Barang tidak ditemukan");
    var jumlah = Number(data.jumlah);
    if (!(jumlah > 0)) throw new Error("Jumlah masuk harus lebih dari 0");
    var total = Number(data.hargaBeliTotal) || 0;
    var tgl = tgl_(data.tanggal);
    cekBukaGudang_(tgl);

    appendRows_("STOK_MASUK", [[id_("IN"), tgl, stok.id, stok.nama, jumlah, total, data.keterangan || ""]]);
    if (total > 0 && data.catatPengeluaran) {
      appendRows_("PENGELUARAN", [[id_("OUT"), tgl, "Belanja Bahan Baku", "Bahan Baku",
        stok.nama + " x" + jumlah, total, metode_(data.metodeBayar), PUSAT]]);
    }
    return true;
  });
}

/** Transfer gudang -> area sales CABANG tujuan. data: {tanggal, cabang, idBarang, jumlah, keterangan} */
function addTransfer(data) {
  return withLock_(function () {
    var stok = hitungStok_().map[data.idBarang];
    if (!stok) throw new Error("Barang tidak ditemukan");
    var cab = cabReq_(data.cabang);
    var jumlah = Number(data.jumlah);
    if (!(jumlah > 0)) throw new Error("Jumlah transfer harus lebih dari 0");
    var tgl = tgl_(data.tanggal);
    cekBuka_(tgl, cab);
    appendRows_("TRANSFER_STOK", [[id_("TRF"), tgl, stok.id, stok.nama, jumlah, data.keterangan || "", cab]]);
    var warnings = [];
    if (stok.gSisa - jumlah < 0) warnings.push(stok.nama + " di gudang akan minus (" + (stok.gSisa - jumlah) + ")");
    return { ok: true, warnings: warnings };
  });
}

/**
 * data: {tanggal, cabang, metodeBayar, items:[{idMenu, ukuran, jumlah, hargaJual?}]}
 * Stok AREA SALES cabang tsb berkurang sesuai resep.
 */
function addPenjualan(data) {
  return withLock_(function () {
    var items = (data.items || []).filter(function (it) { return Number(it.jumlah) > 0; });
    if (!items.length) throw new Error("Tidak ada item penjualan");
    var cab = cabReq_(data.cabang);
    var tgl = tgl_(data.tanggal), metode = metode_(data.metodeBayar);
    cekBuka_(tgl, cab);

    var menuMap = {};
    readRows_("MENU").forEach(function (r) { menuMap[r[0]] = r; });
    var resep = resepMap_();
    var stok = hitungStok_().map;
    var idTrx = id_("TRX");

    var jualRows = [], pakaiRows = [], need = {}, warnings = [];
    var sumJual = 0, sumHpp = 0;

    items.forEach(function (it) {
      var m = menuMap[it.idMenu];
      if (!m) throw new Error("Menu tidak ditemukan: " + it.idMenu);
      var uk = it.ukuran;
      if (UKURAN.indexOf(uk) === -1) throw new Error("Ukuran gelas harus 18oz atau 22oz");
      var jml = Number(it.jumlah);
      var hargaMenu = num_(uk === "22oz" ? m[4] : m[3]);
      var harga = (it.hargaJual !== undefined && it.hargaJual !== "" && Number(it.hargaJual) >= 0) ? Number(it.hargaJual) : hargaMenu;
      var hpp = num_(uk === "22oz" ? m[6] : m[5]) * jml;

      var bahanList = (resep[m[0]] || []).filter(function (b) { return b.ukuran === "Semua" || b.ukuran === uk; });
      bahanList.forEach(function (b) {
        var bahan = stok[b.idBahan];
        if (!bahan) { warnings.push("Bahan resep " + m[1] + " tidak ditemukan di STOK (dilewati)"); return; }
        var qty = b.jumlah * jml;
        need[bahan.id] = (need[bahan.id] || 0) + qty;
        pakaiRows.push([id_("USE"), tgl, idTrx, bahan.id, bahan.nama, qty, cab]);
      });
      if (!bahanList.length) warnings.push("Menu " + m[1] + " (" + uk + ") belum punya resep, stok tidak berkurang");

      var total = jml * harga;
      sumJual += total; sumHpp += hpp;
      jualRows.push([idTrx, tgl, m[0], m[1], uk, jml, harga, metode, total, hpp, total - hpp, cab]);
    });

    Object.keys(need).forEach(function (id) {
      var sisa = stok[id].area[cab].sisa - need[id];
      if (sisa < 0) warnings.push(stok[id].nama + " di area sales " + cab + " akan minus (" + sisa + ")");
    });

    appendRows_("PENJUALAN", jualRows);
    appendRows_("PEMAKAIAN_STOK", pakaiRows);
    return { ok: true, idTransaksi: idTrx, totalPenjualan: sumJual, totalHPP: sumHpp,
             profit: sumJual - sumHpp, warnings: warnings };
  });
}

/**
 * Hapus satu transaksi penjualan utuh (semua item ber-ID sama).
 * Stok kembali otomatis karena baris PEMAKAIAN_STOK ikut dihapus.
 * Ditolak jika buku tanggal tsb (untuk cabang tsb) sudah ditutup.
 */
function hapusPenjualan(idTrx) {
  return withLock_(function () {
    var sp = sheet_("PENJUALAN");
    var vals = sp.getDataRange().getValues();
    var rows = [], tgl = "", cab = "";
    for (var i = 1; i < vals.length; i++) {
      if (vals[i][0] === idTrx) {
        rows.push(i + 1);
        tgl = normDate_(vals[i][1]);
        cab = cabOf_(vals[i], CI.PENJUALAN);
      }
    }
    if (!rows.length) throw new Error("Transaksi tidak ditemukan");
    if (tutup_(tgl, cab)) throw new Error("Buku " + cab + " tanggal " + tgl + " sudah ditutup. Transaksi tidak bisa dihapus.");

    for (var k = rows.length - 1; k >= 0; k--) sp.deleteRow(rows[k]);

    var su = sheet_("PEMAKAIAN_STOK");
    var u = su.getDataRange().getValues();
    for (var j = u.length - 1; j >= 1; j--) {
      if (u[j][2] === idTrx) su.deleteRow(j + 1);
    }
    return true;
  });
}

/** data: {tanggal, cabang, kategori, jenis, keterangan, jumlah, metodeBayar} */
function addPengeluaran(data) {
  return withLock_(function () {
    var cab = cabReq_(data.cabang);
    var tgl = tgl_(data.tanggal);
    cekBuka_(tgl, cab);
    var jenis = getJenisPengeluaran().indexOf(data.jenis) >= 0 ? data.jenis : "Operasional";
    appendRows_("PENGELUARAN", [[id_("OUT"), tgl, data.kategori || "", jenis,
      data.keterangan || "", Number(data.jumlah) || 0, metode_(data.metodeBayar), cab]]);
    return true;
  });
}

/**
 * data: {tanggal, jenis:'Stok'|'Kas', idBarang + lokasi:'Gudang'|'Area Sales' (Stok),
 *        cabang (wajib untuk Kas & Area Sales), jumlah (+/-), metodeBayar (Kas), keterangan}
 */
function addPenyesuaian(data) {
  return withLock_(function () {
    var jumlah = Number(data.jumlah);
    if (!jumlah) throw new Error("Jumlah penyesuaian tidak boleh 0");
    var tgl = tgl_(data.tanggal);
    if (data.jenis === "Stok") {
      var b = hitungStok_().map[data.idBarang];
      if (!b) throw new Error("Barang tidak ditemukan");
      var lok = lokasi_(data.lokasi), cab = "";
      if (lok === "Gudang") cekBukaGudang_(tgl);
      else { cab = cabReq_(data.cabang); cekBuka_(tgl, cab); }
      appendRows_("PENYESUAIAN", [[id_("ADJ"), tgl, "Stok", b.id, b.nama, jumlah, "", data.keterangan || "", lok, cab]]);
    } else if (data.jenis === "Kas") {
      var c2 = cabReq_(data.cabang);
      cekBuka_(tgl, c2);
      appendRows_("PENYESUAIAN", [[id_("ADJ"), tgl, "Kas", "", "", jumlah, metode_(data.metodeBayar), data.keterangan || "", "", c2]]);
    } else {
      throw new Error("Jenis penyesuaian harus Stok atau Kas");
    }
    return true;
  });
}

// ================= DASHBOARD & LAPORAN =================

function ringkas_(sales, expenses, adjKas) {
  var s = { cup: 0, omset: 0, tunai: 0, qris: 0, hpp: 0, grossProfit: 0,
            outcomeOperasional: 0, outcomeBahanBaku: 0, saldoTunai: 0, saldoQris: 0, perJenis: {} };
  sales.forEach(function (r) {
    s.cup += num_(r[5]); s.omset += num_(r[8]); s.hpp += num_(r[9]); s.grossProfit += num_(r[10]);
    if (r[7] === "Qris") s.qris += num_(r[8]); else s.tunai += num_(r[8]);
  });
  var keluarTunai = 0, keluarQris = 0;
  expenses.forEach(function (r) {
    var j = num_(r[5]), jenis = r[3] || "Operasional";
    if (jenis === "Bahan Baku") s.outcomeBahanBaku += j; else s.outcomeOperasional += j;
    s.perJenis[jenis] = (s.perJenis[jenis] || 0) + j;
    if (r[6] === "Qris") keluarQris += j; else keluarTunai += j;
  });
  var adjTunai = 0, adjQris = 0;
  (adjKas || []).forEach(function (r) {
    if (r[6] === "Qris") adjQris += num_(r[5]); else adjTunai += num_(r[5]);
  });
  s.saldoTunai = s.tunai - keluarTunai + adjTunai;
  s.saldoQris = s.qris - keluarQris + adjQris;
  s.adjTunai = adjTunai;
  s.adjQris = adjQris;
  s.adjKas = adjTunai + adjQris;
  // HPP sudah memuat biaya bahan baku -> jenis "Bahan Baku" tidak dikurangkan lagi.
  s.netIncome = s.grossProfit - s.outcomeOperasional;
  return s;
}

/** Ringkasan terpisah untuk tiap cabang. */
function perCabang_(sales, exp, adjKas) {
  var o = {};
  CABANG.forEach(function (c) {
    var s = ringkas_(filterCab_(sales, CI.PENJUALAN, c), filterCab_(exp, CI.KELUAR, c), filterCab_(adjKas, CI.ADJ, c));
    o[c] = { omset: s.omset, cup: s.cup, hpp: s.hpp, grossProfit: s.grossProfit,
             outcomeOperasional: s.outcomeOperasional, netIncome: s.netIncome,
             saldoTunai: s.saldoTunai, saldoQris: s.saldoQris };
  });
  return o;
}

function filterBulan_(rows, bulanTahun) {
  if (!bulanTahun) return rows;
  return rows.filter(function (r) { return normDate_(r[1]).slice(0, 7) === bulanTahun; });
}

/** bulanTahun opsional; cab opsional (kosong = semua cabang + rincian perCabang). */
function getDashboardData(bulanTahun, cab) {
  var sales = filterBulan_(readRows_("PENJUALAN"), bulanTahun);
  var exp = filterBulan_(readRows_("PENGELUARAN"), bulanTahun);
  var adj = filterBulan_(readRows_("PENYESUAIAN"), bulanTahun).filter(function (r) { return r[2] === "Kas"; });
  var stok = getStokData();
  var ringkasan = ringkas_(filterCab_(sales, CI.PENJUALAN, cab), filterCab_(exp, CI.KELUAR, cab), filterCab_(adj, CI.ADJ, cab));
  if (!cab) ringkasan.perCabang = perCabang_(sales, exp, adj);
  ringkasan.totalBarang = stok.length;
  ringkasan.totalMenu = readRows_("MENU").length;
  ringkasan.stokMinus = [];
  stok.forEach(function (s) {
    CABANG.forEach(function (c) {
      if ((!cab || cab === c) && s.area[c].sisa < 0) ringkasan.stokMinus.push({ nama: s.nama, lokasi: "Area " + c, sisa: s.area[c].sisa });
    });
    if (s.gSisa < 0) ringkasan.stokMinus.push({ nama: s.nama, lokasi: "Gudang", sisa: s.gSisa });
  });
  return ringkasan;
}

/** tanggal: "yyyy-MM-dd"; cab opsional. */
function getDailySalesReport(tanggal, cab) {
  var t = tgl_(tanggal);
  var sales = filterCab_(readRows_("PENJUALAN").filter(function (r) { return normDate_(r[1]) === t; }), CI.PENJUALAN, cab);
  var exp = filterCab_(readRows_("PENGELUARAN").filter(function (r) { return normDate_(r[1]) === t; }), CI.KELUAR, cab);
  var adj = filterCab_(readRows_("PENYESUAIAN").filter(function (r) { return normDate_(r[1]) === t && r[2] === "Kas"; }), CI.ADJ, cab);
  return {
    tanggal: t, cabang: cab || "",
    closed: tutup_(t, cab || ""),
    sales: sales.map(function (r) { return fmtRow_(r, 1); }),
    expenses: exp.map(function (r) { return fmtRow_(r, 1); }),
    summary: ringkas_(sales, exp, adj)
  };
}

/** bulanTahun: "yyyy-MM"; cab opsional. */
function getMonthlyReport(bulanTahun, cab) {
  bulanRange_(bulanTahun);
  var salesAll = filterBulan_(readRows_("PENJUALAN"), bulanTahun);
  var expAll = filterBulan_(readRows_("PENGELUARAN"), bulanTahun);
  var adjAll = filterBulan_(readRows_("PENYESUAIAN"), bulanTahun).filter(function (r) { return r[2] === "Kas"; });
  var sales = filterCab_(salesAll, CI.PENJUALAN, cab);
  var exp = filterCab_(expAll, CI.KELUAR, cab);

  var harian = {};
  sales.forEach(function (r) {
    var t = normDate_(r[1]);
    var h = harian[t] = harian[t] || { tanggal: t, cup: 0, omset: 0, tunai: 0, qris: 0, profit: 0 };
    h.cup += num_(r[5]); h.omset += num_(r[8]); h.profit += num_(r[10]);
    if (r[7] === "Qris") h.qris += num_(r[8]); else h.tunai += num_(r[8]);
  });

  return {
    sales: sales.map(function (r) { return fmtRow_(r, 1); }),
    expenses: exp.map(function (r) { return fmtRow_(r, 1); }),
    harian: Object.keys(harian).sort().map(function (k) { return harian[k]; }),
    tutup: Object.keys(tutupMap_()).filter(function (k) {
      return k.slice(0, 7) === bulanTahun && (!cab || k.split("|")[1] === cab);
    }),
    summary: ringkas_(sales, exp, filterCab_(adjAll, CI.ADJ, cab)),
    perCabang: cab ? null : perCabang_(salesAll, expAll, adjAll),
    stokBulanan: getStokBulanan(bulanTahun),
    adj: filterBulan_(readRows_("PENYESUAIAN"), bulanTahun).filter(function (r) {
  return !cab || r[8] === "Gudang" || cabOf_(r, CI.ADJ) === cab;}).map(function (r) { return fmtRow_(r, 1); })
  };
}

// ================= STOK BULANAN =================

function getStokBulanan(bulanTahun) {
  var r = bulanRange_(bulanTahun);
  var h = hitungStok_(r.from, r.to);
  return h.order.map(function (id) {
    var s = h.map[id];
    return { id: s.id, nama: s.nama, satuan: s.satuan,
             gAwal: s.gAwal, gMasuk: s.gMasuk, gKeluar: s.gKeluar, gAdj: s.gAdj, gSisa: s.gSisa, area: s.area };
  });
}

/** Arsip stok bulan tsb ke STOK_BULANAN (timpa bulan yang sama). Panggil dari dalam lock. */
function arsipStok_(bulanTahun) {
  var data = getStokBulanan(bulanTahun);
  var s = sheet_("STOK_BULANAN");
  var vals = s.getDataRange().getValues();
  for (var i = vals.length - 1; i >= 1; i--) {
    if (String(vals[i][0]) === bulanTahun) s.deleteRow(i + 1);
  }
  var rows = [];
  data.forEach(function (d) {
    rows.push([bulanTahun, d.id, d.nama, "Gudang", d.gAwal, d.gMasuk, d.gKeluar, d.gAdj, d.gSisa]);
    CABANG.forEach(function (c) {
      var a = d.area[c];
      rows.push([bulanTahun, d.id, d.nama, "Area Sales - " + c, a.awal, a.masuk, a.terpakai, a.adj, a.sisa]);
    });
  });
  appendRows_("STOK_BULANAN", rows);
  return data.length;
}

// ================= TUTUP BUKU (per cabang) =================

/** tanggal, cab (opsional: kosong = semua cabang tertutup?) -> {closed, info} */
function getStatusBuku(tanggal, cab) {
  var t = tgl_(tanggal), m = tutupMap_();
  return { tanggal: t, cabang: cab || "", closed: tutup_(t, cab || ""), info: cab ? (m[t + "|" + cab] || "") : "" };
}

/** Tutup buku satu tanggal untuk satu cabang. Setelah itu tanggal tsb terkunci untuk cabang itu. */
function tutupBuku(tanggal, cab) {
  return withLock_(function () {
    cab = cabReq_(cab);
    var t = tgl_(tanggal);
    if (tutup_(t, cab)) throw new Error("Buku " + cab + " tanggal " + t + " sudah ditutup");
    var sales = filterCab_(readRows_("PENJUALAN").filter(function (r) { return normDate_(r[1]) === t; }), CI.PENJUALAN, cab);
    var exp = filterCab_(readRows_("PENGELUARAN").filter(function (r) { return normDate_(r[1]) === t; }), CI.KELUAR, cab);
    var adj = filterCab_(readRows_("PENYESUAIAN").filter(function (r) { return normDate_(r[1]) === t && r[2] === "Kas"; }), CI.ADJ, cab);
    var s = ringkas_(sales, exp, adj);
    appendRows_("TUTUP_BUKU", [[t, Utilities.formatDate(new Date(), TZ, "dd/MM/yyyy HH:mm"),
      s.omset, s.hpp, s.grossProfit, s.outcomeOperasional, s.netIncome, s.saldoTunai, s.saldoQris, cab]]);
    arsipStok_(t.slice(0, 7));
    return true;
  });
}

function bukaBuku(tanggal, cab) {
  return withLock_(function () {
    cab = cabReq_(cab);
    var t = tgl_(tanggal);
    var s = sheet_("TUTUP_BUKU");
    var vals = s.getDataRange().getValues();
    var hapus = 0;
    for (var i = vals.length - 1; i >= 1; i--) {
      if (normDate_(vals[i][0]) === t && cabOf_(vals[i], CI.TUTUP) === cab) { s.deleteRow(i + 1); hapus++; }
    }
    if (!hapus) throw new Error("Buku " + cab + " tanggal " + t + " belum ditutup");
    return true;
  });
}

// ================= EDIT / HAPUS MENU =================

function hapusResepRows_(idMenu) {
  var s = sheet_("RESEP");
  var vals = s.getDataRange().getValues();
  for (var i = vals.length - 1; i >= 1; i--) {
    if (vals[i][0] === idMenu) s.deleteRow(i + 1);
  }
}

function cariBarisMenu_(s, id) {
  var vals = s.getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) {
    if (vals[i][0] === id) return i + 1;
  }
  throw new Error("Menu tidak ditemukan");
}

function updateMenu(data) {
  return withLock_(function () {
    cekHargaMenu_(data);
    var s = sheet_("MENU");
    var row = cariBarisMenu_(s, data.id);
    s.getRange(row, 2, 1, 6).setValues([[
      String(data.nama).trim(), data.tipe || "",
      Number(data.harga18), Number(data.harga22), num_(data.hpp18), num_(data.hpp22)
    ]]);
    hapusResepRows_(data.id);
    appendRows_("RESEP", resepRows_(data.id, data.resep));
    return true;
  });
}

function hapusMenu(id) {
  return withLock_(function () {
    var s = sheet_("MENU");
    var row = cariBarisMenu_(s, id);
    s.deleteRow(row);
    hapusResepRows_(id);
    return true;
  });
}

function hapusPengeluaran(id) {
  return withLock_(function () {
    var s = sheet_("PENGELUARAN");
    var vals = s.getDataRange().getValues();
    for (var i = 1; i < vals.length; i++) {
      if (vals[i][0] === id) {
        var tgl = normDate_(vals[i][1]), cab = cabOf_(vals[i], CI.KELUAR);
        if (tutup_(tgl, cab)) throw new Error("Buku " + cab + " tanggal " + tgl + " sudah ditutup. Pengeluaran tidak bisa dihapus.");
        s.deleteRow(i + 1);
        return true;
      }
    }
    throw new Error("Pengeluaran tidak ditemukan");
  });
}

function hapusPenyesuaian(id) {
  return withLock_(function () {
    var s = sheet_("PENYESUAIAN");
    var vals = s.getDataRange().getValues();
    for (var i = 1; i < vals.length; i++) {
      if (vals[i][0] !== id) continue;
      var gudang = vals[i][8] === "Gudang", cab = cabOf_(vals[i], CI.ADJ);
      var t = normDate_(vals[i][1]);
      if (gudang ? tutup_(t, "") : tutup_(t, cab))
        throw new Error("Buku " + (gudang ? "gudang" : cab) + " tanggal " + t + " sudah ditutup. Penyesuaian tidak bisa dihapus.");
      s.deleteRow(i + 1);
      return true;
    }
    throw new Error("Penyesuaian tidak ditemukan");
  });
}

/** data: {id, tanggal, jumlah, metodeBayar (Kas), keterangan}. Jenis, barang, lokasi, cabang tidak diubah. */
function updatePenyesuaian(data) {
  return withLock_(function () {
    var jumlah = Number(data.jumlah);
    if (!jumlah) throw new Error("Jumlah penyesuaian tidak boleh 0");
    var s = sheet_("PENYESUAIAN");
    var vals = s.getDataRange().getValues();
    for (var i = 1; i < vals.length; i++) {
      var r = vals[i];
      if (r[0] !== data.id) continue;
      var gudang = r[8] === "Gudang", cab = cabOf_(r, CI.ADJ);
      var lama = normDate_(r[1]), baru = tgl_(data.tanggal);
      [lama, baru].forEach(function (t) {
        if (gudang ? tutup_(t, "") : tutup_(t, cab))
          throw new Error("Buku " + (gudang ? "gudang" : cab) + " tanggal " + t + " sudah ditutup. Penyesuaian tidak bisa diubah.");
      });
      var metode = r[2] === "Kas" ? metode_(data.metodeBayar) : "";
      s.getRange(i + 1, 2, 1, 1).setValue(baru);
      s.getRange(i + 1, 6, 1, 3).setValues([[jumlah, metode, data.keterangan || ""]]);
      return true;
    }
    throw new Error("Penyesuaian tidak ditemukan");
  });
}

/** data: {id, tanggal, kategori, jenis, keterangan, jumlah, metodeBayar}. Cabang tidak diubah. */
function updatePengeluaran(data) {
  return withLock_(function () {
    if (!(Number(data.jumlah) > 0)) throw new Error("Biaya harus lebih dari 0");
    var s = sheet_("PENGELUARAN");
    var vals = s.getDataRange().getValues();
    for (var i = 1; i < vals.length; i++) {
      if (vals[i][0] !== data.id) continue;
      var cab = cabOf_(vals[i], CI.KELUAR);
      var lama = normDate_(vals[i][1]), baru = tgl_(data.tanggal);
      [lama, baru].forEach(function (t) {
        if (tutup_(t, cab))
          throw new Error("Buku " + cab + " tanggal " + t + " sudah ditutup. Pengeluaran tidak bisa diubah.");
      });
      var jenis = getJenisPengeluaran().indexOf(data.jenis) >= 0 ? data.jenis : "Operasional";
      s.getRange(i + 1, 2, 1, 6).setValues([[baru, data.kategori || "", jenis,
        data.keterangan || "", Number(data.jumlah), metode_(data.metodeBayar)]]);
      return true;
    }
    throw new Error("Pengeluaran tidak ditemukan");
  });
}

function getPenyesuaian(bulanTahun, cab) {
  return filterBulan_(readRows_("PENYESUAIAN"), bulanTahun)
    .filter(function (r) { return !cab || r[8] === "Gudang" || cabOf_(r, CI.ADJ) === cab; })
    .map(function (r) { return fmtRow_(r, 1); });
}