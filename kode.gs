// =========================================================================
// GOOGLE APPS SCRIPT (Kode.gs) - DATABASE & WEB APP BACKEND (MULTI-STORE)
// =========================================================================

/**
 * 1. FUNGSI UTAMA DOGET: Memuat Tampilan HTML atau Merespon API GET
 */
function doGet(e) {
  try {
    const action = (e && e.parameter) ? e.parameter.action : '';
    const callback = (e && e.parameter) ? e.parameter.callback : '';
    
    // Jika dipanggil via API endpoint untuk mengambil data JSON (GET)
    if (action === 'getAllData' || action === 'GET_ALL') {
      const data = getAllDataFromSheets();
      if (callback) {
        return ContentService.createTextOutput(`${callback}(${JSON.stringify({ status: 'success', data: data })})`)
          .setMimeType(ContentService.MimeType.JAVASCRIPT);
      }
      return createJsonResponse({ status: 'success', data: data });
    }
    
    // Tampilan UTAMA: Membuka Aplikasi Web Kasir
    const htmlOutput = HtmlService.createHtmlOutputFromFile('index')
      .setTitle('LAPORAN PENJUALAN - POS & Manajemen Toko')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
      
    return htmlOutput;
  } catch (error) {
    return HtmlService.createHtmlOutput(
      `<div style="font-family:sans-serif; padding:30px; text-align:center; color:#1e293b;">
        <h2 style="color:#e11d48;">⚠️ File index.html Belum Ditemukan</h2>
        <p>Silakan buat file HTML dengan nama <b>index</b> (menjadi index.html) di panel sebelah kiri Google Apps Script dan salin seluruh skrip frontend ke dalamnya.</p>
        <p style="color:#64748b; font-size:12px;">Detail Error: ${error.toString()}</p>
      </div>`
    );
  }
}

/**
 * 2. FUNGSI DOPOST: Menerima Sinkronisasi Data dari Aplikasi (POST)
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
  } catch (errLock) {
    return createJsonResponse({ status: 'error', message: 'Server sibuk, silakan coba beberapa saat lagi.' });
  }

  try {
    let requestData = {};
    if (e && e.postData && e.postData.contents) {
      try {
        requestData = JSON.parse(e.postData.contents);
      } catch (errParse) {
        requestData = {};
      }
    }
    
    const result = processApiAction(requestData);
    return createJsonResponse(result);
  } catch (err) {
    return createJsonResponse({ status: 'error', message: err.toString() });
  } finally {
    lock.releaseLock();
  }
}

/**
 * 3. PROCESS API ACTION: Memproses Setiap Permintaan (CRUD & SYNC)
 */
function processApiAction(requestData) {
  const action = requestData ? requestData.action : '';
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  ensureDatabaseStructure(ss);

  if (action === 'SYNC_ALL') {
    if (requestData.products) saveProductsToSheet(ss, requestData.products);
    if (requestData.transactions) saveTransactionsToSheet(ss, requestData.transactions);
    if (requestData.profile) saveProfileToSheet(ss, requestData.profile, requestData.userPIN);
    return { status: 'success', message: 'Seluruh data tersimpan di Google Sheets!', data: getAllDataFromSheets() };
  }

  if (action === 'ADD_TRANSACTION') {
    if (requestData.transaction) {
      addTransactionToSheet(ss, requestData.transaction);
      if (requestData.transaction.items && requestData.transaction.items.length > 0) {
        updateProductStocksInSheet(ss, requestData.transaction.items);
      }
    }
    return { status: 'success', message: 'Transaksi berhasil disimpan!', data: getAllDataFromSheets() };
  }

  if (action === 'SAVE_PRODUCT') {
    if (requestData.product) {
      saveSingleProductToSheet(ss, requestData.product);
    }
    return { status: 'success', message: 'Data produk tersimpan!', data: getAllDataFromSheets() };
  }

  if (action === 'DELETE_PRODUCT') {
    if (requestData.productId) {
      deleteProductFromSheet(ss, requestData.productId);
    }
    return { status: 'success', message: 'Produk berhasil dihapus!', data: getAllDataFromSheets() };
  }

  if (action === 'UPDATE_PROFILE') {
    if (requestData.profile) {
      saveProfileToSheet(ss, requestData.profile, requestData.userPIN);
    }
    return { status: 'success', message: 'Profil berhasil diperbarui!', data: getAllDataFromSheets() };
  }

  if (action === 'SAVE_SHIFT_CLOSING') {
    if (requestData.shiftClosing) {
      addShiftClosingToSheet(ss, requestData.shiftClosing);
    }
    return { status: 'success', message: 'Rekap tutup kasir tersimpan!', data: getAllDataFromSheets() };
  }

  if (action === 'ADD_EXPENSE') {
    if (requestData.expense) {
      addExpenseToSheet(ss, requestData.expense);
    }
    return { status: 'success', message: 'Pengeluaran operasional tersimpan!', data: getAllDataFromSheets() };
  }

  if (action === 'DELETE_EXPENSE') {
    if (requestData.expenseId) {
      deleteExpenseFromSheet(ss, requestData.expenseId);
    }
    return { status: 'success', message: 'Pengeluaran berhasil dihapus!', data: getAllDataFromSheets() };
  }

  if (action === 'GET_ALL') {
    return { status: 'success', data: getAllDataFromSheets() };
  }

  return { status: 'error', message: 'Action tidak dikenal' };
}

/**
 * 4. FUNGSI PEMBENTUK STRUKTUR DATABASE OTOMATIS
 */
function ensureDatabaseStructure(ss) {
  const sheetConfigs = [
    {
      name: 'Produk',
      headers: ['ID', 'SKU', 'Nama Produk', 'Kategori', 'Harga Beli (HPP)', 'Harga Jual', 'Stok', 'Satuan', 'Waktu Update']
    },
    {
      name: 'Transaksi',
      headers: ['No Struk / ID', 'Waktu', 'Pelanggan', 'Metode Pembayaran', 'Detail Items (JSON)', 'Total Omset', 'Total HPP', 'Total Keuntungan', 'Uang Diterima', 'Kembalian']
    },
    {
      name: 'Profil',
      headers: ['Parameter', 'Nilai']
    },
    {
      name: 'TutupKasir',
      headers: ['ID Shift', 'Waktu', 'Kasir', 'Jumlah Trx', 'Modal Awal', 'Omset Tunai', 'Omset QRIS', 'Omset Transfer', 'Total Omset', 'Ekspektasi Laci', 'Fisik Laci', 'Selisih', 'Status', 'Catatan']
    },
    {
      name: 'Pengeluaran',
      headers: ['ID Pengeluaran', 'Waktu', 'Kategori', 'Keterangan', 'Nominal', 'Metode Bayar', 'ISO Date']
    }
  ];

  sheetConfigs.forEach(cfg => {
    let sheet = ss.getSheetByName(cfg.name);
    if (!sheet) {
      sheet = ss.insertSheet(cfg.name);
      sheet.getRange(1, 1, 1, cfg.headers.length).setValues([cfg.headers]);
      sheet.getRange(1, 1, 1, cfg.headers.length).setFontWeight("bold").setBackground("#0d9488").setFontColor("#ffffff");
    }
  });
}

function setupDatabaseSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureDatabaseStructure(ss);
  Logger.log("Database LAPORAN PENJUALAN berhasil diinisialisasi!");
}

/**
 * 5. FUNGSI PEMBACAAN SELURUH DATA DARI SHEET (GET ALL)
 */
function getAllDataFromSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureDatabaseStructure(ss);

  // A. Produk
  const products = [];
  const sheetProduk = ss.getSheetByName('Produk');
  if (sheetProduk && sheetProduk.getLastRow() > 1) {
    const rows = sheetProduk.getRange(2, 1, sheetProduk.getLastRow() - 1, 8).getValues();
    rows.forEach(r => {
      if (r[0] !== "" && r[0] !== null && r[0] !== undefined) {
        products.push({
          id: Number(r[0]),
          sku: String(r[1] || ''),
          name: String(r[2] || ''),
          category: String(r[3] || 'Umum'),
          buyPrice: Number(r[4] || 0),
          sellPrice: Number(r[5] || 0),
          stock: Number(r[6] || 0),
          unit: String(r[7] || 'Pcs')
        });
      }
    });
  }

  // B. Transaksi
  const transactions = [];
  let sheetTx = ss.getSheetByName('Transaksi');
  if (!sheetTx) {
    const sheets = ss.getSheets();
    for (let s of sheets) {
      if (s.getName().toLowerCase().includes('transaksi')) {
        sheetTx = s;
        break;
      }
    }
  }

  if (sheetTx && sheetTx.getLastRow() > 1) {
    const rows = sheetTx.getRange(2, 1, sheetTx.getLastRow() - 1, 10).getValues();
    rows.forEach(r => {
      const idStr = String(r[0] || '').trim();
      if (idStr !== "" && idStr.toLowerCase() !== "no struk / id" && idStr.toLowerCase() !== "no") {
        let items = [];
        try {
          items = typeof r[4] === 'string' ? JSON.parse(r[4]) : (Array.isArray(r[4]) ? r[4] : []);
        } catch (e) {
          items = [];
        }

        let timeVal = String(r[1] || '');
        let isoVal = '';
        if (r[1] instanceof Date) {
          try {
            isoVal = r[1].toISOString();
            timeVal = Utilities.formatDate(r[1], ss.getSpreadsheetTimeZone() || "GMT+7", "dd/MM/yyyy, HH.mm");
          } catch (eDate) {
            timeVal = String(r[1]);
          }
        }

        transactions.push({
          id: idStr,
          time: timeVal,
          rawDate: isoVal || timeVal,
          customer: String(r[2] || 'Pelanggan Umum'),
          paymentMethod: String(r[3] || 'Tunai'),
          items: items,
          totalAmount: Number(r[5] || 0),
          totalCost: Number(r[6] || 0),
          totalProfit: Number(r[7] || 0),
          cashReceived: Number(r[8] || 0),
          changeAmount: Number(r[9] || 0)
        });
      }
    });
  }

  // C. Profil Toko
  const profile = {
    storeName: "LAPORAN PENJUALAN",
    storeSubtitle: "Kasir & Manajemen Toko",
    ownerName: "Pengelola Toko",
    phone: "0812-3456-7890",
    address: "Jl. Raya Utama, Sukabumi",
    operatingHours: "08:00 - 21:00 WIB",
    receiptFooter: "Terima kasih atas kunjungan Anda!",
    photo: "https://images.unsplash.com/photo-1581092160607-ee22621dd758?w=1200&auto=format&fit=crop&q=80"
  };
  let userPIN = "1234";

  const sheetProfil = ss.getSheetByName('Profil');
  if (sheetProfil && sheetProfil.getLastRow() > 1) {
    const rows = sheetProfil.getRange(2, 1, sheetProfil.getLastRow() - 1, 2).getValues();
    rows.forEach(r => {
      const key = String(r[0]).trim();
      const val = String(r[1]).trim();
      if (key === 'Nama Toko' && val) profile.storeName = val;
      if (key === 'Sub-Judul Toko' && val) profile.storeSubtitle = val;
      if (key === 'Nama Pemilik' && val) profile.ownerName = val;
      if (key === 'Telepon' && val) profile.phone = val.replace(/^'/, '');
      if (key === 'Alamat' && val) profile.address = val;
      if (key === 'Jam Operasional' && val) profile.operatingHours = val;
      if (key === 'Catatan Struk' && val) profile.receiptFooter = val;
      if (key === 'Foto Banner URL' && val) profile.photo = val;
      if (key === 'Batas Minimal Stok' && val !== '') profile.lowStockThreshold = Math.max(0, Number(val) || 5);
      if (key === 'PIN Akses' && val) {
        let parsedPin = String(val).replace(/^'/, '').trim().split('.')[0];
        if (/^\d+$/.test(parsedPin) && parsedPin.length < 4) {
          parsedPin = parsedPin.padStart(4, '0');
        }
        userPIN = parsedPin;
      }
    });
  }

  // D. Pengeluaran Operasional
  const expenses = [];
  const sheetExp = ss.getSheetByName('Pengeluaran');
  if (sheetExp && sheetExp.getLastRow() > 1) {
    const rows = sheetExp.getRange(2, 1, sheetExp.getLastRow() - 1, 7).getValues();
    rows.forEach(r => {
      const idStr = String(r[0] || '').trim();
      if (idStr !== "" && idStr.toLowerCase() !== "id pengeluaran") {
        let timeVal = String(r[1] || '');
        let isoVal = String(r[6] || '');
        if (r[1] instanceof Date) {
          try {
            if (!isoVal) isoVal = r[1].toISOString();
            timeVal = Utilities.formatDate(r[1], ss.getSpreadsheetTimeZone() || "GMT+7", "dd/MM/yyyy, HH.mm");
          } catch (eDate) {
            timeVal = String(r[1]);
          }
        }
        expenses.push({
          id: idStr,
          time: timeVal,
          rawDate: isoVal || timeVal,
          category: String(r[2] || 'Operasional'),
          description: String(r[3] || '-'),
          amount: Number(r[4] || 0),
          paymentMethod: String(r[5] || 'Tunai')
        });
      }
    });
  }

  // E. Rekap Tutup Kasir Harian
  const shiftClosings = [];
  const sheetShift = ss.getSheetByName('TutupKasir');
  if (sheetShift && sheetShift.getLastRow() > 1) {
    const rows = sheetShift.getRange(2, 1, sheetShift.getLastRow() - 1, 14).getValues();
    rows.forEach(r => {
      const idStr = String(r[0] || '').trim();
      if (idStr !== "" && idStr.toLowerCase() !== "id shift") {
        let timeVal = String(r[1] || '');
        if (r[1] instanceof Date) {
          try {
            timeVal = Utilities.formatDate(r[1], ss.getSpreadsheetTimeZone() || "GMT+7", "dd/MM/yyyy, HH.mm");
          } catch (eDate) {
            timeVal = String(r[1]);
          }
        }
        shiftClosings.push({
          id: idStr,
          time: timeVal,
          cashierName: String(r[2] || 'Kasir Utama'),
          txCount: Number(r[3] || 0),
          openingCash: Number(r[4] || 0),
          cashSales: Number(r[5] || 0),
          qrisSales: Number(r[6] || 0),
          transferSales: Number(r[7] || 0),
          totalGross: Number(r[8] || 0),
          expectedCash: Number(r[9] || 0),
          actualCash: Number(r[10] || 0),
          variance: Number(r[11] || 0),
          status: String(r[12] || 'Klop'),
          notes: String(r[13] || '-')
        });
      }
    });
  }

  return {
    status: 'success',
    products: products,
    transactions: transactions,
    profile: profile,
    userPIN: userPIN,
    expenses: expenses,
    shiftClosings: shiftClosings
  };
}

function addExpenseToSheet(ss, exp) {
  let sheet = ss.getSheetByName('Pengeluaran');
  if (!sheet) {
    ensureDatabaseStructure(ss);
    sheet = ss.getSheetByName('Pengeluaran');
  }
  if (!sheet) return;

  const rowData = [
    String(exp.id || ''),
    String(exp.time || ''),
    String(exp.category || 'Operasional'),
    String(exp.description || '-'),
    Number(exp.amount || 0),
    String(exp.paymentMethod || 'Tunai'),
    String(exp.rawDate || new Date().toISOString())
  ];

  if (sheet.getLastRow() <= 1) {
    sheet.appendRow(rowData);
  } else {
    sheet.insertRowBefore(2);
    const range = sheet.getRange(2, 1, 1, 7);
    range.setValues([rowData]);
    range.setBackground(null).setFontWeight("normal").setFontColor("#000000");
  }
}

function deleteExpenseFromSheet(ss, expenseId) {
  const sheet = ss.getSheetByName('Pengeluaran');
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(expenseId)) {
        sheet.deleteRow(i + 2);
        break;
      }
    }
  }
}

function addShiftClosingToSheet(ss, sc) {
  let sheet = ss.getSheetByName('TutupKasir');
  if (!sheet) {
    ensureDatabaseStructure(ss);
    sheet = ss.getSheetByName('TutupKasir');
  }
  if (!sheet) return;

  const rowData = [
    String(sc.id || ''),
    String(sc.time || ''),
    String(sc.cashierName || 'Kasir Utama'),
    Number(sc.txCount || 0),
    Number(sc.openingCash || 0),
    Number(sc.cashSales || 0),
    Number(sc.qrisSales || 0),
    Number(sc.transferSales || 0),
    Number(sc.totalGross || 0),
    Number(sc.expectedCash || 0),
    Number(sc.actualCash || 0),
    Number(sc.variance || 0),
    String(sc.status || 'Klop'),
    String(sc.notes || '-')
  ];

  if (sheet.getLastRow() <= 1) {
    sheet.appendRow(rowData);
  } else {
    sheet.insertRowBefore(2);
    const range = sheet.getRange(2, 1, 1, 14);
    range.setValues([rowData]);
    range.setBackground(null).setFontWeight("normal").setFontColor("#000000");
  }
}

// ==================== OPERASI PENULISAN DATABASE ====================
function saveProductsToSheet(ss, products) {
  let sheet = ss.getSheetByName('Produk');
  if (!sheet) return;
  
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).clearContent();
  }

  if (products && products.length > 0) {
    const nowStr = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone() || "GMT+7", "dd/MM/yyyy, HH.mm");
    const rows = products.map(p => [
      p.id,
      p.sku || '',
      p.name || '',
      p.category || 'Umum',
      p.buyPrice || 0,
      p.sellPrice || 0,
      p.stock || 0,
      p.unit || 'Pcs',
      nowStr
    ]);
    sheet.getRange(2, 1, rows.length, 9).setValues(rows);
  }
}

function saveSingleProductToSheet(ss, product) {
  let sheet = ss.getSheetByName('Produk');
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  const nowStr = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone() || "GMT+7", "dd/MM/yyyy, HH.mm");
  let found = false;

  if (lastRow > 1) {
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(product.id)) {
        sheet.getRange(i + 2, 1, 1, 9).setValues([[
          product.id,
          product.sku || '',
          product.name || '',
          product.category || 'Umum',
          product.buyPrice || 0,
          product.sellPrice || 0,
          product.stock || 0,
          product.unit || 'Pcs',
          nowStr
        ]]);
        found = true;
        break;
      }
    }
  }

  if (!found) {
    sheet.appendRow([
      product.id,
      product.sku || '',
      product.name || '',
      product.category || 'Umum',
      product.buyPrice || 0,
      product.sellPrice || 0,
      product.stock || 0,
      product.unit || 'Pcs',
      nowStr
    ]);
  }
}

function saveTransactionsToSheet(ss, transactions) {
  let sheet = ss.getSheetByName('Transaksi');
  if (!sheet) return;
  if (!transactions || transactions.length === 0) return;

  const lastRow = sheet.getLastRow();
  const existingIds = new Set();
  if (lastRow > 1) {
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    ids.forEach(r => { 
      if (r[0] !== "" && r[0] !== null && r[0] !== undefined) {
        existingIds.add(String(r[0])); 
      }
    });
  }

  const nowStr = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone() || "GMT+7", "dd/MM/yyyy, HH.mm");
  const newRows = [];
  transactions.forEach(tx => {
    if (tx && tx.id && !existingIds.has(String(tx.id))) {
      newRows.push([
        String(tx.id),
        tx.time || nowStr,
        tx.customer || 'Pelanggan Umum',
        tx.paymentMethod || 'Tunai',
        JSON.stringify(tx.items || []),
        tx.totalAmount || 0,
        tx.totalCost || 0,
        tx.totalProfit || 0,
        tx.cashReceived || 0,
        tx.changeAmount || 0
      ]);
    }
  });

  if (newRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, 10).setValues(newRows);
  }
}

function deleteProductFromSheet(ss, productId) {
  const sheet = ss.getSheetByName('Produk');
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(productId)) {
        sheet.deleteRow(i + 2);
        break;
      }
    }
  }
}

function addTransactionToSheet(ss, tx) {
  let sheet = ss.getSheetByName('Transaksi');
  if (!sheet) return;

  const nowStr = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone() || "GMT+7", "dd/MM/yyyy, HH.mm");
  const rowData = [
    String(tx.id),
    tx.time || nowStr,
    tx.customer || 'Pelanggan Umum',
    tx.paymentMethod || 'Tunai',
    JSON.stringify(tx.items || []),
    tx.totalAmount || 0,
    tx.totalCost || 0,
    tx.totalProfit || 0,
    tx.cashReceived || 0,
    tx.changeAmount || 0
  ];

  if (sheet.getLastRow() <= 1) {
    sheet.appendRow(rowData);
  } else {
    sheet.insertRowBefore(2);
    const range = sheet.getRange(2, 1, 1, 10);
    range.setValues([rowData]);
    range.setBackground(null).setFontWeight("normal").setFontColor("#000000");
  }
}

function updateProductStocksInSheet(ss, items) {
  const sheet = ss.getSheetByName('Produk');
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  const dataRange = sheet.getRange(2, 1, lastRow - 1, 7);
  const rows = dataRange.getValues();

  items.forEach(item => {
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][0]) === String(item.id)) {
        let currentStock = Number(rows[i][6] || 0);
        let newStock = Math.max(0, currentStock - item.qty);
        sheet.getRange(i + 2, 7).setValue(newStock);
        break;
      }
    }
  });
}

function saveProfileToSheet(ss, profile, userPIN) {
  let sheet = ss.getSheetByName('Profil');
  if (!sheet) return;

  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).clearContent();
  }

  let safePhotoUrl = profile.photo || '';
  if (safePhotoUrl.length > 48000) {
    safePhotoUrl = "https://images.unsplash.com/photo-1581092160607-ee22621dd758?w=1200&auto=format&fit=crop&q=80";
  }

  const cleanPin = String(userPIN || '1234').replace(/^'/, '').trim().split('.')[0];

  const data = [
    ['Nama Toko', profile.storeName || 'LAPORAN PENJUALAN'],
    ['Sub-Judul Toko', profile.storeSubtitle || 'Kasir & Manajemen Toko'],
    ['Nama Pemilik', profile.ownerName || 'Pengelola Toko'],
    ['Telepon', "'" + String(profile.phone || '0812-3456-7890').replace(/^'/, '')],
    ['Alamat', profile.address || 'Alamat Toko'],
    ['Jam Operasional', profile.operatingHours || '08:00 - 21:00 WIB'],
    ['Catatan Struk', profile.receiptFooter || 'Terima kasih atas kunjungan Anda!'],
    ['PIN Akses', "'" + cleanPin],
    ['Foto Banner URL', safePhotoUrl],
    ['Batas Minimal Stok', Number(profile.lowStockThreshold !== undefined ? profile.lowStockThreshold : 5)]
  ];

  sheet.getRange(2, 1, data.length, 2).setValues(data);
}

function createJsonResponse(dataObject) {
  return ContentService.createTextOutput(JSON.stringify(dataObject))
    .setMimeType(ContentService.MimeType.JSON);
}
