function doGet(e) {
  return HtmlService
    .createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Apotek Indocement')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService
    .createHtmlOutputFromFile(filename)
    .getContent();
}

// 1. AUTENTIKASI & HELPER EMAIL
function prosesLogin(username, password) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  if(!sheet) return { success: false, message: "Sheet 'Users' tidak ditemukan!" };
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] && data[i][1].toString().trim() === username.trim() && data[i][2] && data[i][2].toString().trim() === password.trim()) {
      return { success: true, nama: data[i][3] || username, role: data[i][5] || "Staff", email: data[i][4] || "" };
    }
  }
  return { success: false, message: "Username atau Password salah." };
}

function getSemuaEmailKaryawan() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  if(!sheet) return [];
  const data = sheet.getDataRange().getValues();
  let emails = [];
  for(let i=1; i<data.length; i++) { if(data[i][4]) emails.push(data[i][4].toString().trim()); }
  return emails.join(","); // Gabungkan semua email
}

// 2. AGREGASI & SORTING
function getAppData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetStok = ss.getSheetByName('Stok_Master');
  const sheetMutasi = ss.getSheetByName('Mutasi_Log');
  let sheetTransaksi = ss.getSheetByName('Transaksi_Apotek');
  
  // Auto create sheet Transaksi jika belum ada
  if(!sheetTransaksi) {
    sheetTransaksi = ss.insertSheet('Transaksi_Apotek');
    sheetTransaksi.appendRow(["ID Transaksi", "Tanggal", "Rincian Obat", "Total Item", "Kasir"]);
  }

  const dataStok = sheetStok ? sheetStok.getDataRange().getDisplayValues() : [];
  const dataMutasi = sheetMutasi ? sheetMutasi.getDataRange().getDisplayValues() : [];
  const dataTrans = sheetTransaksi.getDataRange().getDisplayValues();
  
  let stokList = [], mutasiList = [], transList = [];
  let stats = { aman: 0, bln6: 0, bln3: 0, bln1: 0, expired: 0, totalGudang: 0, totalApotek: 0, pendingMutasi: 0 };

  if (dataStok.length  > 1) {
    for (let i = 1; i < dataStok.length; i++) {
      let id = dataStok[i][0], nama = dataStok[i][1], batch = dataStok[i][2];
      let sGudang = parseInt(dataStok[i][3]) || 0, sApotek = parseInt(dataStok[i][4]) || 0;
      let satuan = dataStok[i][5], tglExpStr = dataStok[i][6];
      
      stats.totalGudang += sGudang; stats.totalApotek += sApotek;
      let diffMonths = 999, diffDays = 999;
      let badgeColor = "bg-blue", statusLabel = "Aman";

      if (tglExpStr) {
        let parts = tglExpStr.split(/[-/]/);
        if (parts.length === 3) {
          let expDate = parts[0].length === 4 ? new Date(parts[0], parts[1] - 1, parts[2]) : new Date(parts[2], parts[1] - 1, parts[0]);
          diffDays = Math.ceil((expDate - new Date()) / (1000 * 60 * 60 * 24));
          diffMonths = diffDays / 30.44;
        }
        if (diffDays <= 0) { statusLabel = "0 (Expired)"; badgeColor = "bg-black"; stats.expired++; }
        else if (diffMonths <= 1) { statusLabel = "1 Bulan"; badgeColor = "bg-red"; stats.bln1++; }
        else if (diffMonths <= 3) { statusLabel = "3 Bulan"; badgeColor = "bg-yellow"; stats.bln3++; }
        else if (diffMonths <= 6) { statusLabel = "6 Bulan"; badgeColor = "bg-green"; stats.bln6++; }
        else { statusLabel = "Aman"; badgeColor = "bg-blue"; stats.aman++; }
      }
      stokList.push({ id, nama, batch, stokGudang: sGudang, stokApotek: sApotek, satuan, tglExp: tglExpStr, status: statusLabel, color: badgeColor, sisaHari: diffDays });
    }
  }
  stokList.sort((a, b) => a.sisaHari - b.sisaHari);

  if (dataMutasi.length > 1) {
    for (let i = dataMutasi.length - 1; i > 0; i--) {
      if (dataMutasi[i][7] === "Pending") stats.pendingMutasi++;
      let match = stokList.find(s => s.id === dataMutasi[i][3]);
      mutasiList.push({
        idMutasi: dataMutasi[i][0], tglKirim: dataMutasi[i][1], tglTerima: dataMutasi[i][2],
        idObat: dataMutasi[i][3], namaObat: match ? match.nama : "Unknown", jumlah: dataMutasi[i][4],
        pengirim: dataMutasi[i][5], penerima: dataMutasi[i][6], status: dataMutasi[i][7]
      });
    }
  }

  if (dataTrans.length > 1) {
    for (let i = dataTrans.length - 1; i > 0; i--) {
      transList.push({ idTrans: dataTrans[i][0], tanggal: dataTrans[i][1], rincian: dataTrans[i][2], totalQty: dataTrans[i][3], kasir: dataTrans[i][4] });
    }
  }

  return { stok: stokList, mutasi: mutasiList, transaksi: transList, stats: stats };
}

// 3. MASTER STOK (DENGAN ALERT EMAIL JIKA < 6 BULAN)
function inputBarangSupplier(namaObat, batch, qty, tglExp, satuan) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Stok_Master');
  const idBaru = "OBT-" + Utilities.formatDate(new Date(), "GMT+7", "mmssSSS").slice(-5);
  let expDate = new Date(tglExp);
  let formattedDate = Utilities.formatDate(expDate, "GMT+7", "dd/MM/yyyy");
  
  sheet.appendRow([idBaru, namaObat, batch, parseInt(qty), 0, satuan, formattedDate]);

  // Cek apakah expired < 6 Bulan, langsung kirim email
  let diffDays = Math.ceil((expDate - new Date()) / (1000 * 60 * 60 * 24));
  let diffMonths = diffDays / 30.44;
  
  if(diffMonths <= 6) {
    let targetEmails = getSemuaEmailKaryawan();
    if(targetEmails) {
      let subj = `[PERINGATAN] Stok Obat Baru Hampir Kedaluwarsa (${namaObat})`;
      let body = `Sistem mencatat input obat baru dari supplier dengan masa kedaluwarsa Kritis (<= 6 Bulan).\n\nRincian:\n- Nama Obat: ${namaObat}\n- Batch: ${batch}\n- Qty Masuk: ${qty}\n- Tanggal Expired: ${formattedDate} (Sisa ${diffDays} Hari)\n\nHarap jadikan prioritas distribusi.`;
      MailApp.sendEmail(targetEmails, subj, body);
    }
  }

  return { success: true, message: "Sukses menginput stok! Jika expired < 6 bulan, email peringatan otomatis terkirim." };
}

// [Fungsi Edit & Hapus Master tetap sama, diringkas di sini]
function updateBarang(id, n, b, q, t) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Stok_Master');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.getRange(i+1, 2).setValue(n); 
      sheet.getRange(i+1, 3).setValue(b);
      sheet.getRange(i+1, 4).setValue(parseInt(q));
      
      // === SISTEM VALIDASI TANGGAL AMAN ===
      if (t) {
        let dateObj = null;
        let strDate = t.toString().trim();
        
        if (strDate.includes("-")) {
          const parts = strDate.split("-");
          if (parts[0].length === 4) {
            // Format YYYY-MM-DD (dari input date HTML)
            dateObj = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
          } else {
            // Format DD-MM-YYYY
            dateObj = new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
          }
        } else if (strDate.includes("/")) {
          // Format DD/MM/YYYY (dari tampilan tabel sheet)
          const parts = strDate.split("/");
          dateObj = new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
        } else {
          dateObj = new Date(strDate);
        }
        
        // Tulis ke sel jika konversi tanggal berhasil & valid
        const cellTanggal = sheet.getRange(i+1, 7);
        if (dateObj && !isNaN(dateObj.getTime())) {
          cellTanggal.setValue(dateObj);
          cellTanggal.setNumberFormat("dd/MM/yyyy");
        } else {
          cellTanggal.setValue(t);
        }
      }
      // ====================================
      
      return { success: true, message: "Data obat diperbarui!" };
    }
  }
}

// 4. KIRIM & TERIMA MUTASI [Tetap sama logicnya]
function kirimMutasiMassal(cartItems, pengirim) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetStok = ss.getSheetByName('Stok_Master');
  const sheetMutasi = ss.getSheetByName('Mutasi_Log');
  const dataStok = sheetStok.getDataRange().getValues();
  const tglKirim = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd");
  let berhasilCount = 0; let baseId = Utilities.formatDate(new Date(), "GMT+7", "HHmmss");

  cartItems.forEach((item, idx) => {
    let r = dataStok.findIndex(row => row[0] === item.id);
    if (r > 0) {
      let sLama = parseInt(dataStok[r][3]) || 0; let q = parseInt(item.qty);
      if (sLama >= q) {
        dataStok[r][3] = sLama - q; sheetStok.getRange(r + 1, 4).setValue(sLama - q);
        sheetMutasi.appendRow(["MUT-" + baseId + idx, tglKirim, "Kosong", item.id, q, pengirim, "", "Pending"]);
        berhasilCount++;
      }
    }
  });
  return berhasilCount > 0 ? { success: true, message: "Dikirim ke antrean Apotek!" } : { success: false, message: "Gagal memproses." };
}

function terimaMutasi(idMutasi, penerima) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetStok = ss.getSheetByName('Stok_Master'), sheetMutasi = ss.getSheetByName('Mutasi_Log');
  const dataMutasi = sheetMutasi.getDataRange().getValues();
  for (let i = 1; i < dataMutasi.length; i++) {
    if (dataMutasi[i][0] === idMutasi && dataMutasi[i][7] === "Pending") {
      sheetMutasi.getRange(i+1, 3).setValue(Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd"));
      sheetMutasi.getRange(i+1, 7).setValue(penerima); sheetMutasi.getRange(i+1, 8).setValue("Diterima");
      const dataStok = sheetStok.getDataRange().getValues();
      for (let j = 1; j < dataStok.length; j++) {
        if (dataStok[j][0] === dataMutasi[i][3]) {
          sheetStok.getRange(j+1, 5).setValue((parseInt(dataStok[j][4])||0) + parseInt(dataMutasi[i][4]));
          return { success: true, message: "Diterima! Stok Klinik bertambah." };
        }
      }
    }
  }
}
function hapusMutasi(id) { /* Kode sebelumnya untuk batalkan mutasi (dijaga) */
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sMutasi = ss.getSheetByName('Mutasi_Log'), sStok = ss.getSheetByName('Stok_Master');
  const dMut = sMutasi.getDataRange().getValues();
  for(let i = dMut.length - 1; i >= 1; i--) {
    if(dMut[i][0] === id) {
      if(dMut[i][7] === "Pending") {
         let dStok = sStok.getDataRange().getValues();
         for(let j = 1; j < dStok.length; j++) {
            if(dStok[j][0] === dMut[i][3]) { sStok.getRange(j+1, 4).setValue((parseInt(dStok[j][3])||0) + parseInt(dMut[i][4])); break; }
         }
      }
      sMutasi.deleteRow(i+1); return { success: true, message: "Mutasi dibatalkan/dihapus." };
    }
  }
}

// 5. TRANSAKSI APOTEK (BARU)
function submitTransaksiApotek(cartItems, kasir) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetStok = ss.getSheetByName('Stok_Master');
  const sheetTrans = ss.getSheetByName('Transaksi_Apotek');
  const dataStok = sheetStok.getDataRange().getValues();
  
  let idTrans = "TRX-" + Utilities.formatDate(new Date(), "GMT+7", "ddMMyyHHmmss");
  let tgl = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm");
  
  let rincianArr = [];
  let totalQty = 0;

  cartItems.forEach(item => {
    let r = dataStok.findIndex(row => row[0] === item.id);
    if(r > 0) {
      let sApotekLama = parseInt(dataStok[r][4]) || 0;
      let q = parseInt(item.qty);
      if(sApotekLama >= q) {
        // Kurangi Stok Apotek
        sheetStok.getRange(r + 1, 5).setValue(sApotekLama - q);
        rincianArr.push(`${item.nama} (${q})`);
        totalQty += q;
      }
    }
  });

  if(rincianArr.length > 0) {
    sheetTrans.appendRow([idTrans, tgl, rincianArr.join(", "), totalQty, kasir]);
    return { success: true, message: "Transaksi berhasil dicatat & Stok Klinik telah terpotong!" };
  }
  return { success: false, message: "Gagal memproses transaksi." };
}

// 6. CRON JOB: CEK EMAIL HARIAN (EXP 1, 3, 6 BULAN)
// CATATAN: Atur fungsi ini di Triggers Apps Script untuk berjalan otomatis setiap hari.
function cekExpiredOtomatis() {
  const { stok } = getAppData(); // Ambil list stok yang sudah diagregasi
  
  let list6 = [], list3 = [], list1 = [];
  
  stok.forEach(s => {
    if(s.stokGudang > 0 || s.stokApotek > 0) {
      let d = s.sisaHari;
      if(d > 90 && d <= 180) list6.push(`- ${s.nama} | Batch: ${s.batch} | Exp: ${s.tglExp}`);
      else if(d > 30 && d <= 90) list3.push(`- ${s.nama} | Batch: ${s.batch} | Exp: ${s.tglExp}`);
      else if(d > 0 && d <= 30) list1.push(`- ${s.nama} | Batch: ${s.batch} | Exp: ${s.tglExp}`);
    }
  });

  let hasData = list6.length > 0 || list3.length > 0 || list1.length > 0;
  if(hasData) {
    let targetEmails = getSemuaEmailKaryawan();
    if(targetEmails) {
      let subj = `[PENGINGAT SISTEM] Ringkasan Stok Obat Mendekati Kedaluwarsa`;
      let body = `Berikut adalah rekapitulasi obat-obatan yang mendekati batas kedaluwarsa hari ini:\n\n`;
      
      if(list1.length > 0) body += `=== KRITIS (1 BULAN ATAU KURANG) ===\n${list1.join("\n")}\n\n`;
      if(list3.length > 0) body += `=== WASPADA (3 BULAN) ===\n${list3.join("\n")}\n\n`;
      if(list6.length > 0) body += `=== PERHATIAN (6 BULAN) ===\n${list6.join("\n")}\n\n`;
      
      body += `\nMohon lakukan pengecekan fisik dan prioritaskan pengeluaran stok di atas.\n\nSalam,\nSistem Apotek Indocement`;
      
      MailApp.sendEmail(targetEmails, subj, body);
    }
  }
}