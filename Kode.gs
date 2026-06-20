// ============================================
// KODE.GS — FISIOACTIVE
// Sistem Manajemen Klinik Fisioterapi
// Paste seluruh isi file ini ke Kode.gs
// di Google Apps Script
// ============================================

// Fungsi utama: serve halaman website
function doGet(e) {
  if (e && e.parameter && e.parameter.action) {
    var action = e.parameter.action;
    var params = {};
    for (var key in e.parameter) {
      if (key !== 'action') {
        params[key] = e.parameter[key];
      }
    }
    var result = runAction(action, params);
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('FisioActive — Sistem Manajemen Klinik Fisioterapi')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e) {
  var result;
  try {
    var postData = JSON.parse(e.postData.contents);
    var action = postData.action;
    result = runAction(action, postData);
  } catch (err) {
    result = { error: err.toString() };
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// Router Action untuk Apps Script & Web API
function runAction(action, params) {
  try {
    if (action === 'getAllData') {
      return getAllData();
    } else if (action === 'upsertData') {
      var sheetName = params.sheetName;
      var rowData = typeof params.data === 'string' ? JSON.parse(params.data) : params.data;
      return upsertData(sheetName, rowData);
    } else if (action === 'deleteData') {
      var sheetName = params.sheetName;
      var keyData = typeof params.data === 'string' ? JSON.parse(params.data) : params.data;
      return deleteData(sheetName, keyData);
    } else if (action === 'getSheetStats') {
      return getSheetStats();
    } else if (action === 'getPatientStatus') {
      var idPasien = params.idPasien;
      return getPatientStatus(idPasien);
    } else {
      return { error: 'Unknown action: ' + action };
    }
  } catch (err) {
    return { error: err.toString() };
  }
}


// ============================================
// KONFIGURASI SHEET & HEADER
// ============================================

var SHEET_CONFIG = {
  'Pasien': {
    key: 'ID_Pasien',
    headers: ['ID_Pasien','Nama','KTP_NIK','Tgl_Lahir','Jenis_Kelamin','No_HP','Email','Keluhan','Tgl_Daftar','Status']
  },
  'RekamMedis': {
    key: 'ID_Rekam',
    headers: ['ID_Rekam','ID_Pasien','Diagnosa','Rekomendasi_Terapi','Catatan','Pemeriksa','Tgl_Periksa','Status']
  },
  'Jadwal': {
    key: 'ID_Jadwal',
    headers: ['ID_Jadwal','ID_Pasien','Nama_Pasien','Jenis_Terapi','Terapis','Tanggal','Jam','Ruang','Status']
  },
  'Users': {
    key: 'Username',
    headers: ['Username','Password','Nama_Lengkap','Role']
  },
  'Terapis': {
    key: 'ID_Terapis',
    headers: ['ID_Terapis','Nama','Spesialisasi','Deskripsi','Shift_Mulai','Shift_Selesai','Total_Sesi','Pasien_Hari_Ini']
  },
  'JenisTerapi': {
    key: 'Kode_Terapi',
    headers: ['Kode_Terapi','Nama_Terapi','Deskripsi','Durasi_Menit']
  }
};

// ============================================
// INISIALISASI SHEET + DATA DUMMY (RUN SEKALI)
// ============================================

function initSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var report = [];

  for (var name in SHEET_CONFIG) {
    var config = SHEET_CONFIG[name];
    var sheet = ss.getSheetByName(name);

    if (!sheet) {
      sheet = ss.insertSheet(name);
      report.push(name + ': CREATED');
    } else {
      if (sheet.getLastRow() > 1) {
        sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clear();
      }
    }

    var headerRange = sheet.getRange(1, 1, 1, config.headers.length);
    headerRange.setValues([config.headers]);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4f46e5');
    headerRange.setFontColor('#ffffff');

    report.push(name + ': Headers OK (' + config.headers.length + ' kolom)');
  }

  // ============================
  // DATA DUMMY — SEMUA ORIGINAL
  // ============================
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  // --- USERS ---
  var usersSheet = ss.getSheetByName('Users');
  var usersData = [
    ['admin', 'admin123', 'Administrator', 'admin'],
    ['terapis', 'terapis123', "M. Belva I'zaz Sahitya, S. Ft", 'terapis']
  ];
  usersData.forEach(function(row) { usersSheet.appendRow(row); });
  report.push('Users: 2 users added');

  // --- JENIS TERAPI (5 jenis) ---
  var jtSheet = ss.getSheetByName('JenisTerapi');
  var jtData = [
    ['KT01', 'Mobilisasi Sendi & Jaringan', 'Pendekatan manual untuk mengembalikan gerak sendi dan fleksibilitas jaringan lunak sekitar persendian', '40'],
    ['KT02', 'Stimulasi Elektrik Fungsional', 'Penggunaan arus listrik terkontrol untuk mengaktifkan otot yang lemah atau lumpuh', '35'],
    ['KT03', 'Latihan Keseimbangan Dinamis', 'Program latihan untuk meningkatkan stabilitas tubuh saat bergerak dan mencegah jatuh', '30'],
    ['KT04', 'Terapi Ultrasound Terapeutik', 'Gelombang suara frekuensi tinggi untuk mempercepat penyembuhan jaringan dalam', '25'],
    ['KT05', 'Program Penguatan Progresif', 'Latihan bertahap untuk membangun kembali kekuatan otot pasca cedera atau operasi', '50']
  ];
  jtData.forEach(function(row) { jtSheet.appendRow(row); });
  report.push('JenisTerapi: 5 jenis terapi added');

  // --- TERAPIS (4 terapis) ---
  var tSheet = ss.getSheetByName('Terapis');
  var tData = [
    ['FT01', "M. Belva I'zaz Sahitya, S. Ft", 'Gangguan Gerak & Postur', 'Fokus pada koreksi postur dan pemulihan gerak fungsional selama 6 tahun praktik', '07:30', '14:30', '167', '4'],
    ['FT02', 'Galang Mahardika Putra, S.Ft', 'Rehabilitasi Cedera Olahraga', 'Menangani atlet dan individu aktif dengan cedera muskuloskeletal kompleks', '08:00', '15:00', '124', '3'],
    ['FT03', 'Melati Ayu Purnama, S.Ft', 'Tumbuh Kembang Anak', 'Terapis khusus anak dengan sertifikasi Bobath dan Sensory Integration', '09:00', '16:00', '93', '5'],
    ['FT04', 'Rangga Dwi Satriya, S.Ft', 'Pemulihan Pasca Bedah', 'Ahli dalam program rehabilitasi sistematis setelah prosedur bedah ortopedi', '08:30', '15:30', '108', '2']
  ];
  tData.forEach(function(row) { tSheet.appendRow(row); });
  report.push('Terapis: 4 terapis added');

  // --- PASIEN (15 pasien) ---
  var pSheet = ss.getSheetByName('Pasien');
  var pasienData = [
    ['FA001', 'Bagus Tirta Wijaya', '3271080503910001', '1991-03-05', 'Laki-laki', '082145670011', 'bagus.tirta@gmail.com', 'Kaku pada pergelangan tangan kanan setelah jatuh', '2025-01-08', 'Aktif'],
    ['FA002', 'Citra Pramesti Hapsari', '3271094107930002', '1993-07-01', 'Perempuan', '085378120022', 'citra.pramesti@yahoo.com', 'Nyeri tumit saat berjalan pagi hari', '2025-01-20', 'Aktif'],
    ['FA003', 'Dimas Arya Nugraha', '3271021209870003', '1987-09-12', 'Laki-laki', '081290450033', 'dimas.arya87@gmail.com', 'Ketegangan otot leher berulang', '2025-02-03', 'Aktif'],
    ['FA004', 'Elvira Putri Ramadhani', '3271035604960004', '1996-04-16', 'Perempuan', '087854310044', 'elvira.pr@outlook.com', 'Keterbatasan gerak bahu pasca kecelakaan motor', '2025-02-14', 'Aktif'],
    ['FA005', 'Farhan Maulana Akbar', '3271042811840005', '1984-11-28', 'Laki-laki', '082367890055', 'farhan.ma84@gmail.com', 'Nyeri pinggang menjalar ke tungkai kiri', '2025-02-25', 'Aktif'],
    ['FA006', 'Gita Nandya Paramitha', '3271051503950006', '1995-03-15', 'Perempuan', '085612340066', 'gita.nandya@gmail.com', 'Jari tangan kiri sering mati rasa', '2025-03-04', 'Aktif'],
    ['FA007', 'Haikal Baskara Pratama', '3271060208890007', '1989-08-02', 'Laki-laki', '081478560077', 'haikal.bp@proton.me', 'Kelemahan lengan dan kaki sisi kanan pasca rawat inap', '2025-03-10', 'Aktif'],
    ['FA008', 'Intan Maharani Putri', '3271071906970008', '1997-06-19', 'Perempuan', '087823450088', 'intan.mp97@gmail.com', 'Tulang punggung terasa miring sejak remaja', '2025-03-18', 'Aktif'],
    ['FA009', 'Joko Susanto Prawiranegara', '3271083010800009', '1980-10-30', 'Laki-laki', '082190780099', 'joko.sp80@yahoo.co.id', 'Lutut bengkak dan kaku terutama saat naik tangga', '2025-03-26', 'Aktif'],
    ['FA010', 'Kartini Dwi Lestari', '3271090704920010', '1992-04-07', 'Perempuan', '085734560110', 'kartini.dl@gmail.com', 'Pergelangan tangan nyeri saat mengetik lama', '2025-04-05', 'Aktif'],
    ['FA011', 'Lutfi Hakim Pradipta', '3271101201860011', '1986-01-12', 'Laki-laki', '081356780121', 'lutfi.hp86@gmail.com', 'Sakit punggung bawah tidak bisa membungkuk', '2025-04-12', 'Aktif'],
    ['FA012', 'Mayang Sari Utami', '3271112508940012', '1994-08-25', 'Perempuan', '087945120132', 'mayang.su@outlook.com', 'Wajah sebelah kanan tidak bisa bergerak tiba-tiba', '2025-04-19', 'Aktif'],
    ['FA013', 'Naufal Rizky Maulana', '3271121807830013', '1983-07-18', 'Laki-laki', '082278900143', 'naufal.rm@gmail.com', 'Lutut berbunyi dan nyeri saat ditekuk penuh', '2025-04-28', 'Aktif'],
    ['FA014', 'Olivia Zahra Azzahra', '3271130903180014', '2018-03-09', 'Perempuan', '085890120154', 'mama.olivia@gmail.com', 'Anak belum bisa berdiri sendiri di usia 3 tahun', '2025-05-06', 'Aktif'],
    ['FA015', 'Putra Aditya Wicaksono', '3271142206910015', '1991-06-22', 'Laki-laki', '081467230165', 'putra.aw91@gmail.com', 'Kaki kanan kaku setelah operasi patah tulang paha', '2025-05-14', 'Aktif']
  ];
  pasienData.forEach(function(row) { pSheet.appendRow(row); });
  report.push('Pasien: 15 pasien added');

  // --- JADWAL HARI INI (15 jadwal) ---
  var jSheet = ss.getSheetByName('Jadwal');
  var jadwalData = [
    ['SJ001', 'FA001', 'Bagus Tirta Wijaya', 'Mobilisasi Sendi & Jaringan', "M. Belva I'zaz Sahitya, S. Ft", today, '07:30', 'Studio A', 'Selesai'],
    ['SJ002', 'FA005', 'Farhan Maulana Akbar', 'Terapi Ultrasound Terapeutik', "M. Belva I'zaz Sahitya, S. Ft", today, '08:15', 'Studio A', 'Selesai'],
    ['SJ003', 'FA003', 'Dimas Arya Nugraha', 'Mobilisasi Sendi & Jaringan', 'Galang Mahardika Putra, S.Ft', today, '08:00', 'Studio B', 'Selesai'],
    ['SJ004', 'FA007', 'Haikal Baskara Pratama', 'Stimulasi Elektrik Fungsional', 'Rangga Dwi Satriya, S.Ft', today, '08:30', 'Studio C', 'Dalam Proses'],
    ['SJ005', 'FA002', 'Citra Pramesti Hapsari', 'Program Penguatan Progresif', 'Galang Mahardika Putra, S.Ft', today, '09:00', 'Studio B', 'Dalam Proses'],
    ['SJ006', 'FA014', 'Olivia Zahra Azzahra', 'Latihan Keseimbangan Dinamis', 'Melati Ayu Purnama, S.Ft', today, '09:00', 'Studio D', 'Dalam Proses'],
    ['SJ007', 'FA004', 'Elvira Putri Ramadhani', 'Mobilisasi Sendi & Jaringan', "M. Belva I'zaz Sahitya, S. Ft", today, '09:30', 'Studio A', 'Menunggu'],
    ['SJ008', 'FA009', 'Joko Susanto Prawiranegara', 'Terapi Ultrasound Terapeutik', 'Rangga Dwi Satriya, S.Ft', today, '10:00', 'Studio C', 'Menunggu'],
    ['SJ009', 'FA012', 'Mayang Sari Utami', 'Stimulasi Elektrik Fungsional', 'Galang Mahardika Putra, S.Ft', today, '10:30', 'Studio B', 'Menunggu'],
    ['SJ010', 'FA006', 'Gita Nandya Paramitha', 'Terapi Ultrasound Terapeutik', "M. Belva I'zaz Sahitya, S. Ft", today, '10:45', 'Studio A', 'Menunggu'],
    ['SJ011', 'FA010', 'Kartini Dwi Lestari', 'Mobilisasi Sendi & Jaringan', 'Rangga Dwi Satriya, S.Ft', today, '11:00', 'Studio C', 'Menunggu'],
    ['SJ012', 'FA008', 'Intan Maharani Putri', 'Program Penguatan Progresif', 'Galang Mahardika Putra, S.Ft', today, '13:00', 'Studio B', 'Menunggu'],
    ['SJ013', 'FA011', 'Lutfi Hakim Pradipta', 'Terapi Ultrasound Terapeutik', "M. Belva I'zaz Sahitya, S. Ft", today, '13:30', 'Studio A', 'Menunggu'],
    ['SJ014', 'FA013', 'Naufal Rizky Maulana', 'Latihan Keseimbangan Dinamis', 'Rangga Dwi Satriya, S.Ft', today, '14:00', 'Studio C', 'Menunggu'],
    ['SJ015', 'FA015', 'Putra Aditya Wicaksono', 'Program Penguatan Progresif', 'Galang Mahardika Putra, S.Ft', today, '14:30', 'Studio B', 'Menunggu']
  ];
  jadwalData.forEach(function(row) { jSheet.appendRow(row); });
  report.push('Jadwal: 15 jadwal hari ini added');

  // --- REKAM MEDIS (10 rekam) ---
  var rmSheet = ss.getSheetByName('RekamMedis');
  var rmData = [
    ['MR001', 'FA001', 'Kekakuan Pasca Fraktur Radius Distal', 'Mobilisasi sendi 3x/minggu selama 4 minggu', 'Grip strength meningkat 40% dari baseline', "M. Belva I'zaz Sahitya, S. Ft", '2025-03-08', 'Aktif'],
    ['MR002', 'FA005', 'Lumbar Disc Bulging L5-S1', 'Kombinasi ultrasound dan latihan stabilisasi core', 'VAS nyeri turun dari 8 ke 3 setelah 6 sesi', "M. Belva I'zaz Sahitya, S. Ft", '2025-03-14', 'Aktif'],
    ['MR003', 'FA003', 'Cervical Myofascial Pain Syndrome', 'Mobilisasi cervical dan stretching rutin di rumah', 'Trigger point berkurang signifikan', 'Galang Mahardika Putra, S.Ft', '2025-03-20', 'Aktif'],
    ['MR004', 'FA007', 'Hemiparesis Kanan ec CVA Infark', 'Stimulasi elektrik dan latihan fungsional harian', 'Kekuatan otot naik dari grade 2 ke 3+', 'Rangga Dwi Satriya, S.Ft', '2025-03-25', 'Aktif'],
    ['MR005', 'FA002', 'Plantar Fasciitis Bilateral', 'Penguatan intrinsik kaki dan modifikasi alas kaki', 'Morning pain berkurang 60% setelah 8 sesi', 'Galang Mahardika Putra, S.Ft', '2025-04-02', 'Aktif'],
    ['MR006', 'FA009', 'Gonarthrosis Bilateral Grade III', 'Ultrasound dan latihan penguatan quadriceps', 'Jarak jalan meningkat dari 100m ke 350m', 'Rangga Dwi Satriya, S.Ft', '2025-04-10', 'Aktif'],
    ['MR007', 'FA012', 'Paresis Nervus Fasialis Perifer Dextra', 'Elektrostimulasi wajah dan mirror therapy', 'Asimetri berkurang dari grade 5 ke grade 3', 'Galang Mahardika Putra, S.Ft', '2025-04-22', 'Aktif'],
    ['MR008', 'FA014', 'Keterlambatan Motorik Kasar', 'Latihan keseimbangan dan koordinasi berbasis bermain', 'Sudah mampu berdiri 5 detik tanpa pegangan', 'Melati Ayu Purnama, S.Ft', '2025-05-08', 'Aktif'],
    ['MR009', 'FA011', 'Herniasi Diskus Lumbal dengan Sciatica', 'Traksi manual dan McKenzie exercise', 'Straight leg raise meningkat dari 30 ke 65 derajat', "M. Belva I'zaz Sahitya, S. Ft", '2025-05-12', 'Aktif'],
    ['MR010', 'FA015', 'Post ORIF Fraktur Shaft Femur Dextra', 'Penguatan progresif dan gait training', 'Sudah bisa partial weight bearing dengan kruk', 'Rangga Dwi Satriya, S.Ft', '2025-05-18', 'Aktif']
  ];
  rmData.forEach(function(row) { rmSheet.appendRow(row); });
  report.push('RekamMedis: 10 rekam medis added');

  report.push('');
  report.push('=== SELESAI! Data dummy berhasil ditambahkan ===');
  report.push('Login: admin/admin123 atau terapis/terapis123');

  return report.join('\n');
}

// ============================================
// AMBIL SEMUA DATA
// ============================================

function getAllData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var result = {};

  for (var name in SHEET_CONFIG) {
    var sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastRow() < 2) {
      result[name] = [];
      continue;
    }
    var data = sheet.getDataRange().getValues();
    var headers = SHEET_CONFIG[name].headers;
    var rows = [];
    for (var i = 1; i < data.length; i++) {
      var row = {};
      for (var j = 0; j < headers.length; j++) {
        row[headers[j]] = (j < data[i].length) ? String(data[i][j]) : '';
      }
      rows.push(row);
    }
    result[name] = rows;
  }

  return result;
}

// ============================================
// UPSERT DATA (Insert / Update)
// ============================================

function upsertData(sheetName, rowData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return 'Sheet not found: ' + sheetName;

  var config = SHEET_CONFIG[sheetName];
  if (!config) return 'Config not found: ' + sheetName;

  var headers = config.headers;
  var keyColumnName = config.key;
  var keyIndex = headers.indexOf(keyColumnName);
  var keyValue = rowData[keyColumnName];
  var existingRowIndex = -1;

  if (keyIndex !== -1 && keyValue) {
    var values = sheet.getDataRange().getValues();
    for (var i = 1; i < values.length; i++) {
      if (String(values[i][keyIndex]) === String(keyValue)) {
        existingRowIndex = i + 1;
        break;
      }
    }
  }

  var newRow = [];
  headers.forEach(function(h) {
    newRow.push(rowData[h] !== undefined ? rowData[h] : '');
  });

  if (existingRowIndex !== -1) {
    sheet.getRange(existingRowIndex, 1, 1, newRow.length).setValues([newRow]);
    return 'Updated';
  } else {
    sheet.appendRow(newRow);
    return 'Appended';
  }
}

// ============================================
// HAPUS DATA
// ============================================

function deleteData(sheetName, keyData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return 'Sheet not found: ' + sheetName;

  var config = SHEET_CONFIG[sheetName];
  if (!config) return 'Config not found: ' + sheetName;

  var headers = config.headers;
  var keyColumnName = config.key;
  var keyIndex = headers.indexOf(keyColumnName);
  var keyVal = keyData[keyColumnName];

  if (keyIndex !== -1 && keyVal) {
    var values = sheet.getDataRange().getValues();
    for (var i = values.length - 1; i >= 1; i--) {
      if (String(values[i][keyIndex]) === String(keyVal)) {
        sheet.deleteRow(i + 1);
      }
    }
    return 'Deleted';
  }
  return 'Key not found';
}

// ============================================
// STATISTIK UNTUK DASHBOARD & LAPORAN
// ============================================

function getSheetStats() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var stats = {};

  var pSheet = ss.getSheetByName('Pasien');
  stats.totalPasien = pSheet ? Math.max(0, pSheet.getLastRow() - 1) : 0;

  var jSheet = ss.getSheetByName('Jadwal');
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  stats.sesiHariIni = 0;
  stats.menunggu = 0;
  stats.selesai = 0;
  stats.dalamProses = 0;

  if (jSheet && jSheet.getLastRow() >= 2) {
    var jData = jSheet.getDataRange().getValues();
    for (var i = 1; i < jData.length; i++) {
      var tgl = '';
      if (jData[i][5] instanceof Date) {
        tgl = Utilities.formatDate(jData[i][5], Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } else {
        tgl = String(jData[i][5]);
      }
      if (tgl === today) {
        stats.sesiHariIni++;
        var status = String(jData[i][8]).toLowerCase();
        if (status === 'menunggu') stats.menunggu++;
        else if (status === 'dalam proses') stats.dalamProses++;
        else if (status === 'selesai') stats.selesai++;
      }
    }
  }

  return stats;
}

// ============================================
// CEK STATUS PASIEN (PUBLIK)
// ============================================

function getPatientStatus(idPasien) {
  if (!idPasien) return { error: 'ID Pasien wajib diisi' };
  idPasien = String(idPasien).trim();

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Jika input diawali dengan 'MR' (misal MR001), cari ID_Pasien di sheet RekamMedis
  if (idPasien.toUpperCase().indexOf('MR') === 0) {
    var rSheet = ss.getSheetByName('RekamMedis');
    if (rSheet && rSheet.getLastRow() >= 2) {
      var rData = rSheet.getDataRange().getValues();
      for (var i = 1; i < rData.length; i++) {
        if (String(rData[i][0]).toUpperCase() === idPasien.toUpperCase()) {
          idPasien = String(rData[i][1]); // Dapatkan ID_Pasien
          break;
        }
      }
    }
  }

  var result = { found: false };

  var pSheet = ss.getSheetByName('Pasien');
  if (pSheet && pSheet.getLastRow() >= 2) {
    var pData = pSheet.getDataRange().getValues();
    var pHeaders = SHEET_CONFIG['Pasien'].headers;
    for (var i = 1; i < pData.length; i++) {
      if (String(pData[i][0]).toUpperCase() === idPasien.toUpperCase()) {
        result.found = true;
        result.pasien = {};
        for (var j = 0; j < pHeaders.length; j++) {
          if (pHeaders[j] !== 'KTP_NIK') {
            result.pasien[pHeaders[j]] = String(pData[i][j] || '');
          }
        }
        break;
      }
    }
  }

  if (!result.found) return { found: false, error: 'ID Pasien atau Rekam Medis tidak ditemukan' };

  result.jadwal = [];
  var jSheet = ss.getSheetByName('Jadwal');
  if (jSheet && jSheet.getLastRow() >= 2) {
    var jData = jSheet.getDataRange().getValues();
    var jHeaders = SHEET_CONFIG['Jadwal'].headers;
    for (var i = 1; i < jData.length; i++) {
      if (String(jData[i][1]).toUpperCase() === idPasien.toUpperCase()) {
        var row = {};
        for (var j = 0; j < jHeaders.length; j++) {
          row[jHeaders[j]] = String(jData[i][j] || '');
        }
        result.jadwal.push(row);
      }
    }
  }

  result.rekamMedis = [];
  var rSheet = ss.getSheetByName('RekamMedis');
  if (rSheet && rSheet.getLastRow() >= 2) {
    var rData = rSheet.getDataRange().getValues();
    var rHeaders = SHEET_CONFIG['RekamMedis'].headers;
    for (var i = 1; i < rData.length; i++) {
      if (String(rData[i][1]).toUpperCase() === idPasien.toUpperCase()) {
        var row = {};
        for (var j = 0; j < rHeaders.length; j++) {
          if (rHeaders[j] !== 'Catatan') {
            row[rHeaders[j]] = String(rData[i][j] || '');
          }
        }
        result.rekamMedis.push(row);
      }
    }
  }

  return result;
}
