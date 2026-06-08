/**
 * OKR & KPI Tracker System - Google Apps Script Backend
 * File: Code.gs
 */

// Menampilkan halaman utama (Web App Entry Point)
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('SaaS OKR & KPI Tracker')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Helper untuk menyertakan file HTML lain jika diperlukan (meskipun kita satukan di Index.html)
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * 1. SETUP DATABASE OTOMATIS
 * Membuat spreadsheet baru jika belum ada, lalu menginisialisasi sheet & data awal admin.
 */
function setupDatabase() {
  const scriptProperties = PropertiesService.getScriptProperties();
  let spreadsheetId = scriptProperties.getProperty('SPREADSHEET_ID');
  let ss;

  if (!spreadsheetId) {
    // Buat spreadsheet baru
    ss = SpreadsheetApp.create('Database OKR & KPI Tracker');
    spreadsheetId = ss.getId();
    scriptProperties.setProperty('SPREADSHEET_ID', spreadsheetId);
  } else {
    try {
      ss = SpreadsheetApp.openById(spreadsheetId);
    } catch (e) {
      // Jika ID yang tersimpan tidak valid/terhapus, buat baru
      ss = SpreadsheetApp.create('Database OKR & KPI Tracker');
      spreadsheetId = ss.getId();
      scriptProperties.setProperty('SPREADSHEET_ID', spreadsheetId);
    }
  }

  // Definisikan lembar kerja (Sheets) dan header masing-masing
  const sheetsConfig = {
    'Users': ['Username', 'Password', 'Role'],
    'Employees': ['EmployeeID', 'Nama', 'Jabatan', 'Divisi', 'Email'],
    'Objectives': ['ObjectiveID', 'Objective', 'PIC', 'StartDate', 'EndDate', 'Status'],
    'KeyResults': ['KRID', 'ObjectiveID', 'KeyResult', 'Target', 'Progress', 'Status'],
    'KPI': ['KPIID', 'EmployeeID', 'KPIName', 'Target', 'Realisasi', 'Bobot', 'NilaiKPI']
  };

  for (let sheetName in sheetsConfig) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.getRange(1, 1, 1, sheetsConfig[sheetName].length)
           .setValues([sheetsConfig[sheetName]])
           .setFontWeight('bold')
           .setBackground('#2563eb')
           .setFontColor('#ffffff');
      
      // Data awal khusus untuk Users
      if (sheetName === 'Users') {
        sheet.appendRow(['admin', 'admin123', 'Admin']);
      }
      
      // Contoh data awal untuk demonstrasi
      if (sheetName === 'Employees') {
        sheet.appendRow(['EMP-001', 'Budi Santoso', 'Software Engineer', 'IT', 'budi@company.com']);
        sheet.appendRow(['EMP-002', 'Siti Aminah', 'Product Manager', 'Product', 'siti@company.com']);
      }
      if (sheetName === 'Objectives') {
        sheet.appendRow(['OBJ-001', 'Meningkatkan Kualitas Rilis Sistem', 'Budi Santoso', '2026-01-01', '2026-06-30', 'On Progress']);
      }
      if (sheetName === 'KeyResults') {
        sheet.appendRow(['KR-001', 'OBJ-001', 'Mengurangi bug produksi sebesar 50%', '100', '60', 'On Progress']);
      }
      if (sheetName === 'KPI') {
        sheet.appendRow(['KPI-001', 'EMP-001', 'Ketepatan Waktu Delivery Task', '100', '90', '40', '36']);
        sheet.appendRow(['KPI-002', 'EMP-002', 'Kepuasan Pengguna Aplikasi', '5', '4.5', '60', '54']);
      }
    }
  }
  
  return {
    status: 'success',
    message: 'Database berhasil dikonfigurasi!',
    spreadsheetUrl: ss.getUrl()
  };
}

/**
 * Mendapatkan koneksi ke Spreadsheet Database
 */
function getDatabase() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const spreadsheetId = scriptProperties.getProperty('SPREADSHEET_ID');
  if (!spreadsheetId) {
    throw new Error('Database belum dikonfigurasi. Harap jalankan fungsi setupDatabase() terlebih dahulu.');
  }
  return SpreadsheetApp.openById(spreadsheetId);
}

/**
 * 2. SISTEM LOGIN & UTILITY
 */
function login(username, password) {
  try {
    const ss = getDatabase();
    const sheet = ss.getSheetByName('Users');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === username && data[i][1] === password) {
        return {
          status: 'success',
          user: {
            username: data[i][0],
            role: data[i][2]
          }
        };
      }
    }
    return { status: 'error', message: 'Username atau Password salah!' };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

/**
 * 3. MODUL EMPLOYEES (CRUD)
 */
function getEmployees() {
  try {
    const ss = getDatabase();
    const sheet = ss.getSheetByName('Employees');
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const result = [];
    
    for (let i = 1; i < data.length; i++) {
      let row = {};
      for (let j = 0; j < headers.length; j++) {
        row[headers[j]] = data[i][j];
      }
      result.push(row);
    }
    return { status: 'success', data: result };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

function saveEmployee(employee) {
  try {
    const ss = getDatabase();
    const sheet = ss.getSheetByName('Employees');
    const data = sheet.getDataRange().getValues();
    
    // Auto generate ID (EMP-001, dst)
    let nextId = 'EMP-001';
    if (data.length > 1) {
      const lastId = data[data.length - 1][0];
      const num = parseInt(lastId.split('-')[1]) + 1;
      nextId = 'EMP-' + String(num).padStart(3, '0');
    }
    
    sheet.appendRow([
      nextId,
      employee.nama,
      employee.jabatan,
      employee.divisi,
      employee.email
    ]);
    return { status: 'success', message: 'Karyawan berhasil ditambahkan.' };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

function updateEmployee(employee) {
  try {
    const ss = getDatabase();
    const sheet = ss.getSheetByName('Employees');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === employee.EmployeeID) {
        const row = i + 1;
        sheet.getRange(row, 2).setValue(employee.nama);
        sheet.getRange(row, 3).setValue(employee.jabatan);
        sheet.getRange(row, 4).setValue(employee.divisi);
        sheet.getRange(row, 5).setValue(employee.email);
        return { status: 'success', message: 'Data karyawan berhasil diperbarui.' };
      }
    }
    return { status: 'error', message: 'Karyawan tidak ditemukan.' };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

function deleteEmployee(employeeId) {
  try {
    const ss = getDatabase();
    const sheet = ss.getSheetByName('Employees');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === employeeId) {
        sheet.deleteRow(i + 1);
        return { status: 'success', message: 'Karyawan berhasil dihapus.' };
      }
    }
    return { status: 'error', message: 'Karyawan tidak ditemukan.' };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

/**
 * 4. MODUL OBJECTIVES (CRUD)
 */
function getObjectives() {
  try {
    const ss = getDatabase();
    const sheet = ss.getSheetByName('Objectives');
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const result = [];
    
    for (let i = 1; i < data.length; i++) {
      let row = {};
      for (let j = 0; j < headers.length; j++) {
        row[headers[j]] = data[i][j];
      }
      result.push(row);
    }
    return { status: 'success', data: result };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

function saveObjective(obj) {
  try {
    const ss = getDatabase();
    const sheet = ss.getSheetByName('Objectives');
    const data = sheet.getDataRange().getValues();
    
    let nextId = 'OBJ-001';
    if (data.length > 1) {
      const lastId = data[data.length - 1][0];
      const num = parseInt(lastId.split('-')[1]) + 1;
      nextId = 'OBJ-' + String(num).padStart(3, '0');
    }
    
    sheet.appendRow([
      nextId,
      obj.objective,
      obj.pic,
      obj.startDate,
      obj.endDate,
      obj.status || 'On Progress'
    ]);
    return { status: 'success', message: 'Objective berhasil disimpan.' };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

function updateObjective(obj) {
  try {
    const ss = getDatabase();
    const sheet = ss.getSheetByName('Objectives');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === obj.ObjectiveID) {
        const row = i + 1;
        sheet.getRange(row, 2).setValue(obj.objective);
        sheet.getRange(row, 3).setValue(obj.pic);
        sheet.getRange(row, 4).setValue(obj.startDate);
        sheet.getRange(row, 5).setValue(obj.endDate);
        sheet.getRange(row, 6).setValue(obj.status);
        return { status: 'success', message: 'Objective berhasil diperbarui.' };
      }
    }
    return { status: 'error', message: 'Objective tidak ditemukan.' };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

function deleteObjective(objId) {
  try {
    const ss = getDatabase();
    const sheet = ss.getSheetByName('Objectives');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === objId) {
        sheet.deleteRow(i + 1);
        
        // Hapus juga Key Results yang bergantung pada Objective ini
        const krSheet = ss.getSheetByName('KeyResults');
        const krData = krSheet.getDataRange().getValues();
        for (let j = krData.length - 1; j >= 1; j--) {
          if (krData[j][1] === objId) {
            krSheet.deleteRow(j + 1);
          }
        }
        
        return { status: 'success', message: 'Objective dan Key Results terkait berhasil dihapus.' };
      }
    }
    return { status: 'error', message: 'Objective tidak ditemukan.' };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

/**
 * 5. MODUL KEY RESULTS (CRUD)
 */
function getKeyResults() {
  try {
    const ss = getDatabase();
    const sheet = ss.getSheetByName('KeyResults');
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const result = [];
    
    for (let i = 1; i < data.length; i++) {
      let row = {};
      for (let j = 0; j < headers.length; j++) {
        row[headers[j]] = data[i][j];
      }
      result.push(row);
    }
    return { status: 'success', data: result };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

function saveKeyResult(kr) {
  try {
    const ss = getDatabase();
    const sheet = ss.getSheetByName('KeyResults');
    const data = sheet.getDataRange().getValues();
    
    let nextId = 'KR-001';
    if (data.length > 1) {
      const lastId = data[data.length - 1][0];
      const num = parseInt(lastId.split('-')[1]) + 1;
      nextId = 'KR-' + String(num).padStart(3, '0');
    }
    
    sheet.appendRow([
      nextId,
      kr.objectiveId,
      kr.keyResult,
      Number(kr.target),
      Number(kr.progress),
      kr.status || 'On Progress'
    ]);
    
    // Auto-update status objective berdasarkan total KR terkait
    recalculateObjectiveProgress(kr.objectiveId);
    
    return { status: 'success', message: 'Key Result berhasil disimpan.' };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

function updateKeyResult(kr) {
  try {
    const ss = getDatabase();
    const sheet = ss.getSheetByName('KeyResults');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === kr.KRID) {
        const row = i + 1;
        sheet.getRange(row, 2).setValue(kr.objectiveId);
        sheet.getRange(row, 3).setValue(kr.keyResult);
        sheet.getRange(row, 4).setValue(Number(kr.target));
        sheet.getRange(row, 5).setValue(Number(kr.progress));
        sheet.getRange(row, 6).setValue(kr.status);
        
        // Auto-update status objective
        recalculateObjectiveProgress(kr.objectiveId);
        
        return { status: 'success', message: 'Key Result berhasil diperbarui.' };
      }
    }
    return { status: 'error', message: 'Key Result tidak ditemukan.' };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

function deleteKeyResult(krId, objectiveId) {
  try {
    const ss = getDatabase();
    const sheet = ss.getSheetByName('KeyResults');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === krId) {
        sheet.deleteRow(i + 1);
        
        // Recalculate status
        recalculateObjectiveProgress(objectiveId);
        
        return { status: 'success', message: 'Key Result berhasil dihapus.' };
      }
    }
    return { status: 'error', message: 'Key Result tidak ditemukan.' };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

/**
 * Logika Otomatis: Menghitung persentase progres Objective dari Key Results
 */
function recalculateObjectiveProgress(objectiveId) {
  const ss = getDatabase();
  const krSheet = ss.getSheetByName('KeyResults');
  const krData = krSheet.getDataRange().getValues();
  
  let totalTarget = 0;
  let totalProgress = 0;
  let count = 0;
  
  for (let i = 1; i < krData.length; i++) {
    if (krData[i][1] === objectiveId) {
      totalTarget += Number(krData[i][3]) || 0;
      totalProgress += Number(krData[i][4]) || 0;
      count++;
    }
  }
  
  let finalStatus = 'On Progress';
  if (count > 0) {
    const pct = (totalProgress / totalTarget) * 100;
    if (pct >= 100) {
      finalStatus = 'Achieved';
    } else if (pct <= 0) {
      finalStatus = 'Not Started';
    }
  } else {
    finalStatus = 'Not Started';
  }
  
  const objSheet = ss.getSheetByName('Objectives');
  const objData = objSheet.getDataRange().getValues();
  for (let j = 1; j < objData.length; j++) {
    if (objData[j][0] === objectiveId) {
      objSheet.getRange(j + 1, 6).setValue(finalStatus);
      break;
    }
  }
}

/**
 * 6. MODUL KPI (CRUD & PERHITUNGAN)
 */
function getKPIs() {
  try {
    const ss = getDatabase();
    const sheet = ss.getSheetByName('KPI');
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const result = [];
    
    for (let i = 1; i < data.length; i++) {
      let row = {};
      for (let j = 0; j < headers.length; j++) {
        row[headers[j]] = data[i][j];
      }
      result.push(row);
    }
    return { status: 'success', data: result };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

function calculateKPI(realisasi, target, bobot) {
  const t = Number(target) || 1;
  const r = Number(realisasi) || 0;
  const b = Number(bobot) || 0;
  // Rumus: (Realisasi / Target) * Bobot
  const score = (r / t) * b;
  return Number(score.toFixed(2));
}

function saveKPI(kpi) {
  try {
    const ss = getDatabase();
    const sheet = ss.getSheetByName('KPI');
    const data = sheet.getDataRange().getValues();
    
    let nextId = 'KPI-001';
    if (data.length > 1) {
      const lastId = data[data.length - 1][0];
      const num = parseInt(lastId.split('-')[1]) + 1;
      nextId = 'KPI-' + String(num).padStart(3, '0');
    }
    
    const nilaiKpi = calculateKPI(kpi.realisasi, kpi.target, kpi.bobot);
    
    sheet.appendRow([
      nextId,
      kpi.employeeId,
      kpi.kpiName,
      Number(kpi.target),
      Number(kpi.realisasi),
      Number(kpi.bobot),
      nilaiKpi
    ]);
    return { status: 'success', message: 'KPI Karyawan berhasil disimpan.' };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

function updateKPI(kpi) {
  try {
    const ss = getDatabase();
    const sheet = ss.getSheetByName('KPI');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === kpi.KPIID) {
        const row = i + 1;
        const nilaiKpi = calculateKPI(kpi.realisasi, kpi.target, kpi.bobot);
        
        sheet.getRange(row, 2).setValue(kpi.employeeId);
        sheet.getRange(row, 3).setValue(kpi.kpiName);
        sheet.getRange(row, 4).setValue(Number(kpi.target));
        sheet.getRange(row, 5).setValue(Number(kpi.realisasi));
        sheet.getRange(row, 6).setValue(Number(kpi.bobot));
        sheet.getRange(row, 7).setValue(nilaiKpi);
        return { status: 'success', message: 'KPI Karyawan berhasil diperbarui.' };
      }
    }
    return { status: 'error', message: 'KPI tidak ditemukan.' };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

function deleteKPI(kpiId) {
  try {
    const ss = getDatabase();
    const sheet = ss.getSheetByName('KPI');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === kpiId) {
        sheet.deleteRow(i + 1);
        return { status: 'success', message: 'KPI berhasil dihapus.' };
      }
    }
    return { status: 'error', message: 'KPI tidak ditemukan.' };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

/**
 * 7. LEADERBOARD
 * Akumulasi nilai KPI total berdasarkan masing-masing karyawan
 */
function getLeaderboard() {
  try {
    const ss = getDatabase();
    const empSheet = ss.getSheetByName('Employees');
    const kpiSheet = ss.getSheetByName('KPI');
    
    const empData = empSheet.getDataRange().getValues();
    const kpiData = kpiSheet.getDataRange().getValues();
    
    const empMap = {};
    for (let i = 1; i < empData.length; i++) {
      empMap[empData[i][0]] = {
        id: empData[i][0],
        nama: empData[i][1],
        jabatan: empData[i][2],
        divisi: empData[i][3],
        totalKPI: 0,
        kpiCount: 0
      };
    }
    
    // Jumlahkan nilai KPI per karyawan
    for (let j = 1; j < kpiData.length; j++) {
      const empId = kpiData[j][1];
      const nilai = Number(kpiData[j][6]) || 0;
      if (empMap[empId]) {
        empMap[empId].totalKPI += nilai;
        empMap[empId].kpiCount += 1;
      }
    }
    
    // Ubah ke array dan urutkan
    const leaderboard = Object.keys(empMap).map(key => {
      // Nilai rata-rata / akumulasi total KPI (asumsi total bobot maksimal diatur dengan baik)
      return {
        id: empMap[key].id,
        nama: empMap[key].nama,
        jabatan: empMap[key].jabatan,
        divisi: empMap[key].divisi,
        score: Number(empMap[key].totalKPI.toFixed(2))
      };
    });
    
    // Sort descending berdasarkan score
    leaderboard.sort((a, b) => b.score - a.score);
    
    return { status: 'success', data: leaderboard };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

/**
 * 8. DASHBOARD DATA
 */
function getDashboardData() {
  try {
    const ss = getDatabase();
    
    const empData = ss.getSheetByName('Employees').getDataRange().getValues();
    const objData = ss.getSheetByName('Objectives').getDataRange().getValues();
    const kpiData = ss.getSheetByName('KPI').getDataRange().getValues();
    
    const totalEmployees = empData.length - 1;
    const totalObjectives = objData.length - 1;
    
    let kpiTercapai = 0;
    let kpiBelumTercapai = 0;
    
    // Asumsi KPI tercapai jika Realisasi >= Target
    for (let i = 1; i < kpiData.length; i++) {
      const target = Number(kpiData[i][3]) || 0;
      const realisasi = Number(kpiData[i][4]) || 0;
      if (realisasi >= target && target > 0) {
        kpiTercapai++;
      } else {
        kpiBelumTercapai++;
      }
    }
    
    // Data Distribusi Status Objective untuk Chart
    const objStatusCount = { 'Achieved': 0, 'On Progress': 0, 'Not Started': 0 };
    for (let j = 1; j < objData.length; j++) {
      const status = objData[j][5] || 'On Progress';
      if (objStatusCount[status] !== undefined) {
        objStatusCount[status]++;
      } else {
        objStatusCount['On Progress']++;
      }
    }
    
    // Ambil top 5 KPI untuk chart performansi
    const topKPIs = [];
    for (let k = 1; k < Math.min(kpiData.length, 6); k++) {
      topKPIs.push({
        name: kpiData[k][2],
        score: Number(kpiData[k][6]) || 0
      });
    }
    
    return {
      status: 'success',
      summary: {
        totalEmployees: Math.max(0, totalEmployees),
        totalObjectives: Math.max(0, totalObjectives),
        kpiTercapai,
        kpiBelumTercapai
      },
      charts: {
        objectiveStatus: objStatusCount,
        topKPIs: topKPIs
      }
    };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}