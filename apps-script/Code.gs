// ============================================================
// ระบบนิเทศภายในโรงเรียนนราศึกษาธิการ
// Google Apps Script Backend
// ============================================================

const SPREADSHEET_ID = '1KHJVpe1w4gRPAckL_gbs0WrGEjnDJylwny0od7jvkQM';
const DRIVE_FOLDER_ID = '1FDj-gDOWuvq8A_kugvPrBhPZ8X4pHVUo';
const ADMIN_PASSWORD = 'admin123';

const SHEETS = {
  BOOKING: 'Booking',
  FILES: 'Files',
  EVALUATION: 'Supervision',
  USERS: 'Users'
};

const STATUS = {
  PENDING: 'รอดำเนินการ',
  CONFIRMED: 'ยืนยันแล้ว',
  REJECTED: 'ปฏิเสธ',
  COMPLETED: 'นิเทศแล้ว'
};

const FILE_STATUS = {
  PENDING: 'รอตรวจสอบ',
  APPROVED: 'ผ่าน',
  REVISE: 'ปรับปรุง'
};

const QUALITY_LEVELS = {
  EXCELLENT: 'ดีมาก',
  GOOD: 'ดี',
  FAIR: 'พอใช้',
  POOR: 'ปรับปรุง'
};

// ============================================================
// Web App Entry Points
// ============================================================

function doGet(e) {
  const page = e.parameter.page || 'index';
  const template = HtmlService.createTemplateFromFile(page);
  const html = template.evaluate()
    .setTitle('ระบบนิเทศภายในโรงเรียนนราศึกษาธิการ')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  return html;
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;

    switch (action) {
      case 'login': return jsonResponse(login(data));
      case 'createBooking': return jsonResponse(createBooking(data));
      case 'updateBooking': return jsonResponse(updateBooking(data));
      case 'deleteBooking': return jsonResponse(deleteBooking(data));
      case 'getBookings': return jsonResponse(getBookings(data));
      case 'getBookingById': return jsonResponse(getBookingById(data));
      case 'uploadFile': return jsonResponse(uploadFile(data));
      case 'updateFileStatus': return jsonResponse(updateFileStatus(data));
      case 'getFiles': return jsonResponse(getFiles(data));
      case 'deleteFile': return jsonResponse(deleteFile(data));
      case 'createEvaluation': return jsonResponse(createEvaluation(data));
      case 'updateEvaluation': return jsonResponse(updateEvaluation(data));
      case 'getEvaluations': return jsonResponse(getEvaluations(data));
      case 'deleteEvaluation': return jsonResponse(deleteEvaluation(data));
      case 'getDashboardStats': return jsonResponse(getDashboardStats(data));
      case 'getCalendarData': return jsonResponse(getCalendarData(data));
      case 'getReport': return jsonResponse(getReport(data));
      case 'getTeacherReport': return jsonResponse(getTeacherReport(data));
      case 'getDepartmentReport': return jsonResponse(getDepartmentReport(data));
      default: return jsonResponse({ success: false, message: 'ไม่รู้จักคำสั่ง' });
    }
  } catch (err) {
    return jsonResponse({ success: false, message: err.toString() });
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================================
// Utility Functions
// ============================================================

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet(name) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    initializeSheet(sheet, name);
  }
  return sheet;
}

function initializeSheet(sheet, name) {
  const headers = {
    [SHEETS.BOOKING]: ['Timestamp', 'Date', 'Time', 'Teacher Name', 'Department', 'Period', 'Subject Name', 'Subject Code', 'Class Level', 'Room', 'Status', 'Notes'],
    [SHEETS.FILES]: ['Timestamp', 'Teacher Name', 'Booking ID', 'File Type', 'File Name', 'File URL', 'Drive File ID', 'Status', 'Admin Note', 'Reviewed By', 'Reviewed Date'],
    [SHEETS.EVALUATION]: ['Timestamp', 'Teacher Name', 'Supervision Date', 'Booking ID', 'Strengths', 'Improvements', 'Suggestions', 'Quality Level', 'Score', 'Evaluated By', 'Booking Reference'],
    [SHEETS.USERS]: ['Username', 'Password', 'Full Name', 'Role', 'Department', 'Active']
  };

  if (headers[name]) {
    sheet.getRange(1, 1, 1, headers[name].length).setValues([headers[name]]);
    sheet.getRange(1, 1, 1, headers[name].length).setFontWeight('bold');
    sheet.getRange(1, 1, 1, headers[name].length).setBackground('#1a5276');
    sheet.getRange(1, 1, 1, headers[name].length).setFontColor('white');
  }
}

function generateId() {
  return Utilities.getUuid().split('-')[0];
}

function formatDate(date) {
  const d = new Date(date);
  return Utilities.formatDate(d, 'Asia/Bangkok', 'yyyy-MM-dd');
}

function formatDateTime(date) {
  const d = new Date(date);
  return Utilities.formatDate(d, 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
}

function findRowByValue(sheet, column, value) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][column - 1] == value) {
      return i + 1;
    }
  }
  return -1;
}

function getDriveFolder() {
  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  return folder;
}

function getSubFolder(name) {
  const parent = getDriveFolder();
  const folders = parent.getFoldersByName(name);
  if (folders.hasNext()) {
    return folders.next();
  }
  return parent.createFolder(name);
}

// ============================================================
// Authentication
// ============================================================

function login(data) {
  try {
    const { username, password } = data;

    if (username === 'admin' && password === ADMIN_PASSWORD) {
      return {
        success: true,
        user: {
          username: 'admin',
          fullName: 'ผู้ดูแลระบบ',
          role: 'admin',
          department: 'ทั้งหมด'
        }
      };
    }

    const sheet = getSheet(SHEETS.USERS);
    const rowData = sheet.getDataRange().getValues();

    for (let i = 1; i < rowData.length; i++) {
      if (rowData[i][0] === username && rowData[i][1] === password && rowData[i][5] !== false) {
        return {
          success: true,
          user: {
            username: rowData[i][0],
            fullName: rowData[i][2],
            role: rowData[i][3],
            department: rowData[i][4]
          }
        };
      }
    }

    return { success: false, message: 'รหัสผ่านไม่ถูกต้อง' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

// ============================================================
// Booking Management
// ============================================================

function createBooking(data) {
  try {
    const sheet = getSheet(SHEETS.BOOKING);
    const id = generateId();

    const conflicting = checkBookingConflict(data.date, data.time, data.period);
    if (conflicting) {
      return { success: false, message: 'วัน-เวลานี้มีการจองแล้ว กรุณาเลือกวัน-เวลาอื่น' };
    }

    const row = [
      formatDateTime(new Date()),
      data.date,
      data.time,
      data.teacherName,
      data.department,
      data.period,
      data.subjectName,
      data.subjectCode,
      data.classLevel,
      data.room,
      STATUS.PENDING,
      data.notes || ''
    ];

    sheet.appendRow(row);

    return {
      success: true,
      message: 'จองวันนิเทศสำเร็จ รอการยืนยันจากผู้ดูแลระบบ',
      bookingId: id
    };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function checkBookingConflict(date, time, period) {
  const sheet = getSheet(SHEETS.BOOKING);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === date && data[i][2] === time && data[i][5] === period) {
      if (data[i][10] !== STATUS.REJECTED) {
        return true;
      }
    }
  }
  return false;
}

function updateBooking(data) {
  try {
    const sheet = getSheet(SHEETS.BOOKING);
    const rowNum = data.rowNumber;

    if (rowNum < 2) {
      return { success: false, message: 'ไม่พบรายการจอง' };
    }

    const currentData = sheet.getRange(rowNum, 1, 1, 12).getValues()[0];

    if (data.status) currentData[10] = data.status;
    if (data.notes !== undefined) currentData[11] = data.notes;
    if (data.date) currentData[1] = data.date;
    if (data.time) currentData[2] = data.time;
    if (data.period) currentData[5] = data.period;

    sheet.getRange(rowNum, 1, 1, 12).setValues([currentData]);

    return { success: true, message: 'อัพเดทการจองสำเร็จ' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function deleteBooking(data) {
  try {
    const sheet = getSheet(SHEETS.BOOKING);
    sheet.deleteRow(data.rowNumber);
    return { success: true, message: 'ลบการจองสำเร็จ' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function getBookings(data) {
  try {
    const sheet = getSheet(SHEETS.BOOKING);
    const allData = sheet.getDataRange().getValues();
    const headers = allData[0];
    let bookings = [];

    for (let i = 1; i < allData.length; i++) {
      const row = {};
      row.rowNumber = i + 1;
      for (let j = 0; j < headers.length; j++) {
        row[headers[j]] = allData[i][j];
      }
      row.bookingIndex = i;

      if (data && data.status && allData[i][10] !== data.status) continue;
      if (data && data.teacherName && allData[i][3] !== data.teacherName) continue;
      if (data && data.date && allData[i][1] !== data.date) continue;
      if (data && data.department && allData[i][4] !== data.department) continue;

      bookings.push(row);
    }

    bookings.reverse();

    return { success: true, bookings: bookings };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function getBookingById(data) {
  try {
    const sheet = getSheet(SHEETS.BOOKING);
    const allData = sheet.getDataRange().getValues();
    const headers = allData[0];

    for (let i = 1; i < allData.length; i++) {
      if (i + 1 === data.rowNumber || allData[i][0] === data.bookingId) {
        const row = {};
        row.rowNumber = i + 1;
        for (let j = 0; j < headers.length; j++) {
          row[headers[j]] = allData[i][j];
        }
        return { success: true, booking: row };
      }
    }

    return { success: false, message: 'ไม่พบรายการจอง' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

// ============================================================
// File Management
// ============================================================

function uploadFile(data) {
  try {
    const sheet = getSheet(SHEETS.FILES);

    const fileFolderMap = {
      'แผนการสอน': 'Plans',
      'สื่อการสอน': 'Media',
      'ภาพกิจกรรม': 'Photos',
      'คลิปวิดีโอ': 'Clips'
    };

    let fileUrl = data.fileUrl || '';
    let driveFileId = '';
    let fileName = data.fileName || '';

    if (data.fileData && data.fileName) {
      const folderName = fileFolderMap[data.fileType] || 'Other';
      const folder = getSubFolder(folderName);

      const decoded = Utilities.base64Decode(data.fileData.split(',')[1] || data.fileData);
      const blob = Utilities.newBlob(decoded, data.mimeType || 'application/octet-stream', data.fileName);
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

      fileUrl = file.getUrl();
      driveFileId = file.getId();
      fileName = file.getName();
    }

    const row = [
      formatDateTime(new Date()),
      data.teacherName,
      data.bookingId || '',
      data.fileType,
      fileName,
      fileUrl,
      driveFileId,
      FILE_STATUS.PENDING,
      '',
      '',
      ''
    ];

    sheet.appendRow(row);

    return {
      success: true,
      message: 'อัพโหลดไฟล์สำเร็จ',
      fileUrl: fileUrl
    };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function updateFileStatus(data) {
  try {
    const sheet = getSheet(SHEETS.FILES);
    const rowNum = data.rowNumber;

    if (rowNum < 2) {
      return { success: false, message: 'ไม่พบรายการไฟล์' };
    }

    const currentData = sheet.getRange(rowNum, 1, 1, 11).getValues()[0];

    if (data.status) currentData[7] = data.status;
    if (data.adminNote !== undefined) currentData[8] = data.adminNote;
    if (data.reviewedBy) currentData[9] = data.reviewedBy;
    currentData[10] = formatDateTime(new Date());

    sheet.getRange(rowNum, 1, 1, 11).setValues([currentData]);

    return { success: true, message: 'อัพเดทสถานะไฟล์สำเร็จ' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function getFiles(data) {
  try {
    const sheet = getSheet(SHEETS.FILES);
    const allData = sheet.getDataRange().getValues();
    const headers = allData[0];
    let files = [];

    for (let i = 1; i < allData.length; i++) {
      const row = {};
      row.rowNumber = i + 1;
      for (let j = 0; j < headers.length; j++) {
        row[headers[j]] = allData[i][j];
      }

      if (data && data.status && allData[i][7] !== data.status) continue;
      if (data && data.teacherName && allData[i][1] !== data.teacherName) continue;
      if (data && data.fileType && allData[i][3] !== data.fileType) continue;
      if (data && data.bookingId && allData[i][2] !== data.bookingId) continue;

      files.push(row);
    }

    files.reverse();

    return { success: true, files: files };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function deleteFile(data) {
  try {
    const sheet = getSheet(SHEETS.FILES);

    if (data.driveFileId) {
      try {
        DriveApp.getFileById(data.driveFileId).setTrashed(true);
      } catch (e) {
        // File may not exist
      }
    }

    sheet.deleteRow(data.rowNumber);
    return { success: true, message: 'ลบไฟล์สำเร็จ' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

// ============================================================
// Evaluation Management
// ============================================================

function createEvaluation(data) {
  try {
    const sheet = getSheet(SHEETS.EVALUATION);

    const row = [
      formatDateTime(new Date()),
      data.teacherName,
      data.supervisionDate,
      data.bookingId || '',
      data.strengths,
      data.improvements,
      data.suggestions,
      data.qualityLevel,
      data.score || '',
      data.evaluatedBy || '',
      data.bookingReference || ''
    ];

    sheet.appendRow(row);

    if (data.bookingId) {
      try {
        const bookingSheet = getSheet(SHEETS.BOOKING);
        const allBookings = bookingSheet.getDataRange().getValues();
        for (let i = 1; i < allBookings.length; i++) {
          if (allBookings[i][0] === data.bookingId) {
            bookingSheet.getRange(i + 1, 11).setValue(STATUS.COMPLETED);
            break;
          }
        }
      } catch (e) {
        // Continue even if booking update fails
      }
    }

    return {
      success: true,
      message: 'บันทึกผลการประเมินสำเร็จ'
    };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function updateEvaluation(data) {
  try {
    const sheet = getSheet(SHEETS.EVALUATION);
    const rowNum = data.rowNumber;

    if (rowNum < 2) {
      return { success: false, message: 'ไม่พบรายการประเมิน' };
    }

    const currentData = sheet.getRange(rowNum, 1, 1, 11).getValues()[0];

    if (data.strengths !== undefined) currentData[4] = data.strengths;
    if (data.improvements !== undefined) currentData[5] = data.improvements;
    if (data.suggestions !== undefined) currentData[6] = data.suggestions;
    if (data.qualityLevel) currentData[7] = data.qualityLevel;
    if (data.score !== undefined) currentData[8] = data.score;

    sheet.getRange(rowNum, 1, 1, 11).setValues([currentData]);

    return { success: true, message: 'อัพเดทผลการประเมินสำเร็จ' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function getEvaluations(data) {
  try {
    const sheet = getSheet(SHEETS.EVALUATION);
    const allData = sheet.getDataRange().getValues();
    const headers = allData[0];
    let evaluations = [];

    for (let i = 1; i < allData.length; i++) {
      const row = {};
      row.rowNumber = i + 1;
      for (let j = 0; j < headers.length; j++) {
        row[headers[j]] = allData[i][j];
      }

      if (data && data.teacherName && allData[i][1] !== data.teacherName) continue;
      if (data && data.qualityLevel && allData[i][7] !== data.qualityLevel) continue;

      evaluations.push(row);
    }

    evaluations.reverse();

    return { success: true, evaluations: evaluations };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function deleteEvaluation(data) {
  try {
    const sheet = getSheet(SHEETS.EVALUATION);
    sheet.deleteRow(data.rowNumber);
    return { success: true, message: 'ลบผลการประเมินสำเร็จ' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

// ============================================================
// Dashboard & Statistics
// ============================================================

function getDashboardStats(data) {
  try {
    const bookingSheet = getSheet(SHEETS.BOOKING);
    const fileSheet = getSheet(SHEETS.FILES);
    const evalSheet = getSheet(SHEETS.EVALUATION);

    const bookingData = bookingSheet.getDataRange().getValues();
    const fileData = fileSheet.getDataRange().getValues();
    const evalData = evalSheet.getDataRange().getValues();

    const today = formatDate(new Date());
    const thisMonth = today.substring(0, 7);

    let stats = {
      totalBookings: bookingData.length - 1,
      pendingBookings: 0,
      confirmedBookings: 0,
      completedBookings: 0,
      rejectedBookings: 0,
      todayBookings: 0,
      thisMonthBookings: 0,

      totalFiles: fileData.length - 1,
      pendingFiles: 0,
      approvedFiles: 0,
      reviseFiles: 0,

      totalEvaluations: evalData.length - 1,
      excellentCount: 0,
      goodCount: 0,
      fairCount: 0,
      poorCount: 0,

      recentBookings: [],
      recentFiles: [],
      recentEvaluations: [],

      departmentStats: {},
      teacherStats: {}
    };

    for (let i = 1; i < bookingData.length; i++) {
      const status = bookingData[i][10];
      const bookingDate = bookingData[i][1];

      if (status === STATUS.PENDING) stats.pendingBookings++;
      if (status === STATUS.CONFIRMED) stats.confirmedBookings++;
      if (status === STATUS.COMPLETED) stats.completedBookings++;
      if (status === STATUS.REJECTED) stats.rejectedBookings++;
      if (bookingDate === today) stats.todayBookings++;
      if (bookingDate && bookingDate.toString().startsWith(thisMonth)) stats.thisMonthBookings++;

      const dept = bookingData[i][4];
      if (dept) {
        if (!stats.departmentStats[dept]) {
          stats.departmentStats[dept] = { total: 0, completed: 0, pending: 0 };
        }
        stats.departmentStats[dept].total++;
        if (status === STATUS.COMPLETED) stats.departmentStats[dept].completed++;
        if (status === STATUS.PENDING) stats.departmentStats[dept].pending++;
      }

      const teacher = bookingData[i][3];
      if (teacher) {
        if (!stats.teacherStats[teacher]) {
          stats.teacherStats[teacher] = { total: 0, completed: 0, files: 0 };
        }
        stats.teacherStats[teacher].total++;
        if (status === STATUS.COMPLETED) stats.teacherStats[teacher].completed++;
      }
    }

    for (let i = 1; i < fileData.length; i++) {
      const status = fileData[i][7];
      if (status === FILE_STATUS.PENDING) stats.pendingFiles++;
      if (status === FILE_STATUS.APPROVED) stats.approvedFiles++;
      if (status === FILE_STATUS.REVISE) stats.reviseFiles++;

      const teacher = fileData[i][1];
      if (teacher && stats.teacherStats[teacher]) {
        stats.teacherStats[teacher].files++;
      }
    }

    for (let i = 1; i < evalData.length; i++) {
      const level = evalData[i][7];
      if (level === QUALITY_LEVELS.EXCELLENT) stats.excellentCount++;
      if (level === QUALITY_LEVELS.GOOD) stats.goodCount++;
      if (level === QUALITY_LEVELS.FAIR) stats.fairCount++;
      if (level === QUALITY_LEVELS.POOR) stats.poorCount++;
    }

    stats.recentBookings = bookingData.slice(-10).reverse().map((row, idx) => ({
      rowNumber: bookingData.length - idx,
      date: row[1],
      time: row[2],
      teacherName: row[3],
      department: row[4],
      period: row[5],
      subjectName: row[6],
      status: row[10]
    }));

    stats.recentFiles = fileData.slice(-10).reverse().map((row, idx) => ({
      rowNumber: fileData.length - idx,
      teacherName: row[1],
      fileType: row[3],
      fileName: row[4],
      status: row[7]
    }));

    stats.recentEvaluations = evalData.slice(-5).reverse().map((row, idx) => ({
      rowNumber: evalData.length - idx,
      teacherName: row[1],
      supervisionDate: row[2],
      qualityLevel: row[7],
      evaluatedBy: row[9]
    }));

    return { success: true, stats: stats };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function getCalendarData(data) {
  try {
    const sheet = getSheet(SHEETS.BOOKING);
    const allData = sheet.getDataRange().getValues();
    let events = [];

    const month = data.month || new Date().getMonth();
    const year = data.year || new Date().getFullYear();

    for (let i = 1; i < allData.length; i++) {
      const bookingDate = allData[i][1];
      if (!bookingDate) continue;

      const d = new Date(bookingDate);
      if (d.getMonth() == month && d.getFullYear() == year) {
        events.push({
          date: bookingDate,
          title: allData[i][3] + ' - ' + allData[i][6],
          time: allData[i][2],
          status: allData[i][10],
          department: allData[i][4]
        });
      }
    }

    return { success: true, events: events };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

// ============================================================
// Reports
// ============================================================

function getReport(data) {
  try {
    const type = data.type || 'all';
    const startDate = data.startDate || '';
    const endDate = data.endDate || '';

    const bookingSheet = getSheet(SHEETS.BOOKING);
    const fileSheet = getSheet(SHEETS.FILES);
    const evalSheet = getSheet(SHEETS.EVALUATION);

    let report = {
      bookings: [],
      files: [],
      evaluations: [],
      summary: {}
    };

    const bookingData = bookingSheet.getDataRange().getValues();
    for (let i = 1; i < bookingData.length; i++) {
      const d = bookingData[i][1];
      if (startDate && d < startDate) continue;
      if (endDate && d > endDate) continue;
      report.bookings.push({
        date: d,
        time: bookingData[i][2],
        teacherName: bookingData[i][3],
        department: bookingData[i][4],
        period: bookingData[i][5],
        subjectName: bookingData[i][6],
        status: bookingData[i][10]
      });
    }

    const fileData = fileSheet.getDataRange().getValues();
    for (let i = 1; i < fileData.length; i++) {
      report.files.push({
        timestamp: fileData[i][0],
        teacherName: fileData[i][1],
        fileType: fileData[i][3],
        fileName: fileData[i][4],
        status: fileData[i][7]
      });
    }

    const evalData = evalSheet.getDataRange().getValues();
    for (let i = 1; i < evalData.length; i++) {
      report.evaluations.push({
        teacherName: evalData[i][1],
        supervisionDate: evalData[i][2],
        strengths: evalData[i][4],
        improvements: evalData[i][5],
        qualityLevel: evalData[i][7]
      });
    }

    report.summary = {
      totalBookings: report.bookings.length,
      completedBookings: report.bookings.filter(b => b.status === STATUS.COMPLETED).length,
      totalFiles: report.files.length,
      approvedFiles: report.files.filter(f => f.status === FILE_STATUS.APPROVED).length,
      totalEvaluations: report.evaluations.length
    };

    return { success: true, report: report };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function getTeacherReport(data) {
  try {
    const teacherName = data.teacherName;

    const bookingSheet = getSheet(SHEETS.BOOKING);
    const fileSheet = getSheet(SHEETS.FILES);
    const evalSheet = getSheet(SHEETS.EVALUATION);

    let report = {
      teacherName: teacherName,
      bookings: [],
      files: [],
      evaluations: [],
      summary: {}
    };

    const bookingData = bookingSheet.getDataRange().getValues();
    for (let i = 1; i < bookingData.length; i++) {
      if (bookingData[i][3] === teacherName) {
        report.bookings.push({
          date: bookingData[i][1],
          time: bookingData[i][2],
          subjectName: bookingData[i][6],
          status: bookingData[i][10]
        });
      }
    }

    const fileData = fileSheet.getDataRange().getValues();
    for (let i = 1; i < fileData.length; i++) {
      if (fileData[i][1] === teacherName) {
        report.files.push({
          fileType: fileData[i][3],
          fileName: fileData[i][4],
          status: fileData[i][7]
        });
      }
    }

    const evalData = evalSheet.getDataRange().getValues();
    for (let i = 1; i < evalData.length; i++) {
      if (evalData[i][1] === teacherName) {
        report.evaluations.push({
          supervisionDate: evalData[i][2],
          strengths: evalData[i][4],
          improvements: evalData[i][5],
          suggestions: evalData[i][6],
          qualityLevel: evalData[i][7]
        });
      }
    }

    report.summary = {
      totalBookings: report.bookings.length,
      completedBookings: report.bookings.filter(b => b.status === STATUS.COMPLETED).length,
      totalFiles: report.files.length,
      approvedFiles: report.files.filter(f => f.status === FILE_STATUS.APPROVED).length,
      totalEvaluations: report.evaluations.length,
      avgQuality: getAvgQuality(report.evaluations)
    };

    return { success: true, report: report };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function getDepartmentReport(data) {
  try {
    const department = data.department;

    const bookingSheet = getSheet(SHEETS.BOOKING);
    const evalSheet = getSheet(SHEETS.EVALUATION);

    let report = {
      department: department,
      teachers: {},
      summary: {}
    };

    const bookingData = bookingSheet.getDataRange().getValues();
    for (let i = 1; i < bookingData.length; i++) {
      if (bookingData[i][4] === department) {
        const teacher = bookingData[i][3];
        if (!report.teachers[teacher]) {
          report.teachers[teacher] = {
            totalBookings: 0,
            completedBookings: 0,
            files: 0,
            evaluations: []
          };
        }
        report.teachers[teacher].totalBookings++;
        if (bookingData[i][10] === STATUS.COMPLETED) {
          report.teachers[teacher].completedBookings++;
        }
      }
    }

    const evalData = evalSheet.getDataRange().getValues();
    for (let i = 1; i < evalData.length; i++) {
      const teacher = evalData[i][1];
      if (report.teachers[teacher]) {
        report.teachers[teacher].evaluations.push({
          date: evalData[i][2],
          qualityLevel: evalData[i][7]
        });
      }
    }

    const teacherList = Object.keys(report.teachers);
    report.summary = {
      totalTeachers: teacherList.length,
      totalBookings: teacherList.reduce((sum, t) => sum + report.teachers[t].totalBookings, 0),
      completedBookings: teacherList.reduce((sum, t) => sum + report.teachers[t].completedBookings, 0)
    };

    return { success: true, report: report };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function getAvgQuality(evaluations) {
  if (evaluations.length === 0) return 'ไม่มีข้อมูล';

  const scores = {
    [QUALITY_LEVELS.EXCELLENT]: 4,
    [QUALITY_LEVELS.GOOD]: 3,
    [QUALITY_LEVELS.FAIR]: 2,
    [QUALITY_LEVELS.POOR]: 1
  };

  let total = 0;
  let count = 0;
  evaluations.forEach(e => {
    if (scores[e.qualityLevel]) {
      total += scores[e.qualityLevel];
      count++;
    }
  });

  if (count === 0) return 'ไม่มีข้อมูล';
  const avg = total / count;
  if (avg >= 3.5) return QUALITY_LEVELS.EXCELLENT;
  if (avg >= 2.5) return QUALITY_LEVELS.GOOD;
  if (avg >= 1.5) return QUALITY_LEVELS.FAIR;
  return QUALITY_LEVELS.POOR;
}

// ============================================================
// User Management (Admin)
// ============================================================

function createUser(data) {
  try {
    const sheet = getSheet(SHEETS.USERS);

    const existing = findRowByValue(sheet, 1, data.username);
    if (existing !== -1) {
      return { success: false, message: 'ชื่อผู้ใช้นี้มีอยู่แล้ว' };
    }

    sheet.appendRow([
      data.username,
      data.password,
      data.fullName,
      data.role || 'teacher',
      data.department || '',
      true
    ]);

    return { success: true, message: 'สร้างผู้ใช้สำเร็จ' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function getUsers() {
  try {
    const sheet = getSheet(SHEETS.USERS);
    const allData = sheet.getDataRange().getValues();
    let users = [];

    for (let i = 1; i < allData.length; i++) {
      users.push({
        rowNumber: i + 1,
        username: allData[i][0],
        fullName: allData[i][2],
        role: allData[i][3],
        department: allData[i][4],
        active: allData[i][5]
      });
    }

    return { success: true, users: users };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

// ============================================================
// Get Unique Values for Dropdowns
// ============================================================

function getDepartments() {
  const departments = [
    'คณิตศาสตร์', 'วิทยาศาสตร์', 'ภาษาไทย', 'ภาษาอังกฤษ',
    'สังคมศึกษา', 'พลศึกษา', 'ศิลปะ', 'การงานอาชีพ',
    'เทคโนโลยี', 'แนะแนว', '其他'
  ];
  return { success: true, departments: departments };
}

function getTeachers() {
  try {
    const sheet = getSheet(SHEETS.BOOKING);
    const allData = sheet.getDataRange().getValues();
    const teachers = new Set();

    for (let i = 1; i < allData.length; i++) {
      if (allData[i][3]) teachers.add(allData[i][3]);
    }

    return { success: true, teachers: Array.from(teachers).sort() };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function getSubjects() {
  const subjects = [
    { name: 'คณิตศาสตร์', code: 'ค21101' },
    { name: 'วิทยาศาสตร์', code: 'ว21101' },
    { name: 'ภาษาไทย', code: 'ท21101' },
    { name: 'ภาษาอังกฤษ', code: 'อ21101' },
    { name: 'สังคมศึกษา', code: 'ส21101' },
    { name: 'พลศึกษา', code: 'พ21101' },
    { name: 'ศิลปะ', code: 'ศ21101' },
    { name: 'การงานอาชีพ', code: 'ก21101' },
    { name: 'เทคโนโลยี', code: 'ท21102' }
  ];
  return { success: true, subjects: subjects };
}

function getTimeSlots() {
  const slots = [
    { time: '08:30-09:00', period: 'คาบที่ 1' },
    { time: '09:00-09:30', period: 'คาบที่ 2' },
    { time: '09:30-10:00', period: 'คาบที่ 3' },
    { time: '10:00-10:30', period: 'พักรับประทานอาหาร' },
    { time: '10:30-11:00', period: 'คาบที่ 4' },
    { time: '11:00-11:30', period: 'คาบที่ 5' },
    { time: '11:30-12:00', period: 'คาบที่ 6' },
    { time: '13:00-13:30', period: 'คาบที่ 7' },
    { time: '13:30-14:00', period: 'คาบที่ 8' },
    { time: '14:00-14:30', period: 'คาบที่ 9' }
  ];
  return { success: true, timeSlots: slots };
}

function getClassLevels() {
  const levels = [
    'ม.1', 'ม.2', 'ม.3', 'ม.4', 'ม.5', 'ม.6'
  ];
  return { success: true, classLevels: levels };
}
