// ======================================
// CONFIGURATION VARIABLES
// ======================================
const SPREADSHEET_ID = '1xCWWNcJdTPtKDzbRTpbtzeGZtb-gKbJj4YnOnnSZx50';
const NAME_SHEET = 'name';
const DATA_SHEET = 'data_check';
const DRIVE_FOLDER_ID = '1w6kvH6Og847MXv8BJjVuJiCeme9L7H0c';
const REPORT_FOLDER_ID = '1f58T3IbvOyHX6pRzYHGHWnTw06iwj79p';
const REPORT_PASSWORD = '2550'; // รหัสผ่านสำหรับเข้าถึงหน้ารายงาน


// ======================================
// THAI CONSTANTS
// ======================================
const THAI_MONTHS = [
  'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน',
  'พฤษภาคม', 'มิถุนายน', 'กรกฎาคม', 'สิงหาคม',
  'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'
];


const THAI_DAYS = [
  'วันอาทิตย์', 'วันจันทร์', 'วันอังคาร', 'วันพุธ',
  'วันพฤหัสบดี', 'วันศุกร์', 'วันเสาร์'
];


// ======================================
// ROUTING - แสดงหน้าหลักที่มีทั้ง 2 แท็บ
// ======================================
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('main')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setTitle('ระบบลงชื่อเข้า-ออกเวร');
}


function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


// ======================================
// PASSWORD VERIFICATION
// ======================================
function verifyReportPassword(password) {
  return password === REPORT_PASSWORD;
}


// ======================================
// FUNCTIONS FOR CHECK-IN/OUT SYSTEM
// ======================================
function getNames() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(NAME_SHEET);
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1).getValues();
  return data.map(row => row[0]);
}


function checkDuplicateCheckIn(name) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(DATA_SHEET);
  const lastRow = sheet.getLastRow();
 
  if (lastRow < 2) {
    return { hasDuplicate: false };
  }
 
  const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  const today = new Date();
  const todayDateString = Utilities.formatDate(today, "Asia/Bangkok", "yyyy-MM-dd");
 
  const todayCheckIns = data.filter(row => {
    const rowDate = new Date(row[0]);
    const rowDateString = Utilities.formatDate(rowDate, "Asia/Bangkok", "yyyy-MM-dd");
   
    return row[1] === name &&
           row[2] === "เข้างาน" &&
           rowDateString === todayDateString;
  });
 
  if (todayCheckIns.length > 0) {
    const lastCheckInTime = todayCheckIns[todayCheckIns.length - 1][0];
    return {
      hasDuplicate: true,
      lastCheckInTime: Utilities.formatDate(lastCheckInTime, "Asia/Bangkok", "HH:mm:ss"),
      message: `คุณได้เข้างานในวันนี้แล้ว เวลา ${Utilities.formatDate(lastCheckInTime, "Asia/Bangkok", "HH:mm:ss")}`
    };
  }

  return { hasDuplicate: false };
}


// ฟังก์ชันตรวจสอบการออกงานซ้ำในวันเดียวกัน
function checkDuplicateCheckOut(name) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(DATA_SHEET);
  const lastRow = sheet.getLastRow();
 
  if (lastRow < 2) {
    return { hasDuplicate: false };
  }
 
  const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  const today = new Date();
  const todayDateString = Utilities.formatDate(today, "Asia/Bangkok", "yyyy-MM-dd");
 
  const todayCheckOuts = data.filter(row => {
    const rowDate = new Date(row[0]);
    const rowDateString = Utilities.formatDate(rowDate, "Asia/Bangkok", "yyyy-MM-dd");
   
    return row[1] === name &&
           row[2] === "ออกงาน" &&
           rowDateString === todayDateString;
  });
 
  if (todayCheckOuts.length > 0) {
    const lastCheckOutTime = todayCheckOuts[todayCheckOuts.length - 1][0];
    return {
      hasDuplicate: true,
      lastCheckOutTime: Utilities.formatDate(lastCheckOutTime, "Asia/Bangkok", "HH:mm:ss"),
      message: `คุณได้ออกงานในวันนี้แล้ว เวลา ${Utilities.formatDate(lastCheckOutTime, "Asia/Bangkok", "HH:mm:ss")}`
    };
  }

 
  return { hasDuplicate: false };
}


function checkLastCheckIn(name) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(DATA_SHEET);
  const lastRow = sheet.getLastRow();
 
  if (lastRow < 2) {
    return { hasCheckIn: false, message: "คุณไม่ได้ลงชื่อเข้างาน 24 ชั่วโมงที่ผ่านมา" };
  }
 
  const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  const userCheckIns = data.filter(row => row[1] === name && row[2] === "เข้างาน");
 
  if (userCheckIns.length === 0) {
    return { hasCheckIn: false, message: "คุณไม่ได้ลงชื่อเข้างาน 24 ชั่วโมงที่ผ่านมา" };
  }
 
  userCheckIns.sort((a, b) => new Date(b[0]) - new Date(a[0]));
  const lastCheckIn = userCheckIns[0][0];
  const now = new Date();
  const timeDiff = now - lastCheckIn;
  const hoursDiff = timeDiff / (1000 * 60 * 60);
 
  if (hoursDiff > 24) {
    return { hasCheckIn: false, message: "คุณไม่ได้ลงชื่อเข้างาน 24 ชั่วโมงที่ผ่านมา" };
  }
 
  return {
    hasCheckIn: true,
    lastCheckIn: lastCheckIn,
    checkInTime: Utilities.formatDate(lastCheckIn, "Asia/Bangkok", "HH:mm:ss")
  };
}


function calculateWorkTime(checkInTime, checkOutTime) {
  const diff = checkOutTime - checkInTime;
  const hours = Math.floor(diff / (1000 * 60 * 60));
  const minutes = Math.floor((diff % (1000 * 60 * 60)) / (1000 * 60));
  const seconds = Math.floor((diff % (1000 * 60)) / 1000);
 
  return {
    hours: hours,
    minutes: minutes,
    seconds: seconds,
    display: `${hours} ชั่วโมง ${minutes} นาที ${seconds} วินาที`
  };
}



function uploadImage(base64Data) {
  try {
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'image/jpeg', 'checkin_image.jpg');
    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    const file = folder.createFile(blob);
    file.setName(`Image_${new Date().getTime()}.jpg`);
   
    return file.getUrl();
  } catch (error) {
    Logger.log('Error uploading image: ' + error.toString());
    throw new Error('ไม่สามารถอัพโหลดภาพได้');
  }
}


function saveData(data) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(DATA_SHEET);
 
  if (!data.name || !data.checkType || !data.latitude || !data.longitude) {
    throw new Error("กรุณากรอกข้อมูลให้ครบถ้วน");
  }
 

 
  let checkInInfo = null;
  let workTimeInfo = null;
 
  if (data.checkType === "ออกงาน") {
    const checkInStatus = checkLastCheckIn(data.name);
    if (!checkInStatus.hasCheckIn) {
      throw new Error(checkInStatus.message);
    }
    checkInInfo = checkInStatus;
   
    const now = new Date();
    workTimeInfo = calculateWorkTime(checkInInfo.lastCheckIn, now);
  }
 
  let displayMessage = "";
  if (data.checkType === "เข้างาน") {
    displayMessage = "บันทึกการเข้างานสำเร็จ";
  } else {
    displayMessage = `คุณได้ออกงานเรียบร้อยแล้ว\nคุณเข้างานเวลา ${checkInInfo.checkInTime}\nคุณใช้เวลาทำงาน ${workTimeInfo.display}`;
  }
 
  const now = new Date();
  const thaiDay = THAI_DAYS[now.getDay()];
  const thaiDate = now.getDate();
  const thaiMonth = THAI_MONTHS[now.getMonth()];
  const thaiYear = now.getFullYear() + 543;
  const thaiTime = Utilities.formatDate(now, "Asia/Bangkok", "HH:mm:ss");
  const thaiDateTime = `${thaiDay}ที่ ${thaiDate} ${thaiMonth} ${thaiYear} เวลา ${thaiTime} น.`;
 
  const workTimeDisplay = data.checkType === "ออกงาน" ? workTimeInfo.display : "";
 
  const rowData = [
    now,
    data.name,
    data.checkType,
    data.latitude,
    data.longitude,
    thaiDateTime,
    data.imageUrl || "",
    workTimeDisplay,
    "" // คอลัมน์สถานะเวลา - เว้นว่างไว้เนื่องจากไม่ตรวจสอบแล้ว
  ];
 
  sheet.appendRow(rowData);
 
  return {
    success: true,
    message: displayMessage,
    timestamp: thaiDateTime,
    timeStatus: "", // ไม่มีการตรวจสอบสถานะเวลาแล้ว
    checkInTime: data.checkType === "ออกงาน" ? checkInInfo.checkInTime : null,
    workTime: workTimeDisplay
  };
}


// ======================================
// FUNCTIONS FOR REPORT SYSTEM
// ======================================
function getData() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(DATA_SHEET);
  const allData = sheet.getDataRange().getValues();


  if (allData.length <= 1) return { rows: [], uniqueEValues: [] };


  const rowsWithHeader = allData.slice(1);


  const uniqueEValues = [...new Set(rowsWithHeader.map(row => row[2]).filter(Boolean))].sort();


  let sortedRows = rowsWithHeader.sort((a, b) => new Date(b[0]) - new Date(a[0]));
  sortedRows = sortedRows.slice(0, 300);


  const thaiMonths = [
    "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
    "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"
  ];


  const formattedRows = sortedRows.map(row => {
    const date = new Date(row[0]);
    const day = date.getDate();
    const month = thaiMonths[date.getMonth()];
    const year = date.getFullYear() + 543;
    const time = date.toLocaleTimeString('th-TH', { hour: '2-digit', minute: '2-digit' });
    const thaiDate = `${day} ${month} ${year} เวลา ${time} น.`;


    const yyyy = date.getFullYear();
    const mm = String(date.getMonth() + 1).padStart(2, '0');
    const dd = String(date.getDate()).padStart(2, '0');
    const isoDate = `${yyyy}-${mm}-${dd}`;


    let thumbnailUrl = '';
    if (row[6] && typeof row[6] === 'string' && row[6].includes('drive.google.com')) {
      const fileId = extractFileId(row[6]);
      if (fileId) {
        thumbnailUrl = `https://drive.google.com/thumbnail?sz=w300-h400&id=${fileId}`;
      }
    }


    return [
      thaiDate,           // 0
      row[1] || '',       // 1 - ชื่อ
      row[2] || '',       // 2 - ประเภท
      row[8] || '',       // 3 - สถานะ
      row[6] || '',       // 4 - fileUrl
      thumbnailUrl,       // 5 - thumbnailUrl
      isoDate,            // 6 - isoDate
      row[7] || '',       // 7 - workSummary
      row[3] || '',       // 8 - latitude
      row[4] || ''        // 9 - longitude
    ];
  });


   return {
    rows: formattedRows,
    uniqueEValues: uniqueEValues
  };
}


function extractFileId(url) {
  const patterns = [
    /\/open\?id=([^&]+)/,
    /\/file\/d\/([^\/]+)/,
    /\/uc\?id=([^&]+)/,
    /[&?]id=([^&]+)/
  ];
  for (const pattern of patterns) {
    const match = url.match(pattern);
    if (match && match[1]) {
      return match[1];
    }
  }
  return null;
}


function formatThaiDateTime(date) {
  const thaiDay = THAI_DAYS[date.getDay()];
  const thaiDate = date.getDate();
  const thaiMonth = THAI_MONTHS[date.getMonth()];
  const thaiYear = date.getFullYear() + 543;
  const thaiTime = Utilities.formatDate(date, "Asia/Bangkok", "HH:mm:ss");
  return `${thaiDay}ที่ ${thaiDate} ${thaiMonth} ${thaiYear} เวลา ${thaiTime} น.`;
}


function generateReport(filteredData, filterOptions) {
  try {
    const folder = DriveApp.getFolderById(REPORT_FOLDER_ID);
   
    const today = new Date();
    const doc = DocumentApp.create(`รายงาน Check-in - ${formatThaiDateTime(today)}`);
    const body = doc.getBody();
   
    body.setMarginTop(50);
    body.setMarginBottom(50);
    body.setMarginLeft(70);
    body.setMarginRight(70);
   
    // ตั้งค่าหัวเรื่อง
    const titleStyle = {};
    titleStyle[DocumentApp.Attribute.FONT_SIZE] = 20;
    titleStyle[DocumentApp.Attribute.BOLD] = true;
    titleStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Prompt';
    titleStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
    titleStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#000000';
   
    const title = body.appendParagraph('รายงานระบบ Check-in');
    title.setAttributes(titleStyle);
    title.setSpacingAfter(6);
   
    // หัวเรื่องรอง
    const subtitle = body.appendParagraph('เรือนจำอำเภอสีคิ้ว');
    const subtitleStyle = {};
    subtitleStyle[DocumentApp.Attribute.FONT_SIZE] = 14;
    subtitleStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Prompt';
    subtitleStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
    subtitleStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#000000';
    subtitle.setAttributes(subtitleStyle);
    subtitle.setSpacingAfter(12);
   
    body.appendParagraph('');
   
    // เงื่อนไขการกรอง
    const filterSection = body.appendParagraph('เงื่อนไขการกรองข้อมูล');
    const filterSectionStyle = {};
    filterSectionStyle[DocumentApp.Attribute.FONT_SIZE] = 12;
    filterSectionStyle[DocumentApp.Attribute.BOLD] = true;
    filterSectionStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Prompt';
    filterSectionStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#000000';
    filterSection.setAttributes(filterSectionStyle);
    filterSection.setSpacingBefore(6);
    filterSection.setSpacingAfter(6);
   
    const normalStyle = {};
    normalStyle[DocumentApp.Attribute.FONT_SIZE] = 11;
    normalStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Prompt';
    normalStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#000000';
   
    if (filterOptions.searchTerm) {
      const searchPara = body.appendParagraph(`คำค้นหา: ${filterOptions.searchTerm}`);
      searchPara.setAttributes(normalStyle);
    }
    if (filterOptions.agency) {
      const agencyPara = body.appendParagraph(`หน่วยงาน: ${filterOptions.agency}`);
      agencyPara.setAttributes(normalStyle);
    }
    if (filterOptions.date) {
      const datePara = body.appendParagraph(`วันที่: ${filterOptions.date}`);
      datePara.setAttributes(normalStyle);
    }
   
    body.appendParagraph('');
   
    const resultSection = body.appendParagraph(`จำนวนรายการทั้งหมด: ${filteredData.length} รายการ`);
    const resultStyle = {};
    resultStyle[DocumentApp.Attribute.FONT_SIZE] = 11;
    resultStyle[DocumentApp.Attribute.BOLD] = true;
    resultStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Prompt';
    resultStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#000000';
    resultSection.setAttributes(resultStyle);
    resultSection.setSpacingAfter(12);
   
    body.appendHorizontalRule();
    body.appendParagraph('');
   
    // สร้างตาราง
    const table = body.appendTable();
    const headerRow = table.appendTableRow();
    const headers = ['ลำดับ', 'วันที่-เวลา', 'ชื่อ', 'ประเภท', 'เวลาทำงาน'];
   
    // สไตล์ Header
    const headerStyle = {};
    headerStyle[DocumentApp.Attribute.BOLD] = true;
    headerStyle[DocumentApp.Attribute.FONT_SIZE] = 10;
    headerStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Prompt';
    headerStyle[DocumentApp.Attribute.BACKGROUND_COLOR] = '#4A90E2';
    headerStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#FFFFFF';
   
    headers.forEach(headerText => {
      const cell = headerRow.appendTableCell(headerText);
      cell.setAttributes(headerStyle);
      cell.setPaddingTop(8);
      cell.setPaddingBottom(8);
      cell.setPaddingLeft(5);
      cell.setPaddingRight(5);
    });
   
    // สไตล์เนื้อหาในตาราง
    const cellStyle = {};
    cellStyle[DocumentApp.Attribute.FONT_SIZE] = 9;
    cellStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Prompt';
    cellStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#000000';
    cellStyle[DocumentApp.Attribute.BACKGROUND_COLOR] = '#FFFFFF';
   
    filteredData.forEach((row, index) => {
      const dataRow = table.appendTableRow();
      const cells = [
        (index + 1).toString(),
        row[0] || '',
        row[1] || '',
        row[2] || '',
        row[7] || ''
      ];
     
      cells.forEach(cellText => {
        const cell = dataRow.appendTableCell(cellText);
        cell.setAttributes(cellStyle);
        cell.setPaddingTop(6);
        cell.setPaddingBottom(6);
        cell.setPaddingLeft(5);
        cell.setPaddingRight(5);
      });
    });
   
    body.appendParagraph('');
    body.appendHorizontalRule();
   
    // Footer
    const footerPara = body.appendParagraph(`สร้างรายงานเมื่อ: ${formatThaiDateTime(today)}`);
    const footerStyle = {};
    footerStyle[DocumentApp.Attribute.FONT_SIZE] = 9;
    footerStyle[DocumentApp.Attribute.ITALIC] = true;
    footerStyle[DocumentApp.Attribute.FONT_FAMILY] = 'Prompt';
    footerStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = '#666666';
    footerPara.setAttributes(footerStyle);
   
    doc.saveAndClose();
   
    // สร้าง PDF
    const docFile = DriveApp.getFileById(doc.getId());
    const pdfBlob = docFile.getAs('application/pdf');
    const pdfFile = folder.createFile(pdfBlob);
    pdfFile.setName(`รายงาน_Check-in_${Utilities.formatDate(today, "Asia/Bangkok", "yyyyMMdd_HHmmss")}.pdf`);
   
    // ลบไฟล์ Google Docs
    DriveApp.getFileById(doc.getId()).setTrashed(true);
   
    return {
      success: true,
      pdfUrl: pdfFile.getUrl(),
      pdfId: pdfFile.getId(),
      fileName: pdfFile.getName()
    };
   
  } catch (error) {
    Logger.log('Error in generateReport: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

