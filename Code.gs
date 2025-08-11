// Code.gs (สมบูรณ์)

// --------- Core helpers ----------
function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['id', 'date', 'detail', 'floor', 'who', 'status', 'fileName', 'fileId', 'fileUrl', 'createdAt', 'updatedAt', 'updatedBy']);
  }
  return sh;
}

function list() {
  const sh = getSheet_();
  const values = sh.getDataRange().getValues();
  const headers = values.shift();
  const idx = {};
  headers.forEach((h, i) => { idx[h] = i; });
  const out = [];
  values.forEach(row => {
    if (!row[idx.id]) return;
    out.push({
      id: row[idx.id],
      date: row[idx.date],
      detail: row[idx.detail],
      floor: row[idx.floor],
      who: row[idx.who],
      status: row[idx.status],
      fileName: row[idx.fileName],
      fileId: row[idx.fileId],
      fileUrl: row[idx.fileUrl],
      fileType: "", // not stored in sheet; optional
      fileDataURL: "", // intentionally omitted
      createdAt: row[idx.createdAt],
      updatedAt: row[idx.updatedAt],
      updatedBy: row[idx.updatedBy]
    });
  });
  return out;
}

function create(rec) {
  // if rec has base64 file, save to Drive first
  let fileId = "", fileUrl = "";
  if (rec.fileDataURL && rec.fileName) {
    const saved = saveFileFromDataURL_(rec.fileDataURL, rec.fileName);
    fileId = saved.id; fileUrl = saved.url;
  }
  const sh = getSheet_();
  const now = new Date().toISOString();
  const row = [
    rec.id, rec.date, rec.detail, rec.floor, rec.who, rec.status || 'รอดำเนินการ',
    rec.fileName || "", fileId, fileUrl,
    rec.createdAt || now, rec.updatedAt || now, rec.updatedBy || ""
  ];
  sh.appendRow(row);
  rec.fileId = fileId; rec.fileUrl = fileUrl; rec.fileDataURL = ""; // do not echo back base64
  return rec;
}

function update(rec) {
  const sh = getSheet_();
  const values = sh.getDataRange().getValues();
  const headers = values.shift();
  const idx = {};
  headers.forEach((h, i) => { idx[h] = i; });
  let rowIndex = -1;
  for (let r = 0; r < values.length; r++) {
    if (values[r][idx.id] === rec.id) { rowIndex = r + 2; break; } // +2 for header and 1-based
  }
  if (rowIndex < 0) throw new Error('ไม่พบรายการ');

  // if new file provided
  if (rec.fileDataURL && rec.fileName) {
    const saved = saveFileFromDataURL_(rec.fileDataURL, rec.fileName);
    rec.fileId = saved.id; rec.fileUrl = saved.url;
  } else {
    // keep existing file info
    rec.fileName = rec.fileName || sh.getRange(rowIndex, idx.fileName + 1).getValue();
    rec.fileId = rec.fileId || sh.getRange(rowIndex, idx.fileId + 1).getValue();
    rec.fileUrl = rec.fileUrl || sh.getRange(rowIndex, idx.fileUrl + 1).getValue();
  }

  const now = new Date().toISOString();
  sh.getRange(rowIndex, idx.date + 1).setValue(rec.date);
  sh.getRange(rowIndex, idx.detail + 1).setValue(rec.detail);
  sh.getRange(rowIndex, idx.floor + 1).setValue(rec.floor);
  sh.getRange(rowIndex, idx.who + 1).setValue(rec.who);
  sh.getRange(rowIndex, idx.status + 1).setValue(rec.status || 'รอดำเนินการ');
  sh.getRange(rowIndex, idx.fileName + 1).setValue(rec.fileName || "");
  sh.getRange(rowIndex, idx.fileId + 1).setValue(rec.fileId || "");
  sh.getRange(rowIndex, idx.fileUrl + 1).setValue(rec.fileUrl || "");
  sh.getRange(rowIndex, idx.updatedAt + 1).setValue(rec.updatedAt || now);
  sh.getRange(rowIndex, idx.updatedBy + 1).setValue(rec.updatedBy || "");

  return true;
}

function delete_(id) {
  const sh = getSheet_();
  const values = sh.getDataRange().getValues();
  const headers = values.shift();
  const idx = {};
  headers.forEach((h, i) => { idx[h] = i; });
  for (let r = 0; r < values.length; r++) {
    if (values[r][idx.id] === id) {
      sh.deleteRow(r + 2);
      return true;
    }
  }
  return false;
}

function saveFileFromDataURL_(dataURL, fileName) {
  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const parts = dataURL.split(',');
  const meta = parts[0]; const b64 = parts[1];
  const contentType = (meta.match(/data:(.*?);base64/) || [])[1] || MimeType.BINARY;
  const blob = Utilities.newBlob(Utilities.base64Decode(b64), contentType, fileName);
  const file = folder.createFile(blob).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { id: file.getId(), url: file.getUrl() };
}

function exportExcel() {
  // Build a temp Spreadsheet with all rows, then convert to XLSX and save into folder
  const sh = getSheet_();
  const data = sh.getDataRange().getValues(); // includes header
  const tmp = SpreadsheetApp.create("SOC Export " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd' 'HH:mm:ss"));
  const tsh = tmp.getSheets()[0];
  tsh.getRange(1, 1, data.length, data[0].length).setValues(data);
  // get as XLSX
  const xlsx = DriveApp.getFileById(tmp.getId()).getAs(MimeType.MICROSOFT_EXCEL);
  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const saved = folder.createFile(xlsx).setName("SOC_All_" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd") + ".xlsx");
  // cleanup temp spreadsheet
  DriveApp.getFileById(tmp.getId()).setTrashed(true);
  return "สร้างไฟล์ในโฟลเดอร์แล้ว: " + saved.getName();
}

// --------- Sample Data Creator ----------
function createSampleData() {
  const sh = getSheet_();
  const now = new Date().toISOString();
  const samples = [
    { id: 'sample1', date: '2025-08-01', detail: 'พบปัญหา SOC ที่ชั้น 9', floor: 'ชั้น9', who: 'Operation', status: 'รอดำเนินการ', fileName: '', fileId: '', fileUrl: '', createdAt: now, updatedAt: now, updatedBy: '' },
    { id: 'sample2', date: '2025-08-02', detail: 'ตรวจพบ SOC ที่ Packer', floor: 'Packer', who: 'Local', status: 'เสร็จสิ้น', fileName: '', fileId: '', fileUrl: '', createdAt: now, updatedAt: now, updatedBy: 'Admin วอ' },
    { id: 'sample3', date: '2025-08-03', detail: 'SOC ที่ชั้น 5 จากแม่บ้าน', floor: 'ชั้น5', who: 'แม่บ้าน', status: 'รอดำเนินการ', fileName: '', fileId: '', fileUrl: '', createdAt: now, updatedAt: now, updatedBy: '' }
  ];
  samples.forEach(rec => {
    const row = [
      rec.id, rec.date, rec.detail, rec.floor, rec.who, rec.status,
      rec.fileName, rec.fileId, rec.fileUrl,
      rec.createdAt, rec.updatedAt, rec.updatedBy
    ];
    sh.appendRow(row);
  });
  return 'Created sample data successfully.';
}

// --------- Web App Entry Point ----------
const SHEET_NAME = "SOC";
const DRIVE_FOLDER_ID = "1pqA2G9jsf0WBfaorAikqV9FUAHVkxm8j"; // Keep your Drive folder ID

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Soc App ระบบจัดการ Soc ที่เกิดขึ้น')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
