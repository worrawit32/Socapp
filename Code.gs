/**
 * Soc App - Google Apps Script backend
 * Sheet ID: 1CVktJUFaYVQpLHHDiKDFanZzjx5J_BxQzMnTKwnLD48 (Sheet name: Sheet1)
 * Drive Folder for uploads: 170CE-whM4WsW4vtmgdJVNMam0oELQAsh
 */

const SHEET_ID = "1CVktJUFaYVQpLHHDiKDFanZzjx5J_BxQzMnTKwnLD48";
const SHEET_NAME = "Sheet1";
const DRIVE_FOLDER_ID = "170CE-whM4WsW4vtmgdJVNMam0oELQAsh";

function doGet(e) {
  const mode = (e && e.parameter && e.parameter.mode) || "ping";
  if (mode === "list") {
    return json_({ ok: true, rows: getAllRows_() });
  }
  return json_({ ok: true, msg: "pong" });
}

function doPost(e) {
  try {
    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    let payload = {};
    let fileUrl = "";

    if (e.postData && e.postData.type === "multipart/form-data") {
      // รับจาก FormData (มีไฟล์หรือไม่มีก็ได้)
      if (e.parameters && e.parameters.payload && e.parameters.payload.length) {
        payload = JSON.parse(e.parameters.payload[0]);
      }
      if (e.files && e.files.file) {
        const up = e.files.file; // { length, name, type, bytes }
        const blob = Utilities.newBlob(
          up.bytes,
          up.type || "application/octet-stream",
          up.name || ("upload_" + Date.now())
        );
        const created = folder.createFile(blob);
        fileUrl = created.getUrl();
      }
    } else if (e.postData && e.postData.contents) {
      // รับเป็น JSON ทั้งก้อน
      payload = JSON.parse(e.postData.contents);
    }

    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
    ensureHeader_(sheet);

    const id = "SOC-" + new Date().getTime();
    const now = new Date();
    const row = [
      id,
      payload.date || "",
      payload.detail || "",
      payload.floor || "",
      payload.reporter || "",
      payload.status || "รอดำเนินการ",
      fileUrl,
      now,
      payload.updater || ""
    ];
    sheet.appendRow(row);

    return json_({ ok: true, id, fileUrl });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/** สร้าง/เช็คหัวตารางให้ครบ */
function ensureHeader_(sheet) {
  const header = [
    "ID",
    "วันที่ตรวจเช็ค",
    "รายละเอียดที่พบ",
    "ชั้น",
    "ผู้ตรวจเช็ค/ผู้พบเจอ",
    "สถานะ",
    "ไฟล์แนบ(URL)",
    "วันที่บันทึก",
    "ผู้อัปเดตสถานะ"
  ];
  const range = sheet.getRange(1, 1, 1, header.length);
  const values = range.getValues()[0];
  const hasHeader = values.some(String);
  if (!hasHeader) {
    range.setValues([header]);
    sheet.setFrozenRows(1);
  }
}

/** ดึงข้อมูลทั้งหมด */
function getAllRows_() {
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  const data = sh.getDataRange().getValues();
  const header = data.shift();
  const idx = (name) => header.indexOf(name);
  return data.map(r => ({
    id: r[idx("ID")],
    date: r[idx("วันที่ตรวจเช็ค")],
    detail: r[idx("รายละเอียดที่พบ")],
    floor: r[idx("ชั้น")],
    reporter: r[idx("ผู้ตรวจเช็ค/ผู้พบเจอ")],
    status: r[idx("สถานะ")],
    fileUrl: r[idx("ไฟล์แนบ(URL)")],
    createdAt: r[idx("วันที่บันทึก")],
    updater: r[idx("ผู้อัปเดตสถานะ")]
  }));
}

/** helper สร้าง JSON output */
function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
