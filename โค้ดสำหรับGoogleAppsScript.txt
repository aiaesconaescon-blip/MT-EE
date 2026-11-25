/**
 * Google Apps Script for Two-Way Hybrid Sync
 * ผสานข้อมูลอัตโนมัติ: Sheet 1 (Master ส่วนกลาง - ดูงานใหม่) + Sheet 2 (บันทึกสถานะของทีมเรา)
 * 
 * วิธีติดตั้งใน Google Sheet:
 * 1. เปิด Google Sheet แผ่นของทีมเรา
 * 2. ไปที่เมนู "ส่วนขยาย" (Extensions) > "Apps Script"
 * 3. ลบโค้ดเดิมออกทั้งหมด แล้ววางโค้ดนี้แทนที่ทั้งหมด แล้วกด Save (Ctrl+S)
 * 4. กด "ทำให้ใช้งานได้" (Deploy) > "การจัดการการทำให้ใช้งานได้" (Manage deployments) หรือ "การทำให้ใช้งานได้ใหม่" (New deployment)
 * 5. เลือกประเภท: "เว็บแอป" (Web app)
 *    - ดำเนินการในฐานะ: "ฉัน" (Me)
 *    - ผู้มีสิทธิ์เข้าถึง: "ทุกคน" (Anyone)
 * 6. กด Deploy แล้วนำ URL ที่ได้มาใช้งาน
 */

var MASTER_SHEET_ID = "1gB2vq-JhUTgFbOSanF_jAwnzNndzKtYRgFXdEiN8lKU";

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : "getData";
  
  if (action === "getData") {
    return handleGetData();
  }
  
  return ContentService.createTextOutput(JSON.stringify({
    status: "ok",
    message: "GBH Maintenance API is active and ready."
  })).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var raw = e.postData.contents;
    var payload = JSON.parse(raw);
    
    if (payload.action === "updateJob" || payload.no) {
      return handleUpdateJob(payload);
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: "Unknown action"
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * ดึงข้อมูลงานทั้งหมดจาก Master Sheet 1 และนำข้อมูลการแก้ไขจาก Sheet 2 มาประกบกัน
 */
function handleGetData() {
  // 1. ดึงข้อมูลการแก้ไขของทีมเราจาก Sheet 2 (Active Spreadsheet)
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var teamSheet = ss.getActiveSheet();
  var teamData = teamSheet.getDataRange().getValues();
  var teamEdits = {};
  
  if (teamData.length > 1) {
    var teamColMap = getHeaderColumnMap(teamData[0]);
    var noCol = (teamColMap["no"] !== undefined) ? teamColMap["no"] : 0;
    
    for (var i = 1; i < teamData.length; i++) {
      var tRow = teamData[i];
      var tNo = String(tRow[noCol] || "").trim();
      if (!tNo) continue;
      
      teamEdits[tNo] = {
        job_no: getVal(tRow, teamColMap["job_no"]),
        responsible: getVal(tRow, teamColMap["responsible"]),
        contractor: getVal(tRow, teamColMap["contractor"]),
        site_status: getVal(tRow, teamColMap["site_status"]),
        contact: getVal(tRow, teamColMap["contact"]),
        notes: getVal(tRow, teamColMap["notes"]),
        updated_at: getVal(tRow, teamColMap["updated_at"])
      };
    }
  }
  
  // 2. ดึงข้อมูลงานล่าสุดจาก Master Sheet 1
  var masterRows = fetchMasterSheetData();
  var result = [];
  
  if (masterRows && masterRows.length > 1) {
    var mHeaders = masterRows[0];
    var mColMap = getHeaderColumnMap(mHeaders);
    var closedStatuses = ["จัดส่งบัญชี KAROON", "ปิดเรียบร้อย", "CANCELLED", "ยกเลิกงาน", "สาขาดำเนินการเอง"];
    
    for (var m = 1; m < masterRows.length; m++) {
      var row = masterRows[m];
      var noVal = getVal(row, mColMap["no"]);
      if (!noVal) continue;
      var noStr = String(noVal).trim();
      
      var status = String(getVal(row, mColMap["status"]) || "รอประเมินBOQ").trim();
      var closeDate = String(getVal(row, mColMap["close_date"]) || "").trim();
      var isClosed = (closedStatuses.indexOf(status) !== -1) || Boolean(closeDate && closeDate !== "-");
      
      if (isClosed) continue; // ส่งเฉพาะงานที่ยังคงค้าง เพื่อให้ Dashboard ทำงานได้รวดเร็วที่สุด
      
      // ค่าตั้งต้นจาก Master Sheet 1
      var jobNo = String(getVal(row, mColMap["job_no"]) || "").trim();
      var responsible = String(getVal(row, mColMap["responsible"]) || "งานระบบ").trim();
      var contractor = String(getVal(row, mColMap["contractor"]) || "").trim();
      var siteStatus = String(getVal(row, mColMap["site_status"]) || "").trim();
      var contact = String(getVal(row, mColMap["contact"]) || "").trim();
      var notes = String(getVal(row, mColMap["notes"]) || "").trim();
      var updatedAt = "";
      
      // หากมีการแก้ไขใน Sheet 2 ให้ใช้ค่าจาก Sheet 2 แทน
      if (teamEdits[noStr]) {
        var ed = teamEdits[noStr];
        if (ed.job_no) jobNo = ed.job_no;
        if (ed.responsible) responsible = ed.responsible;
        if (ed.contractor) contractor = ed.contractor;
        if (ed.site_status) siteStatus = ed.site_status;
        if (ed.contact) contact = ed.contact;
        if (ed.notes) notes = ed.notes;
        if (ed.updated_at) updatedAt = ed.updated_at;
      }
      
      result.push({
        row_idx: m + 1,
        no: noStr,
        job_no: jobNo,
        date: String(getVal(row, mColMap["date"]) || "").trim(),
        province: String(getVal(row, mColMap["province"]) || "").trim(),
        branch: String(getVal(row, mColMap["branch"]) || "").trim(),
        zone: String(getVal(row, mColMap["zone"]) || "").trim(),
        category: String(getVal(row, mColMap["category"]) || "").trim(),
        nature: String(getVal(row, mColMap["nature"]) || "").trim(),
        desc: String(getVal(row, mColMap["desc"]) || "").trim(),
        quotation: String(getVal(row, mColMap["quotation"]) || "").trim(),
        budget: String(getVal(row, mColMap["budget"]) || "").trim(),
        tender: String(getVal(row, mColMap["tender"]) || "").trim(),
        status: status,
        responsible: responsible || "งานระบบ",
        contractor: contractor,
        site_status: siteStatus,
        contact: contact,
        notes: notes,
        updated_at: updatedAt
      });
    }
  } else {
    // หากไม่สามารถเชื่อมต่อ Master Sheet 1 ได้ ให้ใช้ข้อมูลใน Sheet 2 เป็นข้อมูลสำรอง
    for (var t = 1; t < teamData.length; t++) {
      var row2 = teamData[t];
      var no2 = String(getVal(row2, teamColMap["no"]) || "").trim();
      if (!no2) continue;
      result.push({
        row_idx: t + 1,
        no: no2,
        job_no: String(getVal(row2, teamColMap["job_no"]) || "").trim(),
        date: String(getVal(row2, teamColMap["date"]) || "").trim(),
        province: String(getVal(row2, teamColMap["province"]) || "").trim(),
        branch: String(getVal(row2, teamColMap["branch"]) || "").trim(),
        zone: String(getVal(row2, teamColMap["zone"]) || "").trim(),
        category: String(getVal(row2, teamColMap["category"]) || "").trim(),
        nature: String(getVal(row2, teamColMap["nature"]) || "").trim(),
        desc: String(getVal(row2, teamColMap["desc"]) || "").trim(),
        quotation: String(getVal(row2, teamColMap["quotation"]) || "").trim(),
        budget: String(getVal(row2, teamColMap["budget"]) || "").trim(),
        tender: String(getVal(row2, teamColMap["tender"]) || "").trim(),
        status: String(getVal(row2, teamColMap["status"]) || "รอประเมินBOQ").trim(),
        responsible: String(getVal(row2, teamColMap["responsible"]) || "งานระบบ").trim(),
        contractor: String(getVal(row2, teamColMap["contractor"]) || "").trim(),
        site_status: String(getVal(row2, teamColMap["site_status"]) || "").trim(),
        contact: String(getVal(row2, teamColMap["contact"]) || "").trim(),
        notes: String(getVal(row2, teamColMap["notes"]) || "").trim(),
        updated_at: String(getVal(row2, teamColMap["updated_at"]) || "").trim()
      });
    }
  }
  
  return ContentService.createTextOutput(JSON.stringify({
    status: "success",
    source: "hybrid_master_and_team",
    count: result.length,
    last_synced: Utilities.formatDate(new Date(), "Asia/Bangkok", "dd/MM/yyyy HH:mm:ss"),
    data: result
  })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * ดึงข้อมูลสดจาก Master Sheet 1
 */
function fetchMasterSheetData() {
  try {
    var masterSs = SpreadsheetApp.openById(MASTER_SHEET_ID);
    if (masterSs) {
      return masterSs.getActiveSheet().getDataRange().getValues();
    }
  } catch (e) {
    try {
      var csvUrl = "https://docs.google.com/spreadsheets/d/" + MASTER_SHEET_ID + "/export?format=csv&gid=0";
      var res = UrlFetchApp.fetch(csvUrl, { muteHttpExceptions: true });
      if (res.getResponseCode() === 200) {
        return Utilities.parseCsv(res.getContentText());
      }
    } catch (err) {}
  }
  return null;
}

/**
 * อัปเดตข้อมูลของใบแจ้งซ่อมนั้นๆ ลงใน Sheet 2 (ของทีมเรา) ทันที
 */
function handleUpdateJob(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  
  if (data.length < 2) {
    throw new Error("Sheet is empty");
  }
  
  var headers = data[0];
  var colMap = getHeaderColumnMap(headers);
  var targetNo = String(payload.no || "").trim();
  
  if (!targetNo) {
    throw new Error("Target 'no' is required");
  }
  
  var noColIdx = (colMap["no"] !== undefined) ? colMap["no"] : 0;
  var matchedRowIdx = -1;
  
  for (var i = 1; i < data.length; i++) {
    var cellNo = String(data[i][noColIdx]).trim();
    if (cellNo === targetNo) {
      matchedRowIdx = i + 1; // 1-based index in Sheet
      break;
    }
  }
  
  var nowStr = Utilities.formatDate(new Date(), "Asia/Bangkok", "dd/MM/yyyy HH:mm:ss");

  if (matchedRowIdx === -1) {
    var newRow = [];
    var maxCols = Math.max(headers.length, 20);
    for (var c = 0; c < maxCols; c++) newRow.push("");
    if (colMap["no"] !== undefined) newRow[colMap["no"]] = targetNo;
    if (colMap["job_no"] !== undefined && payload.job_no !== undefined) newRow[colMap["job_no"]] = payload.job_no;
    if (colMap["responsible"] !== undefined && payload.responsible !== undefined) newRow[colMap["responsible"]] = payload.responsible;
    if (colMap["contractor"] !== undefined && payload.contractor !== undefined) newRow[colMap["contractor"]] = payload.contractor;
    if (colMap["site_status"] !== undefined && payload.site_status !== undefined) newRow[colMap["site_status"]] = payload.site_status;
    if (colMap["contact"] !== undefined && payload.contact !== undefined) newRow[colMap["contact"]] = payload.contact;
    if (colMap["notes"] !== undefined && payload.notes !== undefined) newRow[colMap["notes"]] = payload.notes;
    if (colMap["updated_at"] !== undefined) newRow[colMap["updated_at"]] = nowStr;
    
    sheet.appendRow(newRow);
    matchedRowIdx = sheet.getLastRow();
  } else {
    if (colMap["job_no"] !== undefined && payload.job_no !== undefined) {
      sheet.getRange(matchedRowIdx, colMap["job_no"] + 1).setValue(payload.job_no);
    }
    if (colMap["responsible"] !== undefined && payload.responsible !== undefined) {
      sheet.getRange(matchedRowIdx, colMap["responsible"] + 1).setValue(payload.responsible);
    }
    if (colMap["contractor"] !== undefined && payload.contractor !== undefined) {
      sheet.getRange(matchedRowIdx, colMap["contractor"] + 1).setValue(payload.contractor);
    }
    if (colMap["site_status"] !== undefined && payload.site_status !== undefined) {
      sheet.getRange(matchedRowIdx, colMap["site_status"] + 1).setValue(payload.site_status);
    }
    if (colMap["contact"] !== undefined && payload.contact !== undefined) {
      sheet.getRange(matchedRowIdx, colMap["contact"] + 1).setValue(payload.contact);
    }
    if (colMap["notes"] !== undefined && payload.notes !== undefined) {
      sheet.getRange(matchedRowIdx, colMap["notes"] + 1).setValue(payload.notes);
    }
    if (colMap["updated_at"] !== undefined) {
      sheet.getRange(matchedRowIdx, colMap["updated_at"] + 1).setValue(nowStr);
    }
  }
  
  return ContentService.createTextOutput(JSON.stringify({
    status: "success",
    message: "บันทึกข้อมูลใบแจ้งซ่อม No. " + targetNo + " ลงใน Google Sheet สำเร็จ!",
    updated_at: nowStr,
    row: matchedRowIdx
  })).setMimeType(ContentService.MimeType.JSON);
}

function getHeaderColumnMap(headers) {
  var map = {};
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i]).trim();
    if (h.includes("เลขแจ้งซ่อม") || h.includes("No.")) map["no"] = i;
    else if (h.includes("JOB NO.") || h.includes("JOB")) map["job_no"] = i;
    else if (h.includes("วันที่รับ") || h.includes("Date")) map["date"] = i;
    else if (h.includes("จังหวัด")) map["province"] = i;
    else if (h.includes("สาขา")) map["branch"] = i;
    else if (h.includes("เขต")) map["zone"] = i;
    else if (h.includes("ประเภทงาน") || h.includes("หมวดหมู่")) map["category"] = i;
    else if (h.includes("ลักษณะงาน")) map["nature"] = i;
    else if (h.includes("รายละเอียด")) map["desc"] = i;
    else if (h.includes("ใบเสนอราคา")) map["quotation"] = i;
    else if (h.includes("BUDGET")) map["budget"] = i;
    else if (h.includes("TENDER")) map["tender"] = i;
    else if (h.includes("ใบปิดงาน") || h.includes("close")) map["close_date"] = i;
    else if (h.includes("สถานะระบบ") || h.includes("STATUS")) map["status"] = i;
    else if (h.includes("ผู้รับผิดชอบ")) map["responsible"] = i;
    else if (h.includes("ชื่อ ผรม.") || h.includes("ผู้ดำเนินการ")) map["contractor"] = i;
    else if (h.includes("สถานะหน้างาน")) map["site_status"] = i;
    else if (h.includes("ช่องทางติดต่อ") || h.includes("เบอร์ติดต่อ")) map["contact"] = i;
    else if (h.includes("หมายเหตุ") || h.includes("บันทึก")) map["notes"] = i;
    else if (h.includes("อัปเดต") || h.includes("Updated")) map["updated_at"] = i;
  }
  return map;
}

function getVal(row, idx) {
  if (idx !== undefined && idx < row.length) {
    return row[idx];
  }
  return "";
}
