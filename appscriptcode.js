const SHEET_ID = '1vl2QdtUazaNVgfiSdklPwqg-qAOurJ9VcVMzOZPUMHg';
const OVERDUE_SHEET = 'Overdue Chromes';

function doGet(e) {
  const action = e.parameter.action;

  if (action === 'report') return getReport();
  if (action === 'overdue') return getOverdue();
  if (action === 'history') return getHistory(e.parameter.cbNum);

  return jsonResponse({ success: false, message: 'Invalid request' });
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);

  if (data.action === 'checkout') return checkOut(data);
  if (data.action === 'checkin') return checkIn(data);

  return jsonResponse({ success: false, message: 'Unknown action' });
}

/* =======================
   SAFE HELPERS
======================= */

function isBlank(v) {
  return v === null || v === undefined || String(v).trim() === '';
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function formatDate(val) {
  if (!val) return '—';
  const d = new Date(val);
  if (isNaN(d)) return val;

  return (
    (d.getMonth() + 1).toString().padStart(2, '0') + '/' +
    d.getDate().toString().padStart(2, '0') + '/' +
    d.getFullYear()
  );
}

function daysSince(dateVal) {
  const d = new Date(dateVal);
  const today = new Date();

  d.setHours(0,0,0,0);
  today.setHours(0,0,0,0);

  return Math.floor((today - d) / (1000 * 60 * 60 * 24));
}

/* =======================
   CHECK OUT (FIXED)
======================= */

function checkOut(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(String(data.cbNum));

  if (!sheet) {
    return jsonResponse({
      success: false,
      message: `No tab found for Chromebook #${data.cbNum}.`
    });
  }

  const rows = sheet.getDataRange().getValues();

  // Find latest real record (skip headers + info rows)
  let latest = null;

  for (let i = rows.length - 1; i >= 3; i--) {
    if (rows[i][0]) {
      latest = rows[i];
      break;
    }
  }

  // If last record exists and has NO return date → already checked out
  if (latest && isBlank(latest[2])) {
    return jsonResponse({
      success: false,
      message: `Chromebook #${data.cbNum} is already checked out to ${latest[0]}.`
    });
  }

  sheet.appendRow([
    data.studentId,
    data.checkoutDate,
    '',
    '',
    '',
    data.notes || ''
  ]);

  return jsonResponse({ success: true });
}

/* =======================
   CHECK IN (FIXED)
======================= */

function checkIn(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(String(data.cbNum));

  if (!sheet) {
    return jsonResponse({
      success: false,
      message: `No tab found for Chromebook #${data.cbNum}.`
    });
  }

  const rows = sheet.getDataRange().getValues();

  for (let i = rows.length - 1; i >= 3; i--) {
    const student = rows[i][0];

    if (student && isBlank(rows[i][2])) {
      sheet.getRange(i + 1, 3).setValue(data.checkinDate);

      removeFromOverdue(data.cbNum);

      return jsonResponse({
        success: true,
        studentId: student
      });
    }
  }

  return jsonResponse({
    success: false,
    message: `Chromebook #${data.cbNum} is not currently checked out.`
  });
}

/* =======================
   REPORT
======================= */

function getReport() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const main = ss.getSheetByName('Main');

  const mainRows = main.getDataRange().getValues();
  const result = [];

  for (let i = 1; i < mainRows.length; i++) {
    const cbNum = mainRows[i][0];
    const barcode = mainRows[i][1];
    const serial = mainRows[i][2];

    if (!cbNum) continue;

    const sheet = ss.getSheetByName(String(cbNum));
    if (!sheet) continue;

    const rows = sheet.getDataRange().getValues();

    let latest = null;

    for (let j = rows.length - 1; j >= 3; j--) {
      if (rows[j][0]) {
        latest = rows[j];
        break;
      }
    }

    if (latest) {
      const isOut = isBlank(latest[2]);

      if (isOut && !isBlank(latest[1]) && daysSince(latest[1]) > 0) {
        addToOverdue(cbNum, barcode, latest[0], latest[1]);
      }

      result.push({
        cbNum,
        barcode,
        serial,
        studentId: latest[0],
        checkoutDate: formatDate(latest[1]),
        checkinDate: latest[2] ? formatDate(latest[2]) : '—',
        notes: latest[5] || '',
        status: isOut ? 'Out' : 'In'
      });
    }
  }

  return jsonResponse({ rows: result });
}

/* =======================
   OVERDUE
======================= */

function getOverdue() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(OVERDUE_SHEET);

  if (!sheet) return jsonResponse({ rows: [] });

  const rows = sheet.getDataRange().getValues();
  const result = [];

  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;

    result.push({
      cbNum: rows[i][0],
      barcode: rows[i][1] || '—',
      studentId: rows[i][2],
      checkoutDate: formatDate(rows[i][3]),
      daysOverdue: daysSince(rows[i][3])
    });
  }

  return jsonResponse({ rows: result });
}

/* =======================
   OVERDUE HELPERS
======================= */

function addToOverdue(cbNum, barcode, studentId, checkoutDate) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(OVERDUE_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(OVERDUE_SHEET);
    sheet.appendRow(['CB #', 'Barcode', 'Student ID', 'Date Signed Out']);
  }

  const rows = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(cbNum)) return;
  }

  sheet.appendRow([cbNum, barcode, studentId, checkoutDate]);
}

function removeFromOverdue(cbNum) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(OVERDUE_SHEET);
  if (!sheet) return;

  const rows = sheet.getDataRange().getValues();

  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][0]) === String(cbNum)) {
      sheet.deleteRow(i + 1);
    }
  }
}

/* =======================
   HISTORY
======================= */

function getHistory(cbNum) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(String(cbNum));

  if (!sheet) return jsonResponse({ rows: [] });

  const rows = sheet.getDataRange().getValues();
  const result = [];

  for (let i = 3; i < rows.length; i++) {
    if (!rows[i][0]) continue;

    result.push({
      studentId: rows[i][0],
      checkoutDate: formatDate(rows[i][1]),
      checkinDate: rows[i][2] ? formatDate(rows[i][2]) : '—',
      notes: rows[i][5] || ''
    });
  }

  return jsonResponse({ rows: result });
}
