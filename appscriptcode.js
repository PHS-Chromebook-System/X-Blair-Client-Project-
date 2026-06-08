const SHEET_ID = '1vl2QdtUazaNVgfiSdklPwqg-qAOurJ9VcVMzOZPUMHg';
const OVERDUE_SHEET = 'Overdue Chromes';

function doGet(e) {
  const action = e.parameter.action;
  if (action === 'report')  return getReport();
  if (action === 'overdue') return getOverdue();
  if (action === 'history') return getHistory(e.parameter.cbNum);
  return ContentService.createTextOutput('Invalid').setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  if (data.action === 'checkout') return checkOut(data);
  if (data.action === 'checkin')  return checkIn(data);
  return jsonResponse({ success: false, message: 'Unknown action' });
}

function formatDate(val) {
  if (!val) return '—';
  const d = new Date(val);
  if (isNaN(d)) return val;
  return (d.getMonth()+1).toString().padStart(2,'0') + '/' +
         d.getDate().toString().padStart(2,'0') + '/' +
         d.getFullYear();
}

function isOverdue(checkoutDate) {
  if (!checkoutDate) return false;
  const co = new Date(checkoutDate);
  const today = new Date();
  co.setHours(0,0,0,0);
  today.setHours(0,0,0,0);
  return co < today;
}

function getHistory(cbNum) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(String(cbNum));

  if (!sheet) {
    return jsonResponse({ rows: [] });
  }

  const data = sheet.getDataRange().getValues();
  const rows = [];

  for (let i = data.length - 1; i >= 3; i--) {
    if (!data[i][0]) continue;

    rows.push({
      studentId: data[i][0],
      checkoutDate: formatDate(data[i][1]),
      checkinDate: data[i][2] ? formatDate(data[i][2]) : '—',
      notes: data[i][5] || ''
    });
  }

  return jsonResponse({ rows: rows });
}

function daysSince(checkoutDate) {
  if (!checkoutDate) return 0;
  const co = new Date(checkoutDate);
  const today = new Date();
  co.setHours(0,0,0,0);
  today.setHours(0,0,0,0);
  return Math.floor((today - co) / (1000 * 60 * 60 * 24));
}

function addToOverdue(cbNum, barcode, studentId, checkoutDate) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  let sheet   = ss.getSheetByName(OVERDUE_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(OVERDUE_SHEET);
    sheet.appendRow(['CB #', 'Barcode', 'Student ID', 'Date Signed Out']);
  }
  const rows = sheet.getDataRange().getValues();
  // Don't add duplicates
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(cbNum)) return;
  }
  sheet.appendRow([cbNum, barcode, studentId, checkoutDate]);
}

function removeFromOverdue(cbNum) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(OVERDUE_SHEET);
  if (!sheet) return;
  const rows = sheet.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][0]) === String(cbNum)) {
      sheet.deleteRow(i + 1);
    }
  }
}

function checkOut(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  const main = ss.getSheetByName('Main');
  const mainRows = main.getDataRange().getValues();

  let actualCbNum = data.cbNum;

  for (let i = 1; i < mainRows.length; i++) {
    if (String(mainRows[i][1]).trim() === String(data.cbNum).trim()) {
      actualCbNum = String(mainRows[i][0]);
      break;
    }
  }

  const sheet = ss.getSheetByName(String(actualCbNum));

  if (!sheet) {
    return jsonResponse({
      success: false,
      message: `No Chromebook found for barcode ${data.cbNum}.`
    });
  }

  const rows = sheet.getDataRange().getValues();

  let latest = null;

for (let i = rows.length - 1; i >= 3; i--) {
  if (rows[i][0]) {
    latest = rows[i];
    break;
  }
}

if (latest && !latest[2]) {
  return jsonResponse({
    success: false,
    message: `Chromebook #${actualCbNum} is already checked out to ${latest[0]}.`
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


function checkIn(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  const main = ss.getSheetByName('Main');
  const mainRows = main.getDataRange().getValues();

  let actualCbNum = data.cbNum;

  for (let i = 1; i < mainRows.length; i++) {
    if (String(mainRows[i][1]).trim() === String(data.cbNum).trim()) {
      actualCbNum = String(mainRows[i][0]);
      break;
    }
  }

  const sheet = ss.getSheetByName(String(actualCbNum));

  if (!sheet) {
    return jsonResponse({
      success: false,
      message: `No Chromebook found for barcode ${data.cbNum}.`
    });
  }

  const rows = sheet.getDataRange().getValues();
for (let i = rows.length - 1; i >= 3; i--) {
  if (rows[i][0] && !rows[i][2]) {
      const rowNum = i + 1;

      sheet.getRange(rowNum, 3).setValue(data.checkinDate);

      removeFromOverdue(actualCbNum);

      return jsonResponse({
        success: true,
        studentId: rows[i][0]
      });
    }
  }

  return jsonResponse({
    success: false,
    message: `Chromebook #${actualCbNum} is not currently checked out.`
  });
}


function getReport() {
  const ss       = SpreadsheetApp.openById(SHEET_ID);
  const main     = ss.getSheetByName('Main');
  const mainRows = main.getDataRange().getValues();
  const result   = [];
  for (let i = 1; i < mainRows.length; i++) {
    const cbNum   = mainRows[i][0];
    const barcode = mainRows[i][1];
    const serial  = mainRows[i][2];
    if (!cbNum) continue;
    const sheet = ss.getSheetByName(String(cbNum));
    if (!sheet) continue;
    const rows = sheet.getDataRange().getValues();
    let latest = null;
    for (let j = 3; j < rows.length; j++) {
      if (rows[j][0]) latest = rows[j];
    }
    if (latest) {
      const isOut = !latest[2];
      // Auto-add to overdue if out and overdue
      if (isOut && isOverdue(latest[1])) {
        addToOverdue(cbNum, barcode, latest[0], formatDate(latest[1]));
      }
      result.push({
        cbNum, barcode, serial,
        studentId:    latest[0],
        checkoutDate: formatDate(latest[1]),
        checkinDate:  latest[2] ? formatDate(latest[2]) : '—',
        notes:        latest[5] || '',
        status:       isOut ? 'Out' : 'In'
      });
    } else {
      result.push({
        cbNum, barcode, serial,
        studentId: '—', checkoutDate: '—', checkinDate: '—',
        notes: '', status: 'In'
      });
    }
  }
  return jsonResponse({ rows: result });
}

function getOverdue() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const main = ss.getSheetByName('Main');
  const mainRows = main.getDataRange().getValues();

  const result = [];

  for (let i = 1; i < mainRows.length; i++) {
    const cbNum = mainRows[i][0];
    const barcode = mainRows[i][1];

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

    if (!latest) continue;

    const isOut = !latest[2];

    if (isOut) {
      result.push({
        cbNum: cbNum,
        barcode: barcode || '—',
        studentId: latest[0],
        checkoutDate: formatDate(latest[1]),
        daysOverdue: daysSince(latest[1])
      });
    }
  }

  // Sort longest checked-out first
  result.sort((a, b) => b.daysOverdue - a.daysOverdue);

  return jsonResponse({ rows: result });
}

function jsonResponse(obj) {
  const output = ContentService.createTextOutput(JSON.stringify(obj));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

function scanCheckoutBarcode() {
  const barcode = prompt("Enter Chromebook Barcode:");

  if (!barcode) return;

  document.getElementById("coNum").value = barcode;
}

function scanCheckinBarcode() {
  const barcode = prompt("Enter Chromebook Barcode:");

  if (!barcode) return;

  document.getElementById("ciNum").value = barcode;
}

