const SHEET_ID = '1l7eJ8iqc7zTGliKSqJUZIb055DujaIXnu5WazJyU3_c';
const OVERDUE_SHEET = 'Overdue Chromes';
const CACHE_TIME = 5; // 5 minutes

function doGet(e) {
  const action = e.parameter.action;
  if (action === 'report')  return getReport();
  if (action === 'overdue') return getOverdue();
  if (action === 'history') return getHistory(e.parameter.cbNum);
  if (action === 'summary') return getSummary();
  return ContentService.createTextOutput('Invalid').setMimeType(ContentService.MimeType.TEXT);
}

function resetCbHistory(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(String(data.cbNum));

  if (!sheet) {
    return jsonResponse({ success: false, message: 'Chromebook sheet not found.' });
  }

  const lastRow = sheet.getLastRow();
  if (lastRow > 3) {
    sheet.deleteRows(4, lastRow - 3);
  }

  return jsonResponse({ success: true });
}

function resetAllHistory() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const main = ss.getSheetByName('Main');
    if (!main) return jsonResponse({ success: false, message: 'Main sheet not found.' });

    const mainRows = main.getDataRange().getValues();

    for (let i = 1; i < mainRows.length; i++) {
      const cbNum = mainRows[i][0];
      if (!cbNum) continue;

      const sheet = ss.getSheetByName(String(cbNum));
      if (!sheet) continue;

      const lastRow = sheet.getLastRow();
      if (lastRow > 3) {
        sheet.deleteRows(4, lastRow - 3);
      }
    }

    return jsonResponse({ success: true });
  } catch(err) {
    return jsonResponse({ success: false, message: err.message });
  }
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    if (data.action === "sendOverdueEmails") {
  return sendOverdueEmails();
}

    if (data.action === "checkout") {
      return checkout(data);
    }

    if (data.action === "checkin") {
      return checkIn(data);
    }

    if (data.action === "resetHistory") {
      return resetCbHistory(data);
    }

    if (data.action === "resetAllHistory") {
      return resetAllHistory();
    }

    return jsonResponse({
      success: false,
      message: "Unknown action"
    });

  } catch (err) {
    return jsonResponse({
      success: false,
      message: "Parse error: " + err.message
    });
  }
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

function checkout(data) {
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

function getSummary() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const main = ss.getSheetByName('Main');
  const rows = main.getDataRange().getValues();

  let total = rows.length - 1;
  let out = 0;

  for (let i = 1; i < rows.length; i++) {
    const cbNum = rows[i][0];
    const sheet = ss.getSheetByName(String(cbNum));
    if (!sheet) continue;

    const lastRow = sheet.getLastRow();
    const data = sheet.getRange(lastRow, 1, 1, 6).getValues()[0];

    if (data && !data[2]) out++;
  }

  return jsonResponse({
    total,
    out,
    in: total - out
  });
}


function getReport() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get("report");

  if (cached) {
    return ContentService
      .createTextOutput(cached)
      .setMimeType(ContentService.MimeType.JSON);
  }

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const main = ss.getSheetByName('Main');
  const mainRows = main.getDataRange().getValues();

  const result = [];

  for (let i = 1; i < mainRows.length; i++) {
    const cbNum   = mainRows[i][0];
    const barcode = mainRows[i][1];
    const serial  = mainRows[i][2];
    if (!cbNum) continue;

    const sheet = ss.getSheetByName(String(cbNum));
    if (!sheet) continue;

    const lastRow = sheet.getLastRow();
    if (lastRow < 4) {
      result.push({
        cbNum, barcode, serial,
        studentId: '—',
        checkoutDate: '—',
        checkinDate: '—',
        notes: '',
        status: 'In'
      });
      continue;
    }

    const rows = sheet.getRange(1, 1, lastRow, 6).getValues();

    let latest = null;

    for (let j = rows.length - 1; j >= 3; j--) {
      if (rows[j][0]) {
        latest = rows[j];
        break;
      }
    }

    if (latest) {
      const isOut = !latest[2];

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

  const output = JSON.stringify({ rows: result });

  cache.put("report", output, CACHE_TIME); // 🔥 store result

  return ContentService
    .createTextOutput(output)
    .setMimeType(ContentService.MimeType.JSON);
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

    const lastRow = sheet.getLastRow();
const rows = sheet.getRange(1, 1, lastRow, 6).getValues();

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



function sendOverdueEmails() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const main = ss.getSheetByName('Main');
  const rows = main.getDataRange().getValues();

  let sent = 0;

  for (let i = 1; i < rows.length; i++) {
    const cbNum = rows[i][0];
    const sheet = ss.getSheetByName(String(cbNum));

    if (!sheet) continue;

    const data = sheet.getDataRange().getValues();

    let latest = null;

    for (let j = data.length - 1; j >= 3; j--) {
      if (data[j][0]) {
        latest = data[j];
        break;
      }
    }

    if (!latest) continue;

    const studentId = latest[0];
    const checkoutDate = latest[1];
    const checkinDate = latest[2];

    if (!studentId || checkinDate) continue;

    const daysOut = Math.floor(
      (new Date() - new Date(checkoutDate)) / (1000 * 60 * 60 * 24)
    );

    if (daysOut < 1) continue;

    const email = `${studentId}@mcpsmd.net`;

    MailApp.sendEmail({
      to: email,
      subject: `Overdue Chromebook Reminder (#${cbNum})`,
      htmlBody: `
        <p>Chromebook <b>#${cbNum}</b> is overdue by <b>${daysOut} days</b>.</p>
        <p>Please return it to the media center.</p>
      `
    });

    sent++;
  }

  return jsonResponse({
    success: true,
    sent
  });
}

