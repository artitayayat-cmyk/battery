const SHEET_NAMES = {
  plans: 'plans',
  logs: 'logs',
  comments: 'comments',
  battery: 'battery',
};

const HEADERS = {
  plans: ['id', 'studentId', 'studentName', 'domain', 'title', 'goal', 'actions', 'resources', 'deadline', 'progress', 'status', 'reflection', 'createdAt', 'updatedAt'],
  logs: ['id', 'planId', 'studentId', 'progress', 'note', 'evidence', 'createdAt'],
  comments: ['id', 'planId', 'studentId', 'facultyName', 'comment', 'createdAt'],
  battery: ['studentId', 'domain', 'self', 'peer', 'teacher', 'updatedAt'],
};

function doGet(e) {
  const action = (e.parameter.action || 'getAll').trim();
  if (action !== 'getAll') {
    return json({ ok: false, error: 'Unknown action' });
  }
  return json(getAllData());
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    const action = body.action;
    const payload = body.payload || {};

    if (action === 'savePlan') return json(saveRecord('plans', payload));
    if (action === 'addLog') return json(saveRecord('logs', payload));
    if (action === 'addComment') return json(saveRecord('comments', payload));
    if (action === 'upsertBattery') return json(upsertBattery(payload));
    if (action === 'deletePlan') return json(deletePlan(payload.planId));

    return json({ ok: false, error: 'Unknown action' });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

function setupSheets() {
  Object.keys(SHEET_NAMES).forEach(function (key) {
    const sheet = getSheet(key);
    ensureHeaders(sheet, HEADERS[key]);
  });
}

function getAllData() {
  setupSheets();
  return {
    plans: readSheet('plans'),
    logs: readSheet('logs'),
    comments: readSheet('comments'),
    battery: readSheet('battery'),
  };
}

function saveRecord(type, record) {
  setupSheets();
  const sheet = getSheet(type);
  const headers = HEADERS[type];
  const rows = sheet.getDataRange().getValues();
  const idIndex = headers.indexOf('id');
  const id = record.id;

  if (id && idIndex >= 0) {
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][idIndex]) === String(id)) {
        sheet.getRange(i + 1, 1, 1, headers.length).setValues([headers.map(h => valueFor(record, h))]);
        return { ok: true, updated: true };
      }
    }
  }

  sheet.appendRow(headers.map(h => valueFor(record, h)));
  return { ok: true, created: true };
}

function upsertBattery(record) {
  setupSheets();
  record.updatedAt = new Date().toISOString();
  const sheet = getSheet('battery');
  const headers = HEADERS.battery;
  const rows = sheet.getDataRange().getValues();
  const sidIndex = headers.indexOf('studentId');
  const domainIndex = headers.indexOf('domain');

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][sidIndex]) === String(record.studentId) && String(rows[i][domainIndex]) === String(record.domain)) {
      sheet.getRange(i + 1, 1, 1, headers.length).setValues([headers.map(h => valueFor(record, h))]);
      return { ok: true, updated: true };
    }
  }

  sheet.appendRow(headers.map(h => valueFor(record, h)));
  return { ok: true, created: true };
}

function deletePlan(planId) {
  setupSheets();
  const sheet = getSheet('plans');
  const rows = sheet.getDataRange().getValues();
  const idIndex = HEADERS.plans.indexOf('id');

  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][idIndex]) === String(planId)) {
      sheet.deleteRow(i + 1);
    }
  }
  return { ok: true };
}

function readSheet(type) {
  const sheet = getSheet(type);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];

  return values.slice(1).filter(row => row.some(cell => cell !== '')).map(function (row) {
    const obj = {};
    headers.forEach(function (header, i) {
      obj[header] = row[i];
    });
    return obj;
  });
}

function getSheet(type) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = SHEET_NAMES[type];
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function ensureHeaders(sheet, headers) {
  const current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const needsHeader = headers.some((header, i) => current[i] !== header);
  if (needsHeader) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
}

function valueFor(record, header) {
  const value = record[header];
  if (value === undefined || value === null) return '';
  return value;
}

function json(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
