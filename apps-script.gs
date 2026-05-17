/**
 * ═══════════════════════════════════════════════════════════════════
 *  TB TRACKING — Google Apps Script Backend (v2)
 * ═══════════════════════════════════════════════════════════════════
 *
 *  วิธี deploy update:
 *  1. เปิด Apps Script project เดิม → เลือกไฟล์ Code.gs
 *  2. ลบโค้ดเดิมทั้งหมด → paste โค้ดนี้ทับ → Save (Ctrl+S)
 *  3. กด Run → เลือก function "migrate" → Run
 *     (ระบบจะเพิ่มคอลัมน์ใหม่ที่ขาดให้ Sheet เดิมโดยอัตโนมัติ
 *      ข้อมูลเดิมไม่หาย)
 *  4. Deploy → Manage deployments → Edit (icon ดินสอ) ของ deployment เดิม
 *      → Version: New version → Deploy
 *      Web app URL จะเหมือนเดิม → ฝั่งเว็บไม่ต้องแก้ค่าอะไร
 *  5. Refresh เว็บ → ทดสอบกรอกข้อมูลใหม่
 *
 *  สิ่งที่เปลี่ยน vs v1:
 *  ─ headers ครบทุกฟิลด์ที่ frontend ส่ง (zone, areaType, addrNo/Moo/...,
 *    tbClass, tbClassDetail, householdContacts, transferTo, testResult,
 *    visitTime, visitType, asmName, generalSymptoms, warningSigns, ...)
 *  ─ ensureHeaders() auto-extend คอลัมน์ใน Sheet เมื่อเจอฟิลด์ใหม่
 *    (ป้องกันปัญหาเดิมตลอดไป — เพิ่ม field ใหม่ที่ frontend ก็ใช้ได้เลย)
 *  ─ JSON serialize/deserialize ฟิลด์ที่เป็น array/object
 *    (householdContacts, symptoms, generalSymptoms, warningSigns, photos)
 *  ─ migrate() เพิ่มคอลัมน์ที่ขาดให้ Sheet เดิม + backfill addr* จาก
 *    address เก่า (แยกที่อยู่รวมกลับเป็นช่อง)
 * ═══════════════════════════════════════════════════════════════════
 */

const PROPS = PropertiesService.getScriptProperties();

const PATIENT_HEADERS = [
  'id', 'name', 'hn', 'age', 'gender', 'phone',
  'address',
  'addrNo', 'addrMoo', 'addrTambon', 'addrAmphoe', 'addrProvince',
  'zone', 'areaType',
  'contactPerson',
  'diagnosisDate', 'treatmentStartDate', 'regimen', 'regimenMonths',
  'afbResult', 'testResult', 'patientType',
  'tbClass', 'tbClassDetail',
  'transferTo',
  'householdContacts',
  'status', 'notes',
  'createdAt', 'createdBy', 'updatedAt', 'updatedBy'
];

const VISIT_HEADERS = [
  'id', 'patientId', 'date', 'visitTime', 'visitType',
  'visitorName', 'visitorRole', 'asmName',
  'medicationTaken', 'missedDoses',
  'symptoms', 'generalSymptoms', 'generalSymptomsOther', 'warningSigns',
  'sideEffects',
  'weight', 'bloodPressure', 'pulse', 'respiration', 'temperature', 'spo2',
  'notes', 'photos',
  'createdAt', 'createdBy'
];

const PATIENT_JSON_FIELDS = ['householdContacts'];
const VISIT_JSON_FIELDS = ['symptoms', 'generalSymptoms', 'warningSigns', 'photos'];

/* ─────────────────────────────────────────────────────────────────
   SETUP — รันครั้งเดียวเพื่อสร้าง Sheet + Drive folder
   ───────────────────────────────────────────────────────────────── */
function setup() {
  let sheetId = PROPS.getProperty('SHEET_ID');
  let folderId = PROPS.getProperty('FOLDER_ID');

  if (!sheetId) {
    const ss = SpreadsheetApp.create('TB Tracking Database');
    sheetId = ss.getId();
    PROPS.setProperty('SHEET_ID', sheetId);

    const ps = ss.getActiveSheet();
    ps.setName('patients');
    ps.appendRow(PATIENT_HEADERS);
    ps.getRange(1, 1, 1, PATIENT_HEADERS.length)
      .setFontWeight('bold').setBackground('#2d4f3f').setFontColor('#ffffff');
    ps.setFrozenRows(1);

    const vs = ss.insertSheet('visits');
    vs.appendRow(VISIT_HEADERS);
    vs.getRange(1, 1, 1, VISIT_HEADERS.length)
      .setFontWeight('bold').setBackground('#2d4f3f').setFontColor('#ffffff');
    vs.setFrozenRows(1);

    const ls = ss.insertSheet('audit_log');
    ls.appendRow(['timestamp', 'user', 'action', 'entity', 'entityId', 'details']);
    ls.getRange(1, 1, 1, 6)
      .setFontWeight('bold').setBackground('#6b7770').setFontColor('#ffffff');
    ls.setFrozenRows(1);
  }

  if (!folderId) {
    const folder = DriveApp.createFolder('TB Tracking Photos');
    folderId = folder.getId();
    PROPS.setProperty('FOLDER_ID', folderId);
  }

  const result = {
    sheetId: sheetId,
    folderId: folderId,
    sheetUrl: 'https://docs.google.com/spreadsheets/d/' + sheetId,
    folderUrl: 'https://drive.google.com/drive/folders/' + folderId,
    message: '✓ Setup complete'
  };
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

/* ─────────────────────────────────────────────────────────────────
   MIGRATE — รันครั้งเดียวเพื่อเพิ่มคอลัมน์ใหม่ให้ Sheet เดิม
   + พยายาม backfill address ที่รวมเป็นสตริงเดียว ให้แยกกลับเป็นช่อง
   ───────────────────────────────────────────────────────────────── */
function migrate() {
  return withLock(() => {
    const ss = getSS();
    const result = { patients: {}, visits: {} };

    result.patients = ensureHeaders(ss.getSheetByName('patients'), PATIENT_HEADERS);
    result.visits = ensureHeaders(ss.getSheetByName('visits'), VISIT_HEADERS);

    result.addressBackfill = backfillAddresses();
    result.garbageCleanup = cleanGarbageJSONFields();

    Logger.log(JSON.stringify(result, null, 2));
    return result;
  });
}

// Clear cells that contain Java-toString garbage like "[Ljava.lang.Object;@hash"
// or "{key=value, ...}" which were written by an earlier Apps Script bug.
// Also re-stringify any cell that holds a non-string JSON value.
function cleanGarbageJSONFields() {
  const result = { patients: 0, visits: 0 };
  result.patients = cleanJsonInSheet(getSS().getSheetByName('patients'), PATIENT_JSON_FIELDS);
  result.visits = cleanJsonInSheet(getSS().getSheetByName('visits'), VISIT_JSON_FIELDS);
  return result;
}

function cleanJsonInSheet(sheet, jsonFields) {
  if (!sheet) return 0;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  let cleaned = 0;
  jsonFields.forEach(f => {
    const colIdx = headers.indexOf(f);
    if (colIdx < 0) return;
    const range = sheet.getRange(2, colIdx + 1, lastRow - 1, 1);
    const values = range.getValues();
    let dirty = false;
    for (let r = 0; r < values.length; r++) {
      const v = values[r][0];
      if (typeof v !== 'string' || !v) continue;
      // Java toString patterns
      if (v.indexOf('[Ljava.') === 0
          || /^\{[^"]*=/.test(v)
          || /=,/.test(v)) {
        values[r][0] = '';
        cleaned++;
        dirty = true;
      }
    }
    if (dirty) range.setValues(values);
  });
  return cleaned;
}

function ensureHeaders(sheet, requiredHeaders) {
  if (!sheet) return { added: [], existing: [] };
  const lastCol = sheet.getLastColumn();
  const existing = lastCol > 0
    ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].filter(h => h !== '')
    : [];
  const missing = requiredHeaders.filter(h => existing.indexOf(h) < 0);
  if (missing.length === 0) return { added: [], existing: existing };
  const startCol = existing.length + 1;
  sheet.getRange(1, startCol, 1, missing.length).setValues([missing]);
  sheet.getRange(1, 1, 1, existing.length + missing.length)
    .setFontWeight('bold').setBackground('#2d4f3f').setFontColor('#ffffff');
  return { added: missing, existing: existing };
}

/**
 * แยก address ที่เก็บรวมเป็นสตริงเดียว ("99/1 ม.5 ต.x อ.y จ.z")
 * กลับเป็นช่อง addrNo / addrMoo / addrTambon / addrAmphoe / addrProvince
 * ทำเฉพาะแถวที่ addrNo ยังว่าง (ไม่ทับข้อมูลที่ user เพิ่งกรอกแยกช่อง)
 */
function backfillAddresses() {
  const sheet = getSS().getSheetByName('patients');
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return { scanned: 0, updated: 0 };
  const headers = values[0];
  const idx = name => headers.indexOf(name);
  const cAddress = idx('address');
  const cNo = idx('addrNo');
  const cMoo = idx('addrMoo');
  const cTambon = idx('addrTambon');
  const cAmphoe = idx('addrAmphoe');
  const cProvince = idx('addrProvince');
  if ([cAddress, cNo, cMoo, cTambon, cAmphoe, cProvince].some(c => c < 0)) {
    return { error: 'missing required columns' };
  }

  let updated = 0;
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const addr = (row[cAddress] || '').toString().trim();
    const hasSplit = row[cNo] || row[cMoo] || row[cTambon] || row[cAmphoe] || row[cProvince];
    if (!addr || hasSplit) continue;
    const parts = parseAddressString(addr);
    if (!parts) continue;
    sheet.getRange(i + 1, cNo + 1).setValue(parts.addrNo || '');
    sheet.getRange(i + 1, cMoo + 1).setValue(parts.addrMoo || '');
    sheet.getRange(i + 1, cTambon + 1).setValue(parts.addrTambon || '');
    sheet.getRange(i + 1, cAmphoe + 1).setValue(parts.addrAmphoe || '');
    sheet.getRange(i + 1, cProvince + 1).setValue(parts.addrProvince || '');
    updated++;
  }
  return { scanned: values.length - 1, updated: updated };
}

function parseAddressString(s) {
  if (!s) return null;
  const out = { addrNo: '', addrMoo: '', addrTambon: '', addrAmphoe: '', addrProvince: '' };
  const re = /^(.*?)\s*(?:ม\.?\s*(\S+))?\s*(?:ต\.?\s*(\S+))?\s*(?:อ\.?\s*(\S+))?\s*(?:จ\.?\s*(\S+))?\s*$/u;
  const m = s.match(re);
  if (!m) { out.addrNo = s; return out; }
  out.addrNo = (m[1] || '').trim();
  out.addrMoo = (m[2] || '').trim();
  out.addrTambon = (m[3] || '').trim();
  out.addrAmphoe = (m[4] || '').trim();
  out.addrProvince = (m[5] || '').trim();
  if (!out.addrNo && !out.addrMoo && !out.addrTambon) { out.addrNo = s; }
  return out;
}

/* ─────────────────────────────────────────────────────────────────
   ROUTING
   ───────────────────────────────────────────────────────────────── */
function doGet(e) {
  const action = (e.parameter && e.parameter.action) || 'ping';
  return jsonResponse(handle(action, e.parameter));
}

function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch (err) {}
  // LINE webhook: body has "events" array (no "action" key) — route to LINE handler
  if (body && Array.isArray(body.events)) {
    handleLineWebhook(body);
    // LINE requires 200 OK with empty/short body
    return ContentService.createTextOutput('OK').setMimeType(ContentService.MimeType.TEXT);
  }
  const action = body.action || 'ping';
  return jsonResponse(handle(action, body));
}

function handle(action, params) {
  try {
    switch (action) {
      case 'ping':         return { ok: true, message: 'Apps Script ทำงานแล้ว', time: new Date().toISOString(), user: getUser() };
      case 'getConfig':    return { ok: true, data: getConfig() };
      case 'listPatients': return { ok: true, data: listPatients() };
      case 'addPatient':   return { ok: true, data: addPatient(params.data || params) };
      case 'updatePatient':return { ok: true, data: updatePatient(params.id, params.data) };
      case 'deletePatient':return { ok: true, data: deletePatient(params.id) };
      case 'listVisits':   return { ok: true, data: listVisits(params.patientId) };
      case 'addVisit':     return { ok: true, data: addVisit(params.data || params) };
      case 'deleteVisit':  return { ok: true, data: deleteVisit(params.id) };
      case 'uploadPhoto':  return { ok: true, data: uploadPhoto(params.filename, params.base64, params.patientId) };
      case 'getStats':     return { ok: true, data: getStats() };
      case 'migrate':      return { ok: true, data: migrate() };
      case 'reset':        return { ok: true, data: resetAll() };
      case 'getLineStatus':return { ok: true, data: getLineStatus() };
      default:             return { ok: false, error: 'Unknown action: ' + action };
    }
  } catch (err) {
    return { ok: false, error: err.toString(), stack: err.stack };
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ─────────────────────────────────────────────────────────────────
   HELPERS
   ───────────────────────────────────────────────────────────────── */
function getUser() {
  try {
    const email = Session.getActiveUser().getEmail();
    return email || 'anonymous';
  } catch (e) { return 'anonymous'; }
}

function getSS()     { return SpreadsheetApp.openById(PROPS.getProperty('SHEET_ID')); }
function getFolder() { return DriveApp.getFolderById(PROPS.getProperty('FOLDER_ID')); }

function getConfig() {
  return {
    sheetId: PROPS.getProperty('SHEET_ID'),
    folderId: PROPS.getProperty('FOLDER_ID'),
    sheetUrl: 'https://docs.google.com/spreadsheets/d/' + PROPS.getProperty('SHEET_ID'),
    folderUrl: 'https://drive.google.com/drive/folders/' + PROPS.getProperty('FOLDER_ID'),
    user: getUser()
  };
}

function getHeaders(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return [];
  return sheet.getRange(1, 1, 1, lastCol).getValues()[0];
}

function rowsToObjects(sheet, jsonFields) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  const jset = jsonFields || [];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    jset.forEach(f => {
      const v = obj[f];
      if (typeof v === 'string' && v.length && (v.charAt(0) === '[' || v.charAt(0) === '{')) {
        try { obj[f] = JSON.parse(v); } catch (e) { /* leave as string */ }
      } else if (v == null || v === '') {
        obj[f] = (jset.indexOf(f) >= 0) ? [] : v;
      }
    });
    return obj;
  }).filter(o => o.id);
}

function objectToRow(obj, headers, jsonFields) {
  const jset = jsonFields || [];
  return headers.map(h => {
    let v = obj[h];
    if (jset.indexOf(h) >= 0) return safeJSONStringify(v);
    return v != null ? v : '';
  });
}

// Deep-clone a value into pure JS structures, breaking any Java bridging.
// Workaround for Apps Script V8 quirk where JSON.parse(body).x can be
// Java-backed and JSON.stringify(it) returns "[Ljava.lang.Object;@hash" /
// "{key=value}" instead of valid JSON.
function plainClone(v) {
  if (v == null) return v;
  var t = typeof v;
  if (t !== 'object') return v;
  // Detect array-ish values (covers Java Object[] which doesn't pass Array.isArray)
  var isArr = (typeof Array.isArray === 'function' && Array.isArray(v))
    || Object.prototype.toString.call(v) === '[object Array]'
    || (typeof v.length === 'number' && t === 'object' && !v.hasOwnProperty('length'));
  if (isArr) {
    var arr = [];
    for (var i = 0; i < v.length; i++) arr.push(plainClone(v[i]));
    return arr;
  }
  var obj = {};
  for (var k in v) {
    try {
      if (Object.prototype.hasOwnProperty.call(v, k)) obj[k] = plainClone(v[k]);
    } catch (e) {}
  }
  return obj;
}

function safeJSONStringify(v) {
  if (v == null || v === '') return '';
  if (typeof v === 'string') return v;
  try { return JSON.stringify(plainClone(v)); } catch (e) { return ''; }
}

function uid(prefix) {
  return prefix + '_' + Date.now().toString(36) + '_' + Math.random().toString(36).slice(2, 7);
}

function withLock(fn) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    return fn();
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function audit(action, entity, entityId, details) {
  try {
    const sheet = getSS().getSheetByName('audit_log');
    sheet.appendRow([new Date(), getUser(), action, entity, entityId || '', details || '']);
  } catch (e) { Logger.log('Audit failed: ' + e); }
}

/* ─────────────────────────────────────────────────────────────────
   PATIENTS
   ───────────────────────────────────────────────────────────────── */
function listPatients() {
  const sheet = getSS().getSheetByName('patients');
  ensureHeaders(sheet, PATIENT_HEADERS);
  return rowsToObjects(sheet, PATIENT_JSON_FIELDS);
}

function addPatient(data) {
  return withLock(() => {
    const sheet = getSS().getSheetByName('patients');
    ensureHeaders(sheet, PATIENT_HEADERS);
    extendHeadersForUnknownKeys(sheet, data);
    const headers = getHeaders(sheet);

    const id = data.id || uid('p');
    const now = new Date().toISOString();
    const row = {
      ...data,
      id: id,
      createdAt: data.createdAt || now,
      createdBy: data.createdBy || getUser(),
      updatedAt: now,
      updatedBy: data.updatedBy || getUser()
    };
    sheet.appendRow(objectToRow(row, headers, PATIENT_JSON_FIELDS));
    audit('create', 'patient', id, row.name || '');
    // Fire LINE notification (best-effort — won't fail patient save if LINE is down)
    try { notifyLineNewPatient(row); } catch (e) { Logger.log('LINE notify error: ' + e); }
    return readbackPatient(id);
  });
}

function updatePatient(id, data) {
  return withLock(() => {
    const sheet = getSS().getSheetByName('patients');
    ensureHeaders(sheet, PATIENT_HEADERS);
    extendHeadersForUnknownKeys(sheet, data);
    const headers = getHeaders(sheet);
    const values = sheet.getDataRange().getValues();
    const idCol = headers.indexOf('id');

    for (let i = 1; i < values.length; i++) {
      if (values[i][idCol] === id) {
        const existingObj = {};
        headers.forEach((h, j) => { existingObj[h] = values[i][j]; });
        // parse existing JSON fields so we don't double-stringify on merge
        PATIENT_JSON_FIELDS.forEach(f => {
          const v = existingObj[f];
          if (typeof v === 'string' && v.length && (v.charAt(0) === '[' || v.charAt(0) === '{')) {
            try { existingObj[f] = JSON.parse(v); } catch (e) {}
          }
        });
        const merged = {
          ...existingObj,
          ...data,
          id: id,
          updatedAt: new Date().toISOString(),
          updatedBy: data.updatedBy || getUser()
        };
        sheet.getRange(i + 1, 1, 1, headers.length)
          .setValues([objectToRow(merged, headers, PATIENT_JSON_FIELDS)]);
        audit('update', 'patient', id, JSON.stringify(Object.keys(data)));
        return readbackPatient(id);
      }
    }
    throw new Error('Patient not found: ' + id);
  });
}

function readbackPatient(id) {
  const all = listPatients();
  for (let i = 0; i < all.length; i++) {
    if (all[i].id === id) return all[i];
  }
  return null;
}

function extendHeadersForUnknownKeys(sheet, data) {
  if (!data || typeof data !== 'object') return;
  const existing = getHeaders(sheet);
  const unknown = Object.keys(data).filter(k => existing.indexOf(k) < 0);
  if (unknown.length === 0) return;
  const startCol = existing.length + 1;
  sheet.getRange(1, startCol, 1, unknown.length).setValues([unknown]);
  sheet.getRange(1, 1, 1, existing.length + unknown.length)
    .setFontWeight('bold').setBackground('#2d4f3f').setFontColor('#ffffff');
}

function deletePatient(id) {
  return withLock(() => {
    const ss = getSS();
    const sheet = ss.getSheetByName('patients');
    const values = sheet.getDataRange().getValues();
    const idCol = values[0].indexOf('id');
    for (let i = 1; i < values.length; i++) {
      if (values[i][idCol] === id) {
        sheet.deleteRow(i + 1);
        const vs = ss.getSheetByName('visits');
        const vData = vs.getDataRange().getValues();
        const vIdCol = vData[0].indexOf('patientId');
        for (let j = vData.length - 1; j >= 1; j--) {
          if (vData[j][vIdCol] === id) vs.deleteRow(j + 1);
        }
        audit('delete', 'patient', id, '');
        return { deleted: id };
      }
    }
    throw new Error('Patient not found: ' + id);
  });
}

/* ─────────────────────────────────────────────────────────────────
   VISITS
   ───────────────────────────────────────────────────────────────── */
function listVisits(patientId) {
  const sheet = getSS().getSheetByName('visits');
  ensureHeaders(sheet, VISIT_HEADERS);
  const all = rowsToObjects(sheet, VISIT_JSON_FIELDS);
  return patientId ? all.filter(v => v.patientId === patientId) : all;
}

function addVisit(data) {
  return withLock(() => {
    const sheet = getSS().getSheetByName('visits');
    ensureHeaders(sheet, VISIT_HEADERS);
    extendHeadersForUnknownKeys(sheet, data);
    const headers = getHeaders(sheet);

    const id = uid('v');
    const row = {
      ...data,
      id: id,
      createdAt: new Date().toISOString(),
      createdBy: data.createdBy || getUser()
    };
    sheet.appendRow(objectToRow(row, headers, VISIT_JSON_FIELDS));
    audit('create', 'visit', id, 'patientId=' + (data.patientId || ''));
    return readbackVisit(id);
  });
}

function readbackVisit(id) {
  const all = listVisits();
  for (let i = 0; i < all.length; i++) {
    if (all[i].id === id) return all[i];
  }
  return null;
}

function deleteVisit(id) {
  return withLock(() => {
    const sheet = getSS().getSheetByName('visits');
    const values = sheet.getDataRange().getValues();
    const idCol = values[0].indexOf('id');
    for (let i = 1; i < values.length; i++) {
      if (values[i][idCol] === id) {
        sheet.deleteRow(i + 1);
        audit('delete', 'visit', id, '');
        return { deleted: id };
      }
    }
    throw new Error('Visit not found: ' + id);
  });
}

/* ─────────────────────────────────────────────────────────────────
   PHOTO UPLOAD
   ───────────────────────────────────────────────────────────────── */
function uploadPhoto(filename, base64, patientId) {
  if (!base64) throw new Error('No base64 data');

  const match = base64.match(/^data:(image\/\w+);base64,(.+)$/);
  let mimeType = 'image/jpeg';
  let data = base64;
  if (match) {
    mimeType = match[1];
    data = match[2];
  }

  const bytes = Utilities.base64Decode(data);
  const blob = Utilities.newBlob(bytes, mimeType, filename || ('photo_' + Date.now() + '.jpg'));

  let folder = getFolder();
  if (patientId) {
    const sub = folder.getFoldersByName(patientId);
    folder = sub.hasNext() ? sub.next() : folder.createFolder(patientId);
  }
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return {
    fileId: file.getId(),
    name: file.getName(),
    size: file.getSize(),
    url: 'https://drive.google.com/uc?id=' + file.getId(),
    viewUrl: file.getUrl()
  };
}

/* ─────────────────────────────────────────────────────────────────
   STATS
   ───────────────────────────────────────────────────────────────── */
function getStats() {
  const patients = listPatients();
  const visits = listVisits();
  const byStatus = {};
  patients.forEach(p => { byStatus[p.status || 'unknown'] = (byStatus[p.status || 'unknown'] || 0) + 1; });
  return {
    totalPatients: patients.length,
    totalVisits: visits.length,
    byStatus: byStatus,
    user: getUser(),
    timestamp: new Date().toISOString()
  };
}

/* ─────────────────────────────────────────────────────────────────
   RESET — ลบข้อมูลทั้งหมด (ระวัง!)
   ───────────────────────────────────────────────────────────────── */
function resetAll() {
  return withLock(() => {
    const ss = getSS();
    ['patients', 'visits', 'audit_log'].forEach(name => {
      const sheet = ss.getSheetByName(name);
      if (sheet && sheet.getLastRow() > 1) {
        sheet.deleteRows(2, sheet.getLastRow() - 1);
      }
    });
    return { ok: true, message: 'Cleared all data' };
  });
}

/* ═══════════════════════════════════════════════════════════════════
   LINE MESSAGING API — แจ้งเตือนเข้ากลุ่ม รพ.สต. ทุกครั้งที่เพิ่มผู้ป่วย
   ═══════════════════════════════════════════════════════════════════

   วิธี setup (ครั้งเดียว):
   1. paste โค้ดทั้งหมดนี้ลง Apps Script editor + Save
   2. เลือก function "setLineToken" → กดเครื่องหมาย ▶ Run
      (ครั้งแรก authorize permission)
      เปลี่ยน 'PASTE_TOKEN_HERE' ในฟังก์ชันเป็น token จริง แล้ว Run
      หรือเรียก setLineToken("token...") ผ่าน console ก็ได้
   3. Deploy → Manage deployments → Edit (ดินสอ) → New version → Deploy
      copy Web App URL
   4. ไป LINE Developers Console → channel ของคุณ → แท็บ Messaging API
      - Webhook URL: ใส่ Web App URL ที่ได้
      - Use webhook: เปิด
      - กด Verify (ควรขึ้น Success)
   5. ในกลุ่ม LINE ที่มี bot อยู่ → พิมพ์ข้อความใดก็ได้
      bot จะดักจับ groupId อัตโนมัติเก็บใน PropertiesService
   6. รัน function "getLineStatus" เพื่อเช็กว่า groupId ถูก save หรือไม่
   7. ทดสอบ: เพิ่มผู้ป่วยใหม่ในเว็บ → ดูข้อความในกลุ่ม LINE
   ═══════════════════════════════════════════════════════════════════ */

const LINE_API = 'https://api.line.me/v2/bot/message/push';

// เรียกครั้งเดียวเพื่อเก็บ token (อย่า hardcode token ไว้ในโค้ด)
function setLineToken(token) {
  if (!token || token === 'PASTE_TOKEN_HERE') {
    throw new Error('กรุณาส่ง token จริงเป็น argument: setLineToken("xxx...")');
  }
  PROPS.setProperty('LINE_TOKEN', token);
  return { ok: true, message: 'LINE_TOKEN saved (' + token.length + ' chars)' };
}

// ดูสถานะการตั้งค่า LINE (เช็กว่า token + groupId พร้อมหรือยัง)
function getLineStatus() {
  const token = PROPS.getProperty('LINE_TOKEN');
  const groupId = PROPS.getProperty('LINE_GROUP_ID');
  return {
    tokenSet: !!token,
    tokenLength: token ? token.length : 0,
    groupId: groupId || null,
    ready: !!(token && groupId),
    hint: !token ? 'รัน setLineToken("...") ก่อน' :
          !groupId ? 'เพิ่ม bot เข้ากลุ่ม + พิมพ์ข้อความในกลุ่ม → bot จะดักจับ groupId' :
          'พร้อมใช้งาน ✓'
  };
}

// (manual setter เผื่อ webhook ไม่ทำงาน — ตั้งค่า groupId เอง)
function setLineGroupId(groupId) {
  PROPS.setProperty('LINE_GROUP_ID', groupId);
  return { ok: true, groupId: groupId };
}

// ดักจับ LINE webhook events (capture groupId เมื่อมีข้อความในกลุ่ม)
function handleLineWebhook(body) {
  if (!body || !Array.isArray(body.events)) return;
  body.events.forEach(ev => {
    try {
      const src = ev.source || {};
      // Group/room message → save groupId
      if (src.type === 'group' && src.groupId) {
        PROPS.setProperty('LINE_GROUP_ID', src.groupId);
        Logger.log('Captured LINE group ID: ' + src.groupId);
      } else if (src.type === 'room' && src.roomId) {
        // Multi-person chat (less common)
        PROPS.setProperty('LINE_GROUP_ID', src.roomId);
        PROPS.setProperty('LINE_TARGET_TYPE', 'room');
        Logger.log('Captured LINE room ID: ' + src.roomId);
      } else if (src.type === 'user' && src.userId) {
        // Save user ID as fallback (1-on-1 chat with bot)
        PROPS.setProperty('LINE_USER_ID', src.userId);
        Logger.log('Captured LINE user ID: ' + src.userId);
      }
    } catch (e) { Logger.log('Webhook parse error: ' + e); }
  });
}

// ส่งข้อความเข้ากลุ่ม (best-effort — ถ้า fail จะ log แต่ไม่ throw)
function sendLineMessage(text) {
  const token = PROPS.getProperty('LINE_TOKEN');
  const groupId = PROPS.getProperty('LINE_GROUP_ID') || PROPS.getProperty('LINE_USER_ID');
  if (!token || !groupId) {
    Logger.log('LINE not configured (token=' + !!token + ', target=' + !!groupId + ')');
    return false;
  }
  try {
    const res = UrlFetchApp.fetch(LINE_API, {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + token },
      payload: JSON.stringify({
        to: groupId,
        messages: [{ type: 'text', text: text.slice(0, 4900) }]
      }),
      muteHttpExceptions: true
    });
    const code = res.getResponseCode();
    if (code !== 200) Logger.log('LINE push fail ' + code + ': ' + res.getContentText());
    return code === 200;
  } catch (e) {
    Logger.log('LINE push exception: ' + e);
    return false;
  }
}

// ฟอร์แมตข้อความและส่งเมื่อเพิ่มผู้ป่วยใหม่
function notifyLineNewPatient(p) {
  if (!p) return;
  const fmtDate = (iso) => {
    if (!iso) return '—';
    try {
      const d = new Date(iso);
      return d.getDate() + '/' + (d.getMonth() + 1) + '/' + (d.getFullYear() + 543);
    } catch (e) { return String(iso); }
  };
  const lines = [
    '🆕 เพิ่มผู้ป่วยวัณโรครายใหม่',
    '━━━━━━━━━━━━━━━━━━',
    '👤 ' + (p.name || '—'),
    '🆔 HN ' + (p.hn || '—') + ' · ' + (p.age || '—') + ' ปี · ' + (p.gender || '—'),
    '📍 เขต ' + (p.zone || '—') + ' (' + (p.areaType || '—') + ')',
    '🩺 ' + (p.tbClass || '—') + (p.tbClassDetail ? ' (' + p.tbClassDetail + ')' : '') + ' · ' + (p.patientType || '—'),
    '🧪 ผลตรวจ ' + (p.testResult || p.afbResult || '—'),
    '💊 เริ่มรักษา ' + fmtDate(p.treatmentStartDate),
    '━━━━━━━━━━━━━━━━━━',
    'บันทึกโดย ' + (p.createdBy || '—')
  ];
  sendLineMessage(lines.join('\n'));
}

// (เสริม) ส่งข้อความทดสอบ — รัน function นี้เพื่อทดสอบว่าระบบทำงาน
function sendLineTest() {
  const ok = sendLineMessage('🔔 ทดสอบการแจ้งเตือนจาก TB Care Trace\nเวลา ' + new Date().toLocaleString('th-TH'));
  return { ok: ok, status: getLineStatus() };
}
