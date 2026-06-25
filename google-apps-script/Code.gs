const SHEET_STUDENTS = 'Students';
const SHEET_HOUSES = 'Houses';
const SHEET_FLATS = 'Flats';
const SHEET_EXPORT = 'IDCardExport';
const MAX_STUDENTS_PER_HOUSE = 6;
const GEMINI_MODEL = 'gemini-2.0-flash';

// Target spreadsheet (opened by id so the script need not be container-bound).
const SPREADSHEET_ID = '1ViLI-JdrVVOAIGGhABH-vzm_6bfquL8FZHXvpGbnZ40';
// Default Drive folder for per-flat subfolders; a IMAGE_FOLDER_ID script property overrides this.
const DEFAULT_IMAGE_FOLDER_ID = '1Jpj5eVuKdXEichmJ_hx54mUA4Z8K20c4';
// Default Annexure 2A template Google Doc; an ANNEXURE_TEMPLATE_ID script property overrides this.
const DEFAULT_ANNEXURE_TEMPLATE_ID = '1quqEh2jLVxPWSg9WeIz3BAiwZtScWwoKoAmEWCMtWe4';

// Canonical header for the Students sheet. Existing columns are preserved in order;
// new columns are appended. ensureStudentColumns_() adds any missing column on the fly
// so older sheets keep working.
const STUDENTS_HEADER = [
  'Timestamp',
  'StudentId',
  'Name',
  'AadharNumber',
  'Phone',
  'HouseId',
  'AgreementStatus',
  'AgreementRef',
  'PhotoUrl',
  'AadharPhotoUrl',
  'Notes',
  'Status',
  // --- added for the Annexure 2A flow ---
  'Email',
  'Sex',
  'ParentName',
  'ParentMobile',
  'Address',
  'CollegeStudentId',
  'CollegeName',
  'AcademicYear',
  'VehicleNumber',
  'CollegeIdPhotoUrl',
  'SignatureUrl',
  'AssignedId',
  'CourseName',
  'IdValidUpto'
];

// Flat-level owner / caretaker block for the Annexure 2A header + page-5 undertaking.
const FLATS_HEADER = [
  'Flat#',
  'OwnerName',
  'OwnerEmail',
  'OwnerMobile',
  'CaretakerName',
  'CaretakerNumber',
  'OwnerSignatureUrl',
  'SigningDate',
  'SigningPlace',
  'TenancyPeriod',
  'AgreementDate'
];

function getSpreadsheet_() {
  return SPREADSHEET_ID
    ? SpreadsheetApp.openById(SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
}

function doGet(e) {
  // Read-only JSON diagnostic (reachable over GET): ?api=flatdebug&flat=1137
  if (e && e.parameter && e.parameter.api === 'flatdebug') {
    return ContentService
      .createTextOutput(JSON.stringify(debugFlat_(e.parameter.flat)))
      .setMimeType(ContentService.MimeType.JSON);
  }
  // Verify the Annexure template's {{tokens}} match what the code fills: ?api=tplcheck
  if (e && e.parameter && e.parameter.api === 'tplcheck') {
    return ContentService
      .createTextOutput(JSON.stringify(checkAnnexureTemplate_()))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Default route serves the per-student self-fill form (the primary flow).
  // ?page=legacy serves the original 6-tab onboarding tool (Index.html).
  const page = (e && e.parameter && e.parameter.page) || '';
  let file = 'StudentForm', title = 'Student Details — Annexure 2A';
  if (page === 'legacy') { file = 'Index'; title = 'Student Onboarding (legacy)'; }
  else if (page === 'admin') { file = 'Admin'; title = 'Admin — Verify & Compare'; }
  const tmpl = HtmlService.createTemplateFromFile(file);
  tmpl.flat = (e && e.parameter && e.parameter.flat) || '';
  return tmpl
    .evaluate()
    .setTitle(title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// JSON API for the native admin app (token-guarded via the API_TOKEN script property).
function doPost(e) {
  let out;
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const token = PropertiesService.getScriptProperties().getProperty('API_TOKEN');
    if (token && String(body.token || '') !== token) throw new Error('Unauthorized.');

    const payload = body.payload || body;
    let result;
    switch (body.action) {
      case 'getFlatForReview':       result = getFlatForReview(body.flat || (payload && payload.flat)); break;
      case 'submitStudentSelf':      result = submitStudentSelf(payload); break;
      case 'getDocsForVerify':       result = getDocsForVerify_(body.studentId || (payload && payload.studentId)); break;
      case 'uploadAgreement':        result = uploadAgreement_(payload); break;
      case 'saveStudentVerification':result = saveStudentVerification(payload); break;
      case 'saveFlatSigning':        result = saveFlatSigning(payload); break;
      case 'generateAnnexurePdf':    result = generateAnnexurePdf(body.flat || (payload && payload.flat)); break;
      // Admin web-page actions (also callable via google.script.run when embedded).
      case 'getLogo':                result = getLogo(); break;
      case 'getCompareDocs':         result = getCompareDocs(payload.studentId); break;
      case 'getFlatAgreementDoc':    result = getFlatAgreementDoc(body.flat || payload.flat); break;
      case 'saveStudentDetails':     result = saveStudentDetails(payload.studentId, payload.fields); break;
      case 'saveFlatDetails':        result = saveFlatDetails(body.flat || payload.flat, payload.fields); break;
      case 'assignStudentId':        result = assignStudentId(payload.studentId, payload.assignedId); break;
      case 'webUploadAgreement':     result = webUploadAgreement(payload.flat, payload.base64, payload.fileName, payload.mimeType); break;
      default: throw new Error('Unknown action: ' + body.action);
    }
    out = { ok: true, result: result };
  } catch (err) {
    out = { ok: false, error: (err && err.message) ? err.message : String(err) };
  }
  return ContentService
    .createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function setupSheets() {
  const ss = getSpreadsheet_();

  if (!ss.getSheetByName(SHEET_HOUSES)) {
    const houses = ss.insertSheet(SHEET_HOUSES);
    houses.appendRow([
      'HouseId',
      'HouseName',
      'OwnerName',
      'OwnerEmail',
      'OwnerPhone',
      'Active'
    ]);
    houses.appendRow(['H001', 'Lotus House', 'Owner 1', 'owner1@example.com', '9876543210', true]);
    houses.appendRow(['H002', 'Maple House', 'Owner 2', 'owner2@example.com', '9876543211', true]);
  }

  if (!ss.getSheetByName(SHEET_STUDENTS)) {
    const students = ss.insertSheet(SHEET_STUDENTS);
    students.appendRow(STUDENTS_HEADER.slice());
  } else {
    ensureStudentColumns_();
  }

  if (!ss.getSheetByName(SHEET_FLATS)) {
    const flats = ss.insertSheet(SHEET_FLATS);
    flats.appendRow(FLATS_HEADER.slice());
  }

  if (!ss.getSheetByName(SHEET_EXPORT)) {
    const exp = ss.insertSheet(SHEET_EXPORT);
    exp.appendRow(['StudentId', 'Name', 'Phone', 'AadharNumber', 'PhotoUrl', 'HouseId']);
  }
}

// Appends any missing column from STUDENTS_HEADER to an existing Students sheet so
// the new Annexure 2A fields can be written without recreating the sheet. Safe to
// re-run; only adds columns that don't already exist.
function ensureStudentColumns_() {
  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName(SHEET_STUDENTS);
  if (!sh) return;

  const lastCol = Math.max(1, sh.getLastColumn());
  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const missing = STUDENTS_HEADER.filter(h => header.indexOf(h) === -1);
  if (!missing.length) return;

  sh.getRange(1, header.length + 1, 1, missing.length).setValues([missing]);
}

// Same migration for the Flats sheet — appends any FLATS_HEADER column that an older
// sheet is missing (e.g. TenancyPeriod / AgreementDate added later), so those values
// can actually be stored and read.
function ensureFlatColumns_() {
  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName(SHEET_FLATS);
  if (!sh) return;
  const lastCol = Math.max(1, sh.getLastColumn());
  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const missing = FLATS_HEADER.filter(h => header.indexOf(h) === -1);
  if (!missing.length) return;
  sh.getRange(1, header.length + 1, 1, missing.length).setValues([missing]);
}

function getActiveHouses() {
  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName(SHEET_HOUSES);
  if (!sh) {
    throw new Error('Houses sheet not found. Run setupSheets() first.');
  }

  const values = sh.getDataRange().getValues();
  const header = values.shift();
  const iHouseId = header.indexOf('HouseId');
  const iHouseName = header.indexOf('HouseName');
  const iOwnerName = header.indexOf('OwnerName');
  const iActive = header.indexOf('Active');

  return values
    .filter(row => row[iActive] === true)
    .map(row => ({
      houseId: row[iHouseId],
      label: row[iHouseName] + ' (' + row[iOwnerName] + ')'
    }));
}

function getStudentsByFlat(flatNumber) {
  const flat = String(flatNumber || '').trim();
  if (!flat) {
    return [];
  }

  const ss = getSpreadsheet_();
  const students = ss.getSheetByName(SHEET_STUDENTS);
  if (!students) {
    throw new Error('Students sheet not found. Run setupSheets() first.');
  }

  const values = students.getDataRange().getValues();
  const header = values.shift();
  const iStudentId = header.indexOf('StudentId');
  const iName = header.indexOf('Name');
  const iAadhar = header.indexOf('AadharNumber');
  const iPhone = header.indexOf('Phone');
  const iHouseId = header.indexOf('HouseId');
  const iAgreementStatus = header.indexOf('AgreementStatus');
  const iAgreementRef = header.indexOf('AgreementRef');
  const iPhoto = header.indexOf('PhotoUrl');
  const iAadharPhoto = header.indexOf('AadharPhotoUrl');
  const iNotes = header.indexOf('Notes');
  const iStatus = header.indexOf('Status');

  const list = values
    .filter(r => String(r[iHouseId]) === flat && r[iStatus] === 'Active')
    .slice(0, MAX_STUDENTS_PER_HOUSE)
    .map(r => ({
      studentId: r[iStudentId],
      name: r[iName],
      aadharNumber: r[iAadhar],
      phone: r[iPhone],
      agreementStatus: r[iAgreementStatus],
      agreementRef: r[iAgreementRef],
      photoUrl: r[iPhoto],
      aadharPhotoUrl: r[iAadharPhoto],
      notes: r[iNotes]
    }));

  return {
    students: list,
    agreementUrl: getFlatAgreementUrl_(flat),
    agreementInfo: getFlatAgreementInfo_(flat)
  };
}

function saveStudent(payload) {
  return payload && payload.studentId
    ? updateStudent_(payload)
    : submitOnboarding(payload);
}

function submitOnboarding(payload) {
  validatePayload_(payload, false);

  const ss = getSpreadsheet_();
  const students = ss.getSheetByName(SHEET_STUDENTS);
  if (!students) {
    throw new Error('Students sheet not found. Run setupSheets() first.');
  }

  const currentCount = countActiveStudentsByHouse_(payload.houseId);
  if (currentCount >= MAX_STUDENTS_PER_HOUSE) {
    throw new Error('This house already has 6 active students.');
  }

  const studentId = generateStudentId_();
  const folder = getOrCreateFlatFolder_(payload.houseId);
  const photoUrl = saveImageInFolder_(folder, payload.studentPhotoBase64, studentId + '.jpg');
  const aadharPhotoUrl = saveImageInFolder_(folder, payload.aadharPhotoBase64, studentId + '_aadhar.jpg');
  if (payload.collegePhotoBase64) {
    saveImageInFolder_(folder, payload.collegePhotoBase64, studentId + '_collegeid.jpg');
  }
  saveDetailTxt_(folder, studentId, payload);

  students.appendRow([
    new Date(),
    studentId,
    payload.name,
    payload.aadharNumber,
    payload.phone,
    payload.houseId,
    payload.agreementStatus || '',
    payload.agreementRef || '',
    photoUrl,
    aadharPhotoUrl,
    payload.notes || '',
    'Active'
  ]);

  const whatsappMsg =
    'Welcome ' + payload.name +
    '. Your student ID is ' + studentId +
    '. Please keep this for ID card collection.';
  const whatsappLink = createWhatsAppLink_(payload.phone, whatsappMsg);

  return {
    ok: true,
    studentId,
    whatsappLink,
    message: 'Saved successfully.'
  };
}

function updateStudent_(payload) {
  validatePayload_(payload, true);

  const ss = getSpreadsheet_();
  const students = ss.getSheetByName(SHEET_STUDENTS);
  if (!students) {
    throw new Error('Students sheet not found. Run setupSheets() first.');
  }

  const found = findStudentRowById_(payload.studentId);
  if (!found) {
    throw new Error('Student not found for id ' + payload.studentId + '.');
  }

  const header = found.header;
  const row = found.row;
  const studentId = payload.studentId;

  const iPhoto = header.indexOf('PhotoUrl');
  const iAadharPhoto = header.indexOf('AadharPhotoUrl');

  const folder = getOrCreateFlatFolder_(payload.houseId);
  const photoUrl = payload.studentPhotoBase64
    ? saveImageInFolder_(folder, payload.studentPhotoBase64, studentId + '.jpg')
    : row[iPhoto];
  const aadharPhotoUrl = payload.aadharPhotoBase64
    ? saveImageInFolder_(folder, payload.aadharPhotoBase64, studentId + '_aadhar.jpg')
    : row[iAadharPhoto];
  if (payload.collegePhotoBase64) {
    saveImageInFolder_(folder, payload.collegePhotoBase64, studentId + '_collegeid.jpg');
  }
  saveDetailTxt_(folder, studentId, payload);

  const updated = row.slice();
  updated[header.indexOf('Name')] = payload.name;
  updated[header.indexOf('AadharNumber')] = payload.aadharNumber;
  updated[header.indexOf('Phone')] = payload.phone;
  updated[header.indexOf('HouseId')] = payload.houseId;
  updated[header.indexOf('AgreementStatus')] = payload.agreementStatus || '';
  updated[header.indexOf('AgreementRef')] = payload.agreementRef || '';
  updated[iPhoto] = photoUrl;
  updated[iAadharPhoto] = aadharPhotoUrl;
  updated[header.indexOf('Notes')] = payload.notes || '';

  students.getRange(found.rowNumber, 1, 1, updated.length).setValues([updated]);

  return {
    ok: true,
    studentId,
    message: 'Updated successfully.'
  };
}

// ===================================================================
//  Annexure 2A flow — student self-fill, admin verify, flat signing,
//  combined PDF generation. (See README / plan for the data model.)
// ===================================================================

// Returns the Students sheet plus its live header (after column migration),
// so code is resilient to column order and to newly-added columns.
function studentsSheetAndHeader_() {
  const ss = getSpreadsheet_();
  // Auto-create the Students sheet (and its header) on first use so submissions
  // never fail just because setupSheets() hasn't been run yet.
  if (!ss.getSheetByName(SHEET_STUDENTS)) {
    ss.insertSheet(SHEET_STUDENTS).appendRow(STUDENTS_HEADER.slice());
  }
  ensureStudentColumns_();
  const sh = ss.getSheetByName(SHEET_STUDENTS);
  const lastCol = Math.max(1, sh.getLastColumn());
  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  return { sheet: sh, header: header };
}

function buildStudentRow_(header, fields) {
  return header.map(h => (fields.hasOwnProperty(h) ? fields[h] : ''));
}

function cleanDigits_(v, max) {
  const d = String(v == null ? '' : v).replace(/\D/g, '');
  return max ? d.slice(0, max) : d;
}

// Base filename for a student's stored files: the student's name (sanitized) with the
// StudentId appended so two students never collide / overwrite each other's photos.
// e.g. "Priya Sharma_STU20260624121530". Falls back to the id (or 'student') if unnamed.
function studentFileBase_(name, studentId) {
  const n = String(name || '').trim().replace(/[\\\/:*?"<>|]+/g, ' ').replace(/\s+/g, ' ').trim();
  const id = String(studentId || '').trim();
  if (!n) return id || 'student';
  return id ? (n + '_' + id) : n;
}

// Student self-service submit from the mobile web form. Appends a new student row
// (Status='Submitted') or merges onto an existing row when a studentId is supplied,
// stores the Aadhaar + college-ID images, and upserts the flat's owner/caretaker row.
function submitStudentSelf(payload) {
  payload = payload || {};
  const flat = String(payload.flat || payload.houseId || '').trim();
  const name = String(payload.name || '').trim();
  const ownerName = String(payload.ownerName || '').trim();
  const ownerEmail = String(payload.ownerEmail || '').trim();
  const aadhar = cleanDigits_(payload.aadharNumber, 12);
  if (!flat) throw new Error('Flat number is required.');
  if (!ownerName) throw new Error('Owner name is required.');
  if (!ownerEmail) throw new Error('Owner email is required.');
  if (!name) throw new Error('Name is required.');
  if (aadhar.length !== 12) throw new Error('A valid 12-digit Aadhaar number is required.');

  const sh = studentsSheetAndHeader_();
  const sheet = sh.sheet;
  const header = sh.header;

  if (!payload.studentId && countFlatStudents_(flat) >= MAX_STUDENTS_PER_HOUSE) {
    throw new Error('This flat already has ' + MAX_STUDENTS_PER_HOUSE + ' students.');
  }

  const studentId = payload.studentId ? String(payload.studentId) : generateStudentId_();
  const folder = getOrCreateFlatFolder_(flat);
  const base = studentFileBase_(name, studentId);

  // Each document may be several files (front + back) or a single PDF.
  const aadharUrls = saveDocFiles_(folder, base, 'aadhar',
    normalizeFiles_(payload.aadharFiles, payload.aadharPhotoBase64));
  const collegeUrls = saveDocFiles_(folder, base, 'collegeid',
    normalizeFiles_(payload.collegeFiles, payload.collegePhotoBase64));
  const aadharPhotoUrl = aadharUrls.join(' , ');
  const collegeIdPhotoUrl = collegeUrls.join(' , ');

  const fields = {
    Timestamp: new Date(),
    StudentId: studentId,
    Name: name,
    AadharNumber: aadhar,
    Phone: cleanDigits_(payload.mobile || payload.phone, 10),
    HouseId: flat,
    Status: 'Submitted',
    Email: String(payload.email || '').trim(),
    Sex: String(payload.sex || '').trim(),
    ParentName: String(payload.parentName || '').trim(),
    ParentMobile: cleanDigits_(payload.parentMobile, 10),
    CollegeStudentId: String(payload.collegeStudentId || '').trim(),
    CollegeName: String(payload.collegeName || '').trim(),
    AcademicYear: String(payload.academicYear || payload.year || '').trim(),
    VehicleNumber: String(payload.vehicleNumber || '').trim(),
    Address: String(payload.address || '').trim(),
    CourseName: String(payload.courseName || '').trim(),
    IdValidUpto: String(payload.idValidUpto || '').trim(),
    AadharPhotoUrl: aadharPhotoUrl,
    CollegeIdPhotoUrl: collegeIdPhotoUrl
  };

  const existing = payload.studentId ? findStudentRowById_(studentId) : null;
  if (existing) {
    const merged = existing.row.slice();
    header.forEach((h, i) => {
      if (fields.hasOwnProperty(h) && fields[h] !== '') merged[i] = fields[h];
    });
    sheet.getRange(existing.rowNumber, 1, 1, merged.length).setValues([merged]);
  } else {
    sheet.appendRow(buildStudentRow_(header, fields));
  }

  upsertFlatRow_(flat, {
    OwnerName: payload.ownerName,
    OwnerEmail: payload.ownerEmail,
    OwnerMobile: payload.ownerMobile,
    CaretakerName: payload.caretakerName,
    CaretakerNumber: payload.caretakerNumber,
    TenancyPeriod: payload.tenancyPeriod,
    AgreementDate: payload.agreementDate
  });

  return { ok: true, studentId: studentId, message: 'Submitted successfully.' };
}

// Admin pull: all students for a flat plus the flat owner/caretaker row.
function getFlatForReview(flat) {
  flat = String(flat || '').trim();
  if (!flat) return { students: [], flat: {} };

  const sh = studentsSheetAndHeader_();
  const header = sh.header;
  const values = sh.sheet.getDataRange().getValues();
  values.shift();

  const idx = {};
  header.forEach((h, i) => { idx[h] = i; });
  const ACTIVEISH = { Submitted: 1, Verified: 1, Active: 1, InForm: 1 };

  const students = values
    .filter(r => String(r[idx.HouseId]).trim() === flat && ACTIVEISH[String(r[idx.Status])])
    .slice(0, MAX_STUDENTS_PER_HOUSE)
    .map(r => {
      const o = {};
      header.forEach(h => { o[h] = cellSafe_(r[idx[h]]); });
      return o;
    });

  return {
    students: students,
    flat: getFlatRow_(flat),
    agreementUrl: getFlatAgreementUrl_(flat)
  };
}

// Coerce a sheet cell to a value google.script.run can serialize. A Date nested in a
// returned array/object silently breaks serialization (the whole response degrades to
// a Java toString), so dates become strings; everything else passes through.
function cellSafe_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Kolkata', 'yyyy-MM-dd HH:mm:ss');
  return (v === null || v === undefined) ? '' : v;
}

// Admin tablet capture: store the student photo + signature, mark Verified.
function saveStudentVerification(payload) {
  payload = payload || {};
  const studentId = String(payload.studentId || '').trim();
  if (!studentId) throw new Error('studentId is required.');

  const found = findStudentRowById_(studentId);
  if (!found) throw new Error('Student not found: ' + studentId);

  const header = found.header;
  const flat = String(found.row[header.indexOf('HouseId')] || '').trim();
  const folder = getOrCreateFlatFolder_(flat);
  const row = found.row.slice();
  const iAssigned = header.indexOf('AssignedId');
  const assigned = iAssigned >= 0 ? String(row[iAssigned] || '').trim() : '';
  // Prefer the admin-assigned ID for filenames; else fall back to name + StudentId.
  const base = assigned
    ? assigned.replace(/[\\\/:*?"<>|]+/g, ' ').replace(/\s+/g, ' ').trim()
    : studentFileBase_(row[header.indexOf('Name')], studentId);

  if (payload.studentPhotoBase64) {
    row[header.indexOf('PhotoUrl')] =
      saveImageInFolder_(folder, payload.studentPhotoBase64, base + '.jpg');
  }
  if (payload.signatureBase64) {
    row[header.indexOf('SignatureUrl')] =
      savePngInFolder_(folder, payload.signatureBase64, base + '_sign.png');
  }
  if (row[header.indexOf('PhotoUrl')] && row[header.indexOf('SignatureUrl')]) {
    row[header.indexOf('Status')] = 'Verified';
  }

  getSpreadsheet_().getSheetByName(SHEET_STUDENTS)
    .getRange(found.rowNumber, 1, 1, row.length).setValues([row]);
  return { ok: true, studentId: studentId };
}

// Admin verify (on-device OCR): return the bytes the app needs to read locally —
// the student's stored Aadhaar + College-ID images, the flat's agreement (image only;
// PDFs are flagged for manual check), plus the form values to compare against.
// No AI here — the OCR + matching happen on the device.
function getDocsForVerify_(studentId) {
  studentId = String(studentId || '').trim();
  if (!studentId) throw new Error('studentId is required.');

  const found = findStudentRowById_(studentId);
  if (!found) throw new Error('Student not found: ' + studentId);

  const header = found.header, row = found.row;
  const get = (h) => String(row[header.indexOf(h)] || '').trim();
  const firstUrl = (v) => v.split(/\s*,\s*/).filter(Boolean)[0];

  const flat = get('HouseId');
  const out = {
    flat: flat,
    formName: get('Name'),
    formAadhar: cleanDigits_(get('AadharNumber'), 12),
    formCollegeId: get('CollegeStudentId'),
    formCollegeName: get('CollegeName')
  };

  const ab = driveBlobFromUrl_(firstUrl(get('AadharPhotoUrl')));
  if (ab) { out.aadharB64 = Utilities.base64Encode(ab.getBytes()); out.aadharMime = ab.getContentType() || 'image/jpeg'; }

  const cb = driveBlobFromUrl_(firstUrl(get('CollegeIdPhotoUrl')));
  if (cb) { out.collegeB64 = Utilities.base64Encode(cb.getBytes()); out.collegeMime = cb.getContentType() || 'image/jpeg'; }

  const agUrl = getFlatAgreementUrl_(flat);
  if (agUrl) {
    out.agreementUrl = agUrl;
    const gb = driveBlobFromUrl_(agUrl);
    const mt = gb ? (gb.getContentType() || '') : '';
    out.agreementMime = mt;
    // Only ship image agreements for OCR; PDFs are flagged so the app shows "check manually".
    if (gb && mt.indexOf('image') === 0) out.agreementB64 = Utilities.base64Encode(gb.getBytes());
  }

  return out;
}

// Web admin compare view: returns the student's form values plus ALL their Aadhaar
// and College-ID images (base64, so they render in the browser without Drive auth),
// and the flat's agreement (base64 if an image; URL + mime for a PDF so the page can
// embed a Drive preview). Public (no underscore) so google.script.run can call it.
function getCompareDocs(studentId) {
  studentId = String(studentId || '').trim();
  if (!studentId) throw new Error('studentId is required.');

  const found = findStudentRowById_(studentId);
  if (!found) throw new Error('Student not found: ' + studentId);

  const header = found.header, row = found.row;
  const get = (h) => String(row[header.indexOf(h)] || '').trim();
  const urls = (v) => v.split(/\s*,\s*/).filter(Boolean);
  const toImg = (u) => {
    const b = driveBlobFromUrl_(u);
    return b ? { base64: Utilities.base64Encode(b.getBytes()), mime: b.getContentType() || 'image/jpeg', url: u } : null;
  };

  const flat = get('HouseId');
  const aadhar = urls(get('AadharPhotoUrl')).map(toImg).filter(Boolean);
  const college = urls(get('CollegeIdPhotoUrl')).map(toImg).filter(Boolean);
  const photo = get('PhotoUrl') ? toImg(get('PhotoUrl')) : null;

  let agreement = null;
  const agUrl = getFlatAgreementUrl_(flat);
  if (agUrl) {
    const b = driveBlobFromUrl_(agUrl);
    const mt = b ? (b.getContentType() || '') : '';
    agreement = {
      url: agUrl,
      mime: mt,
      base64: (b && mt.indexOf('image') === 0) ? Utilities.base64Encode(b.getBytes()) : ''
    };
  }

  return {
    studentId: studentId,
    flat: flat,
    form: {
      name: get('Name'), aadhar: get('AadharNumber'), phone: get('Phone'), email: get('Email'),
      sex: get('Sex'), parentName: get('ParentName'), parentMobile: get('ParentMobile'),
      collegeId: get('CollegeStudentId'), collegeName: get('CollegeName'),
      year: get('AcademicYear'), status: get('Status')
    },
    aadhar: aadhar,
    college: college,
    photo: photo,
    agreement: agreement
  };
}

// Public wrapper so the web admin page (google.script.run) can store an agreement.
function webUploadAgreement(flat, base64, fileName, mimeType) {
  return uploadAgreement_({ flat: flat, base64: base64, fileName: fileName, mimeType: mimeType });
}

// Returns the Estancia logo as a data URL for the page header. Looks for a file named
// 'estancialogo.jpg' in the image base folder first, then anywhere in Drive. '' if none.
function getLogo() {
  try {
    var f = null;
    var inFolder = getBaseFolder_().getFilesByName('estancialogo.jpg');
    if (inFolder.hasNext()) f = inFolder.next();
    if (!f) { var any = DriveApp.getFilesByName('estancialogo.jpg'); if (any.hasNext()) f = any.next(); }
    if (!f) return '';
    var b = f.getBlob();
    return 'data:' + (b.getContentType() || 'image/jpeg') + ';base64,' + Utilities.base64Encode(b.getBytes());
  } catch (e) { return ''; }
}

// Web admin: edit a student's detail fields in place and write them back to the sheet.
// `fields` is keyed by sheet column name; only the allow-listed columns are written.
function saveStudentDetails(studentId, fields) {
  studentId = String(studentId || '').trim();
  if (!studentId) throw new Error('studentId is required.');
  fields = fields || {};

  const found = findStudentRowById_(studentId);
  if (!found) throw new Error('Student not found: ' + studentId);
  const header = found.header, row = found.row.slice();

  const ALLOW = ['Name', 'AadharNumber', 'Phone', 'Email', 'Sex', 'ParentName', 'ParentMobile',
    'Address', 'CollegeStudentId', 'CollegeName', 'CourseName', 'AcademicYear', 'IdValidUpto'];
  const applied = {};
  ALLOW.forEach(function (col) {
    if (!fields.hasOwnProperty(col)) return;
    const i = header.indexOf(col);
    if (i < 0) return;
    let v = String(fields[col] == null ? '' : fields[col]).trim();
    if (col === 'AadharNumber') v = v.replace(/\D/g, '').slice(0, 12);
    else if (col === 'Phone' || col === 'ParentMobile') v = v.replace(/\D/g, '').slice(0, 10);
    row[i] = v;
    applied[col] = v;
  });

  getSpreadsheet_().getSheetByName(SHEET_STUDENTS)
    .getRange(found.rowNumber, 1, 1, row.length).setValues([row]);
  return { ok: true, studentId: studentId, applied: applied };
}

// Web admin: save the manually-editable flat fields (owner name, rent-agreement date,
// tenancy period). Only non-empty values overwrite, so blanks never wipe existing data.
function saveFlatDetails(flat, fields) {
  flat = String(flat || '').trim();
  if (!flat) throw new Error('Flat number is required.');
  fields = fields || {};
  upsertFlatRow_(flat, {
    OwnerName: fields.ownerName,
    OwnerMobile: fields.ownerMobile,
    OwnerEmail: fields.ownerEmail,
    CaretakerName: fields.caretakerName,
    CaretakerNumber: fields.caretakerNumber,
    AgreementDate: fields.agreementDate,
    TenancyPeriod: fields.tenancyPeriod
  });
  return { ok: true, flat: getFlatRow_(flat) };
}

// Read-only diagnostic: shows how the flat's rows actually look in the sheet, so we can
// see why getFlatForReview may return nothing (status filtered out, HouseId stored with
// spaces / different value, etc.). Reachable via GET ?api=flatdebug&flat=XXXX.
function debugFlat_(flat) {
  flat = String(flat || '').trim();
  const sh = studentsSheetAndHeader_();
  const header = sh.header;
  const iHouse = header.indexOf('HouseId');
  const iStatus = header.indexOf('Status');
  const iName = header.indexOf('Name');
  const values = sh.sheet.getDataRange().getValues();
  values.shift();

  const matches = [];
  const houseCounts = {};
  values.forEach(function (r) {
    const raw = String(r[iHouse]);
    houseCounts[raw] = (houseCounts[raw] || 0) + 1;
    if (raw.trim() === flat) {
      matches.push({ name: String(r[iName] || ''), status: String(r[iStatus] || ''), houseRaw: raw, houseLen: raw.length });
    }
  });

  let reviewCount = -1;
  try { reviewCount = (getFlatForReview(flat).students || []).length; } catch (e) { reviewCount = 'ERR:' + e.message; }

  return {
    flat: flat,
    totalDataRows: values.length,
    matchesForFlat: matches,
    reviewReturns: reviewCount,
    allowedStatuses: ['Submitted', 'Verified', 'Active', 'InForm'],
    distinctHouseIds: Object.keys(houseCounts)
  };
}

// Diagnostic: opens the Annexure template and reports which {{tokens}} it contains vs
// what generateAnnexurePdf actually fills. unknownTokens = present in the doc but NOT
// filled by the code (they'd print literally) — fix those. missingTokens = the code can
// fill them but the doc has none (just prints blank; fine for signatures/unused fields).
function checkAnnexureTemplate_() {
  const id = PropertiesService.getScriptProperties().getProperty('ANNEXURE_TEMPLATE_ID') || DEFAULT_ANNEXURE_TEMPLATE_ID;
  if (!id) return { error: 'No template id configured.' };

  let text;
  try { text = DocumentApp.openById(id).getBody().getText(); }
  catch (e) { return { error: 'Cannot open template doc: ' + (e && e.message ? e.message : e), id: id }; }

  const found = Array.from(new Set((text.match(/\{\{[^}]+\}\}/g) || []).map(s => s.trim())));

  const expected = new Set([
    '{{apartment}}', '{{flat_number}}', '{{owner_name}}', '{{owner_mobile}}', '{{owner_email}}',
    '{{caretaker_name}}', '{{caretaker_number}}', '{{signing_date}}', '{{signing_place}}', '{{owner_sig}}',
    '{{tenancy_period}}', '{{tenancy}}', '{{agreement_date}}',
    '{{today}}', '{{date}}', '{{generated_date}}'
  ]);
  for (let n = 1; n <= MAX_STUDENTS_PER_HOUSE; n++) {
    ['name', 'mobile', 'email', 'address', 'parent_name', 'parent_mobile', 'college_id', 'college', 'acad_year',
     'course', 'sig', 'photo', 'aadhar', 'aadhar2', 'collegeid', 'collegeid2',
     'phone', 'parent_phone', 'id', 'assigned_id', 'efoa_id', 'year', 'validto']  // aliases + image tokens included
      .forEach(k => expected.add('{{s' + n + '_' + k + '}}'));
  }

  return {
    id: id,
    tokensInDoc: found,
    unknownTokens: found.filter(t => !expected.has(t)),
    missingTokens: Array.from(expected).filter(t => found.indexOf(t) === -1)
  };
}

// Web admin: the flat's agreement file bytes (base64) so the left panel can render
// it (PDF → page thumbnails via pdf.js, or a single image). null if none uploaded.
function getFlatAgreementDoc(flat) {
  flat = String(flat || '').trim();
  if (!flat) return null;
  const url = getFlatAgreementUrl_(flat);
  if (!url) return null;
  const b = driveBlobFromUrl_(url);
  return {
    url: url,
    mime: b ? (b.getContentType() || '') : '',
    base64: b ? Utilities.base64Encode(b.getBytes()) : ''
  };
}

// Admin: store the flat's agreement (PDF or image) in the flat's Drive folder.
// No AI extraction — name matching is done on-device via OCR at verify time.
function uploadAgreement_(payload) {
  payload = payload || {};
  const flat = String(payload.flat || '').trim();
  if (!flat) throw new Error('Flat number is required.');
  if (!payload.base64) throw new Error('Agreement file is required.');

  const folder = getOrCreateFlatFolder_(flat);
  const name = String(payload.fileName || 'agreement');
  const ext = name.indexOf('.') >= 0 ? name.split('.').pop().toLowerCase() : mimeToExt_(payload.mimeType);
  const finalName = 'agreement.' + ext;

  removeFlatAgreement_(folder);
  const blob = Utilities.newBlob(
    Utilities.base64Decode(payload.base64),
    payload.mimeType || 'application/octet-stream',
    finalName
  );
  const fileUrl = saveFileInFolder_(folder, finalName, blob);

  return { ok: true, fileUrl: fileUrl, isImage: String(payload.mimeType || '').indexOf('image') === 0 };
}

// Admin tablet capture: owner signature + signing date/place on the flat row.
function saveFlatSigning(payload) {
  payload = payload || {};
  const flat = String(payload.flat || '').trim();
  if (!flat) throw new Error('Flat number is required.');

  const folder = getOrCreateFlatFolder_(flat);
  const fields = {
    SigningDate: String(payload.signingDate || '').trim(),
    SigningPlace: String(payload.signingPlace || '').trim()
  };
  if (payload.ownerSignatureBase64) {
    fields.OwnerSignatureUrl = savePngInFolder_(folder, payload.ownerSignatureBase64, 'owner_sign.png');
  }
  upsertFlatRow_(flat, fields);
  return { ok: true };
}

// Counts students for a flat across the new lifecycle statuses (Submitted/Verified/
// Active/InForm) — used for capacity, unlike countActiveStudentsByHouse_ (Active only).
function countFlatStudents_(flat) {
  const sh = studentsSheetAndHeader_();
  const header = sh.header;
  const iHouse = header.indexOf('HouseId');
  const iStatus = header.indexOf('Status');
  const ACTIVEISH = { Submitted: 1, Verified: 1, Active: 1, InForm: 1 };
  const values = sh.sheet.getDataRange().getValues();
  values.shift();
  return values.filter(r => String(r[iHouse]) === String(flat) && ACTIVEISH[String(r[iStatus])]).length;
}

// --- Flats sheet helpers ---

function getFlatRow_(flat) {
  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName(SHEET_FLATS);
  if (!sh) return {};
  ensureFlatColumns_();
  const data = sh.getDataRange().getValues();
  const header = data.shift();
  const iFlat = header.indexOf('Flat#');
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][iFlat]).trim() === String(flat).trim()) {
      const o = {};
      header.forEach((h, j) => { o[h] = cellSafe_(data[i][j]); });
      return o;
    }
  }
  return {};
}

// Creates the flat row if missing; only overwrites cells for which a non-empty
// value is supplied (so a later student / a partial save never blanks prior data).
function upsertFlatRow_(flat, fields) {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(SHEET_FLATS);
  if (!sh) {
    sh = ss.insertSheet(SHEET_FLATS);
    sh.appendRow(FLATS_HEADER.slice());
  }
  ensureFlatColumns_();

  const data = sh.getDataRange().getValues();
  const header = data[0];
  const iFlat = header.indexOf('Flat#');

  let rowNumber = -1;
  let row = null;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][iFlat]) === String(flat)) {
      rowNumber = i + 1;
      row = data[i].slice();
      break;
    }
  }
  if (!row) {
    row = header.map(h => (h === 'Flat#' ? String(flat) : ''));
    sh.appendRow(row);
    rowNumber = sh.getLastRow();
  }

  let changed = false;
  header.forEach((h, i) => {
    if (fields.hasOwnProperty(h)) {
      const val = String(fields[h] == null ? '' : fields[h]).trim();
      if (val !== '') { row[i] = val; changed = true; }
    }
  });
  if (changed) sh.getRange(rowNumber, 1, 1, row.length).setValues([row]);
  return { rowNumber: rowNumber, header: header, row: row };
}

function savePngInFolder_(folder, base64Data, fileName) {
  const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'image/png', fileName);
  return saveFileInFolder_(folder, fileName, blob);
}

// Saves a list of {base64, mimeType, name?} document files (images or PDF) for a
// student and returns their share URLs. Names: <studentId>_<prefix>_<n>.<ext>.
function saveDocFiles_(folder, studentId, prefix, files) {
  if (!files || !files.length) return [];
  return files.map(function (f, i) {
    const ext = mimeToExt_(f.mimeType);
    const name = studentId + '_' + prefix + '_' + (i + 1) + '.' + ext;
    const blob = Utilities.newBlob(
      Utilities.base64Decode(f.base64),
      f.mimeType || 'application/octet-stream',
      name
    );
    return saveFileInFolder_(folder, name, blob);
  });
}

// Normalizes the new files[] array, falling back to a legacy single base64 image.
function normalizeFiles_(arr, legacyBase64) {
  if (Array.isArray(arr) && arr.length) return arr;
  if (legacyBase64) return [{ base64: legacyBase64, mimeType: 'image/jpeg' }];
  return [];
}

function mimeToExt_(m) {
  m = String(m || '').toLowerCase();
  if (m.indexOf('pdf') >= 0) return 'pdf';
  if (m.indexOf('png') >= 0) return 'png';
  return 'jpg';
}

// ===================================================================
//  Combined PDF generation via a Google Doc template.
//  Requires script property ANNEXURE_TEMPLATE_ID (id of the template Doc).
//  Template uses {{tokens}} for text and {{..._photo}}/{{..._sig}} markers
//  (each alone in its own table cell) for images. See ANNEXURE_TEMPLATE.md.
// ===================================================================
function generateAnnexurePdf(flat) {
  flat = String(flat || '').trim();
  if (!flat) throw new Error('Flat number is required.');

  const templateId = PropertiesService.getScriptProperties().getProperty('ANNEXURE_TEMPLATE_ID')
    || DEFAULT_ANNEXURE_TEMPLATE_ID;
  if (!templateId) throw new Error('ANNEXURE_TEMPLATE_ID not set in Script Properties.');

  const data = getFlatForReview(flat);
  const students = data.students || [];
  const flatInfo = data.flat || {};
  const folder = getOrCreateFlatFolder_(flat);

  const copy = DriveApp.getFileById(templateId).makeCopy('annexure_tmp_' + flat, folder);
  const doc = DocumentApp.openById(copy.getId());
  const body = doc.getBody();

  // --- text tokens ---
  const todayStr = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'dd-MMM-yyyy');
  const text = {
    '{{apartment}}': flat,
    '{{flat_number}}': flat,   // alias for the apartment / flat number
    '{{today}}': todayStr,           // PDF generation date (today)
    '{{date}}': todayStr,
    '{{generated_date}}': todayStr,
    '{{owner_name}}': flatInfo.OwnerName || '',
    '{{owner_mobile}}': flatInfo.OwnerMobile || '',
    '{{owner_email}}': flatInfo.OwnerEmail || '',
    '{{caretaker_name}}': flatInfo.CaretakerName || '',
    '{{caretaker_number}}': flatInfo.CaretakerNumber || '',
    '{{signing_date}}': flatInfo.SigningDate || '',
    '{{signing_place}}': flatInfo.SigningPlace || '',
    '{{tenancy_period}}': flatInfo.TenancyPeriod || '',
    '{{tenancy}}': flatInfo.TenancyPeriod || '',
    '{{agreement_date}}': flatInfo.AgreementDate || ''
  };
  for (let n = 1; n <= MAX_STUDENTS_PER_HOUSE; n++) {
    const s = students[n - 1] || {};
    const p = '{{s' + n + '_';
    text[p + 'name}}'] = s.Name || '';
    text[p + 'mobile}}'] = s.Phone || '';
    text[p + 'email}}'] = s.Email || '';
    text[p + 'address}}'] = s.Address || '';
    text[p + 'parent_name}}'] = s.ParentName || '';
    text[p + 'parent_mobile}}'] = s.ParentMobile || '';
    text[p + 'college_id}}'] = s.CollegeStudentId || '';
    text[p + 'college}}'] = s.CollegeName || '';
    text[p + 'acad_year}}'] = s.AcademicYear || '';
    text[p + 'course}}'] = s.CourseName || '';
    // --- aliases matching the template's token names ---
    text[p + 'phone}}'] = s.Phone || '';
    text[p + 'parent_phone}}'] = s.ParentMobile || '';
    text[p + 'id}}'] = s.CollegeStudentId || '';
    text[p + 'year}}'] = s.AcademicYear || '';
    text[p + 'validto}}'] = s.IdValidUpto || '';   // ID "Valid Upto"
    text[p + 'assigned_id}}'] = s.AssignedId || '';   // EFOA-assigned Student ID (admin)
    text[p + 'efoa_id}}'] = s.AssignedId || '';       // same value, preferred token name
  }
  Object.keys(text).forEach(k => body.replaceText(escapeRegex_(k), text[k]));

  // --- image tokens (photos, signatures, and the document scans) ---
  for (let n = 1; n <= MAX_STUDENTS_PER_HOUSE; n++) {
    const s = students[n - 1] || {};
    const tag = '{{s' + n + '_';
    replaceTokenWithImage_(body, tag + 'photo}}', s.PhotoUrl, 90, 110, true);   // fit width to cell, keep height
    replaceTokenWithImage_(body, tag + 'sig}}', s.SignatureUrl, 120, 45);
    // Aadhaar + College/Company ID scans (front = ..1, back = ..2). Each may hold
    // several files joined by " , "; we insert the first two. Missing ones clear blank.
    const aad = String(s.AadharPhotoUrl || '').split(/\s*,\s*/).filter(Boolean);
    replaceTokenWithImage_(body, tag + 'aadhar}}', aad[0], 260, 170);
    replaceTokenWithImage_(body, tag + 'aadhar2}}', aad[1], 260, 170);
    const col = String(s.CollegeIdPhotoUrl || '').split(/\s*,\s*/).filter(Boolean);
    replaceTokenWithImage_(body, tag + 'collegeid}}', col[0], 260, 170);
    replaceTokenWithImage_(body, tag + 'collegeid2}}', col[1], 260, 170);
  }
  replaceTokenWithImage_(body, '{{owner_sig}}', flatInfo.OwnerSignatureUrl, 150, 50);

  // Final sweep: clear any {{token}} that had no value or wasn't mapped, so no raw
  // placeholders ever appear in the PDF. (Runs after all text/image replacements.)
  body.replaceText('\\{\\{[^{}]*\\}\\}', '');

  doc.saveAndClose();

  const pdfBlob = DriveApp.getFileById(copy.getId())
    .getAs('application/pdf')
    .setName('annexure_2A.pdf');
  const url = saveFileInFolder_(folder, 'annexure_2A.pdf', pdfBlob);
  copy.setTrashed(true);

  return { ok: true, fileUrl: url };
}

// Replaces EVERY occurrence of a {{token}} (each assumed alone in its paragraph/cell)
// with a Drive image. Leaves the cell empty if the image is missing so the layout
// still holds. (Handles the same token appearing in both the page-3 signature grid
// and the page-5 detail block.)
function replaceTokenWithImage_(body, token, fileUrlOrId, width, height, widthOnly) {
  const pattern = escapeRegex_(token);
  const blob = driveBlobFromUrl_(fileUrlOrId);

  // Re-search from the start each pass: clearing the token text means the next
  // findText() lands on the following occurrence.
  let found = body.findText(pattern);
  let guard = 0;
  while (found && guard++ < 50) {
    const el = found.getElement();
    el.asText().setText(''); // clear the placeholder text

    if (blob) {
      const parent = el.getParent();
      let img = null;
      if (parent.getType() === DocumentApp.ElementType.PARAGRAPH) {
        img = parent.asParagraph().appendInlineImage(blob);
      } else if (parent.getType() === DocumentApp.ElementType.LIST_ITEM) {
        img = parent.asListItem().appendInlineImage(blob);
      }
      if (img) sizeImageToCell_(img, el, width, height, widthOnly);
    }
    found = body.findText(pattern);
  }
}

// Returns the TableCell an element sits in, or null if it isn't inside a table.
function cellOf_(el) {
  let p = el.getParent();
  while (p) {
    if (p.getType() === DocumentApp.ElementType.TABLE_CELL) return p.asTableCell();
    p = p.getParent();
  }
  return null;
}

// Sizes an inserted image to fit its table cell's width (keeping the image's aspect
// ratio). Falls back to the given fixed width/height when it's not inside a cell or the
// cell width can't be read.
//   widthOnly === true → set ONLY the width to the cell width; leave the height as the
//   image was inserted (height not touched).
function sizeImageToCell_(img, el, fallbackW, fallbackH, widthOnly) {
  let w = fallbackW, h = fallbackH;
  const cell = cellOf_(el);
  if (cell) {
    let cw = 0;
    try { cw = cell.getWidth(); } catch (e) { cw = 0; }
    const iw = img.getWidth(), ih = img.getHeight();
    if (cw) {
      w = Math.max(24, cw - 8);                            // leave a little padding inside the cell
      if (iw && ih) h = Math.round(w * (ih / iw));         // preserve aspect ratio
    }
  }
  if (widthOnly) {
    if (w) img.setWidth(w);                                // fit to cell width; don't touch height
    return;
  }
  if (w && h) { img.setWidth(w); img.setHeight(h); }
}

function driveBlobFromUrl_(urlOrId) {
  if (!urlOrId) return null;
  let id = String(urlOrId);
  const m = id.match(/[-\w]{25,}/); // extract the Drive file id from a share URL
  if (m) id = m[0];
  try { return DriveApp.getFileById(id).getBlob(); } catch (e) { return null; }
}

function driveFileFromUrl_(urlOrId) {
  if (!urlOrId) return null;
  let id = String(urlOrId);
  const m = id.match(/[-\w]{25,}/);
  if (m) id = m[0];
  try { return DriveApp.getFileById(id); } catch (e) { return null; }
}

function fileExt_(name, fallback) {
  const i = String(name || '').lastIndexOf('.');
  return i >= 0 ? String(name).slice(i + 1).toLowerCase() : fallback;
}

// Renames a " , "-joined list of Drive files to <base>_<n>.<ext>, returning the
// (unchanged) URLs. Renaming keeps the file id/URL, so the sheet links stay valid.
function renameDocList_(joined, base) {
  const urls = String(joined || '').split(/\s*,\s*/).filter(Boolean);
  return urls.map(function (u, i) {
    const f = driveFileFromUrl_(u);
    if (!f) return u;
    f.setName(base + '_' + (i + 1) + '.' + fileExt_(f.getName(), 'jpg'));
    return f.getUrl();
  }).join(' , ');
}

// Admin: assign a human Student ID to a student and rename ALL their stored files to it
// (<id>.jpg photo, <id>_sign.png, <id>_aadhar_n.*, <id>_collegeid_n.*). Stores the id in
// the AssignedId column. Public so the web admin page can call it via google.script.run.
function assignStudentId(studentId, assignedId) {
  studentId = String(studentId || '').trim();
  const safe = String(assignedId || '').trim().replace(/[\\\/:*?"<>|]+/g, ' ').replace(/\s+/g, ' ').trim();
  if (!studentId) throw new Error('studentId is required.');
  if (!safe) throw new Error('Please enter a Student ID to assign.');

  studentsSheetAndHeader_();                 // make sure the AssignedId column exists
  const found = findStudentRowById_(studentId);
  if (!found) throw new Error('Student not found: ' + studentId);

  const header = found.header, row = found.row.slice();

  const iPhoto = header.indexOf('PhotoUrl');
  if (iPhoto >= 0 && row[iPhoto]) {
    const f = driveFileFromUrl_(row[iPhoto]);
    if (f) { f.setName(safe + '.' + fileExt_(f.getName(), 'jpg')); row[iPhoto] = f.getUrl(); }
  }
  const iSig = header.indexOf('SignatureUrl');
  if (iSig >= 0 && row[iSig]) {
    const f = driveFileFromUrl_(row[iSig]);
    if (f) { f.setName(safe + '_sign.' + fileExt_(f.getName(), 'png')); row[iSig] = f.getUrl(); }
  }
  const iAad = header.indexOf('AadharPhotoUrl');
  if (iAad >= 0) row[iAad] = renameDocList_(row[iAad], safe + '_aadhar');
  const iCol = header.indexOf('CollegeIdPhotoUrl');
  if (iCol >= 0) row[iCol] = renameDocList_(row[iCol], safe + '_collegeid');

  const iAssigned = header.indexOf('AssignedId');
  if (iAssigned >= 0) row[iAssigned] = safe;

  getSpreadsheet_().getSheetByName(SHEET_STUDENTS)
    .getRange(found.rowNumber, 1, 1, row.length).setValues([row]);

  return { ok: true, assignedId: safe };
}

function escapeRegex_(s) {
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function findStudentRowById_(studentId) {
  const ss = getSpreadsheet_();
  const students = ss.getSheetByName(SHEET_STUDENTS);
  if (!students) {
    return null;
  }

  const values = students.getDataRange().getValues();
  const header = values[0];
  const iStudentId = header.indexOf('StudentId');

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][iStudentId]) === String(studentId)) {
      return { header, row: values[i], rowNumber: i + 1 };
    }
  }
  return null;
}

// Reads the agreement from the selected page images only (small payload).
// Persists the extracted JSON; the full original is stored separately (chunked).
function extractFlatAgreement(payload) {
  const flat = String((payload || {}).flatNumber || '').trim();
  if (!flat) throw new Error('Flat number is required.');

  const images = (payload.pageImages || []).map(b => ({ base64: b, mimeType: 'image/jpeg' }));
  if (!images.length) throw new Error('No agreement pages to read.');

  const folder = getOrCreateFlatFolder_(flat);

  let agreementInfo = {};
  try {
    agreementInfo = extractAgreementFields_(images);
  } catch (e) {
    agreementInfo = { error: e && e.message ? e.message : String(e) };
  }
  saveJsonInFolder_(folder, 'agreement_extract.json', agreementInfo);

  return { ok: true, agreementInfo: agreementInfo };
}

// Chunked upload of the full original agreement file, so large scanned PDFs
// are never sent in a single oversized google.script.run payload.
function agreementUploadInit(flatNumber) {
  const folder = getOrCreateFlatFolder_(String(flatNumber || '').trim());
  const existing = folder.getFilesByName('agreement_upload.tmp');
  while (existing.hasNext()) existing.next().setTrashed(true);
  folder.createFile(Utilities.newBlob('', 'text/plain', 'agreement_upload.tmp'));
  return { ok: true };
}

function agreementUploadChunk(flatNumber, chunkBase64) {
  const folder = getFlatFolder_(String(flatNumber || '').trim());
  if (!folder) throw new Error('Upload not initialized.');
  const files = folder.getFilesByName('agreement_upload.tmp');
  if (!files.hasNext()) throw new Error('Upload not initialized.');
  const tmp = files.next();
  tmp.setContent(tmp.getBlob().getDataAsString() + String(chunkBase64 || ''));
  return { ok: true };
}

function agreementUploadFinish(flatNumber, fileName, mimeType) {
  const flat = String(flatNumber || '').trim();
  const folder = getFlatFolder_(flat);
  if (!folder) throw new Error('Upload not initialized.');
  const files = folder.getFilesByName('agreement_upload.tmp');
  if (!files.hasNext()) throw new Error('Upload not initialized.');

  const tmp = files.next();
  const base64 = tmp.getBlob().getDataAsString();

  const name = String(fileName || '');
  const ext = name.indexOf('.') >= 0 ? name.split('.').pop().toLowerCase() : 'pdf';
  const finalName = 'agreement.' + ext;

  removeFlatAgreement_(folder);
  const blob = Utilities.newBlob(
    Utilities.base64Decode(base64),
    mimeType || 'application/octet-stream',
    finalName
  );
  const fileUrl = saveFileInFolder_(folder, finalName, blob);
  tmp.setTrashed(true);

  return { ok: true, fileUrl: fileUrl };
}

function analyzeAadhar(payload) {
  if (!payload || !payload.aadharPhotoBase64) {
    throw new Error('Aadhar photo is required for AI extraction.');
  }

  const prompt =
    'Read this Indian Aadhaar card image and return JSON only with keys: ' +
    'name, aadharNumber. ' +
    'Rules: aadharNumber must be exactly 12 digits (no spaces). ' +
    'If uncertain, return empty string for that field.';

  const extracted = geminiExtractJson_(prompt, payload.aadharPhotoBase64, 'image/jpeg');

  return {
    ok: true,
    name: String(extracted.name || '').trim(),
    aadharNumber: String(extracted.aadharNumber || '').replace(/\D/g, '').slice(0, 12)
  };
}

function analyzeCollegeId(payload) {
  if (!payload || !payload.collegePhotoBase64) {
    throw new Error('College ID photo is required for AI extraction.');
  }

  const prompt =
    'Read this college / student ID card image and return JSON only with keys: ' +
    'name, idNumber. ' +
    'name is the student full name printed on the card. ' +
    'idNumber is the college/student identification number. ' +
    'If uncertain, return empty string for that field.';

  const extracted = geminiExtractJson_(prompt, payload.collegePhotoBase64, 'image/jpeg');

  return {
    ok: true,
    name: String(extracted.name || '').trim(),
    idNumber: String(extracted.idNumber || '').trim()
  };
}

function extractAgreementFields_(images) {
  const prompt =
    'This is a rental / lease / tenancy agreement made between an owner and the student tenants ' +
    '(there can be up to 6 students in one agreement). ' +
    'Return JSON only with a single key "students" whose value is an array (up to 6) of objects, ' +
    'one per student / tenant named in the agreement, each with keys: ' +
    'name, address, aadharNumber, phone, fatherName. ' +
    'fatherName is the father / guardian name if present. ' +
    'aadharNumber must be 12 digits (no spaces) if present, else empty. ' +
    'phone is a 10-digit number if present, else empty. ' +
    'For any field that is not present, use an empty string.';

  const extracted = geminiCall_(prompt, images);

  const arr = Array.isArray(extracted.students)
    ? extracted.students
    : (Array.isArray(extracted) ? extracted : []);

  const students = arr.slice(0, MAX_STUDENTS_PER_HOUSE).map(s => ({
    name: String((s && s.name) || '').trim(),
    address: String((s && s.address) || '').trim(),
    aadharNumber: String((s && s.aadharNumber) || '').replace(/\D/g, '').slice(0, 12),
    phone: String((s && s.phone) || '').replace(/\D/g, '').slice(0, 10),
    fatherName: String((s && s.fatherName) || '').trim()
  }));

  return { students };
}

function geminiExtractJson_(promptText, base64, mimeType) {
  return geminiCall_(promptText, [{ base64: base64, mimeType: mimeType }]);
}

function geminiCall_(promptText, images) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    throw new Error('GEMINI_API_KEY not configured in Script Properties.');
  }

  if (!images || !images.length) {
    throw new Error('No document image provided for extraction.');
  }

  const parts = [{ text: promptText }];
  images.forEach(img => {
    parts.push({
      inline_data: {
        mime_type: img.mimeType || 'image/jpeg',
        data: img.base64
      }
    });
  });

  const requestBody = {
    contents: [{ parts: parts }],
    generationConfig: {
      temperature: 0.1,
      responseMimeType: 'application/json'
    }
  };

  const url =
    'https://generativelanguage.googleapis.com/v1beta/models/' +
    GEMINI_MODEL +
    ':generateContent?key=' +
    encodeURIComponent(apiKey);

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(requestBody),
    muteHttpExceptions: true
  };

  // Retry transient overload/rate-limit/server errors with exponential backoff.
  const RETRY_STATUSES = [429, 500, 503];
  const MAX_ATTEMPTS = 5;
  let status, body;
  for (let attempt = 1; attempt <= MAX_ATTEMPTS; attempt++) {
    const response = UrlFetchApp.fetch(url, options);
    status = response.getResponseCode();
    body = response.getContentText();

    if (status < 300) break;
    if (RETRY_STATUSES.indexOf(status) === -1 || attempt === MAX_ATTEMPTS) {
      throw new Error('Gemini API error: ' + status + ' ' + body);
    }
    Utilities.sleep(Math.min(8000, 1000 * Math.pow(2, attempt - 1))); // 1s,2s,4s,8s
  }

  const parsed = JSON.parse(body);
  const text = (((parsed || {}).candidates || [])[0] || {}).content;
  const firstPart = (((text || {}).parts || [])[0] || {}).text || '{}';

  try {
    return JSON.parse(firstPart);
  } catch (e) {
    return {};
  }
}

// ---- Diagnostics: run these directly from the Apps Script editor ----

// Confirms the API key + model work. Check View -> Logs after running.
function pingGemini() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  Logger.log('GEMINI_API_KEY present: ' + (apiKey ? 'yes' : 'NO'));
  Logger.log('Model: ' + GEMINI_MODEL);
  const url =
    'https://generativelanguage.googleapis.com/v1beta/models/' +
    GEMINI_MODEL + ':generateContent?key=' + encodeURIComponent(apiKey || '');
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      contents: [{ parts: [{ text: 'Reply with JSON {"ok":true}' }] }],
      generationConfig: { responseMimeType: 'application/json' }
    }),
    muteHttpExceptions: true
  });
  Logger.log('HTTP ' + res.getResponseCode());
  Logger.log(res.getContentText().slice(0, 800));
}

// Lists the models your API key can use with generateContent. Check View -> Logs.
function listGeminiModels() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  const res = UrlFetchApp.fetch(
    'https://generativelanguage.googleapis.com/v1beta/models?key=' + encodeURIComponent(apiKey || ''),
    { muteHttpExceptions: true }
  );
  Logger.log('HTTP ' + res.getResponseCode());
  const data = JSON.parse(res.getContentText() || '{}');
  (data.models || []).forEach(m => {
    const methods = m.supportedGenerationMethods || [];
    if (methods.indexOf('generateContent') >= 0) {
      Logger.log(m.name + '  (' + methods.join(', ') + ')');
    }
  });
}

// Tests agreement extraction on a file already in Drive (upload a sample, pass its id).
// Sends the WHOLE file to Gemini; logs the extracted students and timing.
function testAgreementFromDrive(fileId) {
  const file = DriveApp.getFileById(fileId);
  const blob = file.getBlob();
  const bytes = blob.getBytes();
  Logger.log('File: ' + file.getName() + ' (' + blob.getContentType() + ', ' + Math.round(bytes.length / 1024) + ' KB)');
  const t0 = Date.now();
  const info = extractAgreementFields_([{ base64: Utilities.base64Encode(bytes), mimeType: blob.getContentType() }]);
  Logger.log('Extraction took ' + ((Date.now() - t0) / 1000) + 's');
  Logger.log(JSON.stringify(info, null, 2));
  return info;
}

function exportForIdCard() {
  const ss = getSpreadsheet_();
  const students = ss.getSheetByName(SHEET_STUDENTS);
  const exp = ss.getSheetByName(SHEET_EXPORT);

  if (!students || !exp) {
    throw new Error('Required sheets missing. Run setupSheets() first.');
  }

  const data = students.getDataRange().getValues();
  const header = data.shift();

  const iStudentId = header.indexOf('StudentId');
  const iName = header.indexOf('Name');
  const iPhone = header.indexOf('Phone');
  const iAadhar = header.indexOf('AadharNumber');
  const iPhoto = header.indexOf('PhotoUrl');
  const iHouseId = header.indexOf('HouseId');
  const iStatus = header.indexOf('Status');

  exp.clearContents();
  exp.appendRow(['StudentId', 'Name', 'Phone', 'AadharNumber', 'PhotoUrl', 'HouseId']);

  data
    .filter(r => r[iStatus] === 'Active')
    .forEach(r => {
      exp.appendRow([
        r[iStudentId],
        r[iName],
        r[iPhone],
        r[iAadhar],
        r[iPhoto],
        r[iHouseId]
      ]);
    });

  const csvContent = sheetToCsv_(exp);
  const file = DriveApp.createFile('id-card-export.csv', csvContent, MimeType.CSV);

  return {
    ok: true,
    fileUrl: file.getUrl(),
    message: 'CSV export generated for ID card app.'
  };
}

function countActiveStudentsByHouse_(houseId) {
  const ss = getSpreadsheet_();
  const students = ss.getSheetByName(SHEET_STUDENTS);
  if (!students) {
    return 0;
  }

  const values = students.getDataRange().getValues();
  const header = values.shift();
  const iHouseId = header.indexOf('HouseId');
  const iStatus = header.indexOf('Status');

  return values.filter(r => r[iHouseId] === houseId && r[iStatus] === 'Active').length;
}

function generateStudentId_() {
  // Millisecond precision + a random suffix so two submissions in the same second
  // (or even millisecond) never collide — duplicate StudentIds previously caused
  // saves to target the wrong row and files to overwrite each other.
  const ts = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyyMMddHHmmssSSS');
  const rand = Math.floor(Math.random() * 9000 + 1000); // 4 digits
  return 'STU' + ts + rand;
}

function getBaseFolder_() {
  const folderId =
    PropertiesService.getScriptProperties().getProperty('IMAGE_FOLDER_ID') ||
    DEFAULT_IMAGE_FOLDER_ID;
  return folderId ? DriveApp.getFolderById(folderId) : DriveApp.getRootFolder();
}

function getFlatFolder_(flatNumber) {
  const base = getBaseFolder_();
  const it = base.getFoldersByName(String(flatNumber));
  return it.hasNext() ? it.next() : null;
}

function getOrCreateFlatFolder_(flatNumber) {
  return getFlatFolder_(flatNumber) || getBaseFolder_().createFolder(String(flatNumber));
}

function saveFileInFolder_(folder, fileName, blob) {
  // Replace an existing file of the same name so re-saves don't duplicate.
  const existing = folder.getFilesByName(fileName);
  while (existing.hasNext()) {
    existing.next().setTrashed(true);
  }
  const file = folder.createFile(blob.setName(fileName));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

function saveImageInFolder_(folder, base64Data, fileName) {
  const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'image/jpeg', fileName);
  return saveFileInFolder_(folder, fileName, blob);
}

function saveDetailTxt_(folder, studentId, payload) {
  const lines = [
    'StudentId: ' + studentId,
    'Name: ' + (payload.name || ''),
    'AadharNumber: ' + (payload.aadharNumber || ''),
    'Phone: ' + (payload.phone || ''),
    'Flat: ' + (payload.houseId || ''),
    'AgreementStatus: ' + (payload.agreementStatus || ''),
    'AgreementRef: ' + (payload.agreementRef || ''),
    'Notes: ' + (payload.notes || ''),
    'UpdatedAt: ' + Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd HH:mm:ss')
  ];
  const blob = Utilities.newBlob(lines.join('\n'), 'text/plain', studentId + '.txt');
  return saveFileInFolder_(folder, studentId + '.txt', blob);
}

function saveJsonInFolder_(folder, fileName, obj) {
  const blob = Utilities.newBlob(JSON.stringify(obj), 'application/json', fileName);
  // Stored for reload; keep it private (no public sharing toggle).
  const existing = folder.getFilesByName(fileName);
  while (existing.hasNext()) {
    existing.next().setTrashed(true);
  }
  folder.createFile(blob);
}

function getFlatAgreementUrl_(flatNumber) {
  const folder = getFlatFolder_(flatNumber);
  if (!folder) return '';
  const files = folder.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    if (f.getName().indexOf('agreement.') === 0) return f.getUrl();
  }
  return '';
}

function getFlatAgreementInfo_(flatNumber) {
  const folder = getFlatFolder_(flatNumber);
  if (!folder) return {};
  const files = folder.getFilesByName('agreement_extract.json');
  if (!files.hasNext()) return {};
  try {
    return JSON.parse(files.next().getBlob().getDataAsString());
  } catch (e) {
    return {};
  }
}

function removeFlatAgreement_(folder) {
  const files = folder.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    if (f.getName().indexOf('agreement.') === 0) f.setTrashed(true);
  }
}

function createWhatsAppLink_(phone, message) {
  const digits = String(phone || '').replace(/\D/g, '');
  const text = encodeURIComponent(message);
  return 'https://wa.me/91' + digits + '?text=' + text;
}

function validatePayload_(payload, isUpdate) {
  if (!payload) throw new Error('Missing form data.');
  if (!payload.name) throw new Error('Name is required.');
  if (!payload.aadharNumber) throw new Error('Aadhar number is required.');
  if (!payload.phone) throw new Error('Phone is required.');
  if (!payload.houseId) throw new Error('Flat number is required.');
  if (!isUpdate) {
    if (!payload.studentPhotoBase64) throw new Error('Student photo is required.');
    if (!payload.aadharPhotoBase64) throw new Error('Aadhar photo is required.');
  }
}

function sheetToCsv_(sheet) {
  const data = sheet.getDataRange().getDisplayValues();
  return data
    .map(row =>
      row
        .map(cell => {
          const value = String(cell).replace(/"/g, '""');
          return '"' + value + '"';
        })
        .join(',')
    )
    .join('\n');
}
  