/**
 * YiW Field Report — Google Apps Script
 * Actions: submitWithLinks (primary), legacy fallback
 * On each submission:
 *   1. Appends a row to the master Google Sheet (creates it if needed)
 *   2. Appends a row to the hub-specific sheet tab
 *   3. Sends a formatted HTML email to all recipients with a link to the sheet
 *
 * Deploy as: Web App | Execute as: Me | Access: Anyone
 */

// ── RECIPIENTS ───────────────────────────────────────────────
var TO_EMAIL = "yiw@seghana.net";
var CC_EMAILS = [
  "execdir@seghana.net",
  "merl@seghana.net",
  "finance@seghana.net",
  "salimatu@seghana.net",
  "leticia.antwi@seghana.net",
  "isaac.quansah@seghana.net"
].join(",");

// ── DRIVE / SHEETS CONFIG ────────────────────────────────────
var ROOT_FOLDER_NAME  = "Youth in Work Field Reports Files";
var MASTER_SHEET_NAME = "YiW Daily Field Reports - Master Log";


// ══════════════════════════════════════════════════════════════
//  ENTRY POINTS
// ══════════════════════════════════════════════════════════════

function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var action  = payload.action || 'legacy';
    if (action === 'submitWithLinks') return handleSubmitWithLinks(payload);
    return handleLegacy(payload);
  } catch (err) {
    Logger.log('FATAL: ' + err.toString());
    return jsonOut({ status: 'error', message: err.toString() });
  }
}

function doGet(e) {
  return jsonOut({ status: 'ok', message: 'YiW Script is live.' });
}


// ══════════════════════════════════════════════════════════════
//  SANITISATION — single source of truth, called by every handler
// ══════════════════════════════════════════════════════════════

function sanitise(raw) {
  var d = raw || {};
  var out = {};

  // Plain string fields — default to ''
  var stringFields = [
    'fpName','fpPhone','fpEmail','fpZone',
    'visitDate','visitType','hubName','community','trainingCentre','centreAddress',
    'hubContact','hubContactPhone','tArr','tDep',
    'enrolCourse','empName','empSector','youthNames','highlight','yVoice',
    'challenges','recommendations','urgency','followUpBy',
    'partnerNotes','nextPDate',
    'docNotes','photoCaption','videoCaption','mediaContext',
    'safeConcern','safeTxt','safeAct','safeRep','safeNotes','finalNotes'
  ];
  for (var i = 0; i < stringFields.length; i++) {
    var key = stringFields[i];
    var val = d[key];
    out[key] = (val === null || val === undefined) ? '' : String(val);
  }

  // Numeric fields — default to 0
  var numFields = ['cMale','cFemale','cPWD','cStaff','cTrainer',
                    'aJobs','aIntern','aCoop','aRef','enrolM','enrolF','rating'];
  for (var j = 0; j < numFields.length; j++) {
    var nkey = numFields[j];
    var nval = parseInt(d[nkey], 10);
    out[nkey] = isNaN(nval) ? 0 : nval;
  }

  // Array-of-strings fields — default to []
  var arrayFields = ['quality','issues','facilities','activities','safeChecked'];
  for (var k = 0; k < arrayFields.length; k++) {
    var akey = arrayFields[k];
    out[akey] = Array.isArray(d[akey]) ? d[akey] : [];
  }

  // Partners — array of objects with known shape
  if (Array.isArray(d.partners)) {
    out.partners = [];
    for (var p = 0; p < d.partners.length; p++) {
      var src = d.partners[p];
      if (!src || typeof src !== 'object') src = {};
      out.partners.push({
        name:         src.name         ? String(src.name)         : '',
        location:     src.location     ? String(src.location)     : '',
        sector:       src.sector       ? String(src.sector)       : '',
        profile:      src.profile      ? String(src.profile)      : '',
        skillsNeeded: src.skillsNeeded ? String(src.skillsNeeded) : '',
        contact:      src.contact      ? String(src.contact)      : '',
        phone:        src.phone        ? String(src.phone)        : '',
        status:       src.status       ? String(src.status)       : '',
        slots:        (function(){ var s = parseInt(src.slots, 10); return isNaN(s) ? 0 : s; })()
      });
    }
  } else {
    out.partners = [];
  }

  return out;
}

function sanitiseDriveLinks(raw) {
  var src = raw || {};
  var cats = ['dAtt','dFin','dMou','dTrack','mPhoto','mVideo'];
  var out = {};
  for (var i = 0; i < cats.length; i++) {
    var cat = cats[i];
    var arr = src[cat];
    if (!Array.isArray(arr)) { out[cat] = []; continue; }
    out[cat] = [];
    for (var j = 0; j < arr.length; j++) {
      var f = arr[j];
      if (!f || typeof f !== 'object') continue;
      out[cat].push({
        name:     f.name     ? String(f.name)     : 'file',
        url:      f.url      ? String(f.url)      : '',
        category: f.category ? String(f.category) : ''
      });
    }
  }
  return out;
}


// ══════════════════════════════════════════════════════════════
//  PRIMARY HANDLER
// ══════════════════════════════════════════════════════════════

function handleSubmitWithLinks(payload) {
  var d          = sanitise(payload.formData);
  var driveLinks = sanitiseDriveLinks(payload.driveLinks);
  var totalFiles = parseInt(payload.totalFiles, 10) || 0;

  // 1. Append to master sheet (creates workbook if needed), get share URL
  var sheetUrl = appendToMasterSheet(d, driveLinks, totalFiles);

  // 2. Append to the hub-specific tab in the same workbook
  appendToHubSheet(d, driveLinks);

  // 3. Build email
  var fileLinksHtml = buildFileLinksHtml(driveLinks);
  var htmlBody = buildEmailHtml(d, fileLinksHtml, sheetUrl);
  var subject  = 'YiW Field Report: ' + (d.fpName || '--') +
                  ' -- ' + (d.trainingCentre || d.hubName || '--') +
                  ' (' + (d.visitDate || '--') + ')';

  MailApp.sendEmail({
    to:       TO_EMAIL,
    cc:       CC_EMAILS,
    subject:  subject,
    htmlBody: htmlBody
  });

  Logger.log('Done: ' + subject);
  return jsonOut({ status: 'success', message: 'Report submitted. Sheet updated.' });
}


// ══════════════════════════════════════════════════════════════
//  LEGACY FALLBACK (no files, simplest path)
// ══════════════════════════════════════════════════════════════

function handleLegacy(payload) {
  var d = sanitise(payload.formData);
  var emptyLinks = sanitiseDriveLinks({});

  var sheetUrl = appendToMasterSheet(d, emptyLinks, 0);
  appendToHubSheet(d, emptyLinks);

  var fileLinksHtml = '<p style="color:#718096;font-style:italic;font-size:13px">No files attached for this submission.</p>';
  var htmlBody = buildEmailHtml(d, fileLinksHtml, sheetUrl);
  var subject  = 'YiW Field Report: ' + (d.fpName || '--') + ' (' + (d.visitDate || '--') + ')';

  MailApp.sendEmail({ to: TO_EMAIL, cc: CC_EMAILS, subject: subject, htmlBody: htmlBody });
  return jsonOut({ status: 'success', message: 'Report submitted.' });
}


// ══════════════════════════════════════════════════════════════
//  MASTER SHEET — one row per submission (Google Forms style)
// ══════════════════════════════════════════════════════════════

function appendToMasterSheet(d, driveLinks, totalFiles) {
  try {
    var ss, sheetUrl;
    var existing = DriveApp.getFilesByName(MASTER_SHEET_NAME);

    if (existing.hasNext()) {
      var f = existing.next();
      ss = SpreadsheetApp.open(f);
      sheetUrl = f.getUrl();
    } else {
      ss = SpreadsheetApp.create(MASTER_SHEET_NAME);
      var newFile = DriveApp.getFileById(ss.getId());
      newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      sheetUrl = newFile.getUrl();
    }

    var sheet = ss.getSheetByName('Field Reports');
    if (!sheet) {
      sheet = ss.getActiveSheet();
      sheet.setName('Field Reports');
    }
    if (sheet.getLastRow() === 0) {
      formatMasterSheet(sheet);
    }

    var fc = countDriveLinks(driveLinks);

    var allFileUrls = [];
    var cats = ['dAtt','dFin','dMou','dTrack','mPhoto','mVideo'];
    for (var c = 0; c < cats.length; c++) {
      var arr = driveLinks[cats[c]];
      for (var i = 0; i < arr.length; i++) {
        allFileUrls.push(arr[i].name + ': ' + arr[i].url);
      }
    }

    var partnerNames = [];
    var partnerSkills = [];
    for (var p = 0; p < d.partners.length; p++) {
      var pr = d.partners[p];
      partnerNames.push(pr.name + (pr.status ? ' (' + pr.status + ')' : ''));
      partnerSkills.push(pr.skillsNeeded);
    }

    var dataRow = [
      new Date(),
      d.fpName, d.fpPhone, d.fpEmail, d.fpZone,
      d.visitDate, d.visitType, d.hubName, d.community, d.trainingCentre,
      d.centreAddress, d.hubContact, d.hubContactPhone, d.tArr, d.tDep,
      // Attendance
      d.cMale, d.cFemale, d.cPWD, d.cStaff, d.cTrainer,
      (d.cMale + d.cFemale + d.cPWD),
      // Activation
      d.aJobs, d.aIntern, d.aCoop, d.aRef,
      (d.aJobs + d.aIntern + d.aCoop + d.aRef),
      d.enrolM, d.enrolF, d.enrolCourse, d.empName, d.empSector,
      // Quality
      d.rating,
      d.quality.join('; '),
      d.issues.join('; '),
      d.facilities.join('; '),
      d.activities.join('; '),
      d.challenges, d.recommendations, d.urgency, d.followUpBy,
      // Partners
      d.partners.length,
      partnerNames.join('; '),
      partnerSkills.join('; '),
      // Files
      fc.total, fc.dAtt, fc.dFin, fc.dMou, fc.dTrack, fc.mPhoto, fc.mVideo,
      allFileUrls.join(' | '),
      // Safeguarding
      d.safeChecked.length,
      d.safeChecked.join('; '),
      (d.safeConcern === 'yes' ? 'YES' : 'No'),
      d.safeTxt,
      // Narrative
      d.highlight, d.yVoice, d.finalNotes
    ];

    sheet.appendRow(dataRow);

    var lastRow = sheet.getLastRow();
    if (lastRow % 2 === 0) {
      sheet.getRange(lastRow, 1, 1, dataRow.length).setBackground('#f0f4f0');
    }

    SpreadsheetApp.flush();
    Logger.log('Master sheet row ' + lastRow + ' written. URL: ' + sheetUrl);
    return sheetUrl;

  } catch (err) {
    Logger.log('Master sheet error: ' + err.toString());
    return '';
  }
}

function formatMasterSheet(sheet) {
  var headers = [
    'Submitted At','FP Name','FP Phone','FP Email','Zone',
    'Visit Date','Visit Type','Hub / TSP','Community','Training Centre',
    'Centre Address','Centre Contact','Contact Phone','Time Arrived','Time Departed',
    'Young Men','Young Women','PWD','Staff','Trainers','Total Youth',
    'Formal Jobs','Internships','Cooperatives','Further Training','Total Activations',
    'Enrolments (M)','Enrolments (F)','Course / Trade','Employer','Sector',
    'Hub Rating','Quality Indicators','Issues Flagged','Facilities','Activities',
    'Challenges','Recommendations','Urgency','Follow-up By',
    'Partners Count','Partner Names & Status','Skills Requested by Partners',
    'Total Files','Attendance Docs','Financial Docs','MoUs','Tracking Sheets','Photos','Videos','File Links',
    'Safeguarding Items Confirmed','Safeguarding Details','Concern Raised','Concern Detail',
    'Success Story','Youth Voice','Final Notes'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#1a5c2a');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setFontSize(10);
  headerRange.setWrap(false);

  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length).createFilter();

  sheet.setColumnWidth(1,  160);
  sheet.setColumnWidth(2,  140);
  sheet.setColumnWidth(8,  220);
  sheet.setColumnWidth(9,  120);
  sheet.setColumnWidth(10, 170);
  sheet.setColumnWidth(33, 220);
  sheet.setColumnWidth(34, 200);
  sheet.setColumnWidth(36, 260);
  sheet.setColumnWidth(37, 260);
  sheet.setColumnWidth(42, 200);
  sheet.setColumnWidth(50, 300);
}


// ══════════════════════════════════════════════════════════════
//  HUB-SPECIFIC SHEET — one tab per TSP/Hub
// ══════════════════════════════════════════════════════════════

function appendToHubSheet(d, driveLinks) {
  try {
    var existing = DriveApp.getFilesByName(MASTER_SHEET_NAME);
    if (!existing.hasNext()) return; // master sheet not created yet, skip
    var ss = SpreadsheetApp.open(existing.next());

    var hubName = d.hubName ? d.hubName.toString().trim() : 'Unknown Hub';
    if (!hubName) hubName = 'Unknown Hub';
    var sheetName = hubName.length > 95 ? hubName.substring(0, 95) : hubName;

    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      formatHubSheet(sheet, hubName);
    }
    if (sheet.getLastRow() === 0) {
      formatHubSheet(sheet, hubName);
    }

    var fc = countDriveLinks(driveLinks);

    var allFileUrls = [];
    var cats = ['dAtt','dFin','dMou','dTrack','mPhoto','mVideo'];
    for (var c = 0; c < cats.length; c++) {
      var arr = driveLinks[cats[c]];
      for (var i = 0; i < arr.length; i++) {
        allFileUrls.push(arr[i].name + ': ' + arr[i].url);
      }
    }

    var partnerNames = [];
    var partnerSkills = [];
    for (var p = 0; p < d.partners.length; p++) {
      var pr = d.partners[p];
      partnerNames.push(pr.name + (pr.status ? ' (' + pr.status + ')' : ''));
      partnerSkills.push(pr.skillsNeeded);
    }

    var row = [
      new Date(),
      d.fpName, d.fpPhone, d.fpEmail, d.fpZone,
      d.visitDate, d.visitType, d.community, d.trainingCentre,
      d.hubContact, d.hubContactPhone, d.tArr, d.tDep,
      // Attendance
      d.cMale, d.cFemale, d.cPWD, d.cStaff, d.cTrainer,
      (d.cMale + d.cFemale + d.cPWD),
      // Activation
      d.aJobs, d.aIntern, d.aCoop, d.aRef,
      (d.aJobs + d.aIntern + d.aCoop + d.aRef),
      d.enrolM, d.enrolF, d.enrolCourse, d.empName, d.empSector,
      // Quality
      d.rating,
      d.quality.join('; '),
      d.issues.join('; '),
      d.facilities.join('; '),
      d.activities.join('; '),
      d.challenges, d.recommendations, d.urgency, d.followUpBy,
      // Partners
      d.partners.length,
      partnerNames.join('; '),
      partnerSkills.join('; '),
      // Files
      fc.total,
      allFileUrls.join(' | '),
      // Safeguarding
      d.safeChecked.length,
      (d.safeConcern === 'yes' ? 'YES' : 'No'),
      d.safeTxt,
      // Narrative
      d.highlight, d.yVoice, d.finalNotes
    ];

    sheet.appendRow(row);

    var lastRow = sheet.getLastRow();
    if (lastRow % 2 === 0) {
      sheet.getRange(lastRow, 1, 1, row.length).setBackground('#f0f4f0');
    }

    SpreadsheetApp.flush();
    Logger.log('Hub sheet "' + sheetName + '" row ' + lastRow + ' written.');

  } catch (err) {
    Logger.log('Hub sheet error: ' + err.toString());
  }
}

function formatHubSheet(sheet, hubName) {
  var headers = [
    'Submitted At','FP Name','FP Phone','FP Email','Zone',
    'Visit Date','Visit Type','Community','Training Centre',
    'Centre Contact','Contact Phone','Time Arrived','Time Departed',
    'Young Men','Young Women','PWD','Staff','Trainers','Total Youth',
    'Formal Jobs','Internships','Cooperatives','Further Training','Total Activations',
    'Enrolments (M)','Enrolments (F)','Course / Trade','Employer','Sector',
    'Hub Rating','Quality Indicators','Issues Flagged','Facilities','Activities',
    'Challenges','Recommendations','Urgency','Follow-up By',
    'Partners Count','Partner Names','Skills Requested',
    'Total Files','File Links',
    'Safeguarding Confirmed','Concern Raised','Concern Detail',
    'Success Story','Youth Voice','Final Notes'
  ];

  sheet.getRange(1, 1, 1, headers.length).merge()
       .setValue(hubName + ' - YiW Field Reports')
       .setBackground('#1a5c2a')
       .setFontColor('#ffffff')
       .setFontWeight('bold')
       .setFontSize(12);

  var headerRange = sheet.getRange(2, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#2d7a3a');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setFontSize(10);
  headerRange.setWrap(false);

  sheet.setFrozenRows(2);
  sheet.getRange(2, 1, 1, headers.length).createFilter();

  sheet.setColumnWidth(1,  160);
  sheet.setColumnWidth(2,  140);
  sheet.setColumnWidth(6,  100);
  sheet.setColumnWidth(8,  120);
  sheet.setColumnWidth(9,  170);
  sheet.setColumnWidth(31, 220);
  sheet.setColumnWidth(32, 200);
  sheet.setColumnWidth(35, 260);
  sheet.setColumnWidth(36, 260);
  sheet.setColumnWidth(42, 300);
}


// ══════════════════════════════════════════════════════════════
//  HELPERS
// ══════════════════════════════════════════════════════════════

function countDriveLinks(driveLinks) {
  var cats = ['dAtt','dFin','dMou','dTrack','mPhoto','mVideo'];
  var out = { total: 0, dAtt: 0, dFin: 0, dMou: 0, dTrack: 0, mPhoto: 0, mVideo: 0 };
  for (var i = 0; i < cats.length; i++) {
    var c = cats[i];
    out[c] = driveLinks[c].length;
    out.total += out[c];
  }
  return out;
}

function buildFileLinksHtml(driveLinks) {
  var catOrder = ['dAtt','dFin','dMou','dTrack','mPhoto','mVideo'];
  var catNames = {
    dAtt: 'Attendance Sheet', dFin: 'Financial Document',
    dMou: 'MoU / Agreement',  dTrack: 'Tracking Sheet',
    mPhoto: 'Photo',          mVideo: 'Video'
  };
  var rows = '';
  var count = 0;

  for (var i = 0; i < catOrder.length; i++) {
    var cat = catOrder[i];
    var files = driveLinks[cat];
    for (var j = 0; j < files.length; j++) {
      var f = files[j];
      var label = f.category || catNames[cat] || cat;
      rows += '<tr>' +
        '<td style="padding:8px;border:1px solid #cbd5e1;font-weight:600;color:#1a5c2a;font-size:12px">' + label + '</td>' +
        '<td style="padding:8px;border:1px solid #cbd5e1;font-size:12px"><a href="' + f.url + '" style="color:#1565c0;font-weight:500">Link: ' + f.name + '</a></td>' +
        '</tr>';
      count++;
    }
  }

  if (count === 0) {
    return '<p style="color:#718096;font-style:italic;font-size:13px">No files attached for this submission.</p>';
  }

  return '<table style="width:100%;border-collapse:collapse;font-size:13px;margin-top:6px">' +
    '<tr style="background:#e8f5eb"><th style="padding:8px;border:1px solid #cbd5e1;text-align:left">Category</th>' +
    '<th style="padding:8px;border:1px solid #cbd5e1;text-align:left">File - click to open in Drive</th></tr>' +
    rows + '</table>';
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function safe(val) {
  return (val === '' || val === null || val === undefined) ? '--' : val;
}

function starString(rating) {
  // Avoid String.prototype.repeat for older runtime compatibility
  var full = '';
  var empty = '';
  for (var i = 0; i < rating; i++) full += '*';
  for (var j = 0; j < (5 - rating); j++) empty += '-';
  if (rating <= 0) return 'Not rated';
  return full + empty + ' (' + rating + '/5)';
}


// ══════════════════════════════════════════════════════════════
//  EMAIL BUILDER
// ══════════════════════════════════════════════════════════════

function buildEmailHtml(d, fileLinksHtml, sheetUrl) {

  function pills(arr, bg, color) {
    if (!arr || arr.length === 0) return '<em style="color:#718096">None</em>';
    var out = '';
    for (var i = 0; i < arr.length; i++) {
      out += '<span style="display:inline-block;background:' + bg + ';color:' + color +
             ';padding:3px 9px;border-radius:12px;margin:2px;font-size:12px;font-weight:600">' + arr[i] + '</span>';
    }
    return out;
  }

  function statBox(val, label, bg, color) {
    return '<td style="text-align:center;background:' + bg + ';padding:10px 6px;border-radius:6px;width:20%">' +
           '<div style="font-size:20px;font-weight:800;color:' + color + '">' + (val || 0) + '</div>' +
           '<div style="font-size:11px;color:#718096;margin-top:2px">' + label + '</div></td>';
  }

  var ratingStr = starString(d.rating);
  var ts = new Date().toLocaleString('en-GB', { timeZone: 'Africa/Accra' });

  // Partner rows
  var partnerRows = '';
  if (d.partners.length > 0) {
    for (var i = 0; i < d.partners.length; i++) {
      var p = d.partners[i];
      partnerRows +=
        '<tr>' +
        '<td style="padding:7px 8px;border-bottom:1px solid #dde3ea;font-weight:600;font-size:12px">' + safe(p.name) + '</td>' +
        '<td style="padding:7px 8px;border-bottom:1px solid #dde3ea;font-size:12px">' + safe(p.location) + '</td>' +
        '<td style="padding:7px 8px;border-bottom:1px solid #dde3ea;font-size:12px">' + safe(p.sector) + '</td>' +
        '<td style="padding:7px 8px;border-bottom:1px solid #dde3ea;font-size:11px;color:#555">' + safe(p.profile) + '</td>' +
        '<td style="padding:7px 8px;border-bottom:1px solid #dde3ea;font-size:11px;color:#555">' + safe(p.skillsNeeded) + '</td>' +
        '<td style="padding:7px 8px;border-bottom:1px solid #dde3ea;font-size:12px">' + safe(p.contact) + '</td>' +
        '<td style="padding:7px 8px;border-bottom:1px solid #dde3ea;font-size:12px">' + safe(p.phone) + '</td>' +
        '<td style="padding:7px 8px;border-bottom:1px solid #dde3ea">' +
        '<span style="background:#fff3e0;color:#e65100;padding:2px 7px;border-radius:10px;font-size:11px;font-weight:600">' + safe(p.status) + '</span>' +
        '</td>' +
        '<td style="padding:7px 8px;border-bottom:1px solid #dde3ea;text-align:center;font-weight:700;color:#1a5c2a">' + (p.slots || 0) + '</td>' +
        '</tr>';
    }
  } else {
    partnerRows = '<tr><td colspan="9" style="padding:12px;text-align:center;color:#718096;font-style:italic">No partner companies logged.</td></tr>';
  }

  // Safeguarding
  var safeHtml;
  if (d.safeChecked.length > 0) {
    safeHtml = '';
    for (var s = 0; s < d.safeChecked.length; s++) {
      safeHtml += '<li style="margin-bottom:3px">' + d.safeChecked[s] + '</li>';
    }
  } else {
    safeHtml = '<li style="color:#718096">No safeguarding items confirmed.</li>';
  }

  var concernHtml;
  if (d.safeConcern === 'yes') {
    concernHtml =
      '<div style="background:#ffebee;border-left:4px solid #c62828;padding:12px;border-radius:6px;margin-top:10px;color:#c62828">' +
      '<strong>SAFEGUARDING CONCERN RAISED</strong><br/>' +
      '<strong>Details:</strong> ' + safe(d.safeTxt) + '<br/>' +
      '<strong>Action:</strong> ' + safe(d.safeAct) + '<br/>' +
      '<strong>Reported to:</strong> ' + safe(d.safeRep) + '</div>';
  } else {
    concernHtml = '<p style="color:#1a5c2a;font-size:13px;margin-top:6px">No concerns identified.</p>';
  }

  var sheetBanner = '';
  if (sheetUrl) {
    sheetBanner =
      '<div style="background:#1a5c2a;border-radius:9px;padding:14px 16px;margin-bottom:14px;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:10px">' +
      '<div>' +
      '<div style="color:#fff;font-weight:700;font-size:13px">View &amp; Download Master Data Sheet</div>' +
      '<div style="color:rgba(255,255,255,.75);font-size:11px;margin-top:2px">All submissions in one Google Sheet - download as Excel anytime</div>' +
      '</div>' +
      '<a href="' + sheetUrl + '" style="display:inline-block;background:#fff;color:#1a5c2a;font-weight:700;font-size:12px;padding:8px 16px;border-radius:6px;text-decoration:none">Open Sheet</a>' +
      '</div>';
  }

  var html = '';
  html += '<!DOCTYPE html><html><head><meta charset="UTF-8"/></head>';
  html += '<body style="font-family:\'Segoe UI\',Arial,sans-serif;background:#f3f7f4;margin:0;padding:20px">';
  html += '<div style="max-width:700px;margin:0 auto;background:#fff;border-radius:12px;border:1px solid #dde3ea;overflow:hidden;box-shadow:0 4px 16px rgba(0,0,0,.07)">';

  // Header
  html += '<div style="background:linear-gradient(135deg,#1a5c2a,#2d7a3a);padding:22px 26px;color:#fff">';
  html += '<div style="font-size:10px;opacity:.7;text-transform:uppercase;letter-spacing:.1em;margin-bottom:3px">SEG Ghana</div>';
  html += '<div style="font-size:20px;font-weight:700;margin-bottom:2px">Youth in Work Programme</div>';
  html += '<div style="font-size:13px;opacity:.85;margin-bottom:14px">Daily Field Report</div>';
  html += '<div style="display:flex;flex-wrap:wrap;gap:16px">';
  html += '<div><div style="font-size:10px;opacity:.65;text-transform:uppercase">Submitted by</div><div style="font-size:14px;font-weight:700;margin-top:1px">' + safe(d.fpName) + '</div></div>';
  html += '<div><div style="font-size:10px;opacity:.65;text-transform:uppercase">Phone</div><div style="font-size:14px;font-weight:700;margin-top:1px">' + safe(d.fpPhone) + '</div></div>';
  html += '<div><div style="font-size:10px;opacity:.65;text-transform:uppercase">Date</div><div style="font-size:14px;font-weight:700;margin-top:1px">' + safe(d.visitDate) + '</div></div>';
  html += '<div><div style="font-size:10px;opacity:.65;text-transform:uppercase">Zone</div><div style="font-size:14px;font-weight:700;margin-top:1px">' + safe(d.fpZone) + '</div></div>';
  html += '</div></div>';

  html += '<div style="padding:20px">';

  html += sheetBanner;

  html += '<div style="background:#e8f5eb;border:1px solid #a5d6a7;border-radius:8px;padding:10px 13px;font-size:12px;color:#1a3a1a;margin-bottom:14px">' +
          '<strong>Master data sheet updated</strong> - this report has been added to the central Google Sheet. ' +
          'Open the sheet above and go to <strong>File &gt; Download &gt; Microsoft Excel (.xlsx)</strong> to export all submissions.' +
          '</div>';

  // Visit details
  html += '<div style="border-radius:9px;border:1px solid #dde3ea;border-left:4px solid #1a5c2a;padding:15px;margin-bottom:13px">';
  html += '<div style="font-size:11px;font-weight:700;color:#1a5c2a;text-transform:uppercase;letter-spacing:.06em;margin-bottom:9px">Visit Details</div>';
  html += '<table style="width:100%;border-collapse:collapse;font-size:13px">';
  html += '<tr><td style="padding:4px 0;color:#718096;width:38%">Visit type</td><td style="padding:4px 0;font-weight:600">' + safe(d.visitType) + '</td></tr>';
  html += '<tr><td style="padding:4px 0;color:#718096">Hub / TSP</td><td style="padding:4px 0;font-weight:600">' + safe(d.hubName) + '</td></tr>';
  html += '<tr><td style="padding:4px 0;color:#718096">Community</td><td style="padding:4px 0;font-weight:600">' + safe(d.community) + '</td></tr>';
  html += '<tr><td style="padding:4px 0;color:#718096">Training centre</td><td style="padding:4px 0;font-weight:600">' + safe(d.trainingCentre) + '</td></tr>';
  html += '<tr><td style="padding:4px 0;color:#718096">Address</td><td style="padding:4px 0;font-weight:600">' + safe(d.centreAddress) + '</td></tr>';
  html += '<tr><td style="padding:4px 0;color:#718096">Centre contact</td><td style="padding:4px 0;font-weight:600">' + safe(d.hubContact) + (d.hubContactPhone ? ' - ' + d.hubContactPhone : '') + '</td></tr>';
  html += '<tr><td style="padding:4px 0;color:#718096">Time on site</td><td style="padding:4px 0;font-weight:600">' + safe(d.tArr) + ' to ' + safe(d.tDep) + '</td></tr>';
  html += '</table></div>';

  // Attendance
  html += '<div style="border-radius:9px;border:1px solid #dde3ea;border-left:4px solid #b8860b;padding:15px;margin-bottom:13px">';
  html += '<div style="font-size:11px;font-weight:700;color:#b8860b;text-transform:uppercase;letter-spacing:.06em;margin-bottom:9px">Attendance Count</div>';
  html += '<table style="width:100%;border-collapse:separate;border-spacing:4px"><tr>';
  html += statBox(d.cMale, 'Young men', '#fff8e1', '#b8860b');
  html += statBox(d.cFemale, 'Young women', '#fff8e1', '#b8860b');
  html += statBox(d.cPWD, 'PWD', '#fff8e1', '#b8860b');
  html += statBox(d.cStaff, 'Staff', '#fff8e1', '#b8860b');
  html += statBox(d.cTrainer, 'Trainers', '#fff8e1', '#b8860b');
  html += '</tr></table></div>';

  // Activation
  html += '<div style="border-radius:9px;border:1px solid #dde3ea;border-left:4px solid #1a5c2a;padding:15px;margin-bottom:13px">';
  html += '<div style="font-size:11px;font-weight:700;color:#1a5c2a;text-transform:uppercase;letter-spacing:.06em;margin-bottom:9px">Activation &amp; Employment</div>';
  html += '<table style="width:100%;border-collapse:separate;border-spacing:4px"><tr>';
  html += statBox(d.aJobs, 'Formal jobs', '#e8f5eb', '#1a5c2a');
  html += statBox(d.aIntern, 'Internships', '#e8f5eb', '#1a5c2a');
  html += statBox(d.aCoop, 'Cooperatives', '#e8f5eb', '#1a5c2a');
  html += statBox(d.aRef, 'Further trng', '#e8f5eb', '#1a5c2a');
  html += '<td></td></tr></table>';
  html += '<table style="width:100%;border-collapse:collapse;font-size:13px;margin-top:10px">';
  html += '<tr><td style="padding:4px 0;color:#718096;width:38%">New enrolments (M/F)</td><td style="padding:4px 0;font-weight:600">' + d.enrolM + ' / ' + d.enrolF + '</td></tr>';
  html += '<tr><td style="padding:4px 0;color:#718096">Course enrolled in</td><td style="padding:4px 0;font-weight:600">' + safe(d.enrolCourse) + '</td></tr>';
  html += '<tr><td style="padding:4px 0;color:#718096">Employer / cooperative</td><td style="padding:4px 0;font-weight:600">' + safe(d.empName) + (d.empSector ? ' (' + d.empSector + ')' : '') + '</td></tr>';
  html += '</table>';
  if (d.youthNames) html += '<div style="margin-top:8px;padding:9px;background:#f8fafc;border-radius:7px;font-size:12px;white-space:pre-wrap"><strong style="color:#1a5c2a">Youth placed:</strong><br/>' + d.youthNames + '</div>';
  if (d.highlight)  html += '<div style="margin-top:7px;padding:9px;background:#e8f5eb;border-radius:7px;border-left:3px solid #4caf50;font-size:13px;color:#1a5c2a"><strong>Success story:</strong><br/>' + d.highlight + '</div>';
  if (d.yVoice)     html += '<div style="margin-top:6px;padding:9px;background:#f8fafc;border-radius:7px;border-left:3px solid #90caf9;font-size:13px;color:#4a5568;font-style:italic">"' + d.yVoice + '"</div>';
  html += '</div>';

  // Hub quality
  html += '<div style="border-radius:9px;border:1px solid #dde3ea;border-left:4px solid #00695c;padding:15px;margin-bottom:13px">';
  html += '<div style="font-size:11px;font-weight:700;color:#00695c;text-transform:uppercase;letter-spacing:.06em;margin-bottom:9px">Training Centre Quality</div>';
  html += '<table style="width:100%;border-collapse:collapse;font-size:13px;margin-bottom:9px">';
  html += '<tr><td style="padding:4px 0;color:#718096;width:38%">Overall rating</td><td style="padding:4px 0;font-weight:700;color:#1a5c2a">' + ratingStr + '</td></tr>';
  html += '<tr><td style="padding:4px 0;color:#718096">Urgency</td><td style="padding:4px 0;font-weight:600">' + safe(d.urgency) + '</td></tr>';
  html += '<tr><td style="padding:4px 0;color:#718096">Follow-up by</td><td style="padding:4px 0;font-weight:600">' + safe(d.followUpBy) + '</td></tr>';
  html += '</table>';
  html += '<div style="margin-bottom:7px"><strong style="font-size:11px;color:#00695c;text-transform:uppercase">Quality</strong><br/><div style="margin-top:4px">' + pills(d.quality, '#e8f5eb', '#1a5c2a') + '</div></div>';
  html += '<div style="margin-bottom:7px"><strong style="font-size:11px;color:#c62828;text-transform:uppercase">Issues</strong><br/><div style="margin-top:4px">' + pills(d.issues, '#ffebee', '#c62828') + '</div></div>';
  html += '<div style="margin-bottom:7px"><strong style="font-size:11px;color:#1565c0;text-transform:uppercase">Activities</strong><br/><div style="margin-top:4px">' + pills(d.activities, '#e3f2fd', '#1565c0') + '</div></div>';
  html += '<div style="margin-bottom:7px"><strong style="font-size:11px;color:#455a64;text-transform:uppercase">Facilities</strong><br/><div style="margin-top:4px">' + pills(d.facilities, '#eceff1', '#455a64') + '</div></div>';
  if (d.challenges)      html += '<div style="margin-top:9px;padding:9px;background:#fff3e0;border-radius:7px;font-size:12px;color:#e65100;border-left:3px solid #ff9800"><strong>Challenges:</strong> ' + d.challenges + '</div>';
  if (d.recommendations) html += '<div style="margin-top:6px;padding:9px;background:#e3f2fd;border-radius:7px;font-size:12px;color:#1565c0;border-left:3px solid #90caf9"><strong>Recommendations:</strong> ' + d.recommendations + '</div>';
  html += '</div>';

  // Partners
  html += '<div style="border-radius:9px;border:1px solid #dde3ea;border-left:4px solid #1565c0;padding:15px;margin-bottom:13px">';
  html += '<div style="font-size:11px;font-weight:700;color:#1565c0;text-transform:uppercase;letter-spacing:.06em;margin-bottom:9px">Partner Engagement - ' + d.partners.length + ' company(ies)</div>';
  html += '<div style="overflow-x:auto"><table style="width:100%;border-collapse:collapse;font-size:12px">';
  html += '<thead><tr style="background:#e3f2fd">';
  html += '<th style="padding:7px 8px;border-bottom:2px solid #90caf9;text-align:left;color:#1565c0;font-size:10px;text-transform:uppercase">Company</th>';
  html += '<th style="padding:7px 8px;border-bottom:2px solid #90caf9;text-align:left;color:#1565c0;font-size:10px;text-transform:uppercase">Location</th>';
  html += '<th style="padding:7px 8px;border-bottom:2px solid #90caf9;text-align:left;color:#1565c0;font-size:10px;text-transform:uppercase">Sector</th>';
  html += '<th style="padding:7px 8px;border-bottom:2px solid #90caf9;text-align:left;color:#1565c0;font-size:10px;text-transform:uppercase">Business profile</th>';
  html += '<th style="padding:7px 8px;border-bottom:2px solid #90caf9;text-align:left;color:#1565c0;font-size:10px;text-transform:uppercase">Skills needed</th>';
  html += '<th style="padding:7px 8px;border-bottom:2px solid #90caf9;text-align:left;color:#1565c0;font-size:10px;text-transform:uppercase">Contact</th>';
  html += '<th style="padding:7px 8px;border-bottom:2px solid #90caf9;text-align:left;color:#1565c0;font-size:10px;text-transform:uppercase">Phone</th>';
  html += '<th style="padding:7px 8px;border-bottom:2px solid #90caf9;text-align:left;color:#1565c0;font-size:10px;text-transform:uppercase">Status</th>';
  html += '<th style="padding:7px 8px;border-bottom:2px solid #90caf9;text-align:center;color:#1565c0;font-size:10px;text-transform:uppercase">Slots</th>';
  html += '</tr></thead><tbody>' + partnerRows + '</tbody></table></div>';
  if (d.partnerNotes) html += '<div style="margin-top:8px;padding:8px;background:#f8fafc;border-radius:7px;font-size:12px;color:#4a5568"><strong>Notes:</strong> ' + d.partnerNotes + '</div>';
  if (d.nextPDate)    html += '<div style="margin-top:5px;font-size:12px;color:#4a5568;padding:5px 8px"><strong>Next engagement:</strong> ' + d.nextPDate + '</div>';
  html += '</div>';

  // Documents
  html += '<div style="border-radius:9px;border:1px solid #dde3ea;border-left:4px solid #6a1b9a;padding:15px;margin-bottom:13px">';
  html += '<div style="font-size:11px;font-weight:700;color:#6a1b9a;text-transform:uppercase;letter-spacing:.06em;margin-bottom:9px">Documents &amp; Media</div>';
  html += fileLinksHtml;
  if (d.docNotes)     html += '<div style="margin-top:8px;padding:8px;background:#f8fafc;border-radius:7px;font-size:12px;color:#4a5568"><strong>Doc notes:</strong> ' + d.docNotes + '</div>';
  if (d.photoCaption) html += '<div style="margin-top:5px;padding:8px;background:#f8fafc;border-radius:7px;font-size:12px;color:#4a5568"><strong>Photo caption:</strong> ' + d.photoCaption + '</div>';
  if (d.videoCaption) html += '<div style="margin-top:5px;padding:8px;background:#f8fafc;border-radius:7px;font-size:12px;color:#4a5568"><strong>Video description:</strong> ' + d.videoCaption + '</div>';
  if (d.mediaContext) html += '<div style="margin-top:5px;padding:8px;background:#f8fafc;border-radius:7px;font-size:12px;color:#4a5568"><strong>Media context:</strong> ' + d.mediaContext + '</div>';
  html += '</div>';

  // Safeguarding
  html += '<div style="border-radius:9px;border:1px solid #dde3ea;border-left:4px solid #00695c;padding:15px;margin-bottom:13px">';
  html += '<div style="font-size:11px;font-weight:700;color:#00695c;text-transform:uppercase;letter-spacing:.06em;margin-bottom:9px">Safeguarding</div>';
  html += '<div style="font-size:13px;color:#4a5568;margin-bottom:7px">' + d.safeChecked.length + ' of 8 items confirmed</div>';
  html += '<ul style="font-size:13px;padding-left:18px;margin:5px 0 8px">' + safeHtml + '</ul>';
  html += concernHtml;
  if (d.safeNotes) html += '<div style="margin-top:7px;padding:8px;background:#f8fafc;border-radius:7px;font-size:12px;color:#4a5568">' + d.safeNotes + '</div>';
  html += '</div>';

  // Final notes
  if (d.finalNotes) {
    html += '<div style="border-radius:9px;border:1px solid #dde3ea;border-left:4px solid #455a64;padding:15px;margin-bottom:13px">';
    html += '<div style="font-size:11px;font-weight:700;color:#455a64;text-transform:uppercase;letter-spacing:.06em;margin-bottom:7px">Additional Notes</div>';
    html += '<div style="font-size:13px;color:#4a5568;background:#f8fafc;padding:10px;border-radius:7px;border-left:3px solid #4caf50">' + d.finalNotes + '</div>';
    html += '</div>';
  }

  html += '</div>'; // end padding

  html += '<div style="background:#eceff1;text-align:center;padding:12px;font-size:11px;color:#718096">';
  html += 'Submitted via YiW Field Reporting System - ' + ts + '<br/>SEG Ghana | Youth in Work Programme';
  html += '</div>';

  html += '</div></body></html>';

  return html;
}
