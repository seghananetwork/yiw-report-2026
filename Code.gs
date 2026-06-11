/**
 * YiW Field Report — Google Apps Script
 * Actions: submitWithLinks (primary), legacy fallback
 * On each submission:
 *   1. Sends HTML email to all recipients
 *   2. Appends a row to the master Google Sheet (creates it if needed)
 *   3. Generates a formatted per-report Sheet tab and exports as XLSX attachment
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
var MASTER_SHEET_NAME = "YiW Daily Field Reports — Master Log";

// ── ENTRY POINT ──────────────────────────────────────────────
function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var action  = payload.action || 'legacy';
    if (action === 'submitWithLinks') return handleSubmitWithLinks(payload);
    return handleLegacy(payload);
  } catch(err) {
    Logger.log('FATAL: ' + err.toString());
    return jsonOut({ status:'error', message: err.toString() });
  }
}

function doGet(e) {
  return jsonOut({ status:'ok', message:'YiW Script is live.' });
}

// ── PRIMARY HANDLER ──────────────────────────────────────────
function handleSubmitWithLinks(p) {
  var d          = p.formData;
  var driveLinks = p.driveLinks || {};
  var totalFiles = p.totalFiles || 0;

  // 1. Append row to master Google Sheet AND to the hub-specific sheet
  var sheetUrl = appendToMasterSheet(d, driveLinks, totalFiles);
  appendToHubSheet(d, driveLinks);

  // 2. Build file links HTML for email
  var fileLinksHtml = buildFileLinksHtml(driveLinks);

  // 3. Build and send HTML email — includes link to master sheet
  var htmlBody = buildEmailHtml(d, fileLinksHtml, sheetUrl);
  var subject  = 'YiW Field Report: ' + (d.fpName||'--') +
                 ' -- ' + (d.trainingCentre||d.hubName||'--') +
                 ' (' + (d.visitDate||'--') + ')';

  MailApp.sendEmail({
    to:       TO_EMAIL,
    cc:       CC_EMAILS,
    subject:  subject,
    htmlBody: htmlBody
  });

  Logger.log('Done: ' + subject);
  return jsonOut({ status:'success', message:'Report submitted. Sheet updated.' });
}

// ── LEGACY FALLBACK ──────────────────────────────────────────
function handleLegacy(payload) {
  var d        = payload.formData || {};
  var sheetUrl = appendToMasterSheet(d, {}, 0);
  appendToHubSheet(d, {});
  var htmlBody = buildEmailHtml(d, '<p style="color:#718096;font-style:italic">No files attached.</p>', sheetUrl);
  var subject  = 'YiW Field Report: ' + (d.fpName||'--') + ' (' + (d.visitDate||'--') + ')';
  MailApp.sendEmail({ to:TO_EMAIL, cc:CC_EMAILS, subject:subject, htmlBody:htmlBody });
  return jsonOut({ status:'success', message:'Report submitted.' });
}

// ── HUB-SPECIFIC SHEET ───────────────────────────────────────
// Each TSP/Hub gets its own tab in the master workbook.
// All focal persons reporting for that hub write to the same tab.
// Every column header has auto-filter enabled so data can be sorted.
function appendToHubSheet(d, driveLinks) {
  try {
    // Open (or create) the same master workbook
    var ss;
    var files = DriveApp.getFilesByName(MASTER_SHEET_NAME);
    if (!files.hasNext()) return; // master sheet must exist first
    ss = SpreadsheetApp.open(files.next());

    var hubName = (d.hubName || 'Unknown Hub').toString().trim();
    // Sheet names max 100 chars, no special chars
    var sheetName = hubName.length > 95 ? hubName.substring(0, 95) : hubName;

    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      formatHubSheet(sheet, hubName);
    }

    // If somehow the header row is missing, add it
    if (sheet.getLastRow() === 0) {
      formatHubSheet(sheet, hubName);
    }

    var fileCounts = countDriveLinks(driveLinks);
    var partners   = d.partners || [];
    var allFileUrls = [];
    ['dAtt','dFin','dMou','dTrack','mPhoto','mVideo'].forEach(function(cat){
      (driveLinks[cat]||[]).forEach(function(f){ allFileUrls.push(f.name+': '+f.url); });
    });

    var row = [
      new Date(),
      d.fpName          || '',
      d.fpPhone         || '',
      d.fpEmail         || '',
      d.fpZone          || '',
      d.visitDate       || '',
      d.visitType       || '',
      d.community       || '',
      d.trainingCentre  || '',
      d.hubContact      || '',
      d.hubContactPhone || '',
      d.tArr            || '',
      d.tDep            || '',
      // Attendance
      d.cMale    || 0,
      d.cFemale  || 0,
      d.cPWD     || 0,
      d.cStaff   || 0,
      d.cTrainer || 0,
      (d.cMale||0)+(d.cFemale||0)+(d.cPWD||0),
      // Activation
      d.aJobs    || 0,
      d.aIntern  || 0,
      d.aCoop    || 0,
      d.aRef     || 0,
      (d.aJobs||0)+(d.aIntern||0)+(d.aCoop||0)+(d.aRef||0),
      d.enrolM       || 0,
      d.enrolF       || 0,
      d.enrolCourse  || '',
      d.empName      || '',
      d.empSector    || '',
      // Quality
      d.rating || '',
      (d.quality    ||[]).join('; '),
      (d.issues     ||[]).join('; '),
      (d.facilities ||[]).join('; '),
      (d.activities ||[]).join('; '),
      d.challenges      || '',
      d.recommendations || '',
      d.urgency         || '',
      d.followUpBy      || '',
      // Partners
      partners.length,
      partners.map(function(p){ return p.name+(p.status?' ('+p.status+')':''); }).join('; '),
      partners.map(function(p){ return p.skillsNeeded||''; }).join('; '),
      // Files
      fileCounts.total,
      allFileUrls.join(' | '),
      // Safeguarding
      (d.safeChecked||[]).length,
      d.safeConcern === 'yes' ? 'YES' : 'No',
      d.safeTxt || '',
      // Narrative
      d.highlight  || '',
      d.yVoice     || '',
      d.finalNotes || ''
    ];

    sheet.appendRow(row);

    // Alternate row shading
    var lastRow = sheet.getLastRow();
    if (lastRow % 2 === 0) {
      sheet.getRange(lastRow, 1, 1, row.length).setBackground('#f0f4f0');
    }

    SpreadsheetApp.flush();
    Logger.log('Hub sheet updated: ' + sheetName + ' row ' + lastRow);

  } catch(err) {
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

  // Title row showing hub name
  sheet.getRange(1, 1, 1, headers.length).merge()
       .setValue(hubName + ' — YiW Field Reports')
       .setBackground('#1a5c2a')
       .setFontColor('#ffffff')
       .setFontWeight('bold')
       .setFontSize(12);

  // Header row
  var headerRange = sheet.getRange(2, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setBackground('#2d7a3a');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setFontSize(10);
  headerRange.setWrap(false);

  // Freeze title + header rows, apply auto-filter for sorting
  sheet.setFrozenRows(2);
  sheet.getRange(2, 1, 1, headers.length).createFilter();

  // Column widths
  sheet.setColumnWidth(1,  160); // Submitted At
  sheet.setColumnWidth(2,  140); // FP Name
  sheet.setColumnWidth(6,  100); // Visit Date
  sheet.setColumnWidth(8,  120); // Community
  sheet.setColumnWidth(9,  170); // Training Centre
  sheet.setColumnWidth(31, 220); // Quality
  sheet.setColumnWidth(32, 200); // Issues
  sheet.setColumnWidth(35, 260); // Challenges
  sheet.setColumnWidth(36, 260); // Recommendations
  sheet.setColumnWidth(42, 300); // File links
}
// One permanent Google Sheet — one row per submission, exactly like Google Forms.
// Returns the sheet URL so it can be linked in the email.
function appendToMasterSheet(d, driveLinks, totalFiles) {
  try {
    var ss;
    var sheetUrl = '';
    var files = DriveApp.getFilesByName(MASTER_SHEET_NAME);
    if (files.hasNext()) {
      var f = files.next();
      ss = SpreadsheetApp.open(f);
      sheetUrl = f.getUrl();
    } else {
      ss = SpreadsheetApp.create(MASTER_SHEET_NAME);
      // Share with all recipients so they can open it directly
      var newFile = DriveApp.getFileById(ss.getId());
      newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      sheetUrl = newFile.getUrl();
    }

    var sheet = ss.getSheetByName('Field Reports');
    if (!sheet) {
      sheet = ss.getActiveSheet();
      sheet.setName('Field Reports');
    }

    // Add header row if sheet is empty
    if (sheet.getLastRow() === 0) {
      formatMasterSheet(sheet);
    }

    var fileCounts = countDriveLinks(driveLinks);
    var partners   = d.partners || [];

    // Build all Drive file URLs as a single string for the sheet
    var allFileUrls = [];
    ['dAtt','dFin','dMou','dTrack','mPhoto','mVideo'].forEach(function(cat){
      (driveLinks[cat]||[]).forEach(function(f){ allFileUrls.push(f.name + ': ' + f.url); });
    });

    var dataRow = [
      new Date(),
      d.fpName         || '',
      d.fpPhone        || '',
      d.fpEmail        || '',
      d.fpZone         || '',
      d.visitDate      || '',
      d.visitType      || '',
      d.hubName        || '',
      d.community      || '',
      d.trainingCentre || '',
      d.centreAddress  || '',
      d.hubContact     || '',
      d.hubContactPhone|| '',
      d.tArr           || '',
      d.tDep           || '',
      // Attendance
      d.cMale    || 0,
      d.cFemale  || 0,
      d.cPWD     || 0,
      d.cStaff   || 0,
      d.cTrainer || 0,
      (d.cMale||0)+(d.cFemale||0)+(d.cPWD||0),
      // Activation
      d.aJobs    || 0,
      d.aIntern  || 0,
      d.aCoop    || 0,
      d.aRef     || 0,
      (d.aJobs||0)+(d.aIntern||0)+(d.aCoop||0)+(d.aRef||0),
      d.enrolM      || 0,
      d.enrolF      || 0,
      d.enrolCourse || '',
      d.empName     || '',
      d.empSector   || '',
      // Quality
      d.rating || '',
      (d.quality    ||[]).join('; '),
      (d.issues     ||[]).join('; '),
      (d.facilities ||[]).join('; '),
      (d.activities ||[]).join('; '),
      d.challenges      || '',
      d.recommendations || '',
      d.urgency         || '',
      d.followUpBy      || '',
      // Partners
      partners.length,
      partners.map(function(p){ return p.name + (p.status?' ('+p.status+')':''); }).join('; '),
      partners.map(function(p){ return p.skillsNeeded||''; }).join('; '),
      // Files
      fileCounts.total,
      fileCounts.dAtt,
      fileCounts.dFin,
      fileCounts.dMou,
      fileCounts.dTrack,
      fileCounts.mPhoto,
      fileCounts.mVideo,
      allFileUrls.join(' | '),
      // Safeguarding
      (d.safeChecked||[]).length,
      (d.safeChecked||[]).join('; '),
      d.safeConcern === 'yes' ? 'YES' : 'No',
      d.safeTxt || '',
      // Narrative
      d.highlight  || '',
      d.yVoice     || '',
      d.finalNotes || ''
    ];

    sheet.appendRow(dataRow);

    // Alternate row shading
    var lastRow = sheet.getLastRow();
    if (lastRow % 2 === 0) {
      sheet.getRange(lastRow, 1, 1, dataRow.length).setBackground('#f0f4f0');
    }

    SpreadsheetApp.flush();
    Logger.log('Master sheet updated: row ' + lastRow + ' | URL: ' + sheetUrl);
    return sheetUrl;

  } catch(err) {
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
  // Auto-filter enables clicking any column header to sort
  sheet.getRange(1, 1, 1, headers.length).createFilter();
  sheet.setColumnWidth(1,  160); // Submitted At
  sheet.setColumnWidth(2,  140); // FP Name
  sheet.setColumnWidth(8,  220); // Hub
  sheet.setColumnWidth(9,  120); // Community
  sheet.setColumnWidth(10, 170); // Training Centre
  sheet.setColumnWidth(33, 220); // Quality
  sheet.setColumnWidth(34, 200); // Issues
  sheet.setColumnWidth(36, 260); // Challenges
  sheet.setColumnWidth(37, 260); // Recommendations
  sheet.setColumnWidth(42, 200); // Skills requested
  sheet.setColumnWidth(50, 300); // File links
}

// ── HELPERS ──────────────────────────────────────────────────
function countDriveLinks(driveLinks) {
  var cats = ['dAtt','dFin','dMou','dTrack','mPhoto','mVideo'];
  var out  = { total:0, dAtt:0, dFin:0, dMou:0, dTrack:0, mPhoto:0, mVideo:0 };
  cats.forEach(function(c){ out[c]=(driveLinks[c]||[]).length; out.total+=out[c]; });
  return out;
}

function buildFileLinksHtml(driveLinks) {
  var catOrder = ['dAtt','dFin','dMou','dTrack','mPhoto','mVideo'];
  var catNames = {
    dAtt:'Attendance Sheet', dFin:'Financial Document',
    dMou:'MoU / Agreement',  dTrack:'Tracking Sheet',
    mPhoto:'Photo',          mVideo:'Video'
  };
  var rows = '';
  var count = 0;
  catOrder.forEach(function(cat){
    (driveLinks[cat]||[]).forEach(function(f){
      rows += '<tr>' +
        '<td style="padding:8px;border:1px solid #cbd5e1;font-weight:600;color:#1a5c2a;font-size:12px">'+(f.category||catNames[cat]||cat)+'</td>' +
        '<td style="padding:8px;border:1px solid #cbd5e1;font-size:12px"><a href="'+f.url+'" style="color:#1565c0;font-weight:500">🔗 '+f.name+'</a></td>' +
        '</tr>';
      count++;
    });
  });
  if (!count) return '<p style="color:#718096;font-style:italic;font-size:13px">No files attached for this submission.</p>';
  return '<table style="width:100%;border-collapse:collapse;font-size:13px;margin-top:6px">' +
    '<tr style="background:#e8f5eb"><th style="padding:8px;border:1px solid #cbd5e1;text-align:left">Category</th>' +
    '<th style="padding:8px;border:1px solid #cbd5e1;text-align:left">File — click to open in Drive</th></tr>' +
    rows + '</table>';
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function safe(val) { return val || '—'; }

// ── BUILD HTML EMAIL ─────────────────────────────────────────
function buildEmailHtml(d, fileLinksHtml, sheetUrl) {

  function pills(arr, bg, color) {
    if (!arr || arr.length===0) return '<em style="color:#718096">None</em>';
    return arr.map(function(i){
      return '<span style="display:inline-block;background:'+bg+';color:'+color+
             ';padding:3px 9px;border-radius:12px;margin:2px;font-size:12px;font-weight:600">'+i+'</span>';
    }).join('');
  }

  function statBox(val, label, bg, color) {
    return '<td style="text-align:center;background:'+bg+';padding:10px 6px;border-radius:6px;width:20%">' +
           '<div style="font-size:20px;font-weight:800;color:'+color+'">'+(val||0)+'</div>' +
           '<div style="font-size:11px;color:#718096;margin-top:2px">'+label+'</div></td>';
  }

  var rating    = parseInt(d.rating)||0;
  var ratingStr = rating>0 ? ('★'.repeat(rating)+'☆'.repeat(5-rating)+' ('+rating+'/5)') : 'Not rated';
  var ts        = new Date().toLocaleString('en-GB', {timeZone:'Africa/Accra'});

  var partnerRows = '';
  if (d.partners && d.partners.length>0) {
    d.partners.forEach(function(p){
      partnerRows +=
        '<tr>'+
        '<td style="padding:7px 8px;border-bottom:1px solid #dde3ea;font-weight:600;font-size:12px">'+safe(p.name)+'</td>'+
        '<td style="padding:7px 8px;border-bottom:1px solid #dde3ea;font-size:12px">'+safe(p.location)+'</td>'+
        '<td style="padding:7px 8px;border-bottom:1px solid #dde3ea;font-size:12px">'+safe(p.sector)+'</td>'+
        '<td style="padding:7px 8px;border-bottom:1px solid #dde3ea;font-size:11px;color:#555">'+safe(p.profile)+'</td>'+
        '<td style="padding:7px 8px;border-bottom:1px solid #dde3ea;font-size:11px;color:#555">'+safe(p.skillsNeeded)+'</td>'+
        '<td style="padding:7px 8px;border-bottom:1px solid #dde3ea;font-size:12px">'+safe(p.contact)+'</td>'+
        '<td style="padding:7px 8px;border-bottom:1px solid #dde3ea;font-size:12px">'+safe(p.phone)+'</td>'+
        '<td style="padding:7px 8px;border-bottom:1px solid #dde3ea">'+
          '<span style="background:#fff3e0;color:#e65100;padding:2px 7px;border-radius:10px;font-size:11px;font-weight:600">'+safe(p.status)+'</span>'+
        '</td>'+
        '<td style="padding:7px 8px;border-bottom:1px solid #dde3ea;text-align:center;font-weight:700;color:#1a5c2a">'+(p.slots||0)+'</td>'+
        '</tr>';
    });
  } else {
    partnerRows='<tr><td colspan="9" style="padding:12px;text-align:center;color:#718096;font-style:italic">No partner companies logged.</td></tr>';
  }

  var safeChecked = d.safeChecked||[];
  var safeHtml    = safeChecked.length>0
    ? safeChecked.map(function(s){return '<li style="margin-bottom:3px">'+s+'</li>';}).join('')
    : '<li style="color:#718096">No safeguarding items confirmed.</li>';
  var concernHtml = d.safeConcern==='yes'
    ? '<div style="background:#ffebee;border-left:4px solid #c62828;padding:12px;border-radius:6px;margin-top:10px;color:#c62828">'+
      '<strong>⚠ SAFEGUARDING CONCERN RAISED</strong><br/>'+
      '<strong>Details:</strong> '+safe(d.safeTxt)+'<br/>'+
      '<strong>Action:</strong> '+safe(d.safeAct)+'<br/>'+
      '<strong>Reported to:</strong> '+safe(d.safeRep)+'</div>'
    : '<p style="color:#1a5c2a;font-size:13px;margin-top:6px">✓ No concerns identified.</p>';

  return '<!DOCTYPE html><html><head><meta charset="UTF-8"/></head>'+
  '<body style="font-family:\'Segoe UI\',Arial,sans-serif;background:#f3f7f4;margin:0;padding:20px">'+
  '<div style="max-width:700px;margin:0 auto;background:#fff;border-radius:12px;border:1px solid #dde3ea;overflow:hidden;box-shadow:0 4px 16px rgba(0,0,0,.07)">'+

  // HEADER
  '<div style="background:linear-gradient(135deg,#1a5c2a,#2d7a3a);padding:22px 26px;color:#fff">'+
    '<div style="font-size:10px;opacity:.7;text-transform:uppercase;letter-spacing:.1em;margin-bottom:3px">SEG Ghana</div>'+
    '<div style="font-size:20px;font-weight:700;margin-bottom:2px">Youth in Work Programme</div>'+
    '<div style="font-size:13px;opacity:.85;margin-bottom:14px">Daily Field Report</div>'+
    '<div style="display:flex;flex-wrap:wrap;gap:16px">'+
      '<div><div style="font-size:10px;opacity:.65;text-transform:uppercase">Submitted by</div><div style="font-size:14px;font-weight:700;margin-top:1px">'+safe(d.fpName)+'</div></div>'+
      '<div><div style="font-size:10px;opacity:.65;text-transform:uppercase">Phone</div><div style="font-size:14px;font-weight:700;margin-top:1px">'+safe(d.fpPhone)+'</div></div>'+
      '<div><div style="font-size:10px;opacity:.65;text-transform:uppercase">Date</div><div style="font-size:14px;font-weight:700;margin-top:1px">'+safe(d.visitDate)+'</div></div>'+
      '<div><div style="font-size:10px;opacity:.65;text-transform:uppercase">Zone</div><div style="font-size:14px;font-weight:700;margin-top:1px">'+safe(d.fpZone)+'</div></div>'+
    '</div>'+
  '</div>'+

  '<div style="padding:20px">'+

  // MASTER SHEET LINK — prominent banner
  (sheetUrl ?
    '<div style="background:#1a5c2a;border-radius:9px;padding:14px 16px;margin-bottom:14px;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:10px">'+
      '<div>'+
        '<div style="color:#fff;font-weight:700;font-size:13px">View & Download Master Data Sheet</div>'+
        '<div style="color:rgba(255,255,255,.75);font-size:11px;margin-top:2px">All submissions in one Google Sheet — download as Excel anytime</div>'+
      '</div>'+
      '<a href="'+sheetUrl+'" style="display:inline-block;background:#fff;color:#1a5c2a;font-weight:700;font-size:12px;padding:8px 16px;border-radius:6px;text-decoration:none">Open Sheet &rarr;</a>'+
    '</div>'
  : '') +

  // NOTE ABOUT MASTER SHEET
  '<div style="background:#e8f5eb;border:1px solid #a5d6a7;border-radius:8px;padding:10px 13px;font-size:12px;color:#1a3a1a;margin-bottom:14px">'+
    '📊 <strong>Master data sheet updated</strong> — this report has been added to the central Google Sheet. '+
    'Open the sheet above and go to <strong>File → Download → Microsoft Excel (.xlsx)</strong> to export all submissions.'+
  '</div>'+

  // VISIT DETAILS
  '<div style="border-radius:9px;border:1px solid #dde3ea;border-left:4px solid #1a5c2a;padding:15px;margin-bottom:13px">'+
    '<div style="font-size:11px;font-weight:700;color:#1a5c2a;text-transform:uppercase;letter-spacing:.06em;margin-bottom:9px">📍 Visit Details</div>'+
    '<table style="width:100%;border-collapse:collapse;font-size:13px">'+
      '<tr><td style="padding:4px 0;color:#718096;width:38%">Visit type</td><td style="padding:4px 0;font-weight:600">'+safe(d.visitType)+'</td></tr>'+
      '<tr><td style="padding:4px 0;color:#718096">Hub / TSP</td><td style="padding:4px 0;font-weight:600">'+safe(d.hubName)+'</td></tr>'+
      '<tr><td style="padding:4px 0;color:#718096">Community</td><td style="padding:4px 0;font-weight:600">'+safe(d.community)+'</td></tr>'+
      '<tr><td style="padding:4px 0;color:#718096">Training centre</td><td style="padding:4px 0;font-weight:600">'+safe(d.trainingCentre)+'</td></tr>'+
      '<tr><td style="padding:4px 0;color:#718096">Address</td><td style="padding:4px 0;font-weight:600">'+safe(d.centreAddress)+'</td></tr>'+
      '<tr><td style="padding:4px 0;color:#718096">Centre contact</td><td style="padding:4px 0;font-weight:600">'+safe(d.hubContact)+(d.hubContactPhone?' · '+d.hubContactPhone:'')+'</td></tr>'+
      '<tr><td style="padding:4px 0;color:#718096">Time on site</td><td style="padding:4px 0;font-weight:600">'+safe(d.tArr)+' → '+safe(d.tDep)+'</td></tr>'+
    '</table>'+
  '</div>'+

  // ATTENDANCE
  '<div style="border-radius:9px;border:1px solid #dde3ea;border-left:4px solid #b8860b;padding:15px;margin-bottom:13px">'+
    '<div style="font-size:11px;font-weight:700;color:#b8860b;text-transform:uppercase;letter-spacing:.06em;margin-bottom:9px">👥 Attendance Count</div>'+
    '<table style="width:100%;border-collapse:separate;border-spacing:4px"><tr>'+
      statBox(d.cMale,'Young men','#fff8e1','#b8860b')+
      statBox(d.cFemale,'Young women','#fff8e1','#b8860b')+
      statBox(d.cPWD,'PWD','#fff8e1','#b8860b')+
      statBox(d.cStaff,'Staff','#fff8e1','#b8860b')+
      statBox(d.cTrainer,'Trainers','#fff8e1','#b8860b')+
    '</tr></table>'+
  '</div>'+

  // ACTIVATION
  '<div style="border-radius:9px;border:1px solid #dde3ea;border-left:4px solid #1a5c2a;padding:15px;margin-bottom:13px">'+
    '<div style="font-size:11px;font-weight:700;color:#1a5c2a;text-transform:uppercase;letter-spacing:.06em;margin-bottom:9px">🚀 Activation & Employment</div>'+
    '<table style="width:100%;border-collapse:separate;border-spacing:4px"><tr>'+
      statBox(d.aJobs,'Formal jobs','#e8f5eb','#1a5c2a')+
      statBox(d.aIntern,'Internships','#e8f5eb','#1a5c2a')+
      statBox(d.aCoop,'Cooperatives','#e8f5eb','#1a5c2a')+
      statBox(d.aRef,'Further trng','#e8f5eb','#1a5c2a')+
      '<td></td>'+
    '</tr></table>'+
    '<table style="width:100%;border-collapse:collapse;font-size:13px;margin-top:10px">'+
      '<tr><td style="padding:4px 0;color:#718096;width:38%">New enrolments (M/F)</td><td style="padding:4px 0;font-weight:600">'+(d.enrolM||0)+' / '+(d.enrolF||0)+'</td></tr>'+
      '<tr><td style="padding:4px 0;color:#718096">Course enrolled in</td><td style="padding:4px 0;font-weight:600">'+safe(d.enrolCourse)+'</td></tr>'+
      '<tr><td style="padding:4px 0;color:#718096">Employer / cooperative</td><td style="padding:4px 0;font-weight:600">'+safe(d.empName)+(d.empSector?' ('+d.empSector+')':'')+'</td></tr>'+
    '</table>'+
    (d.youthNames?'<div style="margin-top:8px;padding:9px;background:#f8fafc;border-radius:7px;font-size:12px;white-space:pre-wrap"><strong style="color:#1a5c2a">Youth placed:</strong><br/>'+d.youthNames+'</div>':'')+
    (d.highlight?'<div style="margin-top:7px;padding:9px;background:#e8f5eb;border-radius:7px;border-left:3px solid #4caf50;font-size:13px;color:#1a5c2a"><strong>✨ Success story:</strong><br/>'+d.highlight+'</div>':'')+
    (d.yVoice?'<div style="margin-top:6px;padding:9px;background:#f8fafc;border-radius:7px;border-left:3px solid #90caf9;font-size:13px;color:#4a5568;font-style:italic">"'+d.yVoice+'"</div>':'')+
  '</div>'+

  // HUB QUALITY
  '<div style="border-radius:9px;border:1px solid #dde3ea;border-left:4px solid #00695c;padding:15px;margin-bottom:13px">'+
    '<div style="font-size:11px;font-weight:700;color:#00695c;text-transform:uppercase;letter-spacing:.06em;margin-bottom:9px">⭐ Training Centre Quality</div>'+
    '<table style="width:100%;border-collapse:collapse;font-size:13px;margin-bottom:9px">'+
      '<tr><td style="padding:4px 0;color:#718096;width:38%">Overall rating</td><td style="padding:4px 0;font-weight:700;color:#1a5c2a">'+ratingStr+'</td></tr>'+
      '<tr><td style="padding:4px 0;color:#718096">Urgency</td><td style="padding:4px 0;font-weight:600">'+safe(d.urgency)+'</td></tr>'+
      '<tr><td style="padding:4px 0;color:#718096">Follow-up by</td><td style="padding:4px 0;font-weight:600">'+safe(d.followUpBy)+'</td></tr>'+
    '</table>'+
    '<div style="margin-bottom:7px"><strong style="font-size:11px;color:#00695c;text-transform:uppercase">Quality ✓</strong><br/><div style="margin-top:4px">'+pills(d.quality,'#e8f5eb','#1a5c2a')+'</div></div>'+
    '<div style="margin-bottom:7px"><strong style="font-size:11px;color:#c62828;text-transform:uppercase">Issues ⚠</strong><br/><div style="margin-top:4px">'+pills(d.issues,'#ffebee','#c62828')+'</div></div>'+
    '<div style="margin-bottom:7px"><strong style="font-size:11px;color:#1565c0;text-transform:uppercase">Activities</strong><br/><div style="margin-top:4px">'+pills(d.activities,'#e3f2fd','#1565c0')+'</div></div>'+
    '<div style="margin-bottom:7px"><strong style="font-size:11px;color:#455a64;text-transform:uppercase">Facilities</strong><br/><div style="margin-top:4px">'+pills(d.facilities,'#eceff1','#455a64')+'</div></div>'+
    (d.challenges?'<div style="margin-top:9px;padding:9px;background:#fff3e0;border-radius:7px;font-size:12px;color:#e65100;border-left:3px solid #ff9800"><strong>Challenges:</strong> '+d.challenges+'</div>':'')+
    (d.recommendations?'<div style="margin-top:6px;padding:9px;background:#e3f2fd;border-radius:7px;font-size:12px;color:#1565c0;border-left:3px solid #90caf9"><strong>Recommendations:</strong> '+d.recommendations+'</div>':'')+
  '</div>'+

  // PARTNERS
  '<div style="border-radius:9px;border:1px solid #dde3ea;border-left:4px solid #1565c0;padding:15px;margin-bottom:13px">'+
    '<div style="font-size:11px;font-weight:700;color:#1565c0;text-transform:uppercase;letter-spacing:.06em;margin-bottom:9px">🤝 Partner Engagement — '+(d.partners?d.partners.length:0)+' company(ies)</div>'+
    '<div style="overflow-x:auto">'+
      '<table style="width:100%;border-collapse:collapse;font-size:12px">'+
        '<thead><tr style="background:#e3f2fd">'+
          '<th style="padding:7px 8px;border-bottom:2px solid #90caf9;text-align:left;color:#1565c0;font-size:10px;text-transform:uppercase">Company</th>'+
          '<th style="padding:7px 8px;border-bottom:2px solid #90caf9;text-align:left;color:#1565c0;font-size:10px;text-transform:uppercase">Location</th>'+
          '<th style="padding:7px 8px;border-bottom:2px solid #90caf9;text-align:left;color:#1565c0;font-size:10px;text-transform:uppercase">Sector</th>'+
          '<th style="padding:7px 8px;border-bottom:2px solid #90caf9;text-align:left;color:#1565c0;font-size:10px;text-transform:uppercase">Business profile</th>'+
          '<th style="padding:7px 8px;border-bottom:2px solid #90caf9;text-align:left;color:#1565c0;font-size:10px;text-transform:uppercase">Skills needed</th>'+
          '<th style="padding:7px 8px;border-bottom:2px solid #90caf9;text-align:left;color:#1565c0;font-size:10px;text-transform:uppercase">Contact</th>'+
          '<th style="padding:7px 8px;border-bottom:2px solid #90caf9;text-align:left;color:#1565c0;font-size:10px;text-transform:uppercase">Phone</th>'+
          '<th style="padding:7px 8px;border-bottom:2px solid #90caf9;text-align:left;color:#1565c0;font-size:10px;text-transform:uppercase">Status</th>'+
          '<th style="padding:7px 8px;border-bottom:2px solid #90caf9;text-align:center;color:#1565c0;font-size:10px;text-transform:uppercase">Slots</th>'+
        '</tr></thead>'+
        '<tbody>'+partnerRows+'</tbody>'+
      '</table>'+
    '</div>'+
    (d.partnerNotes?'<div style="margin-top:8px;padding:8px;background:#f8fafc;border-radius:7px;font-size:12px;color:#4a5568"><strong>Notes:</strong> '+d.partnerNotes+'</div>':'')+
    (d.nextPDate?'<div style="margin-top:5px;font-size:12px;color:#4a5568;padding:5px 8px"><strong>Next engagement:</strong> '+d.nextPDate+'</div>':'')+
  '</div>'+

  // DOCUMENTS
  '<div style="border-radius:9px;border:1px solid #dde3ea;border-left:4px solid #6a1b9a;padding:15px;margin-bottom:13px">'+
    '<div style="font-size:11px;font-weight:700;color:#6a1b9a;text-transform:uppercase;letter-spacing:.06em;margin-bottom:9px">📎 Documents & Media</div>'+
    fileLinksHtml+
    (d.docNotes?'<div style="margin-top:8px;padding:8px;background:#f8fafc;border-radius:7px;font-size:12px;color:#4a5568"><strong>Doc notes:</strong> '+d.docNotes+'</div>':'')+
    (d.photoCaption?'<div style="margin-top:5px;padding:8px;background:#f8fafc;border-radius:7px;font-size:12px;color:#4a5568"><strong>Photo caption:</strong> '+d.photoCaption+'</div>':'')+
    (d.videoCaption?'<div style="margin-top:5px;padding:8px;background:#f8fafc;border-radius:7px;font-size:12px;color:#4a5568"><strong>Video description:</strong> '+d.videoCaption+'</div>':'')+
    (d.mediaContext?'<div style="margin-top:5px;padding:8px;background:#f8fafc;border-radius:7px;font-size:12px;color:#4a5568"><strong>Media context:</strong> '+d.mediaContext+'</div>':'')+
  '</div>'+

  // SAFEGUARDING
  '<div style="border-radius:9px;border:1px solid #dde3ea;border-left:4px solid #00695c;padding:15px;margin-bottom:13px">'+
    '<div style="font-size:11px;font-weight:700;color:#00695c;text-transform:uppercase;letter-spacing:.06em;margin-bottom:9px">🛡 Safeguarding</div>'+
    '<div style="font-size:13px;color:#4a5568;margin-bottom:7px">'+safeChecked.length+' of 8 items confirmed</div>'+
    '<ul style="font-size:13px;padding-left:18px;margin:5px 0 8px">'+safeHtml+'</ul>'+
    concernHtml+
    (d.safeNotes?'<div style="margin-top:7px;padding:8px;background:#f8fafc;border-radius:7px;font-size:12px;color:#4a5568">'+d.safeNotes+'</div>':'')+
  '</div>'+

  (d.finalNotes?
    '<div style="border-radius:9px;border:1px solid #dde3ea;border-left:4px solid #455a64;padding:15px;margin-bottom:13px">'+
      '<div style="font-size:11px;font-weight:700;color:#455a64;text-transform:uppercase;letter-spacing:.06em;margin-bottom:7px">📝 Additional Notes</div>'+
      '<div style="font-size:13px;color:#4a5568;background:#f8fafc;padding:10px;border-radius:7px;border-left:3px solid #4caf50">'+d.finalNotes+'</div>'+
    '</div>':'')+

  '</div>'+
  '<div style="background:#eceff1;text-align:center;padding:12px;font-size:11px;color:#718096">'+
    'Submitted via YiW Field Reporting System · '+ts+'<br/>SEG Ghana | Youth in Work Programme'+
  '</div>'+
  '</div></body></html>';
}
