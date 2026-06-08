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

  // 1. Build file links HTML for email
  var fileLinksHtml = buildFileLinksHtml(driveLinks);

  // 2. Build and send HTML email
  var htmlBody = buildEmailHtml(d, fileLinksHtml);
  var subject  = 'YiW Field Report: ' + (d.fpName||'—') +
                 ' — ' + (d.trainingCentre||d.hubName||'—') +
                 ' (' + (d.visitDate||'—') + ')';

  // 3. Build per-report Excel sheet and attach it
  var xlsxBlob = buildReportExcel(d, driveLinks);

  MailApp.sendEmail({
    to:          TO_EMAIL,
    cc:          CC_EMAILS,
    subject:     subject,
    htmlBody:    htmlBody,
    attachments: [xlsxBlob]
  });

  // 4. Append row to master Google Sheet log
  appendToMasterSheet(d, driveLinks, totalFiles);

  Logger.log('Done: ' + subject);
  return jsonOut({ status:'success', message:'Report emailed with Excel attachment.' });
}

// ── LEGACY FALLBACK ──────────────────────────────────────────
function handleLegacy(payload) {
  var d = payload.formData || {};
  var htmlBody = buildEmailHtml(d, '<p style="color:#718096;font-style:italic">No files attached.</p>');
  var subject  = 'YiW Field Report: ' + (d.fpName||'—') + ' (' + (d.visitDate||'—') + ')';
  var xlsxBlob = buildReportExcel(d, {});
  MailApp.sendEmail({ to:TO_EMAIL, cc:CC_EMAILS, subject:subject, htmlBody:htmlBody, attachments:[xlsxBlob] });
  appendToMasterSheet(d, {}, 0);
  return jsonOut({ status:'success', message:'Report submitted.' });
}

// ── MASTER SHEET LOG ─────────────────────────────────────────
// One row per submission in a persistent Google Sheet.
// All recipients can open this sheet to see all submissions in one place.
function appendToMasterSheet(d, driveLinks, totalFiles) {
  try {
    var ss;
    var files = DriveApp.getFilesByName(MASTER_SHEET_NAME);
    if (files.hasNext()) {
      ss = SpreadsheetApp.open(files.next());
    } else {
      ss = SpreadsheetApp.create(MASTER_SHEET_NAME);
      var sheet = ss.getActiveSheet();
      sheet.setName('Field Reports');
      formatMasterSheet(sheet);
    }

    var sheet = ss.getSheetByName('Field Reports') || ss.getActiveSheet();

    // Add header row if sheet is empty
    if (sheet.getLastRow() === 0) {
      formatMasterSheet(sheet);
    }

    // Count drive files
    var fileCounts = countDriveLinks(driveLinks);
    var partners   = d.partners || [];

    var row = [
      new Date(),
      d.fpName        || '',
      d.fpPhone       || '',
      d.fpEmail       || '',
      d.fpZone        || '',
      d.visitDate     || '',
      d.visitType     || '',
      d.hubName       || '',
      d.community     || '',
      d.trainingCentre|| '',
      d.centreAddress || '',
      d.hubContact    || '',
      d.hubContactPhone||'',
      d.tArr          || '',
      d.tDep          || '',
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
      d.enrolM   || 0,
      d.enrolF   || 0,
      d.enrolCourse || '',
      d.empName  || '',
      d.empSector|| '',
      // Quality
      d.rating   || '',
      (d.quality     ||[]).join('; '),
      (d.issues      ||[]).join('; '),
      (d.facilities  ||[]).join('; '),
      (d.activities  ||[]).join('; '),
      d.challenges      || '',
      d.recommendations || '',
      d.urgency         || '',
      d.followUpBy      || '',
      // Partners
      partners.length,
      partners.map(function(p){ return p.name+(p.status?' ('+p.status+')':''); }).join('; '),
      // Files
      fileCounts.total,
      fileCounts.dAtt,
      fileCounts.dFin,
      fileCounts.dMou,
      fileCounts.dTrack,
      fileCounts.mPhoto,
      fileCounts.mVideo,
      // Safeguarding
      (d.safeChecked||[]).length,
      (d.safeChecked||[]).join('; '),
      d.safeConcern === 'yes' ? 'YES' : 'No',
      d.safeTxt  || '',
      // Narrative
      d.highlight   || '',
      d.yVoice      || '',
      d.finalNotes  || ''
    ];

    sheet.appendRow(row);

    // Style the new row alternating
    var lastRow = sheet.getLastRow();
    if (lastRow % 2 === 0) {
      sheet.getRange(lastRow, 1, 1, row.length).setBackground('#f8fafc');
    }

    // Auto-resize columns periodically
    if (lastRow % 10 === 0) {
      sheet.autoResizeColumns(1, row.length);
    }

    Logger.log('Master sheet updated: row ' + lastRow);
  } catch(err) {
    Logger.log('Master sheet error: ' + err.toString());
  }
}

function formatMasterSheet(sheet) {
  var headers = [
    'Submitted At','FP Name','FP Phone','FP Email','Zone',
    'Visit Date','Visit Type','Hub / TSP','Community','Training Centre',
    'Centre Address','Centre Contact','Contact Phone','Time Arrived','Time Departed',
    'Young Men','Young Women','PWD','Staff','Trainers','Total Youth',
    'Formal Jobs','Internships','Cooperatives','Further Training','Total Activations',
    'Enrolments (M)','Enrolments (F)','Course','Employer','Sector',
    'Hub Rating','Quality Indicators','Issues Flagged','Facilities','Activities',
    'Challenges','Recommendations','Urgency','Follow-up By',
    'Partners Count','Partner Names & Status',
    'Total Files','Attendance Docs','Financial Docs','MoUs','Tracking Sheets','Photos','Videos',
    'Safeguarding Items','Safeguarding Details','Concern Raised','Concern Detail',
    'Success Story','Youth Voice','Final Notes'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Header styling — dark green
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#1a5c2a');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setFontSize(10);
  headerRange.setWrap(false);

  // Freeze header row
  sheet.setFrozenRows(1);

  // Set column widths for key columns
  sheet.setColumnWidth(1, 150);  // Submitted At
  sheet.setColumnWidth(2, 140);  // FP Name
  sheet.setColumnWidth(8, 200);  // Hub
  sheet.setColumnWidth(9, 120);  // Community
  sheet.setColumnWidth(10, 160); // Training Centre
  sheet.setColumnWidth(33, 220); // Quality
  sheet.setColumnWidth(34, 200); // Issues
  sheet.setColumnWidth(36, 250); // Challenges
  sheet.setColumnWidth(37, 250); // Recommendations
}

// ── PER-REPORT EXCEL EXPORT ──────────────────────────────────
// Creates a nicely formatted spreadsheet for this one report,
// exports it as .xlsx and returns it as a blob to attach to email.
function buildReportExcel(d, driveLinks) {
  try {
    var ss        = SpreadsheetApp.create('YiW_TempReport_' + new Date().getTime());
    var sheet     = ss.getActiveSheet();
    sheet.setName('Field Report');

    // ── TITLE BLOCK ──
    sheet.getRange('A1').setValue('SEG GHANA - YOUTH IN WORK PROGRAMME');
    sheet.getRange('A1').setFontSize(14).setFontWeight('bold').setFontColor('#1a5c2a');
    sheet.getRange('A1:H1').merge();

    sheet.getRange('A2').setValue('DAILY FIELD REPORT');
    sheet.getRange('A2').setFontSize(11).setFontColor('#2d7a3a').setFontWeight('bold');
    sheet.getRange('A2:H2').merge();

    sheet.getRange('A3').setValue('Generated: ' + new Date().toLocaleString('en-GB', {timeZone:'Africa/Accra'}));
    sheet.getRange('A3').setFontSize(9).setFontColor('#718096');
    sheet.getRange('A3:H3').merge();

    // Title background
    sheet.getRange('A1:H3').setBackground('#e8f5eb');

    var row = 5;

    // ── SECTION HELPER ──
    function sectionHeader(title, color) {
      sheet.getRange(row, 1, 1, 8).merge()
           .setValue(title)
           .setBackground(color||'#1a5c2a')
           .setFontColor('#ffffff')
           .setFontWeight('bold')
           .setFontSize(10);
      row++;
    }

    function dataRow(label, value, labelBg, valueBg) {
      var lCell = sheet.getRange(row, 1, 1, 3);
      var vCell = sheet.getRange(row, 4, 1, 5);
      lCell.merge().setValue(label)
           .setBackground(labelBg||'#f8fafc')
           .setFontWeight('bold').setFontSize(9).setFontColor('#4a5568');
      vCell.merge().setValue(value||'—')
           .setBackground(valueBg||'#ffffff')
           .setFontSize(9).setFontColor('#1a1a2e');
      row++;
    }

    function statRow(labels, values, bg, color) {
      for (var i=0; i<labels.length; i++) {
        sheet.getRange(row, i+1).setValue(labels[i])
             .setBackground('#f8fafc').setFontWeight('bold')
             .setFontSize(9).setFontColor('#718096').setHorizontalAlignment('center');
      }
      row++;
      for (var j=0; j<values.length; j++) {
        sheet.getRange(row, j+1).setValue(values[j]||0)
             .setBackground(bg||'#e8f5eb').setFontWeight('bold')
             .setFontSize(13).setFontColor(color||'#1a5c2a').setHorizontalAlignment('center');
      }
      row+=2;
    }

    // ── SECTION 1: FOCAL PERSON & VISIT ──
    sectionHeader('FOCAL PERSON & VISIT DETAILS', '#1a5c2a');
    dataRow('Focal Person Name', d.fpName);
    dataRow('Phone', d.fpPhone);
    dataRow('Email', d.fpEmail);
    dataRow('Zone / Region', d.fpZone);
    dataRow('Date of Visit', d.visitDate);
    dataRow('Visit Type', d.visitType);
    dataRow('Hub / TSP', d.hubName);
    dataRow('Community', d.community);
    dataRow('Training Centre', d.trainingCentre);
    dataRow('Centre Address', d.centreAddress);
    dataRow('Centre Contact', (d.hubContact||'') + (d.hubContactPhone?' - '+d.hubContactPhone:''));
    dataRow('Time on Site', (d.tArr||'--') + '  to  ' + (d.tDep||'--'));
    row++;

    // ── SECTION 2: ATTENDANCE ──
    sectionHeader('ATTENDANCE COUNT', '#b8860b');
    statRow(
      ['Young Men','Young Women','Persons with Disability','Hub Staff','Trainers / Facilitators'],
      [d.cMale, d.cFemale, d.cPWD, d.cStaff, d.cTrainer],
      '#fff8e1','#b8860b'
    );

    // ── SECTION 3: ACTIVATION ──
    sectionHeader('ACTIVATION & EMPLOYMENT OUTCOMES', '#2d7a3a');
    statRow(
      ['Formal Employment','Internships','Cooperatives','Referred for Training'],
      [d.aJobs, d.aIntern, d.aCoop, d.aRef],
      '#e8f5eb','#1a5c2a'
    );
    dataRow('New Enrolments (Male)', d.enrolM||0);
    dataRow('New Enrolments (Female)', d.enrolF||0);
    dataRow('Course / Trade', d.enrolCourse);
    dataRow('Employer / Cooperative', (d.empName||'') + (d.empSector?' ('+d.empSector+')':''));
    dataRow('Youth Placed (names)', d.youthNames);
    dataRow('Success Story / Highlight', d.highlight);
    dataRow('Youth Voice / Quote', d.yVoice);
    row++;

    // ── SECTION 4: HUB QUALITY ──
    sectionHeader('TRAINING CENTRE QUALITY & COMPLIANCE', '#00695c');
    dataRow('Overall Rating', d.rating ? d.rating+'/5' : '--');
    dataRow('Quality Indicators', (d.quality||[]).join(', '));
    dataRow('Issues Flagged', (d.issues||[]).join(', '));
    dataRow('Activities Observed', (d.activities||[]).join(', '));
    dataRow('Facilities Available', (d.facilities||[]).join(', '));
    dataRow('Challenges', d.challenges);
    dataRow('Recommendations', d.recommendations);
    dataRow('Urgency of Action', d.urgency);
    dataRow('Follow-up By', d.followUpBy);
    row++;

    // ── SECTION 5: PARTNER ENGAGEMENT ──
    sectionHeader('PARTNER ENGAGEMENT', '#1565c0');
    var partners = d.partners || [];
    if (partners.length > 0) {
      var pHeaders = ['Company','Location','Sector','Business Profile','Skills Needed','Contact','Phone','Status','Slots'];
      pHeaders.forEach(function(h,i){
        sheet.getRange(row, i+1).setValue(h)
             .setBackground('#e3f2fd').setFontWeight('bold')
             .setFontSize(9).setFontColor('#1565c0');
      });
      row++;
      partners.forEach(function(p){
        var vals = [p.name,p.location,p.sector,p.profile,p.skillsNeeded,p.contact,p.phone,p.status,p.slots||0];
        vals.forEach(function(v,i){
          sheet.getRange(row, i+1).setValue(v||'')
               .setFontSize(9).setBackground('#f8fafc');
        });
        row++;
      });
    } else {
      dataRow('Partners', 'No partner companies logged for this visit.');
    }
    dataRow('Partner Notes', d.partnerNotes);
    dataRow('Next Engagement Date', d.nextPDate);
    row++;

    // ── SECTION 6: DOCUMENTS & MEDIA ──
    sectionHeader('DOCUMENTS & MEDIA UPLOADED', '#6a1b9a');
    var fc = countDriveLinks(driveLinks);
    dataRow('Attendance Sheets', fc.dAtt + ' file(s)');
    dataRow('Financial Documents', fc.dFin + ' file(s)');
    dataRow('MoUs / Agreements', fc.dMou + ' file(s)');
    dataRow('Tracking Sheets', fc.dTrack + ' file(s)');
    dataRow('Photos', fc.mPhoto + ' file(s)');
    dataRow('Videos', fc.mVideo + ' file(s)');
    dataRow('Total Files', fc.total + ' file(s)');
    dataRow('Document Notes', d.docNotes);
    dataRow('Photo Caption', d.photoCaption);
    dataRow('Video Description', d.videoCaption);
    dataRow('Media Context', d.mediaContext);

    var catOrder = ['dAtt','dFin','dMou','dTrack','mPhoto','mVideo'];
    var catNames = {dAtt:'Attendance',dFin:'Financial',dMou:'MoU',dTrack:'Tracking',mPhoto:'Photo',mVideo:'Video'};
    catOrder.forEach(function(cat){
      var files = driveLinks[cat]||[];
      files.forEach(function(f){
        sheet.getRange(row,1,1,2).merge().setValue(catNames[cat]||cat)
             .setBackground('#f8fafc').setFontWeight('bold').setFontSize(9).setFontColor('#6a1b9a');
        sheet.getRange(row,3,1,6).merge().setValue(f.url||f.name||'')
             .setFontSize(9).setFontColor('#1565c0');
        row++;
      });
    });
    row++;

    // ── SECTION 7: SAFEGUARDING ──
    sectionHeader('SAFEGUARDING', '#00695c');
    dataRow('Items Confirmed', (d.safeChecked||[]).length + ' of 8');
    dataRow('Confirmed Items', (d.safeChecked||[]).join(', '));
    dataRow('Concern Raised?', d.safeConcern==='yes' ? 'YES' : 'No');
    if (d.safeConcern==='yes') {
      dataRow('Concern Details', d.safeTxt, '#fff0f0', '#fff8f8');
      dataRow('Action Taken',    d.safeAct, '#fff0f0', '#fff8f8');
      dataRow('Reported To',     d.safeRep, '#fff0f0', '#fff8f8');
    }
    dataRow('Safeguarding Notes', d.safeNotes);
    row++;

    // ── SECTION 8: ADDITIONAL NOTES ──
    if (d.finalNotes) {
      sectionHeader('ADDITIONAL NOTES', '#455a64');
      dataRow('Notes', d.finalNotes);
    }

    // ── FORMATTING FINAL TOUCHES ──
    sheet.setColumnWidth(1, 190);
    sheet.setColumnWidth(2, 10);
    sheet.setColumnWidth(3, 10);
    sheet.setColumnWidth(4, 320);
    sheet.setColumnWidth(5, 10);
    sheet.setColumnWidth(6, 10);
    sheet.setColumnWidth(7, 10);
    sheet.setColumnWidth(8, 10);
    sheet.getRange(1, 1, row, 8).setWrap(true).setVerticalAlignment('middle');

    // CRITICAL: flush() forces all pending Sheets writes to disk
    // before we attempt to export — without this the sheet exports blank
    SpreadsheetApp.flush();
    Utilities.sleep(2000); // extra safety buffer for large reports

    // Export as XLSX
    var ssId  = ss.getId();
    var exportUrl = 'https://docs.google.com/spreadsheets/d/' + ssId +
                    '/export?format=xlsx&id=' + ssId;
    var token = ScriptApp.getOAuthToken();
    var resp  = UrlFetchApp.fetch(exportUrl, {
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    });

    if (resp.getResponseCode() !== 200) {
      throw new Error('Export failed with HTTP ' + resp.getResponseCode());
    }

    var xlsx = resp.getBlob().setName(
      'YiW_Report_' + (d.fpName||'Unknown').replace(/\s+/g,'_') +
      '_' + (d.visitDate||'nodate') + '.xlsx'
    );

    // Clean up temp spreadsheet
    try { DriveApp.getFileById(ssId).setTrashed(true); } catch(e) {}

    return xlsx;

  } catch(err) {
    Logger.log('Excel build error: ' + err.toString());
    // Return empty blob so email still sends
    return Utilities.newBlob('Excel generation failed: '+err, 'text/plain', 'report_error.txt');
  }
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
function buildEmailHtml(d, fileLinksHtml) {

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

  // NOTE ABOUT EXCEL
  '<div style="background:#e8f5eb;border:1px solid #a5d6a7;border-radius:8px;padding:10px 13px;font-size:12px;color:#1a3a1a;margin-bottom:14px">'+
    '📊 <strong>Excel report attached</strong> — a formatted spreadsheet with all report data is attached to this email.'+
    ' The master log sheet is also updated automatically.'+
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
