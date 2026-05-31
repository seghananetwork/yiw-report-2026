/**
 * YiW Field Report — Google Apps Script
 * Handles chunked file uploads + report email dispatch.
 *
 * Two actions:
 *   uploadChunk  — receives a piece of a file, reassembles in Drive temp store
 *   finalise     — all files done, send email with Drive links
 *
 * Deploy as: Web App | Execute as: Me | Access: Anyone
 */

var ROOT_FOLDER_NAME = "Youth in Work Field Reports Files";
var TO_EMAIL         = "yiw@seghana.net";
var CC_EMAIL         = "execdir@seghana.net";

// ── ENTRY POINT ──────────────────────────────────────────────
function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var action  = payload.action || 'legacy';

    if (action === 'uploadChunk') {
      return handleChunk(payload);
    } else if (action === 'finalise') {
      return handleFinalise(payload);
    } else {
      // Legacy single-POST fallback (no files)
      return handleLegacy(payload);
    }
  } catch(err) {
    Logger.log('FATAL: ' + err.toString());
    return jsonOut({ status:'error', message: err.toString() });
  }
}

function doGet(e) {
  return jsonOut({ status:'ok', message:'YiW Script is live.' });
}

// ── CHUNK HANDLER ─────────────────────────────────────────────
// Stores each chunk as a small file in a temp Drive folder.
// When the last chunk arrives, reassembles into the final file.
function handleChunk(p) {
  var reportId    = p.reportId;
  var category    = p.category;
  var fileName    = p.fileName;
  var fileType    = p.fileType || 'application/octet-stream';
  var chunkIndex  = parseInt(p.chunkIndex);
  var totalChunks = parseInt(p.totalChunks);
  var chunkData   = p.chunkData; // base64 string

  // Get/create a temp folder for this report's in-progress uploads
  var tempFolder = getOrCreateFolder('_YiW_Temp_' + reportId);

  // Save this chunk as its own small file
  var chunkName = fileName + '.chunk.' + chunkIndex;
  // Delete any previous attempt at this chunk
  var existing = tempFolder.getFilesByName(chunkName);
  while (existing.hasNext()) existing.next().setTrashed(true);
  // Save chunk
  var decoded = Utilities.base64Decode(chunkData);
  var blob    = Utilities.newBlob(decoded, 'application/octet-stream', chunkName);
  tempFolder.createFile(blob);

  // Check if all chunks are present
  var allPresent = true;
  for (var i = 0; i < totalChunks; i++) {
    var iter = tempFolder.getFilesByName(fileName + '.chunk.' + i);
    if (!iter.hasNext()) { allPresent = false; break; }
  }

  // If all chunks are here, reassemble into final file
  if (allPresent) {
    reassembleFile(tempFolder, fileName, fileType, totalChunks, reportId, category);
  }

  return jsonOut({ status:'chunk_ok', chunk: chunkIndex, total: totalChunks });
}

function reassembleFile(tempFolder, fileName, fileType, totalChunks, reportId, category) {
  // Collect all chunk bytes in order
  var allBytes = [];
  for (var i = 0; i < totalChunks; i++) {
    var iter = tempFolder.getFilesByName(fileName + '.chunk.' + i);
    if (iter.hasNext()) {
      var bytes = iter.next().getBlob().getBytes();
      allBytes = allBytes.concat(bytes);
    }
  }

  // Write final file into the report folder
  var reportFolder = getReportFolder(reportId);
  var finalBlob    = Utilities.newBlob(allBytes, fileType, fileName);
  var driveFile    = reportFolder.createFile(finalBlob);
  driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // Tag the category as metadata in a small sidecar file
  var sidecarName = '.cat.' + fileName;
  var existing = tempFolder.getFilesByName(sidecarName);
  while (existing.hasNext()) existing.next().setTrashed(true);
  tempFolder.createFile(sidecarName, category, MimeType.PLAIN_TEXT);

  // Clean up chunk files
  for (var j = 0; j < totalChunks; j++) {
    var ci = tempFolder.getFilesByName(fileName + '.chunk.' + j);
    if (ci.hasNext()) ci.next().setTrashed(true);
  }

  Logger.log('Reassembled: ' + fileName + ' (' + allBytes.length + ' bytes) → ' + category);
}

// ── FINALISE HANDLER ─────────────────────────────────────────
// Called after all file chunks are uploaded.
// Reads completed files from report folder, builds email with Drive links.
function handleFinalise(p) {
  var reportId = p.reportId;
  var d        = p.formData;

  var reportFolder = getReportFolder(reportId);
  var tempFolder   = getTempFolder(reportId);

  // Build file links table from whatever is in the report folder
  var fileLinksHtml = buildFileLinksHtml(reportFolder);

  // Clean up temp folder
  if (tempFolder) {
    try { tempFolder.setTrashed(true); } catch(e) {}
  }

  // Send email
  var htmlBody = buildEmailHtml(d, fileLinksHtml, reportFolder.getUrl());
  var subject  = 'YiW Field Report: ' + (d.fpName||'—') +
                 ' — ' + (d.trainingCentre||d.hubName||'—') +
                 ' (' + (d.visitDate||'—') + ')';

  MailApp.sendEmail({
    to:       TO_EMAIL,
    cc:       CC_EMAIL,
    subject:  subject,
    htmlBody: htmlBody
  });

  Logger.log('Email sent: ' + subject);
  return jsonOut({ status:'success', message:'Report submitted and emailed.' });
}

// ── LEGACY HANDLER (no files, just email) ────────────────────
function handleLegacy(payload) {
  var d = payload.formData;
  var noFilesHtml = "<p style='color:#718096;font-style:italic;font-size:13px'>No files were attached.</p>";

  var rootFolder   = getOrCreateFolder(ROOT_FOLDER_NAME);
  var folderName   = (d.visitDate||'no-date') + ' — ' + (d.fpName||'unknown');
  var reportFolder = rootFolder.createFolder(folderName);

  var htmlBody = buildEmailHtml(d, noFilesHtml, reportFolder.getUrl());
  var subject  = 'YiW Field Report: ' + (d.fpName||'—') +
                 ' — ' + (d.trainingCentre||d.hubName||'—') +
                 ' (' + (d.visitDate||'—') + ')';

  MailApp.sendEmail({ to:TO_EMAIL, cc:CC_EMAIL, subject:subject, htmlBody:htmlBody });
  return jsonOut({ status:'success', message:'Report emailed (no files).' });
}

// ── DRIVE HELPERS ────────────────────────────────────────────
function getOrCreateFolder(name) {
  var iter = DriveApp.getFoldersByName(name);
  return iter.hasNext() ? iter.next() : DriveApp.createFolder(name);
}

function getReportFolder(reportId) {
  var root     = getOrCreateFolder(ROOT_FOLDER_NAME);
  var iterExist = root.getFoldersByName(reportId);
  if (iterExist.hasNext()) return iterExist.next();
  var f = root.createFolder(reportId);
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return f;
}

function getTempFolder(reportId) {
  var iter = DriveApp.getFoldersByName('_YiW_Temp_' + reportId);
  return iter.hasNext() ? iter.next() : null;
}

function buildFileLinksHtml(folder) {
  var files = folder.getFiles();
  var rows  = '';
  var count = 0;

  // Category labels
  var catLabels = {
    dAtt:'Attendance Sheet', dFin:'Financial Document',
    dMou:'MoU / Agreement',  dTrack:'Tracking Sheet',
    mPhoto:'Photo',          mVideo:'Video'
  };

  while (files.hasNext()) {
    var f = files.next();
    var name = f.getName();
    if (name.startsWith('.cat.')) continue; // skip sidecar files
    var url = f.getUrl();
    var cat = 'Document';
    // Try to find category from sidecar
    var sidecar = folder.getFilesByName('.cat.' + name);
    if (sidecar.hasNext()) {
      var raw = sidecar.next().getBlob().getDataAsString();
      cat = catLabels[raw] || raw;
    }
    rows += '<tr>' +
      '<td style="padding:8px;border:1px solid #cbd5e1;font-weight:600;color:#1a5c2a;font-size:12px">' + cat + '</td>' +
      '<td style="padding:8px;border:1px solid #cbd5e1;font-size:12px"><a href="' + url + '" style="color:#1565c0;font-weight:500">🔗 ' + name + '</a></td>' +
    '</tr>';
    count++;
  }

  if (count === 0) {
    return '<p style="color:#718096;font-style:italic;font-size:13px">No files were attached during this submission.</p>';
  }

  return '<table style="width:100%;border-collapse:collapse;font-size:13px;margin-top:6px">' +
    '<tr style="background:#e8f5eb"><th style="padding:8px;border:1px solid #cbd5e1;text-align:left">Category</th>' +
    '<th style="padding:8px;border:1px solid #cbd5e1;text-align:left">File — click to open</th></tr>' +
    rows + '</table>' +
    '<div style="margin-top:10px;background:#e3f2fd;padding:10px;border-radius:6px;font-size:12px">' +
    '<strong>All files:</strong> <a href="' + folder.getUrl() + '" style="color:#1565c0;font-weight:600">Browse all files for this visit →</a></div>';
}

// ── JSON OUTPUT ──────────────────────────────────────────────
function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── BUILD HTML EMAIL ─────────────────────────────────────────
function buildEmailHtml(d, fileLinksHtml, driveFolderUrl) {

  function safe(val) { return val || '—'; }

  function pills(arr, bg, color) {
    if (!arr || arr.length === 0) return '<em style="color:#718096">None</em>';
    return arr.map(function(item) {
      return '<span style="display:inline-block;background:' + bg + ';color:' + color +
             ';padding:3px 9px;border-radius:12px;margin:2px;font-size:12px;font-weight:600">' + item + '</span>';
    }).join('');
  }

  function statBox(val, label, bg, color) {
    return '<td style="text-align:center;background:' + bg + ';padding:10px 6px;border-radius:6px;width:20%">' +
           '<div style="font-size:20px;font-weight:800;color:' + color + '">' + (val||0) + '</div>' +
           '<div style="font-size:11px;color:#718096;margin-top:2px">' + label + '</div></td>';
  }

  var rating    = parseInt(d.rating)||0;
  var ratingStr = rating > 0 ? ('★'.repeat(rating)+'☆'.repeat(5-rating)+' ('+rating+'/5)') : 'Not rated';
  var ts        = new Date().toLocaleString('en-GB', { timeZone:'Africa/Accra' });

  var partnerRows = '';
  if (d.partners && d.partners.length > 0) {
    d.partners.forEach(function(p) {
      partnerRows +=
        '<tr>' +
        '<td style="padding:7px 8px;border-bottom:1px solid #dde3ea;font-weight:600;font-size:12px">' + safe(p.name)        + '</td>' +
        '<td style="padding:7px 8px;border-bottom:1px solid #dde3ea;font-size:12px">'                 + safe(p.location)    + '</td>' +
        '<td style="padding:7px 8px;border-bottom:1px solid #dde3ea;font-size:12px">'                 + safe(p.sector)      + '</td>' +
        '<td style="padding:7px 8px;border-bottom:1px solid #dde3ea;font-size:11px;color:#555">'      + safe(p.profile)     + '</td>' +
        '<td style="padding:7px 8px;border-bottom:1px solid #dde3ea;font-size:11px;color:#555">'      + safe(p.skillsNeeded)+ '</td>' +
        '<td style="padding:7px 8px;border-bottom:1px solid #dde3ea;font-size:12px">'                 + safe(p.contact)     + '</td>' +
        '<td style="padding:7px 8px;border-bottom:1px solid #dde3ea;font-size:12px">'                 + safe(p.phone)       + '</td>' +
        '<td style="padding:7px 8px;border-bottom:1px solid #dde3ea">' +
          '<span style="background:#fff3e0;color:#e65100;padding:2px 7px;border-radius:10px;font-size:11px;font-weight:600">' + safe(p.status) + '</span>' +
        '</td>' +
        '<td style="padding:7px 8px;border-bottom:1px solid #dde3ea;text-align:center;font-weight:700;color:#1a5c2a">' + (p.slots||0) + '</td>' +
        '</tr>';
    });
  } else {
    partnerRows = '<tr><td colspan="9" style="padding:12px;text-align:center;color:#718096;font-style:italic">No partner companies logged.</td></tr>';
  }

  var safeChecked = d.safeChecked || [];
  var safeHtml    = safeChecked.length > 0
    ? safeChecked.map(function(s){ return '<li style="margin-bottom:3px">'+s+'</li>'; }).join('')
    : '<li style="color:#718096">No safeguarding items confirmed.</li>';

  var concernHtml = d.safeConcern === 'yes'
    ? '<div style="background:#ffebee;border-left:4px solid #c62828;padding:12px;border-radius:6px;margin-top:10px;color:#c62828">' +
      '<strong>⚠ SAFEGUARDING CONCERN RAISED</strong><br/>' +
      '<strong>Details:</strong> ' + safe(d.safeTxt) + '<br/>' +
      '<strong>Action taken:</strong> ' + safe(d.safeAct) + '<br/>' +
      '<strong>Reported to:</strong> ' + safe(d.safeRep) + '</div>'
    : '<p style="color:#1a5c2a;font-size:13px;margin-top:6px">✓ No safeguarding concerns identified.</p>';

  return '<!DOCTYPE html><html><head><meta charset="UTF-8"/></head>' +
  '<body style="font-family:\'Segoe UI\',Arial,sans-serif;background:#f3f7f4;margin:0;padding:20px">' +
  '<div style="max-width:700px;margin:0 auto;background:#fff;border-radius:12px;border:1px solid #dde3ea;overflow:hidden;box-shadow:0 4px 16px rgba(0,0,0,.07)">' +

  // HEADER
  '<div style="background:linear-gradient(135deg,#1a5c2a,#2d7a3a);padding:22px 26px;color:#fff">' +
    '<div style="font-size:10px;opacity:.7;text-transform:uppercase;letter-spacing:.1em;margin-bottom:3px">SEG Ghana</div>' +
    '<div style="font-size:20px;font-weight:700;margin-bottom:2px">Youth in Work Programme</div>' +
    '<div style="font-size:13px;opacity:.85;margin-bottom:14px">Daily Field Report</div>' +
    '<div style="display:flex;flex-wrap:wrap;gap:16px">' +
      '<div><div style="font-size:10px;opacity:.65;text-transform:uppercase">Submitted by</div><div style="font-size:14px;font-weight:700;margin-top:1px">'+safe(d.fpName)+'</div></div>' +
      '<div><div style="font-size:10px;opacity:.65;text-transform:uppercase">Phone</div><div style="font-size:14px;font-weight:700;margin-top:1px">'+safe(d.fpPhone)+'</div></div>' +
      '<div><div style="font-size:10px;opacity:.65;text-transform:uppercase">Date</div><div style="font-size:14px;font-weight:700;margin-top:1px">'+safe(d.visitDate)+'</div></div>' +
      '<div><div style="font-size:10px;opacity:.65;text-transform:uppercase">Zone</div><div style="font-size:14px;font-weight:700;margin-top:1px">'+safe(d.fpZone)+'</div></div>' +
    '</div>' +
  '</div>' +

  '<div style="padding:20px">' +

  // VISIT DETAILS
  '<div style="border-radius:9px;border:1px solid #dde3ea;border-left:4px solid #1a5c2a;padding:15px;margin-bottom:13px">' +
    '<div style="font-size:11px;font-weight:700;color:#1a5c2a;text-transform:uppercase;letter-spacing:.06em;margin-bottom:9px">📍 Visit Details</div>' +
    '<table style="width:100%;border-collapse:collapse;font-size:13px">' +
      '<tr><td style="padding:4px 0;color:#718096;width:38%">Visit type</td><td style="padding:4px 0;font-weight:600">'+safe(d.visitType)+'</td></tr>' +
      '<tr><td style="padding:4px 0;color:#718096">Hub / TSP</td><td style="padding:4px 0;font-weight:600">'+safe(d.hubName)+'</td></tr>' +
      '<tr><td style="padding:4px 0;color:#718096">Community</td><td style="padding:4px 0;font-weight:600">'+safe(d.community)+'</td></tr>' +
      '<tr><td style="padding:4px 0;color:#718096">Training centre</td><td style="padding:4px 0;font-weight:600">'+safe(d.trainingCentre)+'</td></tr>' +
      '<tr><td style="padding:4px 0;color:#718096">Address</td><td style="padding:4px 0;font-weight:600">'+safe(d.centreAddress)+'</td></tr>' +
      '<tr><td style="padding:4px 0;color:#718096">Centre contact</td><td style="padding:4px 0;font-weight:600">'+safe(d.hubContact)+(d.hubContactPhone?' · '+d.hubContactPhone:'')+'</td></tr>' +
      '<tr><td style="padding:4px 0;color:#718096">Time on site</td><td style="padding:4px 0;font-weight:600">'+safe(d.tArr)+' → '+safe(d.tDep)+'</td></tr>' +
    '</table>' +
  '</div>' +

  // ATTENDANCE
  '<div style="border-radius:9px;border:1px solid #dde3ea;border-left:4px solid #b8860b;padding:15px;margin-bottom:13px">' +
    '<div style="font-size:11px;font-weight:700;color:#b8860b;text-transform:uppercase;letter-spacing:.06em;margin-bottom:9px">👥 Attendance Count</div>' +
    '<table style="width:100%;border-collapse:separate;border-spacing:4px"><tr>' +
      statBox(d.cMale,   'Young men',  '#fff8e1','#b8860b')+
      statBox(d.cFemale, 'Young women','#fff8e1','#b8860b')+
      statBox(d.cPWD,    'PWD',        '#fff8e1','#b8860b')+
      statBox(d.cStaff,  'Staff',      '#fff8e1','#b8860b')+
      statBox(d.cTrainer,'Trainers',   '#fff8e1','#b8860b')+
    '</tr></table>' +
  '</div>' +

  // ACTIVATION
  '<div style="border-radius:9px;border:1px solid #dde3ea;border-left:4px solid #1a5c2a;padding:15px;margin-bottom:13px">' +
    '<div style="font-size:11px;font-weight:700;color:#1a5c2a;text-transform:uppercase;letter-spacing:.06em;margin-bottom:9px">🚀 Activation & Employment</div>' +
    '<table style="width:100%;border-collapse:separate;border-spacing:4px"><tr>' +
      statBox(d.aJobs,  'Formal jobs', '#e8f5eb','#1a5c2a')+
      statBox(d.aIntern,'Internships', '#e8f5eb','#1a5c2a')+
      statBox(d.aCoop,  'Cooperatives','#e8f5eb','#1a5c2a')+
      statBox(d.aRef,   'Further trng','#e8f5eb','#1a5c2a')+
      '<td></td>' +
    '</tr></table>' +
    '<table style="width:100%;border-collapse:collapse;font-size:13px;margin-top:10px">' +
      '<tr><td style="padding:4px 0;color:#718096;width:38%">New enrolments (M/F)</td><td style="padding:4px 0;font-weight:600">'+(d.enrolM||0)+' / '+(d.enrolF||0)+'</td></tr>' +
      '<tr><td style="padding:4px 0;color:#718096">Course enrolled in</td><td style="padding:4px 0;font-weight:600">'+safe(d.enrolCourse)+'</td></tr>' +
      '<tr><td style="padding:4px 0;color:#718096">Employer / cooperative</td><td style="padding:4px 0;font-weight:600">'+safe(d.empName)+(d.empSector?' ('+d.empSector+')':'')+'</td></tr>' +
    '</table>' +
    (d.youthNames?'<div style="margin-top:8px;padding:9px;background:#f8fafc;border-radius:7px;font-size:12px;white-space:pre-wrap"><strong style="color:#1a5c2a">Youth placed:</strong><br/>'+d.youthNames+'</div>':'') +
    (d.highlight ?'<div style="margin-top:7px;padding:9px;background:#e8f5eb;border-radius:7px;border-left:3px solid #4caf50;font-size:13px;color:#1a5c2a"><strong>✨ Success story:</strong><br/>'+d.highlight+'</div>':'') +
    (d.yVoice    ?'<div style="margin-top:6px;padding:9px;background:#f8fafc;border-radius:7px;border-left:3px solid #90caf9;font-size:13px;color:#4a5568;font-style:italic">"'+d.yVoice+'"</div>':'') +
  '</div>' +

  // HUB QUALITY
  '<div style="border-radius:9px;border:1px solid #dde3ea;border-left:4px solid #00695c;padding:15px;margin-bottom:13px">' +
    '<div style="font-size:11px;font-weight:700;color:#00695c;text-transform:uppercase;letter-spacing:.06em;margin-bottom:9px">⭐ Training Centre Quality</div>' +
    '<table style="width:100%;border-collapse:collapse;font-size:13px;margin-bottom:9px">' +
      '<tr><td style="padding:4px 0;color:#718096;width:38%">Overall rating</td><td style="padding:4px 0;font-weight:700;color:#1a5c2a">'+ratingStr+'</td></tr>' +
      '<tr><td style="padding:4px 0;color:#718096">Urgency</td><td style="padding:4px 0;font-weight:600">'+safe(d.urgency)+'</td></tr>' +
      '<tr><td style="padding:4px 0;color:#718096">Follow-up by</td><td style="padding:4px 0;font-weight:600">'+safe(d.followUpBy)+'</td></tr>' +
    '</table>' +
    '<div style="margin-bottom:7px"><strong style="font-size:11px;color:#00695c;text-transform:uppercase">Quality ✓</strong><br/><div style="margin-top:4px">'+pills(d.quality,'#e8f5eb','#1a5c2a')+'</div></div>' +
    '<div style="margin-bottom:7px"><strong style="font-size:11px;color:#c62828;text-transform:uppercase">Issues ⚠</strong><br/><div style="margin-top:4px">'+pills(d.issues,'#ffebee','#c62828')+'</div></div>' +
    '<div style="margin-bottom:7px"><strong style="font-size:11px;color:#1565c0;text-transform:uppercase">Activities</strong><br/><div style="margin-top:4px">'+pills(d.activities,'#e3f2fd','#1565c0')+'</div></div>' +
    '<div style="margin-bottom:7px"><strong style="font-size:11px;color:#455a64;text-transform:uppercase">Facilities</strong><br/><div style="margin-top:4px">'+pills(d.facilities,'#eceff1','#455a64')+'</div></div>' +
    (d.challenges     ?'<div style="margin-top:9px;padding:9px;background:#fff3e0;border-radius:7px;font-size:12px;color:#e65100;border-left:3px solid #ff9800"><strong>Challenges:</strong> '+d.challenges+'</div>':'') +
    (d.recommendations?'<div style="margin-top:6px;padding:9px;background:#e3f2fd;border-radius:7px;font-size:12px;color:#1565c0;border-left:3px solid #90caf9"><strong>Recommendations:</strong> '+d.recommendations+'</div>':'') +
  '</div>' +

  // PARTNERS
  '<div style="border-radius:9px;border:1px solid #dde3ea;border-left:4px solid #1565c0;padding:15px;margin-bottom:13px">' +
    '<div style="font-size:11px;font-weight:700;color:#1565c0;text-transform:uppercase;letter-spacing:.06em;margin-bottom:9px">🤝 Partner Engagement — '+(d.partners?d.partners.length:0)+' company(ies)</div>' +
    '<div style="overflow-x:auto">' +
      '<table style="width:100%;border-collapse:collapse;font-size:12px">' +
        '<thead><tr style="background:#e3f2fd">' +
          '<th style="padding:7px 8px;border-bottom:2px solid #90caf9;text-align:left;color:#1565c0;font-size:10px;text-transform:uppercase">Company</th>' +
          '<th style="padding:7px 8px;border-bottom:2px solid #90caf9;text-align:left;color:#1565c0;font-size:10px;text-transform:uppercase">Location</th>' +
          '<th style="padding:7px 8px;border-bottom:2px solid #90caf9;text-align:left;color:#1565c0;font-size:10px;text-transform:uppercase">Sector</th>' +
          '<th style="padding:7px 8px;border-bottom:2px solid #90caf9;text-align:left;color:#1565c0;font-size:10px;text-transform:uppercase">Business profile</th>' +
          '<th style="padding:7px 8px;border-bottom:2px solid #90caf9;text-align:left;color:#1565c0;font-size:10px;text-transform:uppercase">Skills needed</th>' +
          '<th style="padding:7px 8px;border-bottom:2px solid #90caf9;text-align:left;color:#1565c0;font-size:10px;text-transform:uppercase">Contact</th>' +
          '<th style="padding:7px 8px;border-bottom:2px solid #90caf9;text-align:left;color:#1565c0;font-size:10px;text-transform:uppercase">Phone</th>' +
          '<th style="padding:7px 8px;border-bottom:2px solid #90caf9;text-align:left;color:#1565c0;font-size:10px;text-transform:uppercase">Status</th>' +
          '<th style="padding:7px 8px;border-bottom:2px solid #90caf9;text-align:center;color:#1565c0;font-size:10px;text-transform:uppercase">Slots</th>' +
        '</tr></thead>' +
        '<tbody>'+partnerRows+'</tbody>' +
      '</table>' +
    '</div>' +
    (d.partnerNotes?'<div style="margin-top:8px;padding:8px;background:#f8fafc;border-radius:7px;font-size:12px;color:#4a5568"><strong>Notes:</strong> '+d.partnerNotes+'</div>':'') +
    (d.nextPDate?'<div style="margin-top:5px;font-size:12px;color:#4a5568;padding:5px 8px"><strong>Next engagement:</strong> '+d.nextPDate+'</div>':'') +
  '</div>' +

  // DOCUMENTS & MEDIA
  '<div style="border-radius:9px;border:1px solid #dde3ea;border-left:4px solid #6a1b9a;padding:15px;margin-bottom:13px">' +
    '<div style="font-size:11px;font-weight:700;color:#6a1b9a;text-transform:uppercase;letter-spacing:.06em;margin-bottom:9px">📎 Documents & Media</div>' +
    fileLinksHtml +
    (d.docNotes    ?'<div style="margin-top:8px;padding:8px;background:#f8fafc;border-radius:7px;font-size:12px;color:#4a5568"><strong>Doc notes:</strong> '+d.docNotes+'</div>':'') +
    (d.photoCaption?'<div style="margin-top:5px;padding:8px;background:#f8fafc;border-radius:7px;font-size:12px;color:#4a5568"><strong>Photo caption:</strong> '+d.photoCaption+'</div>':'') +
    (d.videoCaption?'<div style="margin-top:5px;padding:8px;background:#f8fafc;border-radius:7px;font-size:12px;color:#4a5568"><strong>Video description:</strong> '+d.videoCaption+'</div>':'') +
    (d.mediaContext?'<div style="margin-top:5px;padding:8px;background:#f8fafc;border-radius:7px;font-size:12px;color:#4a5568"><strong>Media context:</strong> '+d.mediaContext+'</div>':'') +
  '</div>' +

  // SAFEGUARDING
  '<div style="border-radius:9px;border:1px solid #dde3ea;border-left:4px solid #00695c;padding:15px;margin-bottom:13px">' +
    '<div style="font-size:11px;font-weight:700;color:#00695c;text-transform:uppercase;letter-spacing:.06em;margin-bottom:9px">🛡 Safeguarding</div>' +
    '<div style="font-size:13px;color:#4a5568;margin-bottom:7px">'+safeChecked.length+' of 8 items confirmed</div>' +
    '<ul style="font-size:13px;padding-left:18px;margin:5px 0 8px">'+safeHtml+'</ul>' +
    concernHtml +
    (d.safeNotes?'<div style="margin-top:7px;padding:8px;background:#f8fafc;border-radius:7px;font-size:12px;color:#4a5568">'+d.safeNotes+'</div>':'') +
  '</div>' +

  // FINAL NOTES
  (d.finalNotes ?
    '<div style="border-radius:9px;border:1px solid #dde3ea;border-left:4px solid #455a64;padding:15px;margin-bottom:13px">' +
      '<div style="font-size:11px;font-weight:700;color:#455a64;text-transform:uppercase;letter-spacing:.06em;margin-bottom:7px">📝 Additional Notes</div>' +
      '<div style="font-size:13px;color:#4a5568;background:#f8fafc;padding:10px;border-radius:7px;border-left:3px solid #4caf50">'+d.finalNotes+'</div>' +
    '</div>' : '') +

  '</div>' + // end padding

  // FOOTER
  '<div style="background:#eceff1;text-align:center;padding:12px;font-size:11px;color:#718096">' +
    'Submitted via YiW Field Reporting System · ' + ts + '<br/>SEG Ghana | Youth in Work Programme' +
  '</div>' +

  '</div></body></html>';
}
