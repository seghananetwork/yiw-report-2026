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
  var params = (e && e.parameter) ? e.parameter : {};
  var action = params.action || '';

  if (action === 'dashboard')     return getDashboardData();
  if (action === 'hub')           return getHubData(params.hub || '');
  if (action === 'deleteRow')     return deleteTestRow(params.rowIndex || '');
  if (action === 'listRows')      return listAllRows();
  if (action === 'diagnose')      return diagnoseHeaders();

  return jsonOut({ status: 'ok', message: 'YiW Script is live.', version: 'v4-dashboard-2026-06-15' });
}

// Diagnostic — shows exactly what's in row 1 of your sheet right now,
// with each column's index, so we can see precisely what's misaligned.
function diagnoseHeaders() {
  try {
    var existing = DriveApp.getFilesByName(MASTER_SHEET_NAME);
    if (!existing.hasNext()) return jsonOut({ status:'error', message:'No master sheet found.' });
    var ss    = SpreadsheetApp.open(existing.next());
    var sheet = ss.getSheetByName('Field Reports');
    if (!sheet) return jsonOut({ status:'error', message:'Field Reports tab not found.' });

    var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var sampleRow = sheet.getLastRow() >= 2
      ? sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0]
      : [];

    var cols = [];
    for (var i = 0; i < headerRow.length; i++) {
      cols.push({
        index: i,
        header: String(headerRow[i]),
        sampleValue: sampleRow[i] !== undefined ? String(sampleRow[i]).substring(0,60) : ''
      });
    }

    var dynMap = getDynamicColMap(headerRow);

    return jsonOut({
      status:'success',
      totalColumns: headerRow.length,
      totalRows: sheet.getLastRow() - 1,
      columns: cols,
      resolvedMap: dynMap
    });
  } catch(err) {
    return jsonOut({ status:'error', message: err.toString() });
  }
}


// ══════════════════════════════════════════════════════════════
//  DASHBOARD — general overview data
// ══════════════════════════════════════════════════════════════

function getDashboardData() {
  try {
    var existing = DriveApp.getFilesByName(MASTER_SHEET_NAME);
    if (!existing.hasNext()) return jsonOut({ status: 'error', message: 'No data yet. Submit your first report to get started.' });

    var ss    = SpreadsheetApp.open(existing.next());
    var sheet = ss.getSheetByName('Field Reports');
    if (!sheet || sheet.getLastRow() < 2) return jsonOut({ status: 'error', message: 'No submissions yet.' });

    var data = sheet.getDataRange().getValues();
    var C = getDynamicColMap(data[0]); // read actual headers from row 1
    var now = new Date();
    var thirtyDaysAgo = new Date(now.getTime() - 30*24*60*60*1000);

    var totals = { reports:0, youth:0, men:0, women:0, pwd:0,
                   jobs:0, interns:0, coops:0, furtherTraining:0, activations:0,
                   partners:0, safetyConcerns:0 };

    var byHub={}, byZone={}, byDate={};
    var recentReports=[], urgentItems=[], lowRated=[], successStories=[];
    var hubList = []; // ordered list of hub names

    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      if (!row[C.fpName] && !row[C.hubName]) continue; // skip empty rows

      totals.reports++;
      totals.youth           += parseInt(row[C.totalYouth])        || 0;
      totals.men             += parseInt(row[C.youngMen])          || 0;
      totals.women           += parseInt(row[C.youngWomen])        || 0;
      totals.pwd             += parseInt(row[C.pwd])               || 0;
      totals.jobs            += parseInt(row[C.formalJobs])        || 0;
      totals.interns         += parseInt(row[C.internships])       || 0;
      totals.coops           += parseInt(row[C.coops])             || 0;
      totals.furtherTraining += parseInt(row[C.furtherTraining])   || 0;
      totals.activations     += parseInt(row[C.totalActivations])  || 0;
      totals.partners        += parseInt(row[C.partnersCount])     || 0;
      if (String(row[C.safetyConcern]).toLowerCase() === 'yes') totals.safetyConcerns++;

      var hub       = String(row[C.hubName]  || 'Unknown');
      var zone      = String(row[C.fpZone]   || 'Unknown');
      var fp        = String(row[C.fpName]   || '');
      var rating    = parseInt(row[C.rating]) || 0;
      var visitDate = row[C.visitDate] ? String(row[C.visitDate]).substring(0,10) : '';
      var urgency   = String(row[C.urgency]  || '');
      var issues    = String(row[C.issues]   || '');
      var story     = String(row[C.successStory] || '');
      var submittedAt = row[C.submittedAt];

      // Aggregate by hub
      if (!byHub[hub]) { byHub[hub] = { reports:0, youth:0, activations:0, partners:0, ratings:[], fps:{} }; hubList.push(hub); }
      byHub[hub].reports++;
      byHub[hub].youth       += parseInt(row[C.totalYouth])       || 0;
      byHub[hub].activations += parseInt(row[C.totalActivations]) || 0;
      byHub[hub].partners    += parseInt(row[C.partnersCount])    || 0;
      if (rating > 0) byHub[hub].ratings.push(rating);
      byHub[hub].fps[fp] = (byHub[hub].fps[fp] || 0) + 1;

      // Aggregate by zone
      if (!byZone[zone]) byZone[zone] = { reports:0, youth:0, activations:0 };
      byZone[zone].reports++;
      byZone[zone].youth       += parseInt(row[C.totalYouth])       || 0;
      byZone[zone].activations += parseInt(row[C.totalActivations]) || 0;

      // By date (last 30 days)
      if (submittedAt instanceof Date && submittedAt >= thirtyDaysAgo) {
        if (!byDate[visitDate]) byDate[visitDate] = 0;
        byDate[visitDate]++;
      }

      // Urgent items
      if (urgency.indexOf('Urgent') !== -1 || urgency.indexOf('48') !== -1) {
        urgentItems.push({ hub:hub, fp:fp, date:visitDate, urgency:urgency, issues:issues, rowIndex:r+1 });
      }

      // Low-rated hubs
      if (rating > 0 && rating <= 2) {
        lowRated.push({ hub:hub, fp:fp, date:visitDate, rating:rating });
      }

      // Success stories (max 5)
      if (story && story.length > 10 && successStories.length < 5) {
        successStories.push({ story:story, fp:fp, hub:hub, date:visitDate });
      }

      // Recent reports (last 10)
      if (recentReports.length < 10) {
        recentReports.push({
          fp:fp, hub:hub, zone:zone, date:visitDate,
          youth: parseInt(row[C.totalYouth])||0,
          activations: parseInt(row[C.totalActivations])||0,
          rating:rating, rowIndex:r+1
        });
      }
    }

    // Compute hub stats
    var hubStats = [];
    for (var h in byHub) {
      var hd = byHub[h];
      var ratings = hd.ratings;
      var avgRating = ratings.length > 0
        ? Math.round((ratings.reduce(function(a,b){return a+b;},0)/ratings.length)*10)/10
        : null;
      var fpList = Object.keys(hd.fps);
      hubStats.push({ hub:h, reports:hd.reports, youth:hd.youth,
                      activations:hd.activations, partners:hd.partners,
                      avgRating:avgRating, focalPersons:fpList.length });
    }
    hubStats.sort(function(a,b){ return b.reports - a.reports; });

    var zoneStats = [];
    for (var z in byZone) {
      zoneStats.push({ zone:z, reports:byZone[z].reports,
                       youth:byZone[z].youth, activations:byZone[z].activations });
    }
    zoneStats.sort(function(a,b){ return b.reports - a.reports; });

    return jsonOut({
      status:'success',
      generatedAt: new Date().toLocaleString('en-GB', {timeZone:'Africa/Accra'}),
      totals:totals, hubStats:hubStats, zoneStats:zoneStats,
      byDate:byDate, recentReports:recentReports,
      urgentItems:urgentItems, lowRated:lowRated, successStories:successStories,
      hubNames: hubList.filter(function(v,i,a){return a.indexOf(v)===i;}).sort()
    });

  } catch(err) {
    Logger.log('Dashboard error: ' + err.toString());
    return jsonOut({ status:'error', message:err.toString() });
  }
}


// ══════════════════════════════════════════════════════════════
//  HUB DRILL-DOWN — all metrics for one specific hub
// ══════════════════════════════════════════════════════════════

function getHubData(hubName) {
  try {
    if (!hubName) return jsonOut({ status:'error', message:'No hub specified.' });

    var existing = DriveApp.getFilesByName(MASTER_SHEET_NAME);
    if (!existing.hasNext()) return jsonOut({ status:'error', message:'No data yet.' });

    var ss    = SpreadsheetApp.open(existing.next());
    var sheet = ss.getSheetByName('Field Reports');
    if (!sheet || sheet.getLastRow() < 2) return jsonOut({ status:'error', message:'No data yet.' });

    var data = sheet.getDataRange().getValues();
    var C = getDynamicColMap(data[0]); // dynamic from actual headers

    var rows = [], totals = {
      reports:0, youth:0, men:0, women:0, pwd:0,
      jobs:0, interns:0, coops:0, furtherTraining:0, activations:0,
      partners:0, safetyConcerns:0
    };
    var byFP={}, byDate={}, ratingSum=0, ratingCount=0;
    var qualityCounts={}, issueCounts={}, activityCounts={};
    var partnerEngagements=[];

    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      var hub = String(row[C.hubName] || '');
      if (hub !== hubName) continue;

      totals.reports++;
      var youth      = parseInt(row[C.totalYouth])       || 0;
      var acts       = parseInt(row[C.totalActivations]) || 0;
      var partners   = parseInt(row[C.partnersCount])    || 0;
      var rating     = parseInt(row[C.rating])           || 0;
      var fp         = String(row[C.fpName]   || '');
      var visitDate  = row[C.visitDate] ? String(row[C.visitDate]).substring(0,10) : '';
      var quality    = String(row[C.qualityIndicators] || '');
      var issues     = String(row[C.issues] || '');
      var activities = String(row[C.activities] || '');
      var urgency    = String(row[C.urgency] || '');
      var concern    = String(row[C.safetyConcern] || '');
      var partnerNames = String(row[C.partnerNames] || '');
      var skillsReq  = String(row[C.skillsReq] || '');

      totals.youth           += parseInt(row[C.youngMen])||0 + parseInt(row[C.youngWomen])||0 + parseInt(row[C.pwd])||0;
      totals.men             += parseInt(row[C.youngMen])         || 0;
      totals.women           += parseInt(row[C.youngWomen])       || 0;
      totals.pwd             += parseInt(row[C.pwd])              || 0;
      totals.jobs            += parseInt(row[C.formalJobs])       || 0;
      totals.interns         += parseInt(row[C.internships])      || 0;
      totals.coops           += parseInt(row[C.coops])            || 0;
      totals.furtherTraining += parseInt(row[C.furtherTraining])  || 0;
      totals.activations     += acts;
      totals.partners        += partners;
      if (concern.toLowerCase() === 'yes') totals.safetyConcerns++;

      if (rating > 0) { ratingSum += rating; ratingCount++; }

      // By focal person
      if (!byFP[fp]) byFP[fp] = { reports:0, youth:0, activations:0 };
      byFP[fp].reports++;
      byFP[fp].youth       += youth;
      byFP[fp].activations += acts;

      // By date
      if (!byDate[visitDate]) byDate[visitDate] = 0;
      byDate[visitDate]++;

      // Quality tag counts
      if (quality) quality.split(';').forEach(function(q){ var t=q.trim(); if(t){qualityCounts[t]=(qualityCounts[t]||0)+1;} });
      if (issues)  issues.split(';').forEach(function(i){ var t=i.trim(); if(t){issueCounts[t]=(issueCounts[t]||0)+1;} });
      if (activities) activities.split(';').forEach(function(a){ var t=a.trim(); if(t){activityCounts[t]=(activityCounts[t]||0)+1;} });

      // Partner details
      if (partnerNames) {
        partnerNames.split(';').forEach(function(pn,idx){
          var name = pn.trim();
          if (name) {
            var skills = skillsReq.split(';')[idx] || '';
            partnerEngagements.push({ name:name, date:visitDate, fp:fp, skills:skills.trim() });
          }
        });
      }

      rows.push({
        rowIndex: r+1, date:visitDate, fp:fp,
        community: String(row[C.community]||''),
        trainingCentre: String(row[C.trainingCentre]||''),
        youth:youth, men:parseInt(row[C.youngMen])||0, women:parseInt(row[C.youngWomen])||0,
        pwd:parseInt(row[C.pwd])||0, staff:parseInt(row[C.staff])||0,
        jobs:parseInt(row[C.formalJobs])||0, interns:parseInt(row[C.internships])||0,
        coops:parseInt(row[C.coops])||0, activations:acts,
        partners:partners, rating:rating, urgency:urgency,
        issues:issues, challenges:String(row[C.challenges]||''),
        recommendations:String(row[C.recommendations]||''),
        highlight:String(row[C.successStory]||''),
        safetyConcern:concern
      });
    }

    // Sort rows by date
    rows.sort(function(a,b){ return (b.date||'').localeCompare(a.date||''); });

    var avgRating = ratingCount > 0 ? Math.round((ratingSum/ratingCount)*10)/10 : null;

    // FP summary
    var fpStats = [];
    for (var f in byFP) {
      fpStats.push({ fp:f, reports:byFP[f].reports, youth:byFP[f].youth, activations:byFP[f].activations });
    }
    fpStats.sort(function(a,b){ return b.reports - a.reports; });

    return jsonOut({
      status:'success',
      hub:hubName,
      generatedAt: new Date().toLocaleString('en-GB', {timeZone:'Africa/Accra'}),
      totals:totals, avgRating:avgRating,
      rows:rows, fpStats:fpStats, byDate:byDate,
      qualityCounts:qualityCounts, issueCounts:issueCounts, activityCounts:activityCounts,
      partnerEngagements:partnerEngagements
    });

  } catch(err) {
    Logger.log('Hub data error: ' + err.toString());
    return jsonOut({ status:'error', message:err.toString() });
  }
}


// ══════════════════════════════════════════════════════════════
//  TEST DATA MANAGEMENT — list and delete rows
// ══════════════════════════════════════════════════════════════

function listAllRows() {
  try {
    var existing = DriveApp.getFilesByName(MASTER_SHEET_NAME);
    if (!existing.hasNext()) return jsonOut({ status:'error', message:'No sheet found.' });
    var ss    = SpreadsheetApp.open(existing.next());
    var sheet = ss.getSheetByName('Field Reports');
    if (!sheet || sheet.getLastRow() < 2) return jsonOut({ status:'success', rows:[] });

    var data = sheet.getDataRange().getValues();
    var C = getDynamicColMap(data[0]); // dynamic from actual headers
    var rows = [];

    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      if (!row[C.fpName] && !row[C.hubName]) continue;
      rows.push({
        rowIndex: r+1,
        submittedAt: row[C.submittedAt] ? String(row[C.submittedAt]).substring(0,24) : '',
        fpName:    String(row[C.fpName]   || ''),
        fpZone:    String(row[C.fpZone]   || ''),
        hubName:   String(row[C.hubName]  || ''),
        visitDate: String(row[C.visitDate]|| '').substring(0,10),
        youth:     parseInt(row[C.totalYouth])||0,
        activations: parseInt(row[C.totalActivations])||0
      });
    }

    return jsonOut({ status:'success', rows:rows });
  } catch(err) {
    return jsonOut({ status:'error', message:err.toString() });
  }
}

function deleteTestRow(rowIndexStr) {
  try {
    var rowIndex = parseInt(rowIndexStr);
    if (!rowIndex || rowIndex < 2) return jsonOut({ status:'error', message:'Invalid row index.' });

    var existing = DriveApp.getFilesByName(MASTER_SHEET_NAME);
    if (!existing.hasNext()) return jsonOut({ status:'error', message:'No sheet found.' });
    var ss    = SpreadsheetApp.open(existing.next());
    var sheet = ss.getSheetByName('Field Reports');
    if (!sheet) return jsonOut({ status:'error', message:'Field Reports sheet not found.' });

    // Also delete from hub sheet if it exists
    var hubName = String(sheet.getRange(rowIndex, 8).getValue() || ''); // col 8 = hubName
    if (hubName) {
      var hubSheet = ss.getSheetByName(hubName);
      if (hubSheet) {
        // Find matching row in hub sheet by date + FP name
        var fpName    = String(sheet.getRange(rowIndex, 2).getValue() || '');
        var visitDate = String(sheet.getRange(rowIndex, 6).getValue() || '').substring(0,10);
        var hubData   = hubSheet.getDataRange().getValues();
        for (var h = hubData.length - 1; h >= 2; h--) {
          var hDate = String(hubData[h][5] || '').substring(0,10);
          var hFP   = String(hubData[h][1] || '');
          if (hDate === visitDate && hFP === fpName) {
            hubSheet.deleteRow(h+1);
            break;
          }
        }
      }
    }

    // Delete from master sheet
    sheet.deleteRow(rowIndex);
    SpreadsheetApp.flush();

    return jsonOut({ status:'success', message:'Row ' + rowIndex + ' deleted from master sheet and hub sheet.' });
  } catch(err) {
    Logger.log('Delete error: ' + err.toString());
    return jsonOut({ status:'error', message:err.toString() });
  }
}


// ══════════════════════════════════════════════════════════════
//  COLUMN MAP — single place defining master sheet column indices
// ══════════════════════════════════════════════════════════════

function getColMap() {
  // Static map matching current appendToMasterSheet column order (new submissions)
  return {
    submittedAt:0,        // Submitted At
    fpName:1,             // Field Personnel Name
    fpPhone:2,            // Phone
    fpZone:3,             // Zone
    visitDate:4,          // Visit Date
    visitType:5,          // Visit Type
    hubName:6,            // Hub / TSP
    community:7,          // Community
    trainingCentre:8,     // Training Centre
    tArr:9,               // Time Arrived
    tDep:10,              // Time Departed
    youngMen:11,          // Male
    youngWomen:12,        // Female
    pwd:13,               // PWD
    staff:14,             // Staff
    trainer:15,           // Number of Trainers
    totalYouth:16,        // Total Youth
    formalJobs:17,        // Number of Formal Jobs
    internships:18,       // Internships
    coops:19,             // Cooperatives
    furtherTraining:20,   // Further Training
    totalActivations:21,  // Total Activations
    enrolM:22,            // Enrolments (M)
    enrolF:23,            // Enrolments (F)
    enrolCourse:24,       // Course
    empName:25,           // Employer
    empSector:26,         // Sector
    rating:27,            // Hub Rating
    qualityIndicators:28, // Quality Indicators
    issues:29,            // Issues Flagged
    facilities:30,        // Facilities
    challenges:31,        // Challenges
    partnersCount:32,     // Partners Count
    totalFiles:33,        // Total Files
    attDocs:34,           // Attendance Docs
    finDocs:35,           // Financial Docs
    mous:36,              // MoUs
    trackSheets:37,       // Tracking Sheets
    photos:38,            // Photos
    videos:39,            // Videos
    safeConfirmed:40,     // Safeguarding Items
    safeDetails:41,       // Safeguarding Details
    safetyConcern:42,     // Concern Raised
    safeDetail:43,        // Concern Detail
    finalNotes:44         // Final Notes
  };
}

// Dynamic column map — reads actual header row from sheet
// Used to handle old rows written by previous script versions
function getDynamicColMap(headers) {
  var map = {
    submittedAt:-1, fpName:-1, fpPhone:-1, fpZone:-1,
    visitDate:-1, visitType:-1, hubName:-1, community:-1, trainingCentre:-1,
    tArr:-1, tDep:-1,
    youngMen:-1, youngWomen:-1, pwd:-1, staff:-1, trainer:-1, totalYouth:-1,
    formalJobs:-1, internships:-1, coops:-1, furtherTraining:-1, totalActivations:-1,
    enrolM:-1, enrolF:-1, enrolCourse:-1, empName:-1, empSector:-1,
    rating:-1, qualityIndicators:-1, issues:-1, facilities:-1, challenges:-1,
    partnersCount:-1, totalFiles:-1,
    submissionFolder:-1,
    attFolder:-1, finFolder:-1, mouFolder:-1, trackFolder:-1, photoFolder:-1, videoFolder:-1,
    // legacy count columns (for old rows)
    attDocs:-1, finDocs:-1, mous:-1, trackSheets:-1, photos:-1, videos:-1,
    safeConfirmed:-1, safeDetails:-1, safetyConcern:-1, safeDetail:-1, finalNotes:-1,
    // hub sheet extras
    activities:-1, partnerNames:-1, skillsReq:-1, successStory:-1
  };

  // Header text → map key lookup
  var lookup = {
    'submitted at':                  'submittedAt',
    'field personnel name':          'fpName',
    'fp name':                       'fpName',
    'phone':                         'fpPhone',
    'zone':                          'fpZone',
    'visit date':                    'visitDate',
    'visit type':                    'visitType',
    'hub / tsp':                     'hubName',
    'hub name':                      'hubName',
    'community':                     'community',
    'training centre':               'trainingCentre',
    'time arrived':                  'tArr',
    'time departed':                 'tDep',
    'male':                          'youngMen',
    'young men':                     'youngMen',
    'female':                        'youngWomen',
    'young women':                   'youngWomen',
    'pwd':                           'pwd',
    'staff':                         'staff',
    'number of trainers':            'trainer',
    'trainers':                      'trainer',
    'total youth':                   'totalYouth',
    'number of formal jobs':         'formalJobs',
    'formal jobs':                   'formalJobs',
    'internships':                   'internships',
    'cooperatives':                  'coops',
    'further training':              'furtherTraining',
    'total activations':             'totalActivations',
    'enrolments (m)':                'enrolM',
    'enrolments (f)':                'enrolF',
    'course':                        'enrolCourse',
    'employer':                      'empName',
    'sector':                        'empSector',
    'hub rating':                    'rating',
    'quality indicators':            'qualityIndicators',
    'issues flagged':                'issues',
    'facilities':                    'facilities',
    'challenges':                    'challenges',
    'partners count':                'partnersCount',
    'total files':                   'totalFiles',
    'submission folder':             'submissionFolder',
    'attendance sheets folder':      'attFolder',
    'financial documents folder':    'finFolder',
    'mous & agreements folder':      'mouFolder',
    'tracking sheets folder':        'trackFolder',
    'photos folder':                 'photoFolder',
    'videos folder':                 'videoFolder',
    // legacy count column names
    'attendance docs':               'attDocs',
    'financial docs':                'finDocs',
    'mous':                          'mous',
    'tracking sheets':               'trackSheets',
    'photos':                        'photos',
    'videos':                        'videos',
    'safeguarding items':            'safeConfirmed',
    'safeguarding confirmed':        'safeConfirmed',
    'safeguarding details':          'safeDetails',
    'concern raised':                'safetyConcern',
    'concern detail':                'safeDetail',
    'final notes':                   'finalNotes',
    'activities':                    'activities',
    'partner names':                 'partnerNames',
    'skills requested':              'skillsReq',
    'success story':                 'successStory'
  };

  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i]).toLowerCase().trim();
    if (lookup[h] !== undefined) {
      map[lookup[h]] = i;
    }
  }

  return map;
}


// ══════════════════════════════════════════════════════════════
//  WEEKLY DIGEST TRIGGER SETUP
//  Run setupWeeklyDigestTrigger() once to activate
// ══════════════════════════════════════════════════════════════

function setupWeeklyDigestTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'sendWeeklyDigest') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('sendWeeklyDigest')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(8)
    .create();
  Logger.log('Weekly digest trigger set: every Monday at 8am Ghana time.');
}

function sendWeeklyDigest() {
  try {
    var dashData = JSON.parse(getDashboardData().getContent());
    if (dashData.status !== 'success') { Logger.log('Digest: no data'); return; }

    var t = dashData.totals;
    var now = new Date();
    var weekAgo = new Date(now.getTime() - 7*24*60*60*1000);
    var dateStr = weekAgo.toLocaleDateString('en-GB') + ' - ' + now.toLocaleDateString('en-GB');

    // Filter recent reports to this week only
    var weekReports = (dashData.recentReports||[]).filter(function(r){
      var d = new Date(r.date);
      return d >= weekAgo;
    });

    // Count week-specific totals from byDate
    var weekReportCount = 0;
    var byDate = dashData.byDate || {};
    for (var d in byDate) {
      if (new Date(d) >= weekAgo) weekReportCount += byDate[d];
    }

    var urgents = (dashData.urgentItems||[]).filter(function(u){ return new Date(u.date)>=weekAgo; });
    var stories = (dashData.successStories||[]).filter(function(s){ return new Date(s.date)>=weekAgo; });

    function sc(val,lbl,bg,color){
      return '<td style="text-align:center;padding:5px"><div style="background:'+bg+';border-radius:8px;padding:12px 8px">'+
             '<div style="font-size:24px;font-weight:800;color:'+color+'">'+val+'</div>'+
             '<div style="font-size:11px;color:#718096;margin-top:2px">'+lbl+'</div></div></td>';
    }

    var urgentHtml = '';
    if (urgents.length > 0) {
      urgentHtml = '<div style="background:#ffebee;border-left:4px solid #c62828;border-radius:8px;padding:14px;margin-bottom:14px">'+
        '<div style="font-size:12px;font-weight:700;color:#c62828;text-transform:uppercase;margin-bottom:8px">'+urgents.length+' Urgent Item(s) Needing Action</div>';
      urgents.forEach(function(u){
        urgentHtml += '<div style="background:#fff;border-radius:6px;padding:9px;margin-bottom:6px;font-size:12px">'+
          '<strong>'+u.hub+' — '+u.fp+'</strong> ('+u.date+')<br/>'+
          '<span style="color:#c62828">'+u.urgency+'</span>'+(u.issues?'<br/><span style="color:#718096">'+u.issues+'</span>':'')+
          '</div>';
      });
      urgentHtml += '</div>';
    }

    var storiesHtml = '';
    if (stories.length > 0) {
      storiesHtml = '<div style="background:#e8f5eb;border-left:4px solid #1a5c2a;border-radius:8px;padding:14px;margin-bottom:14px">'+
        '<div style="font-size:12px;font-weight:700;color:#1a5c2a;text-transform:uppercase;margin-bottom:8px">Success Stories</div>';
      stories.slice(0,3).forEach(function(s){
        storiesHtml += '<div style="background:#fff;border-radius:6px;padding:9px;margin-bottom:6px;font-size:12px">'+
          '<strong>'+s.fp+' at '+s.hub+':</strong><br/>'+s.story+'</div>';
      });
      storiesHtml += '</div>';
    }

    var ts = new Date().toLocaleString('en-GB', {timeZone:'Africa/Accra'});

    var html = '<!DOCTYPE html><html><head><meta charset="UTF-8"/></head>'+
      '<body style="font-family:\'Segoe UI\',Arial,sans-serif;background:#f3f7f4;margin:0;padding:20px">'+
      '<div style="max-width:680px;margin:0 auto">'+
      '<div style="background:linear-gradient(135deg,#1a5c2a,#2d7a3a);border-radius:12px;padding:22px;color:#fff;margin-bottom:14px">'+
        '<div style="font-size:10px;opacity:.7;text-transform:uppercase;margin-bottom:3px">SEG Ghana - Weekly Digest</div>'+
        '<div style="font-size:20px;font-weight:700;margin-bottom:2px">Youth in Work Programme</div>'+
        '<div style="font-size:13px;opacity:.85">Week of '+dateStr+'</div>'+
      '</div>'+
      '<div style="background:#fff;border-radius:12px;border:1px solid #dde3ea;padding:16px;margin-bottom:14px">'+
        '<div style="font-size:11px;font-weight:700;color:#1a5c2a;text-transform:uppercase;margin-bottom:12px">Week at a Glance</div>'+
        '<table style="width:100%;border-collapse:collapse"><tr>'+
          sc(weekReportCount,'Reports this week','#e8f5eb','#1a5c2a')+
          sc(t.youth,'Total youth (all time)','#fff8e1','#b8860b')+
          sc(t.activations,'Total activations','#e3f2fd','#1565c0')+
          sc(t.partners,'Total partners','#e0f2f1','#00695c')+
        '</tr></table>'+
      '</div>'+
      urgentHtml + storiesHtml +
      '<div style="text-align:center;padding:14px;font-size:11px;color:#718096">'+
        'YiW Weekly Digest - '+ts+'<br/>SEG Ghana | Youth in Work Programme'+
      '</div></div></body></html>';

    MailApp.sendEmail({
      to: TO_EMAIL, cc: CC_EMAILS,
      subject: 'YiW Weekly Digest - Week of ' + dateStr + ' (' + weekReportCount + ' reports)',
      htmlBody: html
    });

    Logger.log('Weekly digest sent.');
  } catch(err) {
    Logger.log('Digest error: ' + err.toString());
  }
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
    'visitDate','hubName','community','trainingCentre','centreAddress',
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
  var arrayFields = ['quality','issues','facilities','activities','safeChecked','visitTypes'];
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

function sanitiseFolderLinks(raw) {
  var src = raw || {};
  var cats = ['dAtt','dFin','dMou','dTrack','mPhoto','mVideo'];
  var out = {};
  for (var i = 0; i < cats.length; i++) {
    var cat = cats[i];
    out[cat] = src[cat] ? String(src[cat]) : '';
  }
  return out;
}


// ══════════════════════════════════════════════════════════════
//  PRIMARY HANDLER
// ══════════════════════════════════════════════════════════════

function handleSubmitWithLinks(payload) {
  var d                    = sanitise(payload.formData);
  var driveLinks           = sanitiseDriveLinks(payload.driveLinks);
  var driveFolderLinks     = sanitiseFolderLinks(payload.driveFolderLinks);
  var submissionFolderLink = String(payload.submissionFolderLink || '');
  var totalFiles           = parseInt(payload.totalFiles, 10) || 0;

  // 1. Append to master sheet (creates workbook if needed), get share URL
  var sheetUrl = appendToMasterSheet(d, driveLinks, driveFolderLinks, submissionFolderLink, totalFiles);

  // 2. Append to the hub-specific tab in the same workbook
  appendToHubSheet(d, driveLinks, driveFolderLinks, submissionFolderLink);

  // 3. Build email
  var fileLinksHtml = buildFileLinksHtml(driveLinks, driveFolderLinks, submissionFolderLink);
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
  var emptyFolderLinks = sanitiseFolderLinks({});

  var sheetUrl = appendToMasterSheet(d, emptyLinks, emptyFolderLinks, '', 0);
  appendToHubSheet(d, emptyLinks, emptyFolderLinks, '');

  var fileLinksHtml = '<p style="color:#718096;font-style:italic;font-size:13px">No files attached for this submission.</p>';
  var htmlBody = buildEmailHtml(d, fileLinksHtml, sheetUrl);
  var subject  = 'YiW Field Report: ' + (d.fpName || '--') + ' (' + (d.visitDate || '--') + ')';

  MailApp.sendEmail({ to: TO_EMAIL, cc: CC_EMAILS, subject: subject, htmlBody: htmlBody });
  return jsonOut({ status: 'success', message: 'Report submitted.' });
}


// ══════════════════════════════════════════════════════════════
//  MASTER SHEET — one row per submission (Google Forms style)
// ══════════════════════════════════════════════════════════════

function appendToMasterSheet(d, driveLinks, driveFolderLinks, submissionFolderLink, totalFiles) {
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

    var partnerNames = [];
    var partnerSkills = [];
    for (var p = 0; p < d.partners.length; p++) {
      var pr = d.partners[p];
      partnerNames.push(pr.name + (pr.status ? ' (' + pr.status + ')' : ''));
      partnerSkills.push(pr.skillsNeeded);
    }

    var dataRow = [
      new Date(),                                    // 0  Submitted At
      d.fpName,                                      // 1  Field Personnel Name
      d.fpPhone,                                     // 2  Phone
      d.fpZone,                                      // 3  Zone
      d.visitDate,                                   // 4  Visit Date
      d.visitTypes.join('; '),                       // 5  Visit Type
      d.hubName,                                     // 6  Hub / TSP
      d.community,                                   // 7  Community
      d.trainingCentre,                              // 8  Training Centre
      d.tArr,                                        // 9  Time Arrived
      d.tDep,                                        // 10 Time Departed
      // Attendance
      d.cMale,                                       // 11 Male
      d.cFemale,                                     // 12 Female
      d.cPWD,                                        // 13 PWD
      d.cStaff,                                      // 14 Staff
      d.cTrainer,                                    // 15 Number of Trainers
      (d.cMale + d.cFemale + d.cPWD),               // 16 Total Youth
      // Activation
      d.aJobs,                                       // 17 Number of Formal Jobs
      d.aIntern,                                     // 18 Internships
      d.aCoop,                                       // 19 Cooperatives
      d.aRef,                                        // 20 Further Training
      (d.aJobs + d.aIntern + d.aCoop + d.aRef),     // 21 Total Activations
      d.enrolM,                                      // 22 Enrolments (M)
      d.enrolF,                                      // 23 Enrolments (F)
      d.enrolCourse,                                 // 24 Course
      d.empName,                                     // 25 Employer
      d.empSector,                                   // 26 Sector
      // Quality
      d.rating,                                      // 27 Hub Rating
      d.quality.join('; '),                          // 28 Quality Indicators
      d.issues.join('; '),                           // 29 Issues Flagged
      d.facilities.join('; '),                       // 30 Facilities
      d.challenges,                                  // 31 Challenges
      // Partners
      d.partners.length,                             // 32 Partners Count
      // Files — counts
      fc.total,                                      // 33 Total Files
      // Files — folder links (one per category)
      submissionFolderLink,                          // 34 Submission Folder
      driveFolderLinks.dAtt   || '',                 // 35 Attendance Sheets Folder
      driveFolderLinks.dFin   || '',                 // 36 Financial Documents Folder
      driveFolderLinks.dMou   || '',                 // 37 MoUs & Agreements Folder
      driveFolderLinks.dTrack || '',                 // 38 Tracking Sheets Folder
      driveFolderLinks.mPhoto || '',                 // 39 Photos Folder
      driveFolderLinks.mVideo || '',                 // 40 Videos Folder
      // Safeguarding
      d.safeChecked.length,                          // 41 Safeguarding Items
      d.safeChecked.join('; '),                      // 42 Safeguarding Details
      (d.safeConcern === 'yes' ? 'YES' : 'No'),      // 43 Concern Raised
      d.safeTxt,                                     // 44 Concern Detail
      d.finalNotes                                   // 45 Final Notes
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
    'Submitted At',
    'Field Personnel Name',
    'Phone',
    'Zone',
    'Visit Date',
    'Visit Type',
    'Hub / TSP',
    'Community',
    'Training Centre',
    'Time Arrived',
    'Time Departed',
    'Male',
    'Female',
    'PWD',
    'Staff',
    'Number of Trainers',
    'Total Youth',
    'Number of Formal Jobs',
    'Internships',
    'Cooperatives',
    'Further Training',
    'Total Activations',
    'Enrolments (M)',
    'Enrolments (F)',
    'Course',
    'Employer',
    'Sector',
    'Hub Rating',
    'Quality Indicators',
    'Issues Flagged',
    'Facilities',
    'Challenges',
    'Partners Count',
    'Total Files',
    'Submission Folder',
    'Attendance Sheets Folder',
    'Financial Documents Folder',
    'MoUs & Agreements Folder',
    'Tracking Sheets Folder',
    'Photos Folder',
    'Videos Folder',
    'Safeguarding Items',
    'Safeguarding Details',
    'Concern Raised',
    'Concern Detail',
    'Final Notes'
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

  // Column widths
  sheet.setColumnWidth(1,  165); // Submitted At
  sheet.setColumnWidth(2,  150); // Field Personnel Name
  sheet.setColumnWidth(3,  120); // Phone
  sheet.setColumnWidth(4,  120); // Zone
  sheet.setColumnWidth(5,  100); // Visit Date
  sheet.setColumnWidth(6,  200); // Visit Type (wider — may hold multiple)
  sheet.setColumnWidth(7,  240); // Hub / TSP
  sheet.setColumnWidth(8,  130); // Community
  sheet.setColumnWidth(9,  180); // Training Centre
  sheet.setColumnWidth(29, 220); // Quality Indicators
  sheet.setColumnWidth(30, 200); // Issues Flagged
  sheet.setColumnWidth(32, 260); // Challenges
  sheet.setColumnWidth(35, 280); // Submission Folder
  sheet.setColumnWidth(36, 260); // Attendance Sheets Folder
  sheet.setColumnWidth(37, 260); // Financial Documents Folder
  sheet.setColumnWidth(38, 260); // MoUs & Agreements Folder
  sheet.setColumnWidth(39, 260); // Tracking Sheets Folder
  sheet.setColumnWidth(40, 240); // Photos Folder
  sheet.setColumnWidth(41, 240); // Videos Folder
}

// Run this once manually to fix headers on an existing sheet without deleting data
// ══════════════════════════════════════════════════════════════
//  FULL SHEET REPAIR — archives broken old data, rebuilds clean
//  Run this ONCE manually from the Apps Script editor.
//  Old rows are NOT deleted — they're copied to a new tab called
//  "Archived (Pre-Repair)" so nothing is lost, just removed from
//  the dashboard's view.
// ══════════════════════════════════════════════════════════════

// ══════════════════════════════════════════════════════════════
//  RECOVER ARCHIVED DATA — re-aligns old shifted rows and
//  restores them into the clean Field Reports sheet.
//  Old rows had an extra "Email" column after Phone, shifting
//  everything from Zone onward by +1. This function detects
//  that pattern and corrects it before restoring.
//
//  Run this ONCE after repairMasterSheet(). It finds the most
//  recent "Archived (Pre-Repair)" tab automatically.
// ══════════════════════════════════════════════════════════════

// Diagnostic — dumps one raw row from the archive exactly as stored,
// with index numbers, so we can map old→new positions precisely.
function diagnoseArchiveRow() {
  var existing = DriveApp.getFilesByName(MASTER_SHEET_NAME);
  if (!existing.hasNext()) { Logger.log('No master sheet found.'); return; }
  var ss = SpreadsheetApp.open(existing.next());

  var sheets = ss.getSheets();
  var archiveSheet = null, archiveName = '';
  for (var i = 0; i < sheets.length; i++) {
    var n = sheets[i].getName();
    if (n.indexOf('Archived (Pre-Repair)') === 0) {
      if (!archiveSheet || n > archiveName) { archiveSheet = sheets[i]; archiveName = n; }
    }
  }
  if (!archiveSheet) { Logger.log('No archive tab found.'); return; }

  var data = archiveSheet.getDataRange().getValues();
  var headerRow = data[0];
  var sampleRow = data[1];

  var out = [];
  for (var i = 0; i < Math.max(headerRow.length, sampleRow.length); i++) {
    out.push(i + ': [' + String(headerRow[i]||'') + '] = ' + String(sampleRow[i]||''));
  }
  Logger.log('Archive: ' + archiveName + ' | Total columns: ' + headerRow.length);
  Logger.log(out.join('\n'));
}

function recoverArchivedData() {
  var existing = DriveApp.getFilesByName(MASTER_SHEET_NAME);
  if (!existing.hasNext()) { Logger.log('No master sheet found.'); return; }
  var ss = SpreadsheetApp.open(existing.next());

  // Find the most recent archive tab
  var sheets = ss.getSheets();
  var archiveSheet = null;
  var archiveName = '';
  for (var i = 0; i < sheets.length; i++) {
    var n = sheets[i].getName();
    if (n.indexOf('Archived (Pre-Repair)') === 0) {
      if (!archiveSheet || n > archiveName) { archiveSheet = sheets[i]; archiveName = n; }
    }
  }
  if (!archiveSheet) { Logger.log('No archive tab found.'); return; }

  var liveSheet = ss.getSheetByName('Field Reports');
  if (!liveSheet) { Logger.log('Field Reports tab not found — run repairMasterSheet() first.'); return; }

  var archData = archiveSheet.getDataRange().getValues();
  if (archData.length < 2) { Logger.log('Archive is empty, nothing to recover.'); return; }

  // Reset Field Reports to empty before re-inserting (avoid duplicating
  // if this function is run more than once)
  var liveLastRow = liveSheet.getLastRow();
  if (liveLastRow > 1) {
    liveSheet.getRange(2, 1, liveLastRow - 1, liveSheet.getLastColumn()).clearContent();
  }

  var emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  var recovered = 0;
  var skipped = 0;

  for (var r = 1; r < archData.length; r++) {
    var row = archData[r];
    if (!row[1] && !row[6]) { skipped++; continue; } // empty row, skip

    // CONFIRMED layout for old-schema rows (verified against real data with user):
    // Old columns 0-2:   Submitted At, FP Name, Phone           -> no shift
    // Old column  3:     Email (extra, dropped)
    // Old columns 4-9:   Zone, Visit Date, Visit Type, Hub, Community, Training Centre  -> shift -1 (i.e. read from old+1... see below)
    // Old column  10:    Hub Contact Name (extra, dropped)
    // Old column  11:    (was showing contact name in sample - extra, dropped)
    // Old column  12:    Hub Contact Phone (extra, dropped)
    // Old columns 13-14: Time Arrived, Time Departed            -> these map to new tArr/tDep
    // Old columns 15+:   Male, Female, PWD, Staff, Trainers...  -> continue sequentially from old col 15
    var val3 = String(row[3] || '').trim();
    var isOldSchema = emailPattern.test(val3);

    function get(newIdx) {
      if (!isOldSchema) return newIdx < row.length ? row[newIdx] : '';
      // Old-schema explicit position map (0-indexed, confirmed against real sample row)
      var oldSchemaMap = [
        0,  1,  2,           // submittedAt, fpName, phone (cols 0-2, unshifted)
        4,  5,  6,  7,  8,  9, // zone, visitDate, visitType, hubName, community, trainingCentre (old 4-9)
        13, 14,               // tArr, tDep (old 13-14, confirmed via user — actual time values)
        15, 16, 17, 18, 19,    // male, female, pwd, staff, trainer (old 15-19, continuing sequentially)
        20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35,
        36, 37, 38, 39, 40, 41, 42, 43, 44
      ];
      var oldIdx = oldSchemaMap[newIdx];
      return (oldIdx !== undefined && oldIdx < row.length) ? row[oldIdx] : '';
    }

    // Build the clean 45-column row. For old-schema rows, get() uses the
    // confirmed explicit position map. For already-correct rows, get()
    // just reads the value straight through.
    var newRow = [
      get(0),  get(1),  get(2),  get(3),  get(4),  get(5),  get(6),  get(7),
      get(8),  get(9),  get(10), get(11), get(12), get(13), get(14), get(15),
      get(16), get(17), get(18), get(19), get(20), get(21), get(22), get(23),
      get(24), get(25), get(26), get(27), get(28), get(29), get(30), get(31),
      get(32), get(33), get(34),
      // Financial Docs onward: old archive had duplicate header columns
      // (a second copy starting at old index 45). For old-schema rows,
      // prefer that later, more complete duplicate set.
      isOldSchema && row.length > 45 ? row[45] : get(35), // Financial Docs
      isOldSchema && row.length > 46 ? row[46] : get(36), // MoUs
      isOldSchema && row.length > 47 ? row[47] : get(37), // Tracking Sheets
      isOldSchema && row.length > 48 ? row[48] : get(38), // Photos
      isOldSchema && row.length > 49 ? row[49] : get(39), // Videos
      isOldSchema && row.length > 51 ? row[51] : get(40), // Safeguarding Items
      isOldSchema && row.length > 52 ? row[52] : get(41), // Safeguarding Details
      isOldSchema && row.length > 53 ? row[53] : get(42), // Concern Raised
      isOldSchema && row.length > 54 ? row[54] : get(43), // Concern Detail
      isOldSchema && row.length > 57 ? row[57] : get(44)  // Final Notes
    ];

    liveSheet.appendRow(newRow);
    recovered++;
  }

  SpreadsheetApp.flush();
  Logger.log('Recovery complete. Recovered: ' + recovered + ' rows. Skipped (empty): ' + skipped +
             '. Source archive: ' + archiveName);
  Logger.log('IMPORTANT: Please open the Field Reports sheet and spot-check a few rows to ' +
             'confirm columns line up correctly (Zone shows a real zone, Hub shows a real hub name, etc).');
}

function repairMasterSheet() {
  var existing = DriveApp.getFilesByName(MASTER_SHEET_NAME);
  if (!existing.hasNext()) { Logger.log('No master sheet found.'); return; }
  var ss    = SpreadsheetApp.open(existing.next());
  var sheet = ss.getSheetByName('Field Reports');
  if (!sheet) { Logger.log('Field Reports tab not found.'); return; }

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  // 1. Archive everything as-is (raw copy, including the messy duplicate columns)
  if (lastRow > 0) {
    var archiveName = 'Archived (Pre-Repair) ' + Utilities.formatDate(new Date(), 'GMT', 'yyyy-MM-dd HH-mm');
    var oldData = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    var archiveSheet = ss.insertSheet(archiveName);
    archiveSheet.getRange(1, 1, oldData.length, lastCol).setValues(oldData);
    Logger.log('Archived ' + (lastRow - 1) + ' old rows (plus header) to tab: ' + archiveName);
  }

  // 2. Clear the Field Reports sheet completely (remove all columns/rows/filters)
  var filter = sheet.getFilter();
  if (filter) filter.remove();
  sheet.clear();
  sheet.clearFormats();

  // 3. Write a single clean header row with NO duplicates
  formatMasterSheet(sheet);

  SpreadsheetApp.flush();
  Logger.log('Master sheet repaired. Old data archived. Field Reports tab is now clean with ' +
             sheet.getLastColumn() + ' columns and 0 data rows.');
  Logger.log('New submissions will populate this clean sheet going forward.');
}

// Keep the old name working too, but point it at the full repair now —
// a header-only reset doesn't fix shifted data, so we always do a full repair.
function resetMasterSheetHeaders() {
  repairMasterSheet();
}


// ══════════════════════════════════════════════════════════════
//  HUB-SPECIFIC SHEET — one tab per TSP/Hub
// ══════════════════════════════════════════════════════════════

function appendToHubSheet(d, driveLinks, driveFolderLinks, submissionFolderLink) {
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
      d.visitDate, d.visitTypes.join('; '), d.community, d.trainingCentre,
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
      // Files — total count + per-category folder links
      fc.total,
      submissionFolderLink          || '',
      driveFolderLinks.dAtt         || '',
      driveFolderLinks.dFin         || '',
      driveFolderLinks.dMou         || '',
      driveFolderLinks.dTrack       || '',
      driveFolderLinks.mPhoto       || '',
      driveFolderLinks.mVideo       || '',
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
    'Total Files',
    'Submission Folder',
    'Attendance Sheets Folder',
    'Financial Documents Folder',
    'MoUs & Agreements Folder',
    'Tracking Sheets Folder',
    'Photos Folder',
    'Videos Folder',
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
  sheet.setColumnWidth(7,  200); // Visit Type — wider for multi-select
  sheet.setColumnWidth(8,  120);
  sheet.setColumnWidth(9,  170);
  sheet.setColumnWidth(31, 220);
  sheet.setColumnWidth(32, 200);
  sheet.setColumnWidth(35, 260);
  sheet.setColumnWidth(36, 260);
  sheet.setColumnWidth(43, 280); // Submission Folder
  sheet.setColumnWidth(44, 260); // Attendance Sheets Folder
  sheet.setColumnWidth(45, 260); // Financial Documents Folder
  sheet.setColumnWidth(46, 260); // MoUs & Agreements Folder
  sheet.setColumnWidth(47, 260); // Tracking Sheets Folder
  sheet.setColumnWidth(48, 240); // Photos Folder
  sheet.setColumnWidth(49, 240); // Videos Folder
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

function buildFileLinksHtml(driveLinks, driveFolderLinks, submissionFolderLink) {
  var catOrder = ['dAtt','dFin','dMou','dTrack','mPhoto','mVideo'];
  var catNames = {
    dAtt: 'Attendance Sheets', dFin: 'Financial Documents',
    dMou: 'MoUs & Agreements',  dTrack: 'Tracking Sheets',
    mPhoto: 'Photos',           mVideo: 'Videos'
  };

  driveFolderLinks = driveFolderLinks || {};

  // Check if anything was uploaded
  var totalCount = 0;
  for (var i = 0; i < catOrder.length; i++) totalCount += (driveLinks[catOrder[i]] || []).length;

  if (totalCount === 0) {
    return '<p style="color:#718096;font-style:italic;font-size:13px">No files attached for this submission.</p>';
  }

  var html = '';

  // Submission folder banner
  if (submissionFolderLink) {
    html += '<div style="background:#e8f5eb;border:1px solid #4caf50;border-radius:8px;padding:10px 13px;margin-bottom:12px;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px">' +
      '<div style="font-size:12px;font-weight:700;color:#1a5c2a">📁 All files for this report</div>' +
      '<a href="' + submissionFolderLink + '" style="background:#1a5c2a;color:#fff;font-weight:700;font-size:12px;padding:6px 13px;border-radius:6px;text-decoration:none">Open Submission Folder</a>' +
      '</div>';
  }

  // Per-category section: folder link + individual file links
  html += '<table style="width:100%;border-collapse:collapse;font-size:13px;margin-top:4px">';
  html += '<tr style="background:#e8f5eb"><th style="padding:8px;border:1px solid #cbd5e1;text-align:left">Category</th>' +
    '<th style="padding:8px;border:1px solid #cbd5e1;text-align:left">Folder</th>' +
    '<th style="padding:8px;border:1px solid #cbd5e1;text-align:left">Files</th></tr>';

  for (var i = 0; i < catOrder.length; i++) {
    var cat = catOrder[i];
    var files = driveLinks[cat] || [];
    if (files.length === 0) continue;

    var folderLink = driveFolderLinks[cat] || '';
    var folderCell = folderLink
      ? '<a href="' + folderLink + '" style="color:#1565c0;font-weight:600;font-size:12px">📂 Open folder</a>'
      : '<span style="color:#718096;font-size:12px">—</span>';

    var fileLinks = '';
    for (var j = 0; j < files.length; j++) {
      fileLinks += '<div style="margin-bottom:3px"><a href="' + files[j].url + '" style="color:#1565c0;font-size:12px">' + files[j].name + '</a></div>';
    }

    html += '<tr>' +
      '<td style="padding:8px;border:1px solid #cbd5e1;font-weight:600;color:#1a5c2a;font-size:12px;vertical-align:top">' + catNames[cat] + ' (' + files.length + ')</td>' +
      '<td style="padding:8px;border:1px solid #cbd5e1;vertical-align:top">' + folderCell + '</td>' +
      '<td style="padding:8px;border:1px solid #cbd5e1;vertical-align:top">' + fileLinks + '</td>' +
      '</tr>';
  }

  html += '</table>';
  return html;
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
  html += '<tr><td style="padding:4px 0;color:#718096;width:38%">Visit type</td><td style="padding:4px 0;font-weight:600">' + (Array.isArray(d.visitTypes) && d.visitTypes.length ? d.visitTypes.join('; ') : '—') + '</td></tr>';
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
