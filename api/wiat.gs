// ===============================================
// WIAT-2 READING COMPREHENSION - GOOGLE APPS SCRIPT
// (Initials-only edition; PID removed end-to-end)
// ===============================================

// ---------- CONFIG ----------
const CONFIG = {
  RECORDINGS_FOLDER_NAME: 'WIAT-2 Recordings',
  DATA_BACKUP_FOLDER_NAME: 'WIAT-2 Data Backups',
  ITEM_IMAGES_FOLDER_NAME: 'WIAT-2 Stimuli'
};

// (Optional) Central sync to your main Spatial Cognition workbook
const CENTRAL_SYNC = {
  ENABLED: false, // set to true to enable mirroring into your master workbook
  SPREADSHEET_ID: 'PUT_MASTER_SPREADSHEET_ID_HERE', // your master file ID
  TASK_NAME: 'Reading Comprehension Task'
};

// ===============================================
// MAIN HANDLER
// ===============================================
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('No data received');
    }
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    switch (data.action) {
      // Trial-by-trial
      case 'session_start':
        return handleSessionStart(ss, data);

      case 'item_started':
        return handleItemStarted(ss, data);

      case 'item_completed':
        return handleItemCompleted(ss, data);

      case 'item_skipped':
        return handleItemSkipped(ss, data);

      case 'reading_time':
        return handleReadingTime(ss, data);

      case 'video_upload':
        return handleVideoUpload(data);

      case 'upload_blob':
        return handleBlobUpload(data);

      case 'session_complete':
        return handleSessionComplete(ss, data);

      case 'get_session':
        return getSessionData(ss, data.initials || data.pid || '');

      case 'save_backup':
        return saveBackupData(ss, data);

      // Single-payload summary mode (text-only frontend)
      case 'study_completed':
        return handleStudyCompleted(ss, data);

      default:
        logEvent(ss, data);
        return createResponse({ status: 'success' });
    }
  } catch (error) {
    console.error('❌ Error:', error);
    return createResponse({ status: 'error', message: error.toString() });
  }
}

// ===============================================
// SESSION MANAGEMENT (Initials as key)
// ===============================================
function handleSessionStart(ss, data) {
  const sessionsSheet = getOrCreateSheet(ss, 'Sessions', [
    'Initials', 'Education', 'Start Time', 'End Time', 'Duration (min)',
    'Items Completed', 'Total Score', 'Consecutive Zeros',
    'Status', 'Discontinued', 'Gate Items Failed', 'Admin Mode',
    'Recording', 'IP Address', 'User Agent', 'Notes'
  ]);

  const initials = (data.initials || data.pid || '').trim(); // fallback if old client sends pid
  if (!initials) return createResponse({ status: 'error', message: 'Missing initials' });

  const existingRow = findRowByInitials(sessionsSheet, initials);
  if (existingRow > 0) {
    sessionsSheet.getRange(existingRow, 9).setValue('Active'); // Status
    sessionsSheet.getRange(existingRow, 16).setValue('Session resumed at ' + (data.timestamp || new Date().toISOString()));
    wiat_central__touchSession(initials, data.timestamp);
    wiat_central__logEvent(initials, 'WIAT: Session Resumed', '', data.timestamp);

    return createResponse({
      status: 'success',
      message: 'Session resumed',
      existingData: getSessionDataFromRow(sessionsSheet, existingRow)
    });
  }

  sessionsSheet.appendRow([
    initials,
    data.education || '',
    data.timestamp || new Date().toISOString(),
    '', // End time
    0,  // Duration
    0,  // Items completed
    0,  // Total score
    0,  // Consecutive zeros
    'Active',
    'No', // Discontinued
    '',   // Gate items failed
    data.adminMode || false,
    data.hasRecording || false,
    data.ipAddress || '',
    data.userAgent || '',
    'Started'
  ]);

  const recordingsFolder = getOrCreateFolder(CONFIG.RECORDINGS_FOLDER_NAME);
  getOrCreateFolder(`${initials}_${(data.timestamp || new Date().toISOString()).split('T')[0]}`, recordingsFolder);

  logEvent(ss, { ...data, initials, eventType: 'Session Started' });
  wiat_central__touchSession(initials, data.timestamp);
  wiat_central__logEvent(initials, 'WIAT: Session Started', '', data.timestamp);

  return createResponse({ status: 'success', message: 'Session created' });
}

// ===============================================
// ITEM TRACKING
// ===============================================
function handleItemStarted(ss, data) {
  const initials = (data.initials || data.pid || '').trim();
  if (!initials) return createResponse({ status: 'error', message: 'Missing initials' });

  const itemsSheet = getOrCreateSheet(ss, 'Item_Responses', [
    'Timestamp', 'Initials', 'Item Number', 'Image File', 'Question Text',
    'Item Type', 'Start Time', 'End Time', 'Duration (sec)',
    'Response', 'Explanation', 'Auto Score', 'Score Confidence',
    'Needs Review', 'Scoring Notes', 'Final Score', 'Skip Reason'
  ]);

  itemsSheet.appendRow([
    new Date(),
    initials,
    data.itemNumber,
    data.imageFile || '',
    data.questionText || '',
    data.itemType || 'question',
    data.timestamp || new Date().toISOString(),
    '', // end
    0,  // dur
    '', '', '', '',
    '', '', '', ''
  ]);

  const progressSheet = getOrCreateSheet(ss, 'Item_Progress', [
    'Timestamp', 'Initials', 'Item', 'Event', 'Details'
  ]);
  progressSheet.appendRow([
    new Date(),
    initials,
    data.itemNumber,
    'Started',
    `Type: ${data.itemType || 'question'}, Image: ${data.imageFile || ''}`
  ]);

  updateSessionActivity(ss, initials, data.timestamp || new Date().toISOString());
  wiat_central__touchSession(initials, data.timestamp);
  wiat_central__logEvent(initials, 'WIAT: Item Started', 'Item ' + data.itemNumber, data.timestamp);

  return createResponse({ status: 'success' });
}

function handleItemCompleted(ss, data) {
  const initials = (data.initials || data.pid || '').trim();
  if (!initials) return createResponse({ status: 'error', message: 'Missing initials' });

  const itemsSheet = getOrCreateSheet(ss, 'Item_Responses', [
    'Timestamp', 'Initials', 'Item Number', 'Image File', 'Question Text',
    'Item Type', 'Start Time', 'End Time', 'Duration (sec)',
    'Response', 'Explanation', 'Auto Score', 'Score Confidence',
    'Needs Review', 'Scoring Notes', 'Final Score', 'Skip Reason'
  ]);

  const values = itemsSheet.getDataRange().getValues();
  let targetRow = -1;
  for (let i = values.length - 1; i >= 1; i--) {
    if (String(values[i][1]) === initials &&
        String(values[i][2]) === String(data.itemNumber) &&
        !values[i][7]) { // End Time blank
      targetRow = i + 1;
      break;
    }
  }

  const autoScore = data.autoScore !== undefined && data.autoScore !== '' ? Number(data.autoScore) : '';
  const finalScore = data.finalScore !== undefined && data.finalScore !== ''
    ? Number(data.finalScore)
    : (data.autoScore !== undefined && data.autoScore !== '' ? Number(data.autoScore) : 0);
  const needsReview = String(data.needsReview).toLowerCase() === 'true';

  if (targetRow > 0) {
    itemsSheet.getRange(targetRow, 8).setValue(data.endTime || new Date().toISOString()); // End
    itemsSheet.getRange(targetRow, 9).setValue(Number(data.duration) || 0); // Duration
    itemsSheet.getRange(targetRow,10).setValue(data.response || '');
    itemsSheet.getRange(targetRow,11).setValue(data.explanation || '');
    itemsSheet.getRange(targetRow,12).setValue(autoScore);
    itemsSheet.getRange(targetRow,13).setValue(data.scoreConfidence || '');
    itemsSheet.getRange(targetRow,14).setValue(needsReview ? 'YES' : 'NO');
    itemsSheet.getRange(targetRow,15).setValue(data.scoringNotes || '');
    itemsSheet.getRange(targetRow,16).setValue(finalScore);
  } else {
    itemsSheet.appendRow([
      new Date(),
      initials,
      data.itemNumber,
      data.imageFile || '',
      data.questionText || '',
      data.itemType || 'question',
      '',
      data.endTime || new Date().toISOString(),
      Number(data.duration) || 0,
      data.response || '',
      data.explanation || '',
      autoScore,
      data.scoreConfidence || '',
      needsReview ? 'YES' : 'NO',
      data.scoringNotes || '',
      finalScore,
      data.reason || ''
    ]);
  }

  const progressSheet = getOrCreateSheet(ss, 'Item_Progress', [
    'Timestamp', 'Initials', 'Item', 'Event', 'Details'
  ]);
  progressSheet.appendRow([
    new Date(),
    initials,
    data.itemNumber,
    'Completed',
    `Score: ${autoScore}, Confidence: ${data.scoreConfidence}, Review: ${needsReview ? 'YES' : 'NO'}`
  ]);

  updateSessionTotals(ss, initials, Number(finalScore) || 0, Number(data.consecutiveZeros) || 0);

  saveDetailedScoring(ss, { ...data, initials, autoScore: autoScore, needsReview: needsReview });

  wiat_central__touchSession(initials, data.endTime || data.timestamp);
  wiat_central__logEvent(
    initials,
    'WIAT: Item Completed',
    'Item ' + data.itemNumber + ' | autoScore=' + (data.autoScore !== undefined ? data.autoScore : ''),
    data.endTime || data.timestamp
  );

  return createResponse({ status: 'success' });
}

function handleItemSkipped(ss, data) {
  const initials = (data.initials || data.pid || '').trim();
  if (!initials) return createResponse({ status: 'error', message: 'Missing initials' });

  const itemsSheet = getOrCreateSheet(ss, 'Item_Responses', [
    'Timestamp', 'Initials', 'Item Number', 'Image File', 'Question Text',
    'Item Type', 'Start Time', 'End Time', 'Duration (sec)',
    'Response', 'Explanation', 'Auto Score', 'Score Confidence',
    'Needs Review', 'Scoring Notes', 'Final Score', 'Skip Reason'
  ]);

  const ts = data.timestamp || new Date().toISOString();
  itemsSheet.appendRow([
    new Date(),
    initials,
    data.itemNumber,
    data.imageFile || '',
    data.questionText || '',
    data.itemType || 'question',
    ts,
    ts,
    0,
    'SKIPPED',
    '',
    0,
    'N/A',
    'NO',
    'Item skipped',
    0,
    data.reason || 'User choice'
  ]);

  const progressSheet = getOrCreateSheet(ss, 'Item_Progress', [
    'Timestamp', 'Initials', 'Item', 'Event', 'Details'
  ]);
  progressSheet.appendRow([
    new Date(),
    initials,
    data.itemNumber,
    'Skipped',
    data.reason || 'User choice'
  ]);

  updateSessionTotals(ss, initials, 0, Number(data.consecutiveZeros) || 0);

  wiat_central__touchSession(initials, ts);
  wiat_central__logEvent(initials, 'WIAT: Item Skipped', 'Item ' + data.itemNumber + ' | ' + (data.reason || 'User choice'), ts);

  return createResponse({ status: 'success' });
}

// ===============================================
// READING TIME TRACKING
// ===============================================
function handleReadingTime(ss, data) {
  const initials = (data.initials || data.pid || '').trim();
  if (!initials) return createResponse({ status: 'error', message: 'Missing initials' });

  const readingSheet = getOrCreateSheet(ss, 'Reading_Times', [
    'Timestamp', 'Initials', 'Item', 'Image', 'Reading Type',
    'Start Time', 'End Time', 'Duration (sec)', 'Words Count'
  ]);

  readingSheet.appendRow([
    new Date(),
    initials,
    data.itemNumber,
    data.imageFile || '',
    data.readingType || 'silent',
    data.startTime || '',
    data.endTime || '',
    data.duration || '',
    data.wordCount || ''
  ]);

  return createResponse({ status: 'success' });
}

// ===============================================
// VIDEO / BLOB UPLOAD
// ===============================================
function handleVideoUpload(data) {
  try {
    const initials = (data.initials || data.pid || '').trim();
    if (!initials || !data.videoData) throw new Error('Missing required fields');

    const videoBytes = Utilities.base64Decode(data.videoData);
    const maxSize = 25 * 1024 * 1024; // 25MB
    if (videoBytes.length > maxSize) throw new Error(`Video too large (${Math.round(videoBytes.length / 1024 / 1024)}MB)`);

    const recordingsFolder = getOrCreateFolder(CONFIG.RECORDINGS_FOLDER_NAME);
    const participantFolder = getOrCreateFolder(
      `${initials}_${data.sessionDate || new Date().toISOString().split('T')[0]}`,
      recordingsFolder
    );

    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const filename = `${initials}_item${data.itemNumber || 'full'}_${timestamp}.mp4`;

    const blob = Utilities.newBlob(videoBytes, 'video/mp4', filename);
    const file = participantFolder.createFile(blob);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const videoSheet = getOrCreateSheet(ss, 'Video_Recordings', [
      'Timestamp', 'Initials', 'Item Number', 'Filename',
      'File ID', 'File URL', 'File Size (KB)', 'Upload Status'
    ]);

    videoSheet.appendRow([
      new Date(),
      initials,
      data.itemNumber || 'Full Session',
      filename,
      file.getId(),
      file.getUrl(),
      Math.round(videoBytes.length / 1024),
      'Success'
    ]);

    wiat_central__logVideo(initials, data.itemNumber, {
      filename: filename,
      id: file.getId(),
      url: file.getUrl(),
      sizeKb: Math.round(videoBytes.length / 1024)
    });

    return createResponse({
      status: 'success',
      fileId: file.getId(),
      fileUrl: file.getUrl(),
      filename: filename
    });
  } catch (error) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const errorSheet = getOrCreateSheet(ss, 'Upload_Errors', [
      'Timestamp', 'Initials', 'Item', 'Error', 'Type'
    ]);
    errorSheet.appendRow([
      new Date(),
      (data && (data.initials || data.pid)) || 'unknown',
      (data && data.itemNumber) || '',
      error.toString(),
      'Video Upload'
    ]);
    return createResponse({ status: 'error', message: error.toString() });
  }
}

function handleBlobUpload(data) {
  try {
    const initials = (data.initials || data.pid || '').trim();
    if (!initials || !data.data) throw new Error('Missing required fields');

    const bytes = Utilities.base64Decode(data.data);
    const maxSize = 25 * 1024 * 1024; // 25MB
    if (bytes.length > maxSize) {
      throw new Error(`${data.kind || 'blob'} too large (${Math.round(bytes.length / 1024 / 1024)}MB)`);
    }

    const recordingsFolder = getOrCreateFolder(CONFIG.RECORDINGS_FOLDER_NAME);
    const participantFolder = getOrCreateFolder(
      `${initials}_${data.sessionDate || new Date().toISOString().split('T')[0]}`,
      recordingsFolder
    );

    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const extension = data.kind === 'video' ? '.mp4' : '.mp3';
    const mime = data.mime || (data.kind === 'video' ? 'video/mp4' : 'audio/mpeg');
    const filename = `${initials}_item${data.itemNumber || 'full'}_${timestamp}${extension}`;

    const blob = Utilities.newBlob(bytes, mime, filename);
    const file = participantFolder.createFile(blob);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = data.kind === 'video' ? 'Video_Recordings' : 'Audio_Recordings';
    const sheet = getOrCreateSheet(ss, sheetName, [
      'Timestamp', 'Initials', 'Item Number', 'Filename',
      'File ID', 'File URL', 'File Size (KB)', 'Upload Status'
    ]);

    sheet.appendRow([
      new Date(),
      initials,
      data.itemNumber || 'Full Session',
      filename,
      file.getId(),
      file.getUrl(),
      Math.round(file.getSize() / 1024),
      'Success'
    ]);

    if ((data.kind || '').toLowerCase() === 'video') {
      wiat_central__logVideo(initials, data.itemNumber, {
        filename: filename,
        id: file.getId(),
        url: file.getUrl(),
        sizeKb: Math.round(file.getSize() / 1024)
      });
    }

    return createResponse({
      status: 'success',
      fileId: file.getId(),
      fileUrl: file.getUrl(),
      filename: filename
    });
  } catch (error) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const errorSheet = getOrCreateSheet(ss, 'Upload_Errors', [
      'Timestamp', 'Initials', 'Item', 'Error', 'Type'
    ]);

    errorSheet.appendRow([
      new Date(),
      (data && (data.initials || data.pid)) || 'unknown',
      (data && data.itemNumber) || '',
      error.toString(),
      (data && data.kind === 'video') ? 'Video Upload' : 'Audio Upload'
    ]);

    return createResponse({ status: 'error', message: error.toString() });
  }
}

// ===============================================
// SESSION COMPLETION
// ===============================================
function handleSessionComplete(ss, data) {
  const initials = (data.initials || data.pid || '').trim();
  if (!initials) return createResponse({ status: 'error', message: 'Missing initials' });

  const sessionsSheet = getOrCreateSheet(ss, 'Sessions', [
    'Initials', 'Education', 'Start Time', 'End Time', 'Duration (min)',
    'Items Completed', 'Total Score', 'Consecutive Zeros',
    'Status', 'Discontinued', 'Gate Items Failed', 'Admin Mode',
    'Recording', 'IP Address', 'User Agent', 'Notes'
  ]);

  const row = findRowByInitials(sessionsSheet, initials);
  if (row > 0) {
    sessionsSheet.getRange(row, 4).setValue(data.timestamp || new Date());                 // End Time
    sessionsSheet.getRange(row, 5).setValue(Number(data.duration) || 0);                   // Duration (min)
    sessionsSheet.getRange(row, 6).setValue(Number(data.itemsCompleted) || 0);             // Items Completed
    sessionsSheet.getRange(row, 7).setValue(Number(data.totalScore) || 0);                 // Total Score
    sessionsSheet.getRange(row, 8).setValue(Number(data.consecutiveZeros) || 0);           // Consec Zeros
    sessionsSheet.getRange(row, 9).setValue('Complete');                                   // Status
    sessionsSheet.getRange(row,10).setValue(data.discontinued ? 'Yes' : 'No');             // Discontinued
    sessionsSheet.getRange(row,11).setValue(data.gateItemsFailed || '');                   // Gate Items Failed
  }

  saveBackupData(ss, { ...data, initials });
  generateParticipantSummary(ss, initials);
  logEvent(ss, { ...data, initials, eventType: 'Session Complete' });

  // Central: write one consolidated completion row
  const endTs = data.timestamp || new Date().toISOString();
  const elapsedSec = toSeconds_(data.duration); // minutes→seconds if it looks like minutes
  wiat_central__logTask(initials, 'Completed', {
    timestamp: endTs,
    endTime: endTs,
    elapsed: elapsedSec,
    active: elapsedSec,
    details: 'WIAT totalScore=' + (Number(data.totalScore) || 0)
  });
  wiat_central__touchSession(initials, endTs);
  wiat_central__logEvent(initials, 'WIAT: Session Completed',
    'items=' + (Number(data.itemsCompleted)||0) + ', score=' + (Number(data.totalScore)||0),
    endTs);

  return createResponse({ status: 'success', message: 'Session completed' });
}

// ===============================================
// SINGLE-PAYLOAD SUMMARY INGEST
// ===============================================
function handleStudyCompleted(ss, data) {
  const initials = (data.initials || data.pid || '').trim();
  if (!initials) return createResponse({ status: 'error', message: 'Missing initials' });

  const sessions = getOrCreateSheet(ss, 'Sessions', [
    'Initials','Education','Start Time','End Time','Duration (min)',
    'Items Completed','Total Score','Consecutive Zeros',
    'Status','Discontinued','Gate Items Failed','Admin Mode',
    'Recording','IP Address','User Agent','Notes'
  ]);

  const start = data.startedAt ? new Date(data.startedAt) : null;
  const end   = data.finishedAt ? new Date(data.finishedAt) : null;
  const durationMin = (start && end && !isNaN(start) && !isNaN(end)) ? Math.max(0, (end - start) / 1000 / 60) : 0;
  const itemsCompleted = Number((data.totals && data.totals.items) || (data.results ? data.results.length : 0));
  const totalScore = Number((data.totals && data.totals.points) || 0);

  const row = findRowByInitials(sessions, initials);
  if (row > 0) {
    sessions.getRange(row, 1, 1, 16).setValues([[
      initials,
      data.edu || '',
      start || new Date(),
      end || new Date(),
      durationMin,
      itemsCompleted,
      totalScore,
      0,
      'Complete',
      'No',
      '',
      false,
      (data.modality === 'sign' || data.modality === 'speak') ? true : false,
      '',
      '',
      'Summary ingest'
    ]]);
  } else {
    sessions.appendRow([
      initials,
      data.edu || '',
      start || new Date(),
      end || new Date(),
      durationMin,
      itemsCompleted,
      totalScore,
      0,
      'Complete',
      'No',
      '',
      false,
      (data.modality === 'sign' || data.modality === 'speak') ? true : false,
      '',
      '',
      'Summary ingest'
    ]);
  }

  // Write result rows
  const itemsSheet = getOrCreateSheet(ss, 'Item_Responses', [
    'Timestamp', 'Initials', 'Item Number', 'Image File', 'Question Text',
    'Item Type', 'Start Time', 'End Time', 'Duration (sec)',
    'Response', 'Explanation', 'Auto Score', 'Score Confidence',
    'Needs Review', 'Scoring Notes', 'Final Score', 'Skip Reason'
  ]);

  const now = new Date();
  (data.results || []).forEach(r => {
    if (r.type === 'qa') {
      if (r.answers && r.answers.length) {
        r.answers.forEach(a => {
          const review = (a.note || '').toLowerCase().includes('review');
          itemsSheet.appendRow([
            now,
            initials,
            r.item,
            '',
            a.key || '',
            'question',
            '', '',
            '',
            a.answer || '',
            '',
            Number(a.points) || 0,
            '',
            review ? 'YES' : 'NO',
            a.note || '',
            Number(a.points) || 0,
            r.skipped ? 'User choice' : ''
          ]);
        });
      } else {
        itemsSheet.appendRow([
          now, initials, r.item, '', '', 'question',
          '', '', '', 'SKIPPED', '', 0, 'N/A', 'NO', 'Item skipped', 0, 'User choice'
        ]);
      }
    } else if (r.type === 'read-aloud') {
      itemsSheet.appendRow([
        now,
        initials,
        r.item,
        '',
        '',
        'read-aloud',
        '', '',
        r.durationSec || '',
        '',
        '',
        '',
        '',
        'NO',
        r.mediaPresent ? 'media present' : '',
        '',
        ''
      ]);
    }
  });

  saveBackupData(ss, { ...data, initials });
  generateParticipantSummary(ss, initials);
  logEvent(ss, { ...data, initials, eventType: 'Study Completed (summary ingest)' });

  // Central: single completion row
  const endTs2 = new Date().toISOString();
  let elapsedSec2 = 0;
  if (data.totals && data.totals.durationSec != null) {
    elapsedSec2 = Number(data.totals.durationSec) || 0;
  } else if (data.startedAt && data.finishedAt) {
    const st = new Date(data.startedAt), en = new Date(data.finishedAt);
    if (!isNaN(st) && !isNaN(en)) elapsedSec2 = Math.max(0, Math.round((en - st)/1000));
  }
  wiat_central__logTask(initials, 'Completed', {
    timestamp: endTs2,
    endTime: endTs2,
    elapsed: elapsedSec2,
    active: elapsedSec2,
    details: 'WIAT summary ingest; items=' + itemsCompleted
  });
  wiat_central__touchSession(initials, endTs2);
  wiat_central__logEvent(initials, 'WIAT: Study Completed (summary)', '', endTs2);

  return createResponse({ status: 'success', message: 'Summary ingested' });
}

// ===============================================
// DETAILED SCORING TRACKING
// ===============================================
function saveDetailedScoring(ss, data) {
  if (!data || !data.scoringDetails) return;
  const scoringSheet = getOrCreateSheet(ss, 'Scoring_Details', [
    'Timestamp', 'Initials', 'Item', 'Question', 'Response',
    'Matched Patterns', 'Matched Concepts', 'Found Concepts',
    'Required Both', 'Count Based', 'Auto Score',
    'Confidence', 'Needs Review', 'Notes'
  ]);

  const details = data.scoringDetails || {};
  scoringSheet.appendRow([
    new Date(),
    data.initials,
    data.itemNumber,
    data.questionText || '',
    data.response || '',
    details.matchedPattern || '',
    details.matchedConcept || '',
    details.foundConcepts ? details.foundConcepts.join(', ') : '',
    details.requiresBoth || '',
    details.countBased || '',
    data.autoScore || '',
    data.scoreConfidence || '',
    data.needsReview ? 'YES' : 'NO',
    details.notes || ''
  ]);
}

// ===============================================
// DATA BACKUP
// ===============================================
function saveBackupData(ss, data) {
  try {
    const backupFolder = getOrCreateFolder(CONFIG.DATA_BACKUP_FOLDER_NAME);
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const filename = `${(data.initials || 'unknown')}_backup_${timestamp}.json`;

    const blob = Utilities.newBlob(JSON.stringify(data, null, 2), 'application/json', filename);
    const file = backupFolder.createFile(blob);

    return createResponse({
      status: 'success',
      backupId: file.getId(),
      backupUrl: file.getUrl()
    });
  } catch (error) {
    return createResponse({ status: 'error', message: error.toString() });
  }
}

// ===============================================
// HELPERS (Initials as key)
// ===============================================
function getOrCreateSheet(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    if (headers && headers.length > 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f0f0f0');
      sheet.setFrozenRows(1);
    }
  } else if (headers && headers.length) {
    // ensure headers exist in row 1
    const maxc = Math.max(headers.length, sheet.getLastColumn() || headers.length);
    const row = sheet.getRange(1, 1, 1, maxc).getValues()[0].map(v => String(v || ''));
    headers.forEach(h => {
      if (row.indexOf(h) === -1) {
        const newCol = sheet.getLastColumn() + 1;
        sheet.insertColumnAfter(sheet.getLastColumn());
        sheet.getRange(1, newCol).setValue(h).setFontWeight('bold').setBackground('#f0f0f0');
      }
    });
  }
  return sheet;
}

function getOrCreateFolder(folderName, parentFolder = null) {
  const parent = parentFolder || DriveApp.getRootFolder();
  const folders = parent.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  const newFolder = parent.createFolder(folderName);
  return newFolder;
}

function findRowByInitials(sheet, initials) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(initials)) return i + 1;
  }
  return -1;
}

function getSessionDataFromRow(sheet, row) {
  const d = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  return {
    initials: d[0],
    education: d[1],
    itemsCompleted: d[5],
    totalScore: d[6],
    consecutiveZeros: d[7],
    status: d[8]
  };
}

function getSessionData(ss, initials) {
  const sheet = getOrCreateSheet(ss, 'Sessions', [
    'Initials', 'Education', 'Start Time', 'End Time', 'Duration (min)',
    'Items Completed', 'Total Score', 'Consecutive Zeros',
    'Status', 'Discontinued', 'Gate Items Failed', 'Admin Mode',
    'Recording', 'IP Address', 'User Agent', 'Notes'
  ]);
  const row = findRowByInitials(sheet, initials);
  if (row > 0) {
    return createResponse({ status: 'success', session: getSessionDataFromRow(sheet, row) });
  } else {
    return createResponse({ status: 'not_found', session: null });
  }
}

function updateSessionActivity(ss, initials, timestamp) {
  const sheet = getOrCreateSheet(ss, 'Sessions', [
    'Initials', 'Education', 'Start Time', 'End Time', 'Duration (min)',
    'Items Completed', 'Total Score', 'Consecutive Zeros',
    'Status', 'Discontinued', 'Gate Items Failed', 'Admin Mode',
    'Recording', 'IP Address', 'User Agent', 'Notes'
  ]);
  const row = findRowByInitials(sheet, initials);
  if (row > 0) sheet.getRange(row, 16).setValue('Last activity: ' + (timestamp || new Date().toISOString()));
}

function updateSessionTotals(ss, initials, score, consecutiveZeros) {
  const sheet = getOrCreateSheet(ss, 'Sessions', [
    'Initials', 'Education', 'Start Time', 'End Time', 'Duration (min)',
    'Items Completed', 'Total Score', 'Consecutive Zeros',
    'Status', 'Discontinued', 'Gate Items Failed', 'Admin Mode',
    'Recording', 'IP Address', 'User Agent', 'Notes'
  ]);
  const row = findRowByInitials(sheet, initials);
  if (row > 0) {
    const currentItems = Number(sheet.getRange(row, 6).getValue()) || 0;
    sheet.getRange(row, 6).setValue(currentItems + 1);

    const currentScore = Number(sheet.getRange(row, 7).getValue()) || 0;
    sheet.getRange(row, 7).setValue(currentScore + (Number(score) || 0));

    sheet.getRange(row, 8).setValue(Number(consecutiveZeros) || 0);
  }
}

function logEvent(ss, data) {
  const eventSheet = getOrCreateSheet(ss, 'Events_Log', [
    'Timestamp', 'Initials', 'Event Type', 'Details', 'Data'
  ]);
  eventSheet.appendRow([
    new Date(),
    (data.initials || data.pid || 'unknown'),
    data.eventType || data.action || 'unknown',
    data.details || '',
    JSON.stringify(data)
  ]);
}

function createResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ===============================================
// SUMMARY GENERATION
// ===============================================
function generateParticipantSummary(ss, initials) {
  const summarySheet = getOrCreateSheet(ss, 'Participant_Summary', [
    'Initials', 'Education', 'Total Items', 'Total Score', 'Avg Score',
    'Items Needing Review', 'Reading Time Avg (sec)',
    'Discontinued', 'Gate Items Failed', 'Completion Date'
  ]);

  const itemsSheet = getOrCreateSheet(ss, 'Item_Responses', [
    'Timestamp', 'Initials', 'Item Number', 'Image File', 'Question Text',
    'Item Type', 'Start Time', 'End Time', 'Duration (sec)',
    'Response', 'Explanation', 'Auto Score', 'Score Confidence',
    'Needs Review', 'Scoring Notes', 'Final Score', 'Skip Reason'
  ]);
  const itemsData = itemsSheet.getDataRange().getValues();

  let totalItems = 0;
  let totalScore = 0;
  let needsReview = 0;
  let totalReadingTime = 0;
  let readingCount = 0;

  for (let i = 1; i < itemsData.length; i++) {
    if (itemsData[i][1] === initials) {
      totalItems++;
      totalScore += Number(itemsData[i][15] || 0);
      if (itemsData[i][13] === 'YES') needsReview++;
      if (Number(itemsData[i][8] || 0) > 0) {
        totalReadingTime += Number(itemsData[i][8] || 0);
        readingCount++;
      }
    }
  }

  const sessionsSheet = getOrCreateSheet(ss, 'Sessions', [
    'Initials', 'Education', 'Start Time', 'End Time', 'Duration (min)',
    'Items Completed', 'Total Score', 'Consecutive Zeros',
    'Status', 'Discontinued', 'Gate Items Failed', 'Admin Mode',
    'Recording', 'IP Address', 'User Agent', 'Notes'
  ]);
  const sessionRow = findRowByInitials(sessionsSheet, initials);
  const sdat = sessionRow > 0 ? sessionsSheet.getRange(sessionRow, 1, 1, 16).getValues()[0] : [];

  const row = findRowByInitials(summarySheet, initials);
  const summaryValues = [
    initials,
    sdat[1] || '',
    totalItems,
    totalScore,
    totalItems > 0 ? (totalScore / totalItems).toFixed(2) : 0,
    needsReview,
    readingCount > 0 ? (totalReadingTime / readingCount).toFixed(1) : 0,
    sdat[9] || 'No',
    sdat[10] || '',
    new Date()
  ];

  if (row > 0) {
    summarySheet.getRange(row, 1, 1, summaryValues.length).setValues([summaryValues]);
  } else {
    summarySheet.appendRow(summaryValues);
  }
}

// ===============================================
// SETUP / DASHBOARD / ANALYTICS
// ===============================================
function initialSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  getOrCreateSheet(ss, 'Sessions', [
    'Initials', 'Education', 'Start Time', 'End Time', 'Duration (min)',
    'Items Completed', 'Total Score', 'Consecutive Zeros',
    'Status', 'Discontinued', 'Gate Items Failed', 'Admin Mode',
    'Recording', 'IP Address', 'User Agent', 'Notes'
  ]);

  getOrCreateSheet(ss, 'Item_Responses', [
    'Timestamp', 'Initials', 'Item Number', 'Image File', 'Question Text',
    'Item Type', 'Start Time', 'End Time', 'Duration (sec)',
    'Response', 'Explanation', 'Auto Score', 'Score Confidence',
    'Needs Review', 'Scoring Notes', 'Final Score', 'Skip Reason'
  ]);

  getOrCreateSheet(ss, 'Item_Progress', [
    'Timestamp', 'Initials', 'Item', 'Event', 'Details'
  ]);

  getOrCreateSheet(ss, 'Reading_Times', [
    'Timestamp', 'Initials', 'Item', 'Image', 'Reading Type',
    'Start Time', 'End Time', 'Duration (sec)', 'Words Count'
  ]);

  getOrCreateSheet(ss, 'Scoring_Details', [
    'Timestamp', 'Initials', 'Item', 'Question', 'Response',
    'Matched Patterns', 'Matched Concepts', 'Found Concepts',
    'Required Both', 'Count Based', 'Auto Score',
    'Confidence', 'Needs Review', 'Notes'
  ]);

  getOrCreateSheet(ss, 'Video_Recordings', [
    'Timestamp', 'Initials', 'Item Number', 'Filename',
    'File ID', 'File URL', 'File Size (KB)', 'Upload Status'
  ]);

  getOrCreateSheet(ss, 'Audio_Recordings', [
    'Timestamp', 'Initials', 'Item Number', 'Filename',
    'File ID', 'File URL', 'File Size (KB)', 'Upload Status'
  ]);

  getOrCreateSheet(ss, 'Upload_Errors', [
    'Timestamp', 'Initials', 'Item', 'Error', 'Type'
  ]);

  getOrCreateSheet(ss, 'Events_Log', [
    'Timestamp', 'Initials', 'Event Type', 'Details', 'Data'
  ]);

  getOrCreateSheet(ss, 'Participant_Summary', [
    'Initials', 'Education', 'Total Items', 'Total Score', 'Avg Score',
    'Items Needing Review', 'Reading Time Avg (sec)',
    'Discontinued', 'Gate Items Failed', 'Completion Date'
  ]);

  getOrCreateFolder(CONFIG.RECORDINGS_FOLDER_NAME);
  getOrCreateFolder(CONFIG.DATA_BACKUP_FOLDER_NAME);
  getOrCreateFolder(CONFIG.ITEM_IMAGES_FOLDER_NAME);

  createDashboard(ss);
}

function createDashboard(ss) {
  const dashboard = getOrCreateSheet(ss, 'Dashboard', []);
  dashboard.clear();

  dashboard.getRange(1, 1).setValue('WIAT-2 Reading Comprehension Dashboard')
    .setFontSize(20).setFontWeight('bold');
  dashboard.getRange(2, 1).setValue('Last Updated: ' + new Date().toLocaleString());

  dashboard.getRange(4, 1).setValue('Overall Statistics').setFontWeight('bold').setFontSize(14);

  const stats = [
    ['Metric', 'Value'],
    ['Total Participants', '=COUNTA(Sessions!A:A)-1'],
    ['Active Sessions', '=COUNTIF(Sessions!I:I,"Active")'],
    ['Completed Sessions', '=COUNTIF(Sessions!I:I,"Complete")'],
    ['Discontinued', '=COUNTIF(Sessions!J:J,"Yes")'],
    ['Average Score', '=AVERAGE(Sessions!G:G)'],
    ['Total Items Recorded', '=COUNTA(Item_Responses!A:A)-1'],
    ['Items Needing Review', '=COUNTIF(Item_Responses!N:N,"YES")'],
    ['Videos Uploaded', '=COUNTA(Video_Recordings!A:A)-1'],
    ['Audio Uploaded', '=COUNTA(Audio_Recordings!A:A)-1'],
    ['Upload Errors', '=COUNTA(Upload_Errors!A:A)-1'],
    ['Average Reading Time', '=AVERAGE(Reading_Times!H:H)']
  ];
  dashboard.getRange(5, 1, stats.length, 2).setValues(stats);

  dashboard.setColumnWidth(1, 240);
  dashboard.setColumnWidth(4, 220);
  dashboard.setColumnWidth(7, 260);
}

function generateItemStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const itemsSheet = getOrCreateSheet(ss, 'Item_Responses', [
    'Timestamp', 'Initials', 'Item Number', 'Image File', 'Question Text',
    'Item Type', 'Start Time', 'End Time', 'Duration (sec)',
    'Response', 'Explanation', 'Auto Score', 'Score Confidence',
    'Needs Review', 'Scoring Notes', 'Final Score', 'Skip Reason'
  ]);
  const data = itemsSheet.getDataRange().getValues();

  const itemStats = {};
  for (let i = 1; i < data.length; i++) {
    const itemNum = data[i][2];
    if (!itemStats[itemNum]) {
      itemStats[itemNum] = { attempts: 0, totalScore: 0, skipped: 0, needsReview: 0 };
    }
    itemStats[itemNum].attempts++;
    itemStats[itemNum].totalScore += Number(data[i][15] || 0);
    if (data[i][9] === 'SKIPPED') itemStats[itemNum].skipped++;
    if (data[i][13] === 'YES') itemStats[itemNum].needsReview++;
  }

  const statsSheet = getOrCreateSheet(ss, 'Item_Statistics', [
    'Item Number', 'Attempts', 'Average Score', 'Skip Rate', 'Review Rate'
  ]);

  if (statsSheet.getLastRow() > 1) {
    statsSheet.getRange(2, 1, statsSheet.getLastRow() - 1, 5).clear();
  }

  Object.keys(itemStats).sort((a, b) => Number(a) - Number(b)).forEach((itemNum, index) => {
    const s = itemStats[itemNum];
    statsSheet.getRange(index + 2, 1, 1, 5).setValues([[
      itemNum,
      s.attempts,
      s.attempts > 0 ? (s.totalScore / s.attempts).toFixed(2) : 0,
      s.attempts > 0 ? (s.skipped / s.attempts * 100).toFixed(1) + '%' : '0%',
      s.attempts > 0 ? (s.needsReview / s.attempts * 100).toFixed(1) + '%' : '0%'
    ]]]);
  });
}

// ===============================================
// TEST FUNCTION
// ===============================================
function testSetup() {
  initialSetup();

  const testData = {
    action: 'session_start',
    initials: 'TT',
    education: '10',
    timestamp: new Date().toISOString(),
    adminMode: true,
    hasRecording: true
  };

  const result = doPost({ postData: { contents: JSON.stringify(testData) } });
  console.log('Test result:', result.getContent());
}

// ===============================================
// -------- Central Sync Bridge (optional) --------
// Uses INITIALS to look up or create sessions in master workbook
// A "Session Map" sheet in the master workbook may be used:
//   Headers: Initials | Session Code | Email
// If not found, Session Code = Initials.
// ===============================================
function central_(cb) {
  if (!CENTRAL_SYNC.ENABLED || !CENTRAL_SYNC.SPREADSHEET_ID) return null;
  try {
    var ss = SpreadsheetApp.openById(CENTRAL_SYNC.SPREADSHEET_ID);
    return cb(ss);
  } catch (e) {
    console.warn('Central sync unavailable:', e);
    return null;
  }
}
function wiat_central__ensureSheet(ss, name, headers){
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (headers && headers.length) {
      sh.getRange(1,1,1,headers.length).setValues([headers]);
      sh.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#f1f3f4');
      sh.setFrozenRows(1);
    }
  } else if (headers && headers.length) {
    var maxc = Math.max(headers.length, sh.getLastColumn() || headers.length);
    var existing = sh.getRange(1,1,1,maxc).getValues()[0].map(function(v){return String(v||'');});
    headers.forEach(function(h){
      if (existing.indexOf(h) === -1){
        var newCol = sh.getLastColumn() + 1;
        sh.insertColumnAfter(sh.getLastColumn());
        sh.getRange(1,newCol).setValue(h).setFontWeight('bold').setBackground('#f1f3f4');
      }
    });
  }
  return sh;
}
function wiat_central__setByHeader(sh,row,header,value){
  var last = sh.getLastColumn();
  var hdrs = sh.getRange(1,1,1,last).getValues()[0].map(function(v){return String(v||'').trim();});
  var idx = hdrs.indexOf(header);
  if (idx === -1){
    var newCol = last + 1;
    sh.insertColumnAfter(last);
    sh.getRange(1,newCol).setValue(header).setFontWeight('bold').setBackground('#f1f3f4');
    idx = newCol - 1;
  }
  sh.getRange(row, idx + 1).setValue(value);
}
function wiat_central__findRowBySessionCode(sh, code){
  var vals = sh.getDataRange().getValues();
  for (var r=1;r<vals.length;r++){ if (String(vals[r][0]) === String(code)) return r+1; }
  return 0;
}
function wiat_central__getSessionCode(css, initials){
  var map = wiat_central__ensureSheet(css, 'Session Map', ['Initials','Session Code','Email']);
  var vals = map.getDataRange().getValues();
  for (var i=1;i<vals.length;i++){
    if (String(vals[i][0]) === String(initials)) return String(vals[i][1] || initials);
  }
  return String(initials);
}
function wiat_central__touchSession(initials, timestamp){
  central_(function(css){
    var sh = wiat_central__ensureSheet(css,'Sessions',[
      'Session Code','Participant ID','Email','Created Date','Last Activity',
      'Total Time (min)','Active Time (min)','Idle Time (min)','Tasks Completed','Status',
      'Device Type','Consent Status','Consent Source','Consent Code','Consent Timestamp',
      'EEG Status','EEG Scheduled At','EEG Scheduling Source','Hearing Status','Fluency','State JSON'
    ]);
    var code = wiat_central__getSessionCode(css, initials);
    var row = wiat_central__findRowBySessionCode(sh, code);
    if (!row){
      row = sh.getLastRow() + 1;
      sh.insertRowsAfter(sh.getLastRow() || 1, 1);
      wiat_central__setByHeader(sh,row,'Session Code',code);
      wiat_central__setByHeader(sh,row,'Participant ID',initials);
      wiat_central__setByHeader(sh,row,'Status','Active');
      wiat_central__setByHeader(sh,row,'Device Type','Desktop');
      wiat_central__setByHeader(sh,row,'Created Date', timestamp || new Date().toISOString());
    }
    wiat_central__setByHeader(sh,row,'Last Activity', timestamp || new Date().toISOString());
  });
}
function wiat_central__logEvent(initials, type, details, timestamp){
  central_(function(css){
    var ses = wiat_central__ensureSheet(css,'Session Events',
      ['Timestamp','Session Code','Event Type','Details','IP Address','User Agent']);
    var code = wiat_central__getSessionCode(css,initials);
    ses.appendRow([ timestamp || new Date().toISOString(), code, type, details || '', '', '' ]);
  });
}
function toSeconds_(val){
  var n = Number(val) || 0;
  // If val looks like minutes (<1000), treat as minutes; else as seconds
  return (n > 0 && n < 1000) ? Math.round(n * 60) : Math.round(n);
}
function wiat_central__logTask(initials, eventType, opts){
  central_(function(css){
    var tp = wiat_central__ensureSheet(css,'Task Progress',[
      'Timestamp','Session Code','Participant ID','Task Name','Event Type',
      'Start Time','End Time','Elapsed Time (sec)','Active Time (sec)','Pause Count',
      'Inactive Time (sec)','Activity Score (%)','Details','Completed'
    ]);
    var code = wiat_central__getSessionCode(css,initials);
    tp.appendRow([
      opts.timestamp || new Date().toISOString(),
      code,
      initials,
      CENTRAL_SYNC.TASK_NAME,
      eventType,
      opts.startTime || '',
      opts.endTime || '',
      Number(opts.elapsed || 0),
      Number(opts.active  || 0),
      Number(opts.pauseCount || 0),
      Number(opts.inactive || 0),
      Number(opts.activityPct != null ? opts.activityPct : (opts.elapsed ? 100 : 0)),
      opts.details || '',
      eventType === 'Completed'
    ]);
  });
}
function wiat_central__logVideo(initials, itemNumber, file){
  central_(function(css){
    var v = wiat_central__ensureSheet(css,'Video Tracking',[
      'Timestamp','Session Code','Image Number','Filename','File ID','File URL',
      'File Size (KB)','Upload Time','Upload Method','Dropbox Path','Upload Status','Error Message'
    ]);
    var code = wiat_central__getSessionCode(css,initials);
    v.appendRow([
      new Date(),
      code,
      itemNumber || '',
      file.filename || '',
      file.id || file.fileId || '',
      file.url || file.fileUrl || '',
      file.sizeKb != null ? file.sizeKb :
        (file.bytes != null ? Math.round(file.bytes/1024) : ''),
      new Date().toISOString(),
      'google_drive',
      '',
      'success',
      ''
    ]);
  });
}
