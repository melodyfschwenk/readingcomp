// ===============================================
// WIAT-2 READING COMPREHENSION - GOOGLE APPS SCRIPT
// (Merged-ID edition: stores Initials _and_ PID in a single "Initials" value)
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
// ID helpers (merge PID + Initials without losing either)
// ===============================================
function _norm_(s){ return String(s || '').trim(); }
function makeIdKey_(initials, pid){
  const i = _norm_(initials), p = _norm_(pid);
  if (i && p) {
    // avoid double-joining if user already sent merged
    if (i === p || i.indexOf(p) !== -1 || p.indexOf(i) !== -1) return i.length >= p.length ? i : p;
    return i + '_' + p;
  }
  return i || p || 'UNKNOWN';
}
function idFromPayload_(data){
  return makeIdKey_(data.initials, data.pid);
}

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
      case 'session_start':     return handleSessionStart(ss, data);
      case 'item_started':      return handleItemStarted(ss, data);
      case 'item_completed':    return handleItemCompleted(ss, data);
      case 'item_skipped':      return handleItemSkipped(ss, data);
      case 'reading_time':      return handleReadingTime(ss, data);

      // Upload
      case 'video_upload':      return handleVideoUpload(data);
      case 'upload_blob':       return handleBlobUpload(data);

      // Session control
      case 'session_complete':  return handleSessionComplete(ss, data);
      case 'get_session':       return getSessionData(ss, idFromPayload_(data));

      // Backup / summary ingest
      case 'save_backup':       return saveBackupData(ss, data);
      case 'study_completed':   return handleStudyCompleted(ss, data);

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
// SESSION MANAGEMENT (Merged ID as key)
// ===============================================
function handleSessionStart(ss, data) {
  const sessionsSheet = getOrCreateSheet(ss, 'Sessions', SESSIONS_HEADERS());
  const idKey = idFromPayload_(data);
  if (idKey === 'UNKNOWN') return createResponse({ status: 'error', message: 'Missing initials / pid' });

  const existingRow = findRowByInitials(sessionsSheet, idKey);
  if (existingRow > 0) {
    sessionsSheet.getRange(existingRow, 9).setValue('Active'); // Status
    sessionsSheet.getRange(existingRow, 16).setValue('Session resumed at ' + (data.timestamp || new Date().toISOString()));
    wiat_central__touchSession(idKey, data.timestamp);
    wiat_central__logEvent(idKey, 'WIAT: Session Resumed', '', data.timestamp);

    return createResponse({
      status: 'success',
      message: 'Session resumed',
      existingData: getSessionDataFromRow(sessionsSheet, existingRow)
    });
  }

  sessionsSheet.appendRow([
    idKey,
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
  getOrCreateFolder(`${idKey}_${(data.timestamp || new Date().toISOString()).split('T')[0]}`, recordingsFolder);

  logEvent(ss, { ...data, initials: idKey, eventType: 'Session Started' });
  wiat_central__touchSession(idKey, data.timestamp);
  wiat_central__logEvent(idKey, 'WIAT: Session Started', '', data.timestamp);

  return createResponse({ status: 'success', message: 'Session created' });
}

// ===============================================
// ITEM TRACKING
// ===============================================
function handleItemStarted(ss, data) {
  const idKey = idFromPayload_(data);
  if (idKey === 'UNKNOWN') return createResponse({ status: 'error', message: 'Missing initials / pid' });

  const itemsSheet = getOrCreateSheet(ss, 'Item_Responses', ITEM_RESP_HEADERS());

  itemsSheet.appendRow([
    new Date(),
    idKey,
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

  const progressSheet = getOrCreateSheet(ss, 'Item_Progress', ['Timestamp','Initials','Item','Event','Details']);
  progressSheet.appendRow([
    new Date(),
    idKey,
    data.itemNumber,
    'Started',
    `Type: ${data.itemType || 'question'}, Image: ${data.imageFile || ''}`
  ]);

  updateSessionActivity(ss, idKey, data.timestamp || new Date().toISOString());
  wiat_central__touchSession(idKey, data.timestamp);
  wiat_central__logEvent(idKey, 'WIAT: Item Started', 'Item ' + data.itemNumber, data.timestamp);

  return createResponse({ status: 'success' });
}

function handleItemCompleted(ss, data) {
  const idKey = idFromPayload_(data);
  if (idKey === 'UNKNOWN') return createResponse({ status: 'error', message: 'Missing initials / pid' });

  const itemsSheet = getOrCreateSheet(ss, 'Item_Responses', ITEM_RESP_HEADERS());
  const values = itemsSheet.getDataRange().getValues();
  let targetRow = -1;
  for (let i = values.length - 1; i >= 1; i--) {
    if (String(values[i][1]) === idKey &&
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
      idKey,
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

  const progressSheet = getOrCreateSheet(ss, 'Item_Progress', ['Timestamp','Initials','Item','Event','Details']);
  progressSheet.appendRow([
    new Date(),
    idKey,
    data.itemNumber,
    'Completed',
    `Score: ${autoScore}, Confidence: ${data.scoreConfidence}, Review: ${needsReview ? 'YES' : 'NO'}`
  ]);

  updateSessionTotals(ss, idKey, Number(finalScore) || 0, Number(data.consecutiveZeros) || 0);

  saveDetailedScoring(ss, { ...data, initials: idKey, autoScore: autoScore, needsReview: needsReview });

  wiat_central__touchSession(idKey, data.endTime || data.timestamp);
  wiat_central__logEvent(
    idKey,
    'WIAT: Item Completed',
    'Item ' + data.itemNumber + ' | autoScore=' + (data.autoScore !== undefined ? data.autoScore : ''),
    data.endTime || data.timestamp
  );

  return createResponse({ status: 'success' });
}

function handleItemSkipped(ss, data) {
  const idKey = idFromPayload_(data);
  if (idKey === 'UNKNOWN') return createResponse({ status: 'error', message: 'Missing initials / pid' });

  const itemsSheet = getOrCreateSheet(ss, 'Item_Responses', ITEM_RESP_HEADERS());
  const ts = data.timestamp || new Date().toISOString();
  itemsSheet.appendRow([
    new Date(),
    idKey,
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

  const progressSheet = getOrCreateSheet(ss, 'Item_Progress', ['Timestamp','Initials','Item','Event','Details']);
  progressSheet.appendRow([
    new Date(),
    idKey,
    data.itemNumber,
    'Skipped',
    data.reason || 'User choice'
  ]);

  updateSessionTotals(ss, idKey, 0, Number(data.consecutiveZeros) || 0);

  wiat_central__touchSession(idKey, ts);
  wiat_central__logEvent(idKey, 'WIAT: Item Skipped', 'Item ' + data.itemNumber + ' | ' + (data.reason || 'User choice'), ts);

  return createResponse({ status: 'success' });
}

// ===============================================
// READING TIME TRACKING
// ===============================================
function handleReadingTime(ss, data) {
  const idKey = idFromPayload_(data);
  if (idKey === 'UNKNOWN') return createResponse({ status: 'error', message: 'Missing initials / pid' });

  const readingSheet = getOrCreateSheet(ss, 'Reading_Times', [
    'Timestamp', 'Initials', 'Item', 'Image', 'Reading Type',
    'Start Time', 'End Time', 'Duration (sec)', 'Words Count'
  ]);

  readingSheet.appendRow([
    new Date(),
    idKey,
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
    const idKey = idFromPayload_(data);
    if (idKey === 'UNKNOWN' || !data.videoData) throw new Error('Missing required fields');

    const videoBytes = Utilities.base64Decode(data.videoData);
    const maxSize = 25 * 1024 * 1024; // 25MB
    if (videoBytes.length > maxSize) throw new Error(`Video too large (${Math.round(videoBytes.length / 1024 / 1024)}MB)`);

    const recordingsFolder = getOrCreateFolder(CONFIG.RECORDINGS_FOLDER_NAME);
    const participantFolder = getOrCreateFolder(
      `${idKey}_${data.sessionDate || new Date().toISOString().split('T')[0]}`,
      recordingsFolder
    );

    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const filename = `${idKey}_item${data.itemNumber || 'full'}_${timestamp}.mp4`;

    const blob = Utilities.newBlob(videoBytes, 'video/mp4', filename);
    const file = participantFolder.createFile(blob);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const videoSheet = getOrCreateSheet(ss, 'Video_Recordings', [
      'Timestamp', 'Initials', 'Item Number', 'Filename',
      'File ID', 'File URL', 'File Size (KB)', 'Upload Status'
    ]);

    videoSheet.appendRow([
      new Date(),
      idKey,
      data.itemNumber || 'Full Session',
      filename,
      file.getId(),
      file.getUrl(),
      Math.round(videoBytes.length / 1024),
      'Success'
    ]);

    wiat_central__logVideo(idKey, data.itemNumber, {
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
      (idFromPayload_(data) || 'unknown'),
      (data && data.itemNumber) || '',
      error.toString(),
      'Video Upload'
    ]);
    return createResponse({ status: 'error', message: error.toString() });
  }
}

function handleBlobUpload(data) {
  try {
    const idKey = idFromPayload_(data);
    if (idKey === 'UNKNOWN' || !data.data) throw new Error('Missing required fields');

    const bytes = Utilities.base64Decode(data.data);
    const maxSize = 25 * 1024 * 1024; // 25MB
    if (bytes.length > maxSize) {
      throw new Error(`${data.kind || 'blob'} too large (${Math.round(bytes.length / 1024 / 1024)}MB)`);
    }

    const recordingsFolder = getOrCreateFolder(CONFIG.RECORDINGS_FOLDER_NAME);
    const participantFolder = getOrCreateFolder(
      `${idKey}_${data.sessionDate || new Date().toISOString().split('T')[0]}`,
      recordingsFolder
    );

    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const extension = data.kind === 'video' ? '.mp4' : '.mp3';
    const mime = data.mime || (data.kind === 'video' ? 'video/mp4' : 'audio/mpeg');
    const filename = `${idKey}_item${data.itemNumber || 'full'}_${timestamp}${extension}`;

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
      idKey,
      data.itemNumber || 'Full Session',
      filename,
      file.getId(),
      file.getUrl(),
      Math.round(file.getSize() / 1024),
      'Success'
    ]);

    if ((data.kind || '').toLowerCase() === 'video') {
      wiat_central__logVideo(idKey, data.itemNumber, {
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
      (idFromPayload_(data) || 'unknown'),
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
  const idKey = idFromPayload_(data);
  if (idKey === 'UNKNOWN') return createResponse({ status: 'error', message: 'Missing initials / pid' });

  const sessionsSheet = getOrCreateSheet(ss, 'Sessions', SESSIONS_HEADERS());

  const row = findRowByInitials(sessionsSheet, idKey);
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

  saveBackupData(ss, { ...data, initials: idKey });
  generateParticipantSummary(ss, idKey);
  logEvent(ss, { ...data, initials: idKey, eventType: 'Session Complete' });

  // Central: write one consolidated completion row
  const endTs = data.timestamp || new Date().toISOString();
  const elapsedSec = toSeconds_(data.duration); // minutes→seconds if it looks like minutes
  wiat_central__logTask(idKey, 'Completed', {
    timestamp: endTs,
    endTime: endTs,
    elapsed: elapsedSec,
    active: elapsedSec,
    details: 'WIAT totalScore=' + (Number(data.totalScore) || 0)
  });
  wiat_central__touchSession(idKey, endTs);
  wiat_central__logEvent(idKey, 'WIAT: Session Completed',
    'items=' + (Number(data.itemsCompleted)||0) + ', score=' + (Number(data.totalScore)||0),
    endTs);

  return createResponse({ status: 'success', message: 'Session completed' });
}

// ===============================================
// SINGLE-PAYLOAD SUMMARY INGEST
// ===============================================
function handleStudyCompleted(ss, data) {
  const idKey = makeIdKey_(data.initials, data.pid);
  if (idKey === 'UNKNOWN') return createResponse({ status: 'error', message: 'Missing initials / pid' });

  const sessions = getOrCreateSheet(ss, 'Sessions', SESSIONS_HEADERS());

  const start = data.startedAt ? new Date(data.startedAt) : null;
  const end   = data.finishedAt ? new Date(data.finishedAt) : null;
  const durationMin = (start && end && !isNaN(start) && !isNaN(end)) ? Math.max(0, (end - start) / 1000 / 60) : 0;
  const itemsCompleted = Number((data.totals && data.totals.items) || (data.results ? data.results.length : 0));
  const totalScore = Number((data.totals && data.totals.points) || 0);

  const row = findRowByInitials(sessions, idKey);
  if (row > 0) {
    sessions.getRange(row, 1, 1, 16).setValues([[
      idKey,
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
      idKey,
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
  const itemsSheet = getOrCreateSheet(ss, 'Item_Responses', ITEM_RESP_HEADERS());
  const now = new Date();
  (data.results || []).forEach(r => {
    if (r.type === 'qa') {
      if (r.answers && r.answers.length) {
        r.answers.forEach(a => {
          const review = (a.note || '').toLowerCase().includes('review');
          itemsSheet.appendRow([
            now,
            idKey,
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
          now, idKey, r.item, '', '', 'question',
          '', '', '', 'SKIPPED', '', 0, 'N/A', 'NO', 'Item skipped', 0, 'User choice'
        ]);
      }
    } else if (r.type === 'read-aloud') {
      itemsSheet.appendRow([
        now,
        idKey,
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

  saveBackupData(ss, { ...data, initials: idKey });
  generateParticipantSummary(ss, idKey);
  logEvent(ss, { ...data, initials: idKey, eventType: 'Study Completed (summary ingest)' });

  // Central: single completion row
  const endTs2 = new Date().toISOString();
  let elapsedSec2 = 0;
  if (data.totals && data.totals.durationSec != null) {
    elapsedSec2 = Number(data.totals.durationSec) || 0;
  } else if (data.startedAt && data.finishedAt) {
    const st = new Date(data.startedAt), en = new Date(data.finishedAt);
    if (!isNaN(st) && !isNaN(en)) elapsedSec2 = Math.max(0, Math.round((en - st)/1000));
  }
  wiat_central__logTask(idKey, 'Completed', {
    timestamp: endTs2,
    endTime: endTs2,
    elapsed: elapsedSec2,
    active: elapsedSec2,
    details: 'WIAT summary ingest; items=' + itemsCompleted
  });
  wiat_central__touchSession(idKey, endTs2);
  wiat_central__logEvent(idKey, 'WIAT: Study Completed (summary)', '', endTs2);

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
    idFromPayload_(data),
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
    const filename = `${(idFromPayload_(data) || 'unknown')}_backup_${timestamp}.json`;

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
// HELPERS (Initials column holds merged ID)
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

function getSessionData(ss, initialsMerged) {
  const sheet = getOrCreateSheet(ss, 'Sessions', SESSIONS_HEADERS());
  const row = findRowByInitials(sheet, initialsMerged);
  if (row > 0) {
    return createResponse({ status: 'success', session: getSessionDataFromRow(sheet, row) });
  } else {
    return createResponse({ status: 'not_found', session: null });
  }
}

function updateSessionActivity(ss, initialsMerged, timestamp) {
  const sheet = getOrCreateSheet(ss, 'Sessions', SESSIONS_HEADERS());
  const row = findRowByInitials(sheet, initialsMerged);
  if (row > 0) sheet.getRange(row, 16).setValue('Last activity: ' + (timestamp || new Date().toISOString()));
}

function updateSessionTotals(ss, initialsMerged, score, consecutiveZeros) {
  const sheet = getOrCreateSheet(ss, 'Sessions', SESSIONS_HEADERS());
  const row = findRowByInitials(sheet, initialsMerged);
  if (row > 0) {
    const currentItems = Number(sheet.getRange(row, 6).getValue()) || 0;
    sheet.getRange(row, 6).setValue(currentItems + 1);

    const currentScore = Number(sheet.getRange(row, 7).getValue()) || 0;
    sheet.getRange(row, 7).setValue(currentScore + (Number(score) || 0));

    sheet.getRange(row, 8).setValue(Number(consecutiveZeros) || 0);
  }
}

function logEvent(ss, data) {
  const eventSheet = getOrCreateSheet(ss, 'Events_Log', ['Timestamp', 'Initials', 'Event Type', 'Details', 'Data']);
  eventSheet.appendRow([
    new Date(),
    idFromPayload_(data) || 'unknown',
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
function generateParticipantSummary(ss, initialsMerged) {
  const summarySheet = getOrCreateSheet(ss, 'Participant_Summary', [
    'Initials', 'Education', 'Total Items', 'Total Score', 'Avg Score',
    'Items Needing Review', 'Reading Time Avg (sec)',
    'Discontinued', 'Gate Items Failed', 'Completion Date'
  ]);

  const itemsSheet = getOrCreateSheet(ss, 'Item_Responses', ITEM_RESP_HEADERS());
  const itemsData = itemsSheet.getDataRange().getValues();

  let totalItems = 0;
  let totalScore = 0;
  let needsReview = 0;
  let totalReadingTime = 0;
  let readingCount = 0;

  for (let i = 1; i < itemsData.length; i++) {
    if (itemsData[i][1] === initialsMerged) {
      totalItems++;
      totalScore += Number(itemsData[i][15] || 0);
      if (itemsData[i][13] === 'YES') needsReview++;
      if (Number(itemsData[i][8] || 0) > 0) {
        totalReadingTime += Number(itemsData[i][8] || 0);
        readingCount++;
      }
    }
  }

  const sessionsSheet = getOrCreateSheet(ss, 'Sessions', SESSIONS_HEADERS());
  const sessionRow = findRowByInitials(sessionsSheet, initialsMerged);
  const sdat = sessionRow > 0 ? sessionsSheet.getRange(sessionRow, 1, 1, 16).getValues()[0] : [];

  const row = findRowByInitials(summarySheet, initialsMerged);
  const summaryValues = [
    initialsMerged,
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
function SESSIONS_HEADERS() {
  return [
    'Initials', 'Education', 'Start Time', 'End Time', 'Duration (min)',
    'Items Completed', 'Total Score', 'Consecutive Zeros',
    'Status', 'Discontinued', 'Gate Items Failed', 'Admin Mode',
    'Recording', 'IP Address', 'User Agent', 'Notes'
  ];
}
function ITEM_RESP_HEADERS() {
  return [
    'Timestamp', 'Initials', 'Item Number', 'Image File', 'Question Text',
    'Item Type', 'Start Time', 'End Time', 'Duration (sec)',
    'Response', 'Explanation', 'Auto Score', 'Score Confidence',
    'Needs Review', 'Scoring Notes', 'Final Score', 'Skip Reason'
  ];
}

function initialSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  getOrCreateSheet(ss, 'Sessions', SESSIONS_HEADERS());
  getOrCreateSheet(ss, 'Item_Responses', ITEM_RESP_HEADERS());
  getOrCreateSheet(ss, 'Item_Progress', ['Timestamp', 'Initials', 'Item', 'Event', 'Details']);
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
  const itemsSheet = getOrCreateSheet(ss, 'Item_Responses', ITEM_RESP_HEADERS());
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
    ]]);
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
    pid: 'TEST001',
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
function wiat_central__getSessionCode(css, initialsMerged){
  var map = wiat_central__ensureSheet(css, 'Session Map', ['Initials','Session Code','Email']);
  var vals = map.getDataRange().getValues();
  for (var i=1;i<vals.length;i++){
    if (String(vals[i][0]) === String(initialsMerged)) return String(vals[i][1] || initialsMerged);
  }
  return String(initialsMerged);
}
function wiat_central__touchSession(initialsMerged, timestamp){
  central_(function(css){
    var sh = wiat_central__ensureSheet(css,'Sessions',[
      'Session Code','Participant ID','Email','Created Date','Last Activity',
      'Total Time (min)','Active Time (min)','Idle Time (min)','Tasks Completed','Status',
      'Device Type','Consent Status','Consent Source','Consent Code','Consent Timestamp',
      'EEG Status','EEG Scheduled At','EEG Scheduling Source','Hearing Status','Fluency','State JSON'
    ]);
    var code = wiat_central__getSessionCode(css, initialsMerged);
    var row = wiat_central__findRowBySessionCode(sh, code);
    if (!row){
      row = sh.getLastRow() + 1;
      sh.insertRowsAfter(sh.getLastRow() || 1, 1);
      wiat_central__setByHeader(sh,row,'Session Code',code);
      wiat_central__setByHeader(sh,row,'Participant ID',initialsMerged);
      wiat_central__setByHeader(sh,row,'Status','Active');
      wiat_central__setByHeader(sh,row,'Device Type','Desktop');
      wiat_central__setByHeader(sh,row,'Created Date', timestamp || new Date().toISOString());
    }
    wiat_central__setByHeader(sh,row,'Last Activity', timestamp || new Date().toISOString());
  });
}
function wiat_central__logEvent(initialsMerged, type, details, timestamp){
  central_(function(css){
    var ses = wiat_central__ensureSheet(css,'Session Events',
      ['Timestamp','Session Code','Event Type','Details','IP Address','User Agent']);
    var code = wiat_central__getSessionCode(css,initialsMerged);
    ses.appendRow([ timestamp || new Date().toISOString(), code, type, details || '', '', '' ]);
  });
}
function toSeconds_(val){
  var n = Number(val) || 0;
  // If val looks like minutes (<1000), treat as minutes; else as seconds
  return (n > 0 && n < 1000) ? Math.round(n * 60) : Math.round(n);
}
function wiat_central__logTask(initialsMerged, eventType, opts){
  central_(function(css){
    var tp = wiat_central__ensureSheet(css,'Task Progress',[
      'Timestamp','Session Code','Participant ID','Task Name','Event Type',
      'Start Time','End Time','Elapsed Time (sec)','Active Time (sec)','Pause Count',
      'Inactive Time (sec)','Activity Score (%)','Details','Completed'
    ]);
    var code = wiat_central__getSessionCode(css,initialsMerged);
    tp.appendRow([
      opts.timestamp || new Date().toISOString(),
      code,
      initialsMerged,
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
function wiat_central__logVideo(initialsMerged, itemNumber, file){
  central_(function(css){
    var v = wiat_central__ensureSheet(css,'Video Tracking',[
      'Timestamp','Session Code','Image Number','Filename','File ID','File URL',
      'File Size (KB)','Upload Time','Upload Method','Dropbox Path','Upload Status','Error Message'
    ]);
    var code = wiat_central__getSessionCode(css,initialsMerged);
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

/** ===========================================================
 *  WIAT Housekeeping / Migration (Merged-ID, safe & idempotent)
 *  - Backup workbook
 *  - Normalize Sessions (merge PID→Initials by concatenation)
 *  - Rename/Merge PID→Initials across WIAT sheets
 *  - Consolidate duplicate Sessions rows for same merged ID
 *  - (Optional) Migrate uploads → Media_Tracking
 *  - Merge Item_Progress → Events_Log
 *  - Prune empty legacy sheets
 *  - Rebuild Dashboard
 *  Adds "WIAT Admin" menu.
 *  ===========================================================
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('WIAT Admin')
    .addItem('Safe cleanup / migrate (recommended)', 'wiatSafeCleanup')
    .addSeparator()
    .addItem('Normalize Sessions (merge PID→Initials)', 'wiatNormalizeSessions')
    .addItem('Consolidate duplicate Sessions rows', 'wiatConsolidateSessionsDuplicates')
    .addItem('Migrate uploads → Media_Tracking', 'wiatMigrateMediaToUnified')
    .addItem('Merge Item_Progress → Events_Log', 'wiatMergeItemProgress')
    .addSeparator()
    .addItem('Rebuild Dashboard', 'wiatRebuildDashboard')
    .addItem('Prune empty legacy sheets', 'wiatPruneEmptySheets')
    .addToUi();
}

/* ==============================
   Orchestrator
   ============================== */
function wiatSafeCleanup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  wiatBackupWorkbook_(ss);

  // Ensure essential sheets exist
  wiatEnsureSheetWithHeaders_(ss, 'Sessions', SESSIONS_HEADERS());
  wiatEnsureSheetWithHeaders_(ss, 'Item_Responses', ITEM_RESP_HEADERS());
  wiatEnsureSheetWithHeaders_(ss, 'Events_Log', ['Timestamp','Initials','Event Type','Details','Data','Item']);

  // 1) Normalize Sessions + merge PID→Initials across known WIAT sheets
  wiatNormalizeSessions();
  wiatMergePidIntoInitialsAcross_([
    'Sessions',
    'Item_Responses',
    'Item_Progress',
    'Reading_Times',
    'Scoring_Details',
    'Video_Recordings',
    'Audio_Recordings',
    'Upload_Errors',
    'Events_Log',
    'Participant_Summary'
  ]);

  // 2) Consolidate duplicate Sessions rows to a single merged ID (with log)
  wiatConsolidateSessionsDuplicates();

  // 3) Media logs → Media_Tracking (optional)
  wiatMigrateMediaToUnified();

  // 4) Item_Progress → Events_Log
  wiatMergeItemProgress();

  // 5) Rebuild dashboard
  wiatRebuildDashboard();

  // 6) Prune truly empty legacy sheets
  wiatPruneEmptySheets();

  SpreadsheetApp.getUi().alert('WIAT cleanup/migration complete ✅');
}

/* ==============================
   Sheet utils
   ============================== */
function wiatBackupWorkbook_(ss) {
  var name = ss.getName();
  var backup = SpreadsheetApp.create(
    'BACKUP_WIAT_' + name + '_' + new Date().toISOString().replace(/[:.]/g,'-')
  );
  var sheets = ss.getSheets();
  sheets.forEach(function(sh){
    sh.copyTo(backup).setName(sh.getName());
  });
}

function wiatEnsureSheetWithHeaders_(ss, name, headers) {
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (headers && headers.length) {
      sh.getRange(1,1,1,headers.length).setValues([headers]);
      sh.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#f1f3f4');
      sh.setFrozenRows(1);
      sh.autoResizeColumns(1, headers.length);
    }
  } else if (headers && headers.length) {
    // Ensure any missing headers are added at end (non-destructive)
    var last = Math.max(headers.length, sh.getLastColumn() || headers.length);
    var row = sh.getRange(1,1,1,last).getValues()[0].map(function(v){return String(v||'');});
    headers.forEach(function(h){
      if (row.indexOf(h) === -1) {
        var newCol = sh.getLastColumn() + 1;
        sh.insertColumnAfter(sh.getLastColumn());
        sh.getRange(1, newCol)
          .setValue(h)
          .setFontWeight('bold')
          .setBackground('#f1f3f4');
      }
    });
    sh.setFrozenRows(1);
  }
  return sh;
}
function wiatHeaderMap_(sh) {
  var last = sh.getLastColumn();
  if (last < 1) return {};
  var hdrs = sh.getRange(1,1,1,last).getValues()[0].map(function(v){return String(v||'').trim();});
  var map = {};
  for (var i=0;i<hdrs.length;i++) if (hdrs[i]) map[hdrs[i]] = i+1;
  return map;
}
function wiatSetByHeader_(sh, row, header, value) {
  var map = wiatHeaderMap_(sh);
  if (!map[header]) {
    var newCol = sh.getLastColumn() + 1;
    sh.insertColumnAfter(sh.getLastColumn());
    sh.getRange(1, newCol).setValue(header).setFontWeight('bold').setBackground('#f1f3f4');
    map = wiatHeaderMap_(sh);
  }
  sh.getRange(row, map[header]).setValue(value);
}

/* ==============================
   Merge PID→Initials (concatenate safely)
   ============================== */
function wiatNormalizeSessions() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = wiatEnsureSheetWithHeaders_(ss, 'Sessions', SESSIONS_HEADERS());

  // Merge PID into Initials by concatenation
  wiatMergePidColumn_(sh, 'PID', 'Initials');

  // Formats
  var map = wiatHeaderMap_(sh);
  var nRows = Math.max(1, sh.getMaxRows() - 1);
  function fmt(h, f) { if (map[h]) sh.getRange(2, map[h], nRows).setNumberFormat(f); }

  ['Initials','Education','Status','Discontinued','Gate Items Failed','Admin Mode','Recording','IP Address','User Agent','Notes']
    .forEach(function(h){ if (map[h]) sh.getRange(2,map[h],nRows).setNumberFormat('@'); });

  ['Start Time','End Time'].forEach(function(h){ fmt(h, 'yyyy-mm-dd"T"hh:mm:ss.000'); });
  ['Duration (min)','Items Completed','Total Score','Consecutive Zeros']
    .forEach(function(h){ fmt(h, '0'); });

  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, Math.min(sh.getLastColumn(), 20));
}
function wiatMergePidIntoInitialsAcross_(sheetNames) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  sheetNames.forEach(function(name){
    var sh = ss.getSheetByName(name);
    if (!sh) return;
    wiatMergePidColumn_(sh, 'PID', 'Initials');
    // After migration, remove PID col (it’s preserved in the merged Initials string)
    var map = wiatHeaderMap_(sh);
    if (map['PID']) sh.deleteColumn(map['PID']);
  });
}
// Merge data from oldName col into newName by concatenation without losing either
function wiatMergePidColumn_(sh, oldName, newName) {
  var last = sh.getLastColumn();
  if (last < 1) return;
  var hdrs = sh.getRange(1,1,1,last).getValues()[0].map(function(v){return String(v||'');});
  var oldIdx = hdrs.indexOf(oldName);
  if (oldIdx === -1) return;

  var map = wiatHeaderMap_(sh);
  if (!map[newName]) {
    var newCol = sh.getLastColumn() + 1;
    sh.insertColumnAfter(sh.getLastColumn());
    sh.getRange(1, newCol).setValue(newName).setFontWeight('bold').setBackground('#f1f3f4');
    map = wiatHeaderMap_(sh);
  }

  var n = sh.getLastRow();
  for (var r=2;r<=n;r++) {
    var oldVal = _norm_(sh.getRange(r, oldIdx+1).getValue());
    var cur = _norm_(sh.getRange(r, map[newName]).getValue());
    if (!cur && oldVal) {
      sh.getRange(r, map[newName]).setValue(oldVal);
    } else if (cur && oldVal) {
      // If both exist and not already merged, merge safely
      var merged = makeIdKey_(cur, oldVal);
      if (merged !== cur) sh.getRange(r, map[newName]).setValue(merged);
    }
  }
}

/* ==============================
   Consolidate duplicate Sessions rows to one per merged ID
   (keeps a Merge_Log sheet so nothing is lost)
   ============================== */
function wiatConsolidateSessionsDuplicates() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Sessions');
  if (!sh) return;

  var headers = SESSIONS_HEADERS();
  var map = wiatHeaderMap_(sh);
  if (!map['Initials']) return;

  var vals = sh.getDataRange().getValues();
  if (vals.length <= 2) return;

  var byId = {};      // id -> array of rows (2-based index)
  for (var r=2; r<=sh.getLastRow(); r++){
    var id = _norm_(sh.getRange(r, map['Initials']).getValue());
    if (!id) continue;
    if (!byId[id]) byId[id] = [];
    byId[id].push(r);
  }

  var ids = Object.keys(byId).filter(function(k){return byId[k].length > 1;});
  if (!ids.length) return;

  // Merge log
  var log = ss.getSheetByName('Merge_Log') || ss.insertSheet('Merge_Log');
  if (log.getLastRow() === 0) {
    log.getRange(1,1,1,5).setValues([['When','Sheet','Merged ID','Source Rows','Action']]);
    log.setFrozenRows(1);
  }

  // Aggregation helpers
  function num(v){ var n = Number(v); return isNaN(n) ? 0 : n; }
  function earliest(a,b){ if (!a) return b; if (!b) return a; return new Date(a) < new Date(b) ? a : b; }
  function latest(a,b){ if (!a) return b; if (!b) return a; return new Date(a) > new Date(b) ? a : b; }
  function pickStatus(a,b){
    var pri = {'Complete':3, 'Active':2, 'Discontinued':1, '':0};
    var aa = a || '', bb = b || '';
    return (pri[aa] >= pri[bb]) ? aa : bb;
  }
  function pickYes(a,b){ return (String(a).toLowerCase()==='yes' || String(b).toLowerCase()==='yes') ? 'Yes' : (a||b||''); }
  function coalesce(a,b){ return a || b || ''; }
  function joinNotes(a,b){ return [a,b].filter(Boolean).join(' | '); }

  // Build a target map of merged rows
  var mergedValues = {}; // id -> object with header -> value
  ids.forEach(function(id){
    var rows = byId[id];
    var agg = {};
    // seed with first
    var first = rows[0];
    headers.forEach(function(h){ agg[h] = sh.getRange(first, map[h]).getValue(); });

    // merge others
    for (var i=1;i<rows.length;i++){
      var r = rows[i];
      agg['Initials'] = id;
      agg['Education'] = coalesce(agg['Education'], sh.getRange(r, map['Education']).getValue());
      agg['Start Time'] = earliest(agg['Start Time'], sh.getRange(r, map['Start Time']).getValue());
      agg['End Time']   = latest(agg['End Time'],   sh.getRange(r, map['End Time']).getValue());
      agg['Duration (min)'] = num(agg['Duration (min)']) + num(sh.getRange(r, map['Duration (min)']).getValue());
      agg['Items Completed'] = num(agg['Items Completed']) + num(sh.getRange(r, map['Items Completed']).getValue());
      agg['Total Score'] = num(agg['Total Score']) + num(sh.getRange(r, map['Total Score']).getValue());
      // Prefer latest consecutive zeros (or keep max)
      agg['Consecutive Zeros'] = Math.max(num(agg['Consecutive Zeros']), num(sh.getRange(r, map['Consecutive Zeros']).getValue()));
      agg['Status'] = pickStatus(agg['Status'], sh.getRange(r, map['Status']).getValue());
      agg['Discontinued'] = pickYes(agg['Discontinued'], sh.getRange(r, map['Discontinued']).getValue());
      agg['Gate Items Failed'] = coalesce(agg['Gate Items Failed'], sh.getRange(r, map['Gate Items Failed']).getValue());
      agg['Admin Mode'] = coalesce(agg['Admin Mode'], sh.getRange(r, map['Admin Mode']).getValue());
      agg['Recording']  = coalesce(agg['Recording'], sh.getRange(r, map['Recording']).getValue());
      agg['IP Address'] = coalesce(agg['IP Address'], sh.getRange(r, map['IP Address']).getValue());
      agg['User Agent'] = coalesce(agg['User Agent'], sh.getRange(r, map['User Agent']).getValue());
      agg['Notes'] = joinNotes(agg['Notes'], sh.getRange(r, map['Notes']).getValue());
    }
    mergedValues[id] = agg;

    // Log action
    log.appendRow([new Date(), 'Sessions', id, rows.join(','), 'Merged rows into first; deleted extras']);
  });

  // Write merged values back into the first row, delete the rest
  ids.forEach(function(id){
    var rows = byId[id].sort(function(a,b){return a-b;});
    var first = rows.shift();
    var agg = mergedValues[id];
    // write
    var out = headers.map(function(h){ return agg[h]; });
    sh.getRange(first, 1, 1, headers.length).setValues([out]);
    // delete extra rows (bottom-up)
    rows.sort(function(a,b){return b-a;}).forEach(function(r){ sh.deleteRow(r); });
  });
}

/* ==============================
   Media migration → Media_Tracking (optional)
   ============================== */
function MEDIA_HEADERS() {
  return [
    'Timestamp','Initials','Kind', // "video" | "audio"
    'Item Number','Filename','File ID','File URL',
    'File Size (KB)','Upload Status','Error Message'
  ];
}
function wiatMigrateMediaToUnified() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var media = wiatEnsureSheetWithHeaders_(ss, 'Media_Tracking', MEDIA_HEADERS());
  media.setFrozenRows(1);

  function append(kind, row) {
    media.appendRow([
      row.ts || new Date(),
      row.initials || '',
      kind || '',
      row.item || '',
      row.filename || '',
      row.id || '',
      row.url || '',
      row.kb != null ? row.kb : '',
      row.status || 'Success',
      row.error || ''
    ]);
  }

  // Video_Recordings
  var vr = ss.getSheetByName('Video_Recordings');
  if (vr && vr.getLastRow() > 1) {
    var v = vr.getDataRange().getValues();
    var head = v[0].map(String);
    var pidIdx = head.indexOf('Initials') > -1 ? head.indexOf('Initials') : head.indexOf('PID');
    for (var i=1;i<v.length;i++){
      // merge ID if needed
      var rawId = v[i][pidIdx];
      var mergedId = makeIdKey_(rawId, ''); // already merged if cleanup ran
      append('video', {
        ts: v[i][0],
        initials: mergedId,
        item: v[i][2],
        filename: v[i][3],
        id: v[i][4],
        url: v[i][5],
        kb: v[i][6],
        status: v[i][7] || 'Success'
      });
    }
    ss.deleteSheet(vr);
  }

  // Audio_Recordings
  var ar = ss.getSheetByName('Audio_Recordings');
  if (ar && ar.getLastRow() > 1) {
    var a = ar.getDataRange().getValues();
    var headA = a[0].map(String);
    var pidIdxA = headA.indexOf('Initials') > -1 ? headA.indexOf('Initials') : headA.indexOf('PID');
    for (var j=1;j<a.length;j++){
      var rawIdA = a[j][pidIdxA];
      var mergedIdA = makeIdKey_(rawIdA, '');
      append('audio', {
        ts: a[j][0],
        initials: mergedIdA,
        item: a[j][2],
        filename: a[j][3],
        id: a[j][4],
        url: a[j][5],
        kb: a[j][6],
        status: a[j][7] || 'Success'
      });
    }
    ss.deleteSheet(ar);
  }

  // Upload_Errors → Media_Tracking as "Error"
  var ue = ss.getSheetByName('Upload_Errors');
  if (ue && ue.getLastRow() > 1) {
    var e = ue.getDataRange().getValues();
    var headE = e[0].map(String);
    var pidIdxE = headE.indexOf('Initials') > -1 ? headE.indexOf('Initials') : headE.indexOf('PID');
    for (var k=1;k<e.length;k++){
      var kind = (String(e[k][4] || '').toLowerCase().indexOf('audio') !== -1) ? 'audio' : 'video';
      var rawIdE = e[k][pidIdxE];
      var mergedIdE = makeIdKey_(rawIdE, '');
      append(kind, {
        ts: e[k][0],
        initials: mergedIdE,
        item: e[k][2],
        filename: '',
        id: '',
        url: '',
        kb: '',
        status: 'Error',
        error: e[k][3]
      });
    }
    ss.deleteSheet(ue);
  }

  media.autoResizeColumns(1, media.getLastColumn());
}

/* ==============================
   Merge Item_Progress → Events_Log
   ============================== */
function wiatMergeItemProgress() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var events = wiatEnsureSheetWithHeaders_(ss, 'Events_Log', ['Timestamp','Initials','Event Type','Details','Data','Item']);
  var ip = ss.getSheetByName('Item_Progress');
  if (!ip) return;

  if (ip.getLastRow() > 1) {
    var data = ip.getDataRange().getValues();
    var H = data[0].map(String);
    var pidIdx = H.indexOf('Initials') > -1 ? H.indexOf('Initials') : H.indexOf('PID');
    var itemCol = H.indexOf('Item');
    var eventCol = H.indexOf('Event');
    var detailsCol = H.indexOf('Details');

    for (var i=1;i<data.length;i++) {
      var row = data[i];
      var ts = row[0];
      var rawId = row[pidIdx];
      var mergedId = makeIdKey_(rawId, '');
      var item = itemCol > -1 ? row[itemCol] : '';
      var ev = eventCol > -1 ? row[eventCol] : 'Event';
      var det = detailsCol > -1 ? row[detailsCol] : '';

      events.appendRow([ ts || new Date(), mergedId || '', ev || '', det || '', 'source:Item_Progress', item || '' ]);
    }
  }
  ss.deleteSheet(ip);
}

/* ==============================
   Prune truly empty legacy sheets
   ============================== */
function wiatPruneEmptySheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  [
    'Reading_Times',
    'Scoring_Details',
    'Item_Statistics',
    'Video_Recordings',
    'Audio_Recordings',
    'Upload_Errors',
    'Item_Progress'
  ].forEach(function(name){
    var sh = ss.getSheetByName(name);
    if (!sh) return;
    if (sh.getLastRow() <= 1) {
      ss.deleteSheet(sh);
    }
  });
}

/* ==============================
   Dashboard (re)build
   ============================== */
function wiatRebuildDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dash = ss.getSheetByName('Dashboard') || ss.insertSheet('Dashboard');
  dash.clear();

  dash.getRange(1, 1).setValue('WIAT-2 Reading Comprehension Dashboard (Merged-ID)')
      .setFontSize(20).setFontWeight('bold');
  dash.getRange(2, 1).setValue('Last Updated: ' + new Date().toLocaleString());

  dash.getRange(4, 1).setValue('Overall Statistics').setFontWeight('bold').setFontSize(14);

  var stats = [
    ['Metric', 'Value'],
    ['Total Participants', '=COUNTA(Sessions!A:A)-1'],
    ['Active Sessions', '=COUNTIF(Sessions!I:I,"Active")'],
    ['Completed Sessions', '=COUNTIF(Sessions!I:I,"Complete")'],
    ['Discontinued', '=COUNTIF(Sessions!J:J,"Yes")'],
    ['Average Total Score', '=AVERAGE(Sessions!G:G)'],
    ['Total Items Recorded', '=COUNTA(Item_Responses!A:A)-1'],
    ['Items Needing Review', '=COUNTIF(Item_Responses!N:N,"YES")'],
    ['Media Uploads (All)', '=IFERROR(COUNTA(Media_Tracking!A:A)-1, 0)'],
    ['Media Upload Errors', '=IFERROR(COUNTIF(Media_Tracking!I:I,"Error"), 0)']
  ];
  dash.getRange(5, 1, stats.length, 2).setValues(stats);

  dash.setColumnWidth(1, 260);
  dash.autoResizeColumns(1, 2);
}

/* ===========================================================
   OPTIONAL helper: stream future uploads to Media_Tracking
   =========================================================== */
function wiat_appendMedia_(opts){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = wiatEnsureSheetWithHeaders_(ss, 'Media_Tracking', MEDIA_HEADERS());
  var kb = opts.bytesOrKb != null ? (opts.bytesOrKb > 2048 ? Math.round(opts.bytesOrKb/1024) : opts.bytesOrKb) : '';
  sh.appendRow([
    new Date(),
    opts.initials || '',
    (opts.kind || '').toLowerCase(),
    opts.itemNumber || '',
    opts.filename || '',
    opts.fileId || '',
    opts.fileUrl || '',
    kb || '',
    opts.status || 'Success',
    opts.error || ''
  ]);
}
