// ===============================================
// WIAT-2 READING COMPREHENSION - GOOGLE APPS SCRIPT (fixed)
// Supports trial-by-trial logging AND single summary upload
// ===============================================

// CONFIGURATION - Update these as you like
const CONFIG = {
  RECORDINGS_FOLDER_NAME: 'WIAT-2 Recordings',
  DATA_BACKUP_FOLDER_NAME: 'WIAT-2 Data Backups',
  ITEM_IMAGES_FOLDER_NAME: 'WIAT-2 Stimuli'
};

// ===============================================
// MAIN HANDLER
// ===============================================
function doPost(e) {
  try {
    console.log('📨 Received POST request');

    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('No data received');
    }

    const data = JSON.parse(e.postData.contents);
    // Allow requests without PID by falling back to initials
    if (!data.pid && data.initials) data.pid = data.initials;
    console.log('📋 Action:', data.action);

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
        return getSessionData(ss, data.pid);

      case 'save_backup':
        return saveBackupData(ss, data);

      // New single-payload summary mode (text-only frontend)
      case 'study_completed':
        return handleStudyCompleted(ss, data);

      default:
        logEvent(ss, data);
        return createResponse({ status: 'success' });
    }
  } catch (error) {
    console.error('❌ Error:', error);
    return createResponse({
      status: 'error',
      message: error.toString()
    });
  }
}

// ===============================================
// SESSION MANAGEMENT
// ===============================================
function handleSessionStart(ss, data) {
  const sessionsSheet = getOrCreateSheet(ss, 'Sessions', [
    'PID', 'Initials', 'Education', 'Start Time', 'End Time', 'Duration (min)',
    'Items Completed', 'Total Score', 'Consecutive Zeros',
    'Status', 'Discontinued', 'Gate Items Failed', 'Admin Mode',
    'Recording', 'IP Address', 'User Agent', 'Notes'
  ]);

  const existingRow = findRowByPID(sessionsSheet, data.pid);
  if (existingRow > 0) {
    sessionsSheet.getRange(existingRow, 10).setValue('Active');
    sessionsSheet.getRange(existingRow, 17).setValue('Session resumed at ' + data.timestamp);

    logEvent(ss, { ...data, eventType: 'Session Resumed' });

    return createResponse({
      status: 'success',
      message: 'Session resumed',
      existingData: getSessionDataFromRow(sessionsSheet, existingRow)
    });
  }

  sessionsSheet.appendRow([
    data.pid,
    data.initials || '',
    data.education,
    data.timestamp,
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
  getOrCreateFolder(`${data.pid}_${(data.timestamp || new Date().toISOString()).split('T')[0]}`, recordingsFolder);

  logEvent(ss, { ...data, eventType: 'Session Started' });

  return createResponse({ status: 'success', message: 'Session created' });
}

// ===============================================
// ITEM TRACKING
// ===============================================
function handleItemStarted(ss, data) {
  const itemsSheet = getOrCreateSheet(ss, 'Item_Responses', [
    'Timestamp', 'PID', 'Initials', 'Item Number', 'Image File', 'Question Text',
    'Item Type', 'Start Time', 'End Time', 'Duration (sec)',
    'Response', 'Explanation', 'Auto Score', 'Score Confidence',
    'Needs Review', 'Scoring Notes', 'Final Score', 'Skip Reason'
  ]);

  itemsSheet.appendRow([
    new Date(),
    data.pid,
    data.initials || '',
    data.itemNumber,
    data.imageFile || '',
    data.questionText || '',
    data.itemType || 'question',
    data.timestamp || new Date().toISOString(),
    '', // End time
    0,  // Duration
    '', // Response
    '', // Explanation
    '', // Auto score
    '', // Confidence
    '', // Needs review
    '', // Scoring notes
    '', // Final score
    ''  // Skip reason
  ]);

  const progressSheet = getOrCreateSheet(ss, 'Item_Progress', [
    'Timestamp', 'PID', 'Initials', 'Item', 'Event', 'Details'
  ]);
  progressSheet.appendRow([
    new Date(),
    data.pid,
    data.initials || '',
    data.itemNumber,
    'Started',
    `Type: ${data.itemType || 'question'}, Image: ${data.imageFile || ''}`
  ]);

  updateSessionActivity(ss, data.pid, data.timestamp || new Date().toISOString());
  return createResponse({ status: 'success' });
}

function handleItemCompleted(ss, data) {
  const itemsSheet = getOrCreateSheet(ss, 'Item_Responses', [
    'Timestamp', 'PID', 'Initials', 'Item Number', 'Image File', 'Question Text',
    'Item Type', 'Start Time', 'End Time', 'Duration (sec)',
    'Response', 'Explanation', 'Auto Score', 'Score Confidence',
    'Needs Review', 'Scoring Notes', 'Final Score', 'Skip Reason'
  ]);

  const values = itemsSheet.getDataRange().getValues();
  let targetRow = -1;
  for (let i = values.length - 1; i >= 1; i--) {
    if (String(values[i][1]) === String(data.pid) &&
        String(values[i][3]) === String(data.itemNumber) &&
        !values[i][8]) {
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
    itemsSheet.getRange(targetRow, 9).setValue(data.endTime || new Date().toISOString());
    itemsSheet.getRange(targetRow, 10).setValue(Number(data.duration) || 0);
    itemsSheet.getRange(targetRow, 11).setValue(data.response || '');
    itemsSheet.getRange(targetRow, 12).setValue(data.explanation || '');
    itemsSheet.getRange(targetRow, 13).setValue(autoScore);
    itemsSheet.getRange(targetRow, 14).setValue(data.scoreConfidence || '');
    itemsSheet.getRange(targetRow, 15).setValue(needsReview ? 'YES' : 'NO');
    itemsSheet.getRange(targetRow, 16).setValue(data.scoringNotes || '');
    itemsSheet.getRange(targetRow, 17).setValue(finalScore);
  } else {
    itemsSheet.appendRow([
      new Date(),
      data.pid,
      data.initials || '',
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
    'Timestamp', 'PID', 'Initials', 'Item', 'Event', 'Details'
  ]);
  progressSheet.appendRow([
    new Date(),
    data.pid,
    data.initials || '',
    data.itemNumber,
    'Completed',
    `Score: ${autoScore}, Confidence: ${data.scoreConfidence}, Review: ${needsReview ? 'YES' : 'NO'}`
  ]);

  updateSessionTotals(ss, data.pid, Number(finalScore) || 0, Number(data.consecutiveZeros) || 0);

  saveDetailedScoring(ss, { ...data, autoScore: autoScore, needsReview: needsReview });

  return createResponse({ status: 'success' });
}

function handleItemSkipped(ss, data) {
  const itemsSheet = getOrCreateSheet(ss, 'Item_Responses', [
    'Timestamp', 'PID', 'Initials', 'Item Number', 'Image File', 'Question Text',
    'Item Type', 'Start Time', 'End Time', 'Duration (sec)',
    'Response', 'Explanation', 'Auto Score', 'Score Confidence',
    'Needs Review', 'Scoring Notes', 'Final Score', 'Skip Reason'
  ]);

  itemsSheet.appendRow([
    new Date(),
    data.pid,
    data.initials || '',
    data.itemNumber,
    data.imageFile || '',
    data.questionText || '',
    data.itemType || 'question',
    data.timestamp || new Date().toISOString(),
    data.timestamp || new Date().toISOString(),
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
    'Timestamp', 'PID', 'Initials', 'Item', 'Event', 'Details'
  ]);
  progressSheet.appendRow([
    new Date(),
    data.pid,
    data.initials || '',
    data.itemNumber,
    'Skipped',
    data.reason || 'User choice'
  ]);

  updateSessionTotals(ss, data.pid, 0, Number(data.consecutiveZeros) || 0);

  return createResponse({ status: 'success' });
}

// ===============================================
// READING TIME TRACKING
// ===============================================
function handleReadingTime(ss, data) {
  const readingSheet = getOrCreateSheet(ss, 'Reading_Times', [
    'Timestamp', 'PID', 'Initials', 'Item', 'Image', 'Reading Type',
    'Start Time', 'End Time', 'Duration (sec)', 'Words Count'
  ]);

  readingSheet.appendRow([
    new Date(),
    data.pid,
    data.initials || '',
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
// VIDEO UPLOAD
// ===============================================
function handleVideoUpload(data) {
  try {
    console.log('📹 Processing video upload for:', data.pid);

    if (!data.pid || !data.videoData) throw new Error('Missing required fields');

    const videoBytes = Utilities.base64Decode(data.videoData);
    const maxSize = 25 * 1024 * 1024; // 25MB
    if (videoBytes.length > maxSize) throw new Error(`Video too large (${Math.round(videoBytes.length / 1024 / 1024)}MB)`);

    const recordingsFolder = getOrCreateFolder(CONFIG.RECORDINGS_FOLDER_NAME);
    const participantFolder = getOrCreateFolder(
      `${data.pid}_${data.sessionDate || new Date().toISOString().split('T')[0]}`,
      recordingsFolder
    );

    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const filename = `${data.pid}_item${data.itemNumber || 'full'}_${timestamp}.mp4`;

    const blob = Utilities.newBlob(videoBytes, 'video/mp4', filename);
    const file = participantFolder.createFile(blob);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const videoSheet = getOrCreateSheet(ss, 'Video_Recordings', [
      'Timestamp', 'PID', 'Initials', 'Item Number', 'Filename',
      'File ID', 'File URL', 'File Size (KB)', 'Upload Status'
    ]);

    videoSheet.appendRow([
      new Date(),
      data.pid,
      data.initials || '',
      data.itemNumber || 'Full Session',
      filename,
      file.getId(),
      file.getUrl(),
      Math.round(videoBytes.length / 1024),
      'Success'
    ]);

    console.log('✅ Video uploaded:', filename);

    return createResponse({
      status: 'success',
      fileId: file.getId(),
      fileUrl: file.getUrl(),
      filename: filename
    });
  } catch (error) {
    console.error('❌ Video upload error:', error);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const errorSheet = getOrCreateSheet(ss, 'Upload_Errors', [
      'Timestamp', 'PID', 'Initials', 'Item', 'Error', 'Type'
    ]);

    errorSheet.appendRow([
      new Date(),
      data.pid || 'unknown',
      data.initials || '',
      data.itemNumber || '',
      error.toString(),
      'Video Upload'
    ]);

    return createResponse({ status: 'error', message: error.toString() });
  }
}

// ===============================================
// GENERIC BLOB UPLOAD (VIDEO/ AUDIO)
// ===============================================
function handleBlobUpload(data) {
  try {
    console.log(`📦 Processing ${data.kind || 'unknown'} upload for:`, data.pid);

    if (!data.pid || !data.data) throw new Error('Missing required fields');

    const bytes = Utilities.base64Decode(data.data);
    const maxSize = 25 * 1024 * 1024; // 25MB
    if (bytes.length > maxSize) {
      throw new Error(`${data.kind || 'blob'} too large (${Math.round(bytes.length / 1024 / 1024)}MB)`);
    }

    const recordingsFolder = getOrCreateFolder(CONFIG.RECORDINGS_FOLDER_NAME);
    const participantFolder = getOrCreateFolder(
      `${data.pid}_${data.sessionDate || new Date().toISOString().split('T')[0]}`,
      recordingsFolder
    );

    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const extension = data.kind === 'video' ? '.mp4' : '.mp3';
    const mime = data.mime || (data.kind === 'video' ? 'video/mp4' : 'audio/mpeg');
    const filename = `${data.pid}_item${data.itemNumber || 'full'}_${timestamp}${extension}`;

    const blob = Utilities.newBlob(bytes, mime, filename);
    const file = participantFolder.createFile(blob);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = data.kind === 'video' ? 'Video_Recordings' : 'Audio_Recordings';
    const sheet = getOrCreateSheet(ss, sheetName, [
      'Timestamp', 'PID', 'Initials', 'Item Number', 'Filename',
      'File ID', 'File URL', 'File Size (KB)', 'Upload Status'
    ]);

    sheet.appendRow([
      new Date(),
      data.pid,
      data.initials || '',
      data.itemNumber || 'Full Session',
      filename,
      file.getId(),
      file.getUrl(),
      Math.round(file.getSize() / 1024),
      'Success'
    ]);

    console.log('✅ Blob uploaded:', filename);

    return createResponse({
      status: 'success',
      fileId: file.getId(),
      fileUrl: file.getUrl(),
      filename: filename
    });
  } catch (error) {
    console.error('❌ Blob upload error:', error);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const errorSheet = getOrCreateSheet(ss, 'Upload_Errors', [
      'Timestamp', 'PID', 'Initials', 'Item', 'Error', 'Type'
    ]);

    errorSheet.appendRow([
      new Date(),
      data.pid || 'unknown',
      data.initials || '',
      data.itemNumber || '',
      error.toString(),
      data.kind === 'video' ? 'Video Upload' : 'Audio Upload'
    ]);

    return createResponse({ status: 'error', message: error.toString() });
  }
}

// ===============================================
// SESSION COMPLETION
// ===============================================
function handleSessionComplete(ss, data) {
  const sessionsSheet = getOrCreateSheet(ss, 'Sessions', [
    'PID', 'Initials', 'Education', 'Start Time', 'End Time', 'Duration (min)',
    'Items Completed', 'Total Score', 'Consecutive Zeros',
    'Status', 'Discontinued', 'Gate Items Failed', 'Admin Mode',
    'Recording', 'IP Address', 'User Agent', 'Notes'
  ]);

  const row = findRowByPID(sessionsSheet, data.pid);
  if (row > 0) {
    sessionsSheet.getRange(row, 5).setValue(data.timestamp || new Date());
    sessionsSheet.getRange(row, 6).setValue(Number(data.duration) || 0);
    sessionsSheet.getRange(row, 7).setValue(Number(data.itemsCompleted) || 0);
    sessionsSheet.getRange(row, 8).setValue(Number(data.totalScore) || 0);
    sessionsSheet.getRange(row, 9).setValue(Number(data.consecutiveZeros) || 0);
    sessionsSheet.getRange(row, 10).setValue('Complete');
    sessionsSheet.getRange(row, 11).setValue(data.discontinued ? 'Yes' : 'No');
    sessionsSheet.getRange(row, 12).setValue(data.gateItemsFailed || '');
  }

  saveBackupData(ss, data);
  generateParticipantSummary(ss, data.pid);
  logEvent(ss, { ...data, eventType: 'Session Complete' });

  return createResponse({ status: 'success', message: 'Session completed' });
}

// ===============================================
// SINGLE-PAYLOAD SUMMARY INGEST
// ===============================================
function handleStudyCompleted(ss, data) {
  const sessions = getOrCreateSheet(ss, 'Sessions', [
    'PID','Initials','Education','Start Time','End Time','Duration (min)',
    'Items Completed','Total Score','Consecutive Zeros',
    'Status','Discontinued','Gate Items Failed','Admin Mode',
    'Recording','IP Address','User Agent','Notes'
  ]);

  const start = data.startedAt ? new Date(data.startedAt) : null;
  const end   = data.finishedAt ? new Date(data.finishedAt) : null;
  const durationMin = (start && end && !isNaN(start) && !isNaN(end)) ? Math.max(0, (end - start) / 1000 / 60) : 0;
  const itemsCompleted = Number((data.totals && data.totals.items) || (data.results ? data.results.length : 0));
  const totalScore = Number((data.totals && data.totals.points) || 0);

  const row = findRowByPID(sessions, data.pid);
  if (row > 0) {
    sessions.getRange(row, 2).setValue(data.initials || '');
    sessions.getRange(row, 3).setValue(data.edu || '');
    sessions.getRange(row, 4).setValue(start || new Date());
    sessions.getRange(row, 5).setValue(end || new Date());
    sessions.getRange(row, 6).setValue(durationMin);
    sessions.getRange(row, 7).setValue(itemsCompleted);
    sessions.getRange(row, 8).setValue(totalScore);
    sessions.getRange(row, 9).setValue(0);
    sessions.getRange(row,10).setValue('Complete');
    sessions.getRange(row,11).setValue('No');
  } else {
    sessions.appendRow([
      data.pid || '',
      data.initials || '',
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

  const itemsSheet = getOrCreateSheet(ss, 'Item_Responses', [
    'Timestamp', 'PID', 'Initials', 'Item Number', 'Image File', 'Question Text',
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
            data.pid,
            data.initials || '',
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
          now, data.pid, data.initials || '', r.item, '', '', 'question',
          '', '', '', 'SKIPPED', '', 0, 'N/A', 'NO', 'Item skipped', 0, 'User choice'
        ]);
      }
    } else if (r.type === 'read-aloud') {
      itemsSheet.appendRow([
        now,
        data.pid,
        data.initials || '',
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

  saveBackupData(ss, data);
  generateParticipantSummary(ss, data.pid);
  logEvent(ss, { ...data, eventType: 'Study Completed (summary ingest)' });

  return createResponse({ status: 'success', message: 'Summary ingested' });
}

// ===============================================
// DETAILED SCORING TRACKING
// ===============================================
function saveDetailedScoring(ss, data) {
  if (!data || !data.scoringDetails) return;
  const scoringSheet = getOrCreateSheet(ss, 'Scoring_Details', [
    'Timestamp', 'PID', 'Initials', 'Item', 'Question', 'Response',
    'Matched Patterns', 'Matched Concepts', 'Found Concepts',
    'Required Both', 'Count Based', 'Auto Score',
    'Confidence', 'Needs Review', 'Notes'
  ]);

  const details = data.scoringDetails || {};
  scoringSheet.appendRow([
    new Date(),
    data.pid,
    data.initials || '',
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
    const filename = `${(data.pid || 'unknown')}_backup_${timestamp}.json`;

    const blob = Utilities.newBlob(JSON.stringify(data, null, 2), 'application/json', filename);
    const file = backupFolder.createFile(blob);

    console.log('💾 Backup saved:', filename);

    return createResponse({
      status: 'success',
      backupId: file.getId(),
      backupUrl: file.getUrl()
    });
  } catch (error) {
    console.error('Backup error:', error);
    return createResponse({ status: 'error', message: error.toString() });
  }
}

// ===============================================
// HELPERS
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
  }
  return sheet;
}

function getOrCreateFolder(folderName, parentFolder = null) {
  const parent = parentFolder || DriveApp.getRootFolder();
  const folders = parent.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  const newFolder = parent.createFolder(folderName);
  console.log('📁 Created folder:', folderName);
  return newFolder;
}

function findRowByPID(sheet, pid) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === pid) return i + 1;
  }
  return -1;
}

function getSessionDataFromRow(sheet, row) {
  const data = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  return {
    pid: data[0],
    initials: data[1],
    education: data[2],
    itemsCompleted: data[6],
    totalScore: data[7],
    consecutiveZeros: data[8],
    status: data[9]
  };
}

function getSessionData(ss, pid) {
  const sheet = getOrCreateSheet(ss, 'Sessions', [
    'PID', 'Initials', 'Education', 'Start Time', 'End Time', 'Duration (min)',
    'Items Completed', 'Total Score', 'Consecutive Zeros',
    'Status', 'Discontinued', 'Gate Items Failed', 'Admin Mode',
    'Recording', 'IP Address', 'User Agent', 'Notes'
  ]);
  const row = findRowByPID(sheet, pid);
  if (row > 0) {
    return createResponse({ status: 'success', session: getSessionDataFromRow(sheet, row) });
  } else {
    return createResponse({ status: 'not_found', session: null });
  }
}

function updateSessionActivity(ss, pid, timestamp) {
  const sheet = getOrCreateSheet(ss, 'Sessions', [
    'PID', 'Initials', 'Education', 'Start Time', 'End Time', 'Duration (min)',
    'Items Completed', 'Total Score', 'Consecutive Zeros',
    'Status', 'Discontinued', 'Gate Items Failed', 'Admin Mode',
    'Recording', 'IP Address', 'User Agent', 'Notes'
  ]);
  const row = findRowByPID(sheet, pid);
  if (row > 0) sheet.getRange(row, 17).setValue('Last activity: ' + (timestamp || new Date().toISOString()));
}

function updateSessionTotals(ss, pid, score, consecutiveZeros) {
  const sheet = getOrCreateSheet(ss, 'Sessions', [
    'PID', 'Initials', 'Education', 'Start Time', 'End Time', 'Duration (min)',
    'Items Completed', 'Total Score', 'Consecutive Zeros',
    'Status', 'Discontinued', 'Gate Items Failed', 'Admin Mode',
    'Recording', 'IP Address', 'User Agent', 'Notes'
  ]);
  const row = findRowByPID(sheet, pid);
  if (row > 0) {
    const currentItems = Number(sheet.getRange(row, 7).getValue()) || 0;
    sheet.getRange(row, 7).setValue(currentItems + 1);

    const currentScore = Number(sheet.getRange(row, 8).getValue()) || 0;
    sheet.getRange(row, 8).setValue(currentScore + (Number(score) || 0));

    sheet.getRange(row, 9).setValue(Number(consecutiveZeros) || 0);
  }
}

function logEvent(ss, data) {
  const eventSheet = getOrCreateSheet(ss, 'Events_Log', [
    'Timestamp', 'PID', 'Initials', 'Event Type', 'Details', 'Data'
  ]);

  eventSheet.appendRow([
    new Date(),
    data.pid || 'unknown',
    data.initials || '',
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
function generateParticipantSummary(ss, pid) {
  const summarySheet = getOrCreateSheet(ss, 'Participant_Summary', [
    'PID', 'Initials', 'Education', 'Total Items', 'Total Score', 'Avg Score',
    'Items Needing Review', 'Reading Time Avg (sec)',
    'Discontinued', 'Gate Items Failed', 'Completion Date'
  ]);

  const itemsSheet = getOrCreateSheet(ss, 'Item_Responses', [
    'Timestamp', 'PID', 'Initials', 'Item Number', 'Image File', 'Question Text',
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
    if (itemsData[i][1] === pid) {
      totalItems++;
      totalScore += Number(itemsData[i][16] || 0);
      if (itemsData[i][14] === 'YES') needsReview++;
      if (Number(itemsData[i][9] || 0) > 0) {
        totalReadingTime += Number(itemsData[i][9] || 0);
        readingCount++;
      }
    }
  }

  const sessionsSheet = getOrCreateSheet(ss, 'Sessions', [
    'PID', 'Initials', 'Education', 'Start Time', 'End Time', 'Duration (min)',
    'Items Completed', 'Total Score', 'Consecutive Zeros',
    'Status', 'Discontinued', 'Gate Items Failed', 'Admin Mode',
    'Recording', 'IP Address', 'User Agent', 'Notes'
  ]);
  const sessionRow = findRowByPID(sessionsSheet, pid);
  const sessionData = sessionRow > 0 ? sessionsSheet.getRange(sessionRow, 1, 1, 17).getValues()[0] : [];

  const row = findRowByPID(summarySheet, pid);
  const summaryValues = [
    pid,
    sessionData[1] || '',
    sessionData[2] || '',
    totalItems,
    totalScore,
    totalItems > 0 ? (totalScore / totalItems).toFixed(2) : 0,
    needsReview,
    readingCount > 0 ? (totalReadingTime / readingCount).toFixed(1) : 0,
    sessionData[10] || 'No',
    sessionData[11] || '',
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
    'PID', 'Initials', 'Education', 'Start Time', 'End Time', 'Duration (min)',
    'Items Completed', 'Total Score', 'Consecutive Zeros',
    'Status', 'Discontinued', 'Gate Items Failed', 'Admin Mode',
    'Recording', 'IP Address', 'User Agent', 'Notes'
  ]);

  getOrCreateSheet(ss, 'Item_Responses', [
    'Timestamp', 'PID', 'Initials', 'Item Number', 'Image File', 'Question Text',
    'Item Type', 'Start Time', 'End Time', 'Duration (sec)',
    'Response', 'Explanation', 'Auto Score', 'Score Confidence',
    'Needs Review', 'Scoring Notes', 'Final Score', 'Skip Reason'
  ]);

  getOrCreateSheet(ss, 'Item_Progress', [
    'Timestamp', 'PID', 'Initials', 'Item', 'Event', 'Details'
  ]);

  getOrCreateSheet(ss, 'Reading_Times', [
    'Timestamp', 'PID', 'Initials', 'Item', 'Image', 'Reading Type',
    'Start Time', 'End Time', 'Duration (sec)', 'Words Count'
  ]);

  getOrCreateSheet(ss, 'Scoring_Details', [
    'Timestamp', 'PID', 'Initials', 'Item', 'Question', 'Response',
    'Matched Patterns', 'Matched Concepts', 'Found Concepts',
    'Required Both', 'Count Based', 'Auto Score',
    'Confidence', 'Needs Review', 'Notes'
  ]);

  getOrCreateSheet(ss, 'Video_Recordings', [
    'Timestamp', 'PID', 'Initials', 'Item Number', 'Filename',
    'File ID', 'File URL', 'File Size (KB)', 'Upload Status'
  ]);

  getOrCreateSheet(ss, 'Audio_Recordings', [
    'Timestamp', 'PID', 'Initials', 'Item Number', 'Filename',
    'File ID', 'File URL', 'File Size (KB)', 'Upload Status'
  ]);

  getOrCreateSheet(ss, 'Upload_Errors', [
    'Timestamp', 'PID', 'Initials', 'Item', 'Error', 'Type'
  ]);

  getOrCreateSheet(ss, 'Events_Log', [
    'Timestamp', 'PID', 'Initials', 'Event Type', 'Details', 'Data'
  ]);

  getOrCreateSheet(ss, 'Participant_Summary', [
    'PID', 'Initials', 'Education', 'Total Items', 'Total Score', 'Avg Score',
    'Items Needing Review', 'Reading Time Avg (sec)',
    'Discontinued', 'Gate Items Failed', 'Completion Date'
  ]);

  getOrCreateFolder(CONFIG.RECORDINGS_FOLDER_NAME);
  getOrCreateFolder(CONFIG.DATA_BACKUP_FOLDER_NAME);
  getOrCreateFolder(CONFIG.ITEM_IMAGES_FOLDER_NAME);

  createDashboard(ss);
  console.log('✅ Setup complete! All sheets and folders created.');
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

  dashboard.getRange(4, 4).setValue('Grade Distribution').setFontWeight('bold').setFontSize(14);
  const grades = [
    ['Grade', 'Count'],
    ['Grade 9', '=COUNTIF(Sessions!B:B,"9")'],
    ['Grade 10', '=COUNTIF(Sessions!B:B,"10")'],
    ['Grade 11', '=COUNTIF(Sessions!B:B,"11")'],
    ['Grade 12', '=COUNTIF(Sessions!B:B,"12")'],
    ['College+', '=COUNTIF(Sessions!B:B,"13+")']
  ];
  dashboard.getRange(5, 4, grades.length, 2).setValues(grades);

  dashboard.getRange(4, 7).setValue('Item Completion Rates').setFontWeight('bold').setFontSize(14);
  dashboard.getRange(5, 7).setValue('Run generateItemStats() for detailed item analysis');

  dashboard.setColumnWidth(1, 220);
  dashboard.setColumnWidth(4, 200);
  dashboard.setColumnWidth(7, 240);
}

function generateItemStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const itemsSheet = getOrCreateSheet(ss, 'Item_Responses', [
    'Timestamp', 'PID', 'Initials', 'Item Number', 'Image File', 'Question Text',
    'Item Type', 'Start Time', 'End Time', 'Duration (sec)',
    'Response', 'Explanation', 'Auto Score', 'Score Confidence',
    'Needs Review', 'Scoring Notes', 'Final Score', 'Skip Reason'
  ]);
  const data = itemsSheet.getDataRange().getValues();

  const itemStats = {};
  for (let i = 1; i < data.length; i++) {
    const itemNum = data[i][3];
    if (!itemStats[itemNum]) {
      itemStats[itemNum] = { attempts: 0, totalScore: 0, skipped: 0, needsReview: 0 };
    }
    itemStats[itemNum].attempts++;
    itemStats[itemNum].totalScore += Number(data[i][16] || 0);
    if (data[i][10] === 'SKIPPED') itemStats[itemNum].skipped++;
    if (data[i][14] === 'YES') itemStats[itemNum].needsReview++;
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

  console.log('Item statistics generated');
}

// ===============================================
// TEST FUNCTION
// ===============================================
function testSetup() {
  initialSetup();

  const testData = {
    action: 'session_start',
    pid: 'TEST001',
    initials: 'TT',
    education: '10',
    timestamp: new Date().toISOString(),
    adminMode: true,
    hasRecording: true
  };

  const result = doPost({ postData: { contents: JSON.stringify(testData) } });
  console.log('Test result:', result.getContent());
  console.log('✅ Test complete! Check your sheets.');
}
