/**
 * 氣喘控制測驗（C-ACT/ACT）結果寫入 Google 試算表並寄送報告至病患 Email
 *
 * 重要：請先在試算表建立一個名為「ACT Records」的分頁（tab），
 * 並在第一列加入以下欄位標題：
 * 時間 | patient_id | 病歷號 | 姓名 | 出生日期 | 年齡區間 | 測驗類型 | 分數 | 控制等級 | 等級代碼 | 高雄氣溫°C | 高雄濕度% | 天氣建議 | Email | 同意寄送報告及行銷 | Peak Flow (L/min)
 *
 * 使用方式：
 * 1. 開啟您的 ACT Google 試算表
 * 2. 在底部新增一個分頁，命名為「ACT Records」
 * 3. 在該分頁的第一列貼上上述欄位標題
 * 4. 擴充功能 → Apps Script
 * 5. 全選（Ctrl+A）→ 刪除舊程式碼 → 貼上本檔案全部內容
 * 6. 儲存（Ctrl+S）
 * 7. 部署 → 管理部署作業 → 點擊鉛筆圖示 → 版本選「新版本」→ 部署
 *
 * 注意：舊資料保留在原本的分頁（工作表1），不會被影響。
 * 新的提交會寫入「ACT Records」分頁，包含 patient_id、病歷號、出生日期等新欄位。
 */

var ACT_PF_API_VERSION = 2;
var ACT_PF_RELEASE = '2026-11-24-pf-history2';
// Urgent compatibility fallback. This legacy PIN has appeared in public source;
// it is not strong authentication. Set ACT_NURSE_PIN in Script Properties to
// rotate it without publishing a new client-side credential.
var ACT_LEGACY_NURSE_PIN = '3613';

function doGet(e) {
  var params = (e && e.parameter) || {};
  var mode = params.mode || '';

  // Reachability/version only: no spreadsheet, Gmail, or patient-data access.
  if (mode === 'health') {
    return pfJson_({ ok: true, apiVersion: ACT_PF_API_VERSION,
      release: ACT_PF_RELEASE, service: 'act-pf', time: new Date().toISOString() });
  }

  // Credentials and patient data no longer travel in a nurse GET request.
  if (mode === 'nurse_list' || mode === 'update_pf') {
    return pfJson_(pfFailure_('UPGRADE_REQUIRED', '請重新整理網頁後使用管理模式。', false));
  }
  if (mode) {
    return pfJson_(pfFailure_('INVALID_MODE', '不支援的操作。', false));
  }

  // ── Normal submission ──
  if (!String(params.name || '').trim() || !String(params.testType || '').trim() ||
      params.score === undefined || params.score === '') {
    return pfJson_(pfFailure_('INVALID_SUBMISSION', '缺少測驗資料，未新增紀錄。', false));
  }
  var result = 'ok';
  try {
    var name = params.name || '';
    var birthday = params.birthday || '';
    var chartNumber = params.chartNumber || '';
    var patientId = params.patientId || '';
    var ageGroup = params.ageGroup || '';
    var email = params.email || '';
    var consent = params.consent || '';
    var testType = params.testType || '';
    var score = params.score || '';
    var level = params.level || '';
    var levelKey = params.levelKey || '';
    var levelDesc = params.levelDesc || '';
    var timestamp = params.timestamp || '';
    var temp = params.temp || '';
    var humidity = params.humidity || '';
    var weatherSuggestion = params.weatherSuggestion || '';
    var peakFlow = params.peakFlow || '';

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('ACT Records');
    if (!sheet) {
      sheet = ss.getSheets()[0];
    }

    var ageLabel = ageGroup === 'child' ? '4～11歲（兒童）' : ageGroup === 'adult' ? '12歲以上（成人）' : ageGroup;
    sheet.appendRow([
      timestamp,
      patientId,
      chartNumber,
      name,
      birthday,
      ageLabel,
      testType,
      score,
      level,
      levelKey,
      temp,
      humidity,
      weatherSuggestion,
      email,
      consent === 'yes' ? '是' : '',
      peakFlow
    ]);

    if (email && name) {
      sendReportEmail(name, email, testType, score, level, levelDesc, weatherSuggestion);
    }
  } catch (err) {
    result = 'error: ' + err.toString();
  }
  return ContentService.createTextOutput(result).setMimeType(ContentService.MimeType.TEXT);
}


/**
 * Send the report email with error logging. Keep the existing Gmail scope;
 * this urgent PF release does not introduce a MailApp authorization scope.
 * Never throws — sheet write must not fail because email failed.
 * Every attempt is appended to the 'Email Log' tab so failures are visible
 * without having to open the Apps Script executions panel.
 */
function sendReportEmail(name, email, testType, score, level, levelDesc, weatherSuggestion) {
  var status = 'pending';
  var detail = '';
  try {
      var subject = '【徐嘉賢診所】您的氣喘控制測驗報告 - ' + name;
      var htmlBody = buildReportHtml(name, testType, score, level, levelDesc, weatherSuggestion);
      GmailApp.sendEmail(email, subject, '', {
        htmlBody: htmlBody,
        name: '徐嘉賢診所'
      });
      status = 'sent';
      detail = 'accepted_by_gmail';
  } catch (mailErr) {
    status = 'failed';
    detail = String(mailErr);
  }
  logEmailAttempt(name, email, testType, score, status, detail);
}


/**
 * Append one row to the 'Email Log' tab. Creates the tab on first call.
 * Wrapped in try/catch so logging failure never breaks the submission.
 */
function logEmailAttempt(name, email, testType, score, status, detail) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Email Log');
    if (!sheet) {
      sheet = ss.insertSheet('Email Log');
      sheet.appendRow(['Timestamp', 'Name', 'Email', 'TestType', 'Score', 'Status', 'Detail']);
      sheet.setFrozenRows(1);
    }
    sheet.appendRow([new Date(), name, email, testType, score, status, detail]);
  } catch (e) {
    // Silent — never break the main submission flow
  }
}


/**
 * Return today's ACT records as JSON for the nurse list view.
 */
function handleNurseList(params) {
  requireNursePin_(params);
  var today = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd');
  if (params.date && params.date !== today) {
    throw pfError_('DATE_NOT_ALLOWED', '目前僅能查看及修改今天的紀錄。', false);
  }
  var sheet = nurseSheet_();
  var data = sheet.getDataRange().getValues();
  var rows = [];
  var invalidTimestampCount = 0;

  // Row 0 is header, data starts at row 1
  for (var i = 1; i < data.length; i++) {
    if (!data[i].some(function(value) { return value !== '' && value != null; })) continue;
    var recordDate = taipeiRecordDate_(data[i][0]);
    if (!recordDate) { invalidTimestampCount++; continue; }
    if (recordDate !== today) continue;

    rows.push({
      rowIndex: i + 1,  // 1-based sheet row number
      rowKey: nurseRowKey_(i + 1, data[i]),
      name: data[i][3] || '',
      chart: data[i][2] || '',
      testType: data[i][6] || '',
      score: data[i][7] == null ? '' : data[i][7],
      level: data[i][8] || '',
      levelKey: data[i][9] || '',
      peakFlow: data[i][15] == null ? '' : data[i][15]
    });
  }

  return { ok: true, apiVersion: ACT_PF_API_VERSION, date: today,
    rows: rows, invalidTimestampCount: invalidTimestampCount };
}


/**
 * Update peak flow value for a specific row (nurse retroactive entry).
 */
function handleUpdatePF(params, qaSheet) {
  requireNursePin_(params);
  var rowText = String(params.rowIndex || '');
  var pfText = String(params.peakFlow == null ? '' : params.peakFlow).trim();
  if (!/^[1-9]\d*$/.test(rowText) || !Number.isSafeInteger(Number(rowText)) || Number(rowText) < 2) {
    throw pfError_('INVALID_ROW', '紀錄位置無效，請重新整理清單。', false);
  }
  if (!/^\d+$/.test(pfText) || Number(pfText) < 50 || Number(pfText) > 900) {
    throw pfError_('INVALID_PF', '請輸入 50–900 的整數 Peak Flow 數值。', false);
  }
  if (!/^[a-f0-9]{64}$/.test(String(params.rowKey || ''))) {
    throw pfError_('INVALID_RECORD_KEY', '缺少有效的紀錄識別，請重新整理清單。', false);
  }

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    throw pfError_('BUSY', '目前有人正在儲存，請稍後重新整理再試。', true);
  }
  try {
    var rowIndex = Number(rowText);
    var peakFlow = Number(pfText);
    // qaSheet is supplied only by the manually run synthetic self-test below.
    // The public doPost caller never accepts a sheet name or sheet override.
    var sheet = qaSheet || nurseSheet_();
    if (rowIndex > sheet.getLastRow()) {
      throw pfError_('RECORD_NOT_FOUND', '找不到這筆紀錄，請重新整理清單。', false);
    }
    var row = sheet.getRange(rowIndex, 1, 1, 16).getValues()[0];
    if (nurseRowKey_(rowIndex, row) !== params.rowKey) {
      throw pfError_('RECORD_CHANGED', '這筆紀錄已變動，請重新整理並確認病患後再輸入。', false);
    }
    var today = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd');
    if (taipeiRecordDate_(row[0]) !== today || (params.date && params.date !== today)) {
      throw pfError_('DATE_NOT_ALLOWED', '目前僅能修改今天的紀錄，請重新整理清單。', false);
    }

    // The established A:P schema is unchanged; only column P is written.
    sheet.getRange(rowIndex, 16).setValue(peakFlow);
    SpreadsheetApp.flush();
    var confirmedRow = sheet.getRange(rowIndex, 1, 1, 16).getValues()[0];
    if (Number(confirmedRow[15]) !== peakFlow ||
        nurseRowKey_(rowIndex, confirmedRow.slice(0, 15).concat([row[15]])) !== params.rowKey) {
      throw pfError_('WRITE_NOT_CONFIRMED', '未能確認儲存結果，請重新整理查看，勿重複送出。', true);
    }
    return { ok: true, apiVersion: ACT_PF_API_VERSION, rowIndex: rowIndex,
      previousRowKey: params.rowKey, rowKey: nurseRowKey_(rowIndex, confirmedRow), peakFlow: peakFlow };
  } finally {
    lock.releaseLock();
  }
}

/** Nurse operations use simple form POSTs; there is no JSONP or no-cors success fallback. */
function doPost(e) {
  var params = (e && e.parameter) || {};
  try {
    if (params.mode === 'nurse_history') return pfJson_(handleNurseHistory_(params));
    if (params.mode === 'pf_record') return pfJson_(handlePFRecord_(params));
    if (params.mode === 'pf_sync') return pfJson_(handlePFSync_(params));
    if (params.mode === 'pf_events') return pfJson_(handlePFEvents_(params));
    if (params.mode === 'pf_ack') return pfJson_(handlePFAck_(params));
    if (params.mode === 'nurse_list') return pfJson_(handleNurseList(params));
    if (params.mode === 'update_pf') return pfJson_(handleUpdatePF(params));
    return pfJson_(pfFailure_('INVALID_MODE', '不支援的操作。', false));
  } catch (err) {
    return pfJson_(pfFailure_(err.pfCode || 'BACKEND_UNAVAILABLE',
      err.pfCode ? err.message : '無法存取 Google 服務，請確認後台授權與部署設定。',
      err.pfCode ? !!err.pfRetryable : true));
  }
}

function pfJson_(value) {
  return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON);
}

function pfFailure_(code, message, retryable) {
  return { ok: false, apiVersion: ACT_PF_API_VERSION,
    error: { code: code, message: message, retryable: retryable } };
}

function pfError_(code, message, retryable) {
  var error = new Error(message);
  error.pfCode = code;
  error.pfRetryable = retryable;
  return error;
}

function requireNursePin_(params) {
  var configured = PropertiesService.getScriptProperties().getProperty('ACT_NURSE_PIN');
  var expected = configured === null ? ACT_LEGACY_NURSE_PIN : configured;
  if (!expected) throw pfError_('AUTH_NOT_CONFIGURED', '管理密碼尚未設定，請聯絡管理者。', false);
  var supplied = String((params && params.pin) || '');
  var difference = supplied.length ^ expected.length;
  for (var i = 0; i < expected.length; i++) difference |= expected.charCodeAt(i) ^ (supplied.charCodeAt(i) || 0);
  if (difference !== 0) throw pfError_('AUTH_FAILED', '管理密碼錯誤，請重新輸入。', false);
}

function nurseSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss && ss.getSheetByName('ACT Records');
  if (!sheet) throw pfError_('SHEET_NOT_FOUND', '找不到 ACT Records 工作表，請聯絡管理者。', false);
  if (sheet.getLastColumn() < 16) {
    throw pfError_('SCHEMA_MISMATCH', 'ACT Records 欄位不完整，請聯絡管理者。', false);
  }
  return sheet;
}

function nurseRowKey_(rowIndex, row) {
  // A:P is a version key: concurrent PF changes require a fresh list.
  var cells = row.slice(0, 16).map(function(value) {
    return value instanceof Date ? value.toISOString() : String(value == null ? '' : value);
  });
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,
    String(rowIndex) + '\n' + JSON.stringify(cells), Utilities.Charset.UTF_8);
  return digest.map(function(value) { return ('0' + (value & 255).toString(16)).slice(-2); }).join('');
}

function taipeiRecordDate_(value) {
  var date;
  if (value instanceof Date) {
    date = value;
  } else {
    var text = String(value == null ? '' : value).trim();
    // Existing submissions are ISO UTC. Legacy unzoned sheet text is Taipei.
    if (/^\d{4}-\d{2}-\d{2}T.+(?:Z|[+-]\d{2}:?\d{2})$/i.test(text)) {
      date = new Date(text);
    } else {
      var local = /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(text);
      if (!local) return '';
      date = new Date(local[1] + '-' + ('0' + local[2]).slice(-2) + '-' + ('0' + local[3]).slice(-2) +
        'T' + ('0' + (local[4] || '0')).slice(-2) + ':' + (local[5] || '00') + ':' + (local[6] || '00') + '+08:00');
    }
  }
  return isNaN(date.getTime()) ? '' : Utilities.formatDate(date, 'Asia/Taipei', 'yyyy-MM-dd');
}

/** Run in the editor as owner to request/check authorization. Never sends or writes. */
function readOnlyAuthorizationCheck() {
  var report = { release: ACT_PF_RELEASE, sheets: { ok: false }, gmail: { ok: false } };
  try {
    var sheet = nurseSheet_();
    report.sheets = { ok: true, columns: sheet.getLastColumn() };
  } catch (err) { report.sheets.code = err.pfCode || 'AUTH_OR_ACCESS_REQUIRED'; }
  try {
    GmailApp.getAliases(); // Check the Gmail authorization used by existing reports; do not return aliases.
    report.gmail = { ok: true };
  } catch (err) { report.gmail.code = 'AUTH_OR_ACCESS_REQUIRED'; }
  console.log(JSON.stringify(report));
  return report;
}

/** Manually run as owner. Writes only to a dedicated synthetic QA worksheet. */
function runIsolatedPFCloudSelfTest() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var qaName = 'ACT Repair QA 20261124';
  var qa = ss.getSheetByName(qaName) || ss.insertSheet(qaName);
  if (qa.getName() !== qaName) throw new Error('QA sheet guard failed');
  if (qa.getMaxColumns() < 16) qa.insertColumnsAfter(qa.getMaxColumns(), 16 - qa.getMaxColumns());
  if (qa.getLastRow() === 0) {
    qa.appendRow(['Timestamp', 'PatientID', 'Chart', 'Name', 'DOB', 'Age', 'Type', 'Score',
      'Level', 'LevelKey', 'Temp', 'Humidity', 'Weather', 'Email', 'Consent', 'Peak Flow']);
  }
  var row = [new Date().toISOString(), 'QA-PF-NO-PATIENT', 'QA', 'Synthetic PF fixture',
    '', 'QA', 'ACT', 20, 'QA', 'good', '', '', '', '', '', ''];
  qa.appendRow(row);
  SpreadsheetApp.flush();
  var rowIndex = qa.getLastRow();
  var storedRow = qa.getRange(rowIndex, 1, 1, 16).getValues()[0];
  var configured = PropertiesService.getScriptProperties().getProperty('ACT_NURSE_PIN');
  var pin = configured === null ? ACT_LEGACY_NURSE_PIN : configured;
  var result = handleUpdatePF({ pin: pin, rowIndex: String(rowIndex),
    rowKey: nurseRowKey_(rowIndex, storedRow), peakFlow: '321' }, qa);
  var report = { ok: result.ok && result.peakFlow === 321, release: ACT_PF_RELEASE,
    sheet: qaName, fixtureRow: rowIndex, confirmedPeakFlow: result.peakFlow };
  console.log(JSON.stringify(report));
  return report;
}


// Independent measurements. ACT Records A:P is never extended or projected from this registry.
var PF_HISTORY_HEADERS = ['event_id', 'chart_number', 'peak_flow', 'assessment_date',
  'measured_at', 'recorded_at', 'received_at', 'source', 'source_event_id', 'overlay_id',
  'sync_status', 'committed_at', 'schema_version', 'linked_act_identity', 'payload_hash',
  'updated_at', 'error_code', 'error_message'];

function historyHash_(text) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8)
    .map(function(value) { return ('0' + (value & 255).toString(16)).slice(-2); }).join('');
}

function requireHistorySecret_(params, sync) {
  var key = sync ? 'ACT_PF_SYNC_KEY' : 'ACT_NURSE_PIN';
  var configured = PropertiesService.getScriptProperties().getProperty(key);
  // Staff PINs are accepted only from an explicit property, never the legacy fallback.
  // The independently configured service key keeps its stronger minimum length.
  if (!configured || !configured.trim() || configured.length < (sync ? 32 : 4)) {
    throw pfError_('AUTH_NOT_CONFIGURED', '歷史 PF 存取尚未完成憑證設定，請聯絡管理者。', false);
  }
  var supplied = String(params[sync ? 'syncKey' : 'pin'] || '');
  var difference = supplied.length ^ configured.length;
  for (var i = 0; i < configured.length; i++) difference |= configured.charCodeAt(i) ^ (supplied.charCodeAt(i) || 0);
  if (difference !== 0) throw pfError_('AUTH_FAILED', '存取憑證錯誤，請重新輸入。', false);
}

function historyToday_() { return Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd'); }

function historyChart_(value, optional) {
  var text = String(value == null ? '' : value).trim();
  if (!text && optional) return '';
  if (!/^\d{1,7}$/.test(text) || /^0+$/.test(text)) {
    throw pfError_('INVALID_CHART', '請輸入 1–7 位數字病歷號。', false);
  }
  return ('0000000' + text).slice(-7);
}

function historyDate_(value) {
  var text = String(value == null ? '' : value);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(text) || text < '1900-01-01') {
    throw pfError_('INVALID_DATE', '請輸入有效的西元日期。', false);
  }
  var parsed = new Date(text + 'T00:00:00+08:00');
  if (isNaN(parsed.getTime()) || Utilities.formatDate(parsed, 'Asia/Taipei', 'yyyy-MM-dd') !== text) {
    throw pfError_('INVALID_DATE', '請輸入有效的西元日期。', false);
  }
  if (text > historyToday_()) throw pfError_('FUTURE_DATE', '不能登記或查詢未來日期。', false);
  return text;
}

function historyISO_(value, required) {
  var text = String(value == null ? '' : value).trim();
  if (!text && !required) return '';
  if (!/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d+)?(?:Z|[+-]\d{2}:\d{2})$/i.test(text)) {
    throw pfError_('INVALID_TIMESTAMP', '時間必須包含明確時區。', false);
  }
  var parsed = new Date(text);
  var calendar = new Date(text.slice(0, 10) + 'T00:00:00Z');
  if (isNaN(parsed.getTime()) || isNaN(calendar.getTime()) || calendar.toISOString().slice(0, 10) !== text.slice(0, 10) ||
      Utilities.formatDate(parsed, 'Asia/Taipei', 'yyyy-MM-dd') > historyToday_()) {
    throw pfError_('INVALID_TIMESTAMP', '時間無效或位於未來日期。', false);
  }
  return parsed.toISOString();
}

function historyInteger_(value, fallback, min, max) {
  if (value === undefined || value === null || value === '') return fallback;
  var text = String(value);
  if (!/^\d+$/.test(text) || !Number.isSafeInteger(Number(text)) || Number(text) < min || Number(text) > max) {
    throw pfError_('INVALID_PAGINATION', '查詢分頁參數無效。', false);
  }
  return Number(text);
}

function historyRegistry_(create, override) {
  var ss = override ? null : SpreadsheetApp.getActiveSpreadsheet();
  var sheet = override || ss.getSheetByName('PF Records');
  if (!sheet && create) sheet = ss.insertSheet('PF Records');
  if (!sheet) throw pfError_('PF_REGISTRY_NOT_INITIALIZED', 'PF 紀錄尚未初始化，請聯絡管理者。', false);
  if (sheet.getMaxColumns() < PF_HISTORY_HEADERS.length && create) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), PF_HISTORY_HEADERS.length - sheet.getMaxColumns());
  }
  if (sheet.getLastRow() === 0 && create) sheet.getRange(1, 1, 1, PF_HISTORY_HEADERS.length).setValues([PF_HISTORY_HEADERS]);
  if (sheet.getLastColumn() < PF_HISTORY_HEADERS.length) {
    throw pfError_('PF_SCHEMA_MISMATCH', '獨立 PF 欄位不完整，停止存取以保護資料。', false);
  }
  var headers = sheet.getRange(1, 1, 1, PF_HISTORY_HEADERS.length).getValues()[0];
  for (var i = 0; i < PF_HISTORY_HEADERS.length; i++) {
    if (headers[i] !== PF_HISTORY_HEADERS[i]) throw pfError_('PF_SCHEMA_MISMATCH', '獨立 PF 欄位順序不符，停止存取以保護資料。', false);
  }
  return sheet;
}

function historyCellText_(value) { return value instanceof Date ? value.toISOString() : String(value == null ? '' : value); }
function historyCellDate_(value) { return value instanceof Date ? Utilities.formatDate(value, 'Asia/Taipei', 'yyyy-MM-dd') : String(value || ''); }

function historyVersion_(rowIndex, row) {
  return historyHash_(String(rowIndex) + '\n' + JSON.stringify(row.slice(0, PF_HISTORY_HEADERS.length).map(historyCellText_)));
}

function historyRecord_(rowIndex, row) {
  var canonicalSource = String(row[7]);
  var canonical = { eventId: String(row[0]), chartNumber: historyChart_(row[1], false), peakFlow: Number(row[2]),
    assessmentDate: historyCellDate_(row[3]), measuredAt: historyCellText_(row[4]),
    recordedAt: canonicalSource === 'overlay_manual' ? historyCellText_(row[5]) : '',
    source: canonicalSource, sourceEventId: String(row[8]),
    overlayId: canonicalSource === 'overlay_manual' ? String(row[9] || '') : '', linkedActIdentity: String(row[13] || '') };
  if (['mobile_pf', 'overlay_manual'].indexOf(canonicalSource) < 0 ||
      historyHash_(JSON.stringify(canonical)) !== row[14]) {
    throw pfError_('PF_RECORD_INTEGRITY', '獨立 PF 內容與原事件不一致，請聯絡管理者核對。', false);
  }
  return { eventId: String(row[0]), chartNumber: historyChart_(row[1], false), peakFlow: Number(row[2]),
    assessmentDate: historyCellDate_(row[3]), measuredAt: historyCellText_(row[4]),
    recordedAt: historyCellText_(row[5]), receivedAt: historyCellText_(row[6]), source: String(row[7]),
    sourceEventId: String(row[8]), overlayId: String(row[9] || ''), syncStatus: String(row[10]),
    committedAt: historyCellText_(row[11]), schemaVersion: Number(row[12]), linkedActIdentity: String(row[13] || ''),
    rowIndex: rowIndex, versionKey: historyVersion_(rowIndex, row),
    errorCode: String(row[16] || ''), errorMessage: String(row[17] || '') };
}

function historyActChart_(row) {
  var chart = '', fromId = /^CHART-(\d{1,7})$/.exec(String(row[1] || ''));
  try { chart = historyChart_(row[2], true); } catch (err) { return ''; }
  var idChart = fromId ? historyChart_(fromId[1], false) : '';
  return chart && idChart && chart !== idChart ? '' : chart || idChart;
}

function historySnapshotVersion_(act, events, actRows, pfRows) {
  return historyHash_(JSON.stringify([
    act.slice(0, actRows).map(function(row) { return row.slice(0, 16).map(historyCellText_); }),
    events.slice(0, pfRows).map(function(row) {
      // Delivery acknowledgments may change while paging; measurement facts may not.
      return row.slice(0, 9).concat(row.slice(12, 15)).map(historyCellText_);
    })
  ]));
}

function historyCursor_(text, query, act, events) {
  if (!text) return { offset: 0, actRows: act.length, pfRows: events.length, query: query,
    version: historySnapshotVersion_(act, events, act.length, events.length) };
  var value;
  try {
    if (String(text).slice(0, 3) !== 'v1.') throw new Error('Invalid cursor');
    value = JSON.parse(decodeURIComponent(String(text).slice(3)));
  } catch (err) { throw pfError_('INVALID_PAGINATION', '分頁識別無效，請重新查詢。', false); }
  if (!value || value.query !== query || !Number.isSafeInteger(value.offset) || value.offset < 0 || value.offset > 1000000 ||
      !Number.isSafeInteger(value.actRows) || value.actRows < 1 || !Number.isSafeInteger(value.pfRows) || value.pfRows < 1) {
    throw pfError_('INVALID_PAGINATION', '分頁識別與目前查詢不符，請重新查詢。', false);
  }
  if (value.actRows > act.length || value.pfRows > events.length ||
      value.version !== historySnapshotVersion_(act, events, value.actRows, value.pfRows)) {
    throw pfError_('HISTORY_CHANGED_RELOAD', '原查詢資料已變動，請重新查詢以避免遺漏紀錄。', false);
  }
  return value;
}

function handleNurseHistory_(params) {
  requireHistorySecret_(params, false);
  var chart = historyChart_(params.chartNumber, true);
  var today = historyToday_();
  var weekAgo = new Date(today + 'T00:00:00+08:00');
  weekAgo.setUTCDate(weekAgo.getUTCDate() - 6);
  var start = historyDate_(params.startDate || (chart ? Utilities.formatDate(weekAgo, 'Asia/Taipei', 'yyyy-MM-dd') : today));
  var end = historyDate_(params.endDate || today);
  var days = (new Date(end + 'T00:00:00Z') - new Date(start + 'T00:00:00Z')) / 86400000 + 1;
  if (days < 1 || days > 31) throw pfError_('DATE_RANGE_TOO_LARGE', '每次請查詢 1–31 天。', false);
  if (!chart && (start !== today || end !== today)) {
    throw pfError_('CHART_REQUIRED_FOR_HISTORY', '查詢過去紀錄時請先輸入病歷號。', false);
  }
  var limit = historyInteger_(params.limit, 100, 1, 200);
  var act = nurseSheet_().getDataRange().getValues();
  var events = historyRegistry_(false).getDataRange().getValues();
  var snapshot = historyCursor_(params.cursor, historyHash_(chart + '|' + start + '|' + end), act, events);
  var cursor = snapshot.offset;
  var items = [], invalidTimestampCount = 0;
  for (var i = 1; i < snapshot.actRows; i++) {
    var actChart = historyActChart_(act[i]);
    if (chart && actChart !== chart) continue;
    if (!act[i].some(function(value) { return value !== '' && value != null; })) continue;
    var date = taipeiRecordDate_(act[i][0]);
    if (!date) { invalidTimestampCount++; continue; }
    if (date < start || date > end) continue;
    var timestamp = historyCellText_(act[i][0]);
    items.push({ kind: 'ACT', rowIndex: i + 1, rowKey: nurseRowKey_(i + 1, act[i]),
      stableKey: historyHash_(JSON.stringify(act[i].slice(0, 15).map(historyCellText_))),
      name: act[i][3] || '', chart: act[i][2] || '', chartNumber: actChart,
      testType: act[i][6] || '', score: act[i][7] == null ? '' : act[i][7],
      level: act[i][8] || '', levelKey: act[i][9] || '', peakFlow: act[i][15] == null ? '' : act[i][15],
      assessmentDate: date, recordedAt: timestamp, timestamp: timestamp });
  }
  for (var j = 1; j < snapshot.pfRows; j++) {
    if (!events[j][0]) continue;
    var event = historyRecord_(j + 1, events[j]);
    if (chart && event.chartNumber !== chart) continue;
    if (event.assessmentDate < start || event.assessmentDate > end) continue;
    event.kind = 'PF'; items.push(event);
  }
  items.sort(function(a, b) {
    return b.assessmentDate.localeCompare(a.assessmentDate) || b.recordedAt.localeCompare(a.recordedAt) ||
      String(a.kind + ':' + (a.eventId || a.rowIndex)).localeCompare(String(b.kind + ':' + (b.eventId || b.rowIndex)));
  });
  var page = items.slice(cursor, cursor + limit);
  var more = cursor + limit < items.length;
  snapshot.offset = cursor + limit;
  return { ok: true, apiVersion: ACT_PF_API_VERSION, startDate: start, endDate: end, chartNumber: chart,
    rows: page.filter(function(item) { return item.kind === 'ACT'; }),
    pfRecords: page.filter(function(item) { return item.kind === 'PF'; }),
    nextCursor: more ? 'v1.' + encodeURIComponent(JSON.stringify(snapshot)) : null, hasMore: more, invalidTimestampCount: invalidTimestampCount };
}

function historyValidateEvent_(params, source) {
  var id = String(params.eventId || '');
  if (!/^[A-Za-z0-9][A-Za-z0-9:_.-]{7,179}$/.test(id)) throw pfError_('INVALID_EVENT_ID', 'PF 事件識別無效。', false);
  var overlay = source === 'overlay_manual';
  var overlayMatch = /^overlay:clinical_assessments:([1-9]\d*)$/.exec(id);
  if (overlay && (!overlayMatch || !Number.isSafeInteger(Number(overlayMatch[1])))) {
    throw pfError_('INVALID_SOURCE_EVENT', 'Overlay 來源識別無效。', false);
  }
  if (!overlay && /^overlay:/i.test(id)) throw pfError_('INVALID_SOURCE_EVENT', '手機不能冒用 Overlay 來源。', false);
  var sourceId = overlay ? String(params.sourceEventId || '') : id;
  if (overlay && sourceId !== overlayMatch[1] && sourceId !== 'clinical_assessments:' + overlayMatch[1]) {
    throw pfError_('INVALID_SOURCE_EVENT', 'Overlay 來源識別不一致。', false);
  }
  if (overlay) sourceId = overlayMatch[1];
  // 95.0 is the same integer measurement as 95; fractional PF is never rounded.
  var pfText = String(params.peakFlow == null ? '' : params.peakFlow).trim();
  if (!/^\d+(?:\.0+)?$/.test(pfText) || Number(pfText) < 50 || Number(pfText) > 900 || !Number.isInteger(Number(pfText))) {
    throw pfError_('INVALID_PF', '請輸入 50–900 的整數 Peak Flow 數值。', false);
  }
  var date = historyDate_(params.assessmentDate);
  var measured = historyISO_(params.measuredAt, false);
  if (measured && taipeiRecordDate_(measured) !== date) throw pfError_('MEASUREMENT_DATE_MISMATCH', '測量時間與選擇日期不一致。', false);
  var linked = String(params.linkedActIdentity || '');
  if (linked && !/^[a-f0-9]{64}$/.test(linked)) throw pfError_('INVALID_ACT_LINK', '測驗關聯識別無效。', false);
  return { eventId: id, chartNumber: historyChart_(params.chartNumber, false), peakFlow: Number(pfText),
    assessmentDate: date, measuredAt: measured, recordedAt: overlay ? historyISO_(params.recordedAt, true) : '',
    source: source, sourceEventId: sourceId, overlayId: overlay ? overlayMatch[1] : '', linkedActIdentity: linked };
}

function historyHasACTPF_(event) {
  var rows = nurseSheet_().getDataRange().getValues();
  return rows.slice(1).some(function(row) {
    return historyActChart_(row) === event.chartNumber && taipeiRecordDate_(row[0]) === event.assessmentDate &&
      row[15] !== '' && row[15] != null;
  });
}

function historyWriteEvent_(params, source, qaSheet) {
  var event = historyValidateEvent_(params, source);
  var payloadHash = historyHash_(JSON.stringify(event));
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) throw pfError_('BUSY', '目前有人正在儲存，請稍後再試。', true);
  try {
    var sheet = historyRegistry_(false, qaSheet);
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) !== event.eventId) continue;
      if (String(data[i][14]) !== payloadHash) throw pfError_('IDEMPOTENCY_CONFLICT', '同一事件識別的內容不同，未新增或覆寫資料。', false);
      return { ok: true, apiVersion: ACT_PF_API_VERSION, duplicate: true, record: historyRecord_(i + 1, data[i]) };
    }
    if (source === 'mobile_pf' && params.confirmAdditional !== 'yes') {
      var exists = data.slice(1).some(function(row) {
        return historyChart_(row[1], true) === event.chartNumber && historyCellDate_(row[3]) === event.assessmentDate;
      });
      if (exists || (!qaSheet && historyHasACTPF_(event))) {
        throw pfError_('ADDITIONAL_CONFIRMATION_REQUIRED', '此病歷當天已有 PF，請確認這是另一次測量後再新增。', false);
      }
    }
    var now = new Date().toISOString();
    var row = [event.eventId, event.chartNumber, event.peakFlow, event.assessmentDate, event.measuredAt,
      event.recordedAt || now, now, event.source, event.sourceEventId, event.overlayId,
      source === 'overlay_manual' ? 'synced' : 'pending', source === 'overlay_manual' ? event.recordedAt : '',
      1, event.linkedActIdentity, payloadHash, now, '', ''];
    var rowIndex = sheet.getLastRow() + 1;
    if (rowIndex > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), rowIndex - sheet.getMaxRows());
    var range = sheet.getRange(rowIndex, 1, 1, PF_HISTORY_HEADERS.length);
    range.setNumberFormat('@'); range.setValues([row]);
    SpreadsheetApp.flush();
    var confirmed = sheet.getRange(rowIndex, 1, 1, PF_HISTORY_HEADERS.length).getValues()[0];
    if (historyVersion_(rowIndex, confirmed) !== historyVersion_(rowIndex, row)) {
      throw pfError_('WRITE_NOT_CONFIRMED', '未能確認 PF 已儲存，請保留事件識別後重新查詢。', true);
    }
    return { ok: true, apiVersion: ACT_PF_API_VERSION, duplicate: false, record: historyRecord_(rowIndex, confirmed) };
  } finally { lock.releaseLock(); }
}

function handlePFRecord_(params) { requireHistorySecret_(params, false); return historyWriteEvent_(params, 'mobile_pf'); }
function handlePFSync_(params) { requireHistorySecret_(params, true); return historyWriteEvent_(params, 'overlay_manual'); }

function handlePFEvents_(params) {
  requireHistorySecret_(params, true);
  var after = historyInteger_(params.afterRow, 1, 1, 1000000000);
  var limit = historyInteger_(params.limit, 50, 1, 200);
  var data = historyRegistry_(false).getDataRange().getValues();
  var pending = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][7] === 'mobile_pf' && data[i][10] === 'pending') pending.push(historyRecord_(i + 1, data[i]));
  }
  // afterRow is a scan hint, never an acknowledgment: earlier unacked work cannot be lost.
  var page = pending.slice(0, limit);
  return { ok: true, apiVersion: ACT_PF_API_VERSION, events: page, afterRow: after,
    nextAfterRow: page.length ? page[page.length - 1].rowIndex : after,
    hasMore: pending.length > limit, lastScannedRow: data.length };
}

function historyAck_(params, qaSheet) {
  var id = String(params.eventId || '');
  var status = params.status || 'committed_overlay';
  if (!/^[A-Za-z0-9][A-Za-z0-9:_.-]{7,179}$/.test(id)) throw pfError_('INVALID_EVENT_ID', 'PF 事件識別無效。', false);
  if (status !== 'committed_overlay' && status !== 'rejected_patient') throw pfError_('INVALID_ACK', '同步確認狀態無效。', false);
  var overlayId = String(params.overlayId || '');
  var committedAt = '';
  var errorCode = '';
  if (status === 'committed_overlay') {
    if (!/^[1-9]\d*$/.test(overlayId) || !Number.isSafeInteger(Number(overlayId))) throw pfError_('INVALID_ACK', '缺少有效 Overlay 提交識別。', false);
    committedAt = historyISO_(params.committedAt, true);
  } else {
    if (overlayId || params.committedAt) throw pfError_('INVALID_ACK', '未確認病患的事件不能標示 Overlay 提交。', false);
    errorCode = String(params.errorCode || 'PATIENT_NOT_FOUND');
    if (['PATIENT_NOT_FOUND', 'PATIENT_AMBIGUOUS'].indexOf(errorCode) < 0) throw pfError_('INVALID_ACK', '病患核對錯誤代碼無效。', false);
  }
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) throw pfError_('BUSY', '同步確認忙碌，請稍後再試。', true);
  try {
    var sheet = historyRegistry_(false, qaSheet);
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) !== id) continue;
      var row = rows[i].slice(0, PF_HISTORY_HEADERS.length);
      var target = status === 'committed_overlay' ? 'synced' : 'rejected_patient';
      if (row[10] !== 'pending') {
        if (row[10] === target && String(row[9] || '') === overlayId && String(row[16] || '') === errorCode) {
          return { ok: true, apiVersion: ACT_PF_API_VERSION, duplicate: true, eventId: id, record: historyRecord_(i + 1, row) };
        }
        throw pfError_('ACK_CONFLICT', '此事件已有不同的同步結果，停止覆寫。', false);
      }
      row[9] = overlayId; row[10] = target; row[11] = committedAt; row[15] = new Date().toISOString();
      row[16] = errorCode; row[17] = errorCode ? '病歷核對未通過，需人工確認；未寫入 Overlay。' : '';
      sheet.getRange(i + 1, 1, 1, PF_HISTORY_HEADERS.length).setValues([row]);
      SpreadsheetApp.flush();
      var confirmed = sheet.getRange(i + 1, 1, 1, PF_HISTORY_HEADERS.length).getValues()[0];
      if (historyVersion_(i + 1, confirmed) !== historyVersion_(i + 1, row)) throw pfError_('WRITE_NOT_CONFIRMED', '未能確認同步狀態，請以相同事件識別重試。', true);
      return { ok: true, apiVersion: ACT_PF_API_VERSION, duplicate: false, eventId: id, record: historyRecord_(i + 1, confirmed) };
    }
    throw pfError_('RECORD_NOT_FOUND', '找不到 PF 事件。', false);
  } finally { lock.releaseLock(); }
}

function handlePFAck_(params) { requireHistorySecret_(params, true); return historyAck_(params); }

/** Owner only, manual editor invocation. Never logs or returns generated credentials. */
function initializePFHistoryAccess() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) throw new Error('Initialization busy');
  try {
    var properties = PropertiesService.getScriptProperties();
    var nurse = properties.getProperty('ACT_NURSE_PIN');
    var sync = properties.getProperty('ACT_PF_SYNC_KEY');
    var pinCreated = nurse === null;
    var keyCreated = sync === null;
    if (pinCreated) {
      nurse = ('000000000000' + parseInt(Utilities.getUuid().replace(/-/g, '').slice(0, 12), 16).toString()).slice(-12);
      properties.setProperty('ACT_NURSE_PIN', nurse);
    }
    if (keyCreated) {
      sync = 'pf_' + Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
      properties.setProperty('ACT_PF_SYNC_KEY', sync);
    }
    historyRegistry_(true);
    var nurseConfigured = !!nurse.trim() && nurse.length >= 4;
    var syncConfigured = !!sync.trim() && sync.length >= 32;
    var report = { ok: nurseConfigured && syncConfigured,
      release: ACT_PF_RELEASE, nursePinCreated: pinCreated, syncKeyCreated: keyCreated,
      nursePinConfigured: nurseConfigured,
      syncKeyConfigured: syncConfigured, registryReady: true };
    console.log(JSON.stringify(report)); return report;
  } finally { lock.releaseLock(); }
}

/** All writes stay on the named synthetic registry. No ACT lookup, Overlay call, or email. */
function runIsolatedPFHistoryCloudSelfTest() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var name = 'ACT PF History QA 20261124';
  var qa = ss.getSheetByName(name) || ss.insertSheet(name);
  if (qa.getName() !== name) throw new Error('QA registry guard failed');
  historyRegistry_(true, qa);
  var eventId = 'qa-history-' + Utilities.getUuid();
  var params = { eventId: eventId, chartNumber: '0000001', peakFlow: '95', assessmentDate: historyToday_() };
  var first = historyWriteEvent_(params, 'mobile_pf', qa);
  var retry = historyWriteEvent_(params, 'mobile_pf', qa);
  var ack = historyAck_({ eventId: eventId, overlayId: '999999999', committedAt: new Date().toISOString() }, qa);
  var report = { ok: first.ok && retry.duplicate && ack.record.syncStatus === 'synced', release: ACT_PF_RELEASE,
    syntheticOnly: true, sheet: name, registryReadback: first.record.peakFlow === 95,
    idempotentRetry: retry.duplicate, syntheticAckRoundtrip: ack.ok };
  console.log(JSON.stringify(report)); return report;
}

function buildReportHtml(name, testType, score, level, levelDesc, weatherSuggestion) {
  var levelColor = '#22c55e';
  if (level.indexOf('部分') >= 0 || level.indexOf('未完全') >= 0) levelColor = '#eab308';
  if (level.indexOf('不佳') >= 0) levelColor = '#ef4444';

  var suggestion = weatherSuggestion
    ? '<p style="margin:12px 0 0;padding:10px;background:#f0f9ff;border-radius:8px;font-size:14px;color:#0c4a6e;">今日建議：' + weatherSuggestion + '</p>'
    : '';

  var ctaHtml = '<div style="background:#FFB640;padding:24px;text-align:center;border-radius:12px;margin-top:20px;">' +
    '<p style="color:#1a1a1a;font-size:16px;font-weight:bold;margin:0 0 12px;">需要專業諮詢？</p>' +
    '<a href="http://08143.vision.com.tw/Register" style="display:inline-block;background:white;color:#002840;padding:10px 24px;border-radius:24px;text-decoration:none;font-weight:bold;font-size:14px;">線上掛號</a>' +
    '<div style="margin-top:12px;">' +
    '<p style="color:#1a1a1a;font-size:13px;font-weight:bold;margin:0;">徐嘉賢診所</p>' +
    '<p style="color:#1a1a1a;font-size:12px;margin:4px 0;">地址：811高雄市楠梓區藍昌路413號</p>' +
    '<p style="color:#1a1a1a;font-size:12px;margin:4px 0;">電話：07-3613338</p>' +
    '<p style="color:#1a1a1a;font-size:12px;margin:4px 0;">' +
    '<a href="https://drkentsui.com" style="color:#002840;font-weight:bold;text-decoration:underline;">drkentsui.com</a>' +
    '</p></div></div>';

  return '<div style="font-family: \'Noto Sans TC\', sans-serif; max-width: 560px; margin: 0 auto; padding: 24px; color: #1a1a1a;">' +
    '<div style="text-align: center; padding: 20px; background: #002840; border-radius: 12px; margin-bottom: 20px;">' +
    '<h1 style="margin: 0; font-size: 1.25rem; color: white;">徐嘉賢診所</h1>' +
    '<p style="margin: 4px 0 0; font-size: 0.9rem; color: #FFB640; font-weight: bold;">過敏性氣喘｜胸腔專科診所</p>' +
    '</div>' +
    '<p>' + name + ' 您好：</p>' +
    '<p>以下是您的氣喘控制測驗結果，僅供參考，建議與醫師討論後續照護。</p>' +
    '<div style="text-align: center; padding: 24px; margin: 20px 0; border-radius: 16px; background: #f8fafc; border: 2px solid ' + levelColor + ';">' +
    '<p style="margin: 0 0 4px; font-size: 0.9rem; color: #47607A;">' + testType + ' 總分</p>' +
    '<p style="margin: 0; font-size: 2rem; font-weight: 700; color: ' + levelColor + ';">' + score + '</p>' +
    '<p style="margin: 8px 0 0; font-weight: 500;">' + level + '：' + levelDesc + '</p>' +
    '</div>' +
    suggestion +
    ctaHtml +
    '<p style="margin-top: 20px; font-size: 0.8rem; color: #47607A; text-align: center;">本結果僅供參考，請與醫師討論後續照護計畫。<br>© 徐嘉賢診所</p>' +
    '</div>';
}
