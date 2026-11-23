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
var ACT_PF_RELEASE = '2026-11-24-pf1';
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
