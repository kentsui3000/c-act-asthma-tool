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

function doGet(e) {
  var params = e.parameter;
  var mode = params.mode || '';

  // ── Nurse mode: return today's records as JSON ──
  if (mode === 'nurse_list') {
    return handleNurseList();
  }

  // ── Nurse mode: update peak flow for a specific row ──
  if (mode === 'update_pf') {
    return handleUpdatePF(params);
  }

  // ── Normal submission ──
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
      var subject = '【徐嘉賢診所】您的氣喘控制測驗報告 - ' + name;
      var htmlBody = buildReportHtml(name, testType, score, level, levelDesc, weatherSuggestion);
      GmailApp.sendEmail(email, subject, '', {
        htmlBody: htmlBody,
        name: '徐嘉賢診所'
      });
    }
  } catch (err) {
    result = 'error: ' + err.toString();
  }
  return ContentService.createTextOutput(result).setMimeType(ContentService.MimeType.TEXT);
}


/**
 * Return today's ACT records as JSON for the nurse list view.
 */
function handleNurseList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('ACT Records');
  if (!sheet) {
    return ContentService.createTextOutput(JSON.stringify({ rows: [] }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var data = sheet.getDataRange().getValues();
  var today = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd');
  var rows = [];

  // Row 0 is header, data starts at row 1
  for (var i = 1; i < data.length; i++) {
    var timestamp = String(data[i][0] || '');
    // Check if this row is from today
    if (timestamp.substring(0, 10) !== today) continue;

    rows.push({
      rowIndex: i + 1,  // 1-based sheet row number
      name: data[i][3] || '',
      chart: data[i][2] || '',
      testType: data[i][6] || '',
      score: data[i][7] || '',
      level: data[i][8] || '',
      levelKey: data[i][9] || '',
      peakFlow: data[i][15] || ''  // Column P (index 15)
    });
  }

  return ContentService.createTextOutput(JSON.stringify({ rows: rows }))
    .setMimeType(ContentService.MimeType.JSON);
}


/**
 * Update peak flow value for a specific row (nurse retroactive entry).
 */
function handleUpdatePF(params) {
  var result = 'ok';
  try {
    var rowIndex = parseInt(params.rowIndex);
    var peakFlow = params.peakFlow || '';

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('ACT Records');
    if (!sheet) throw new Error('Sheet not found');

    // Column P = column 16 (1-based)
    sheet.getRange(rowIndex, 16).setValue(peakFlow);
  } catch (err) {
    result = 'error: ' + err.toString();
  }
  return ContentService.createTextOutput(result).setMimeType(ContentService.MimeType.TEXT);
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
