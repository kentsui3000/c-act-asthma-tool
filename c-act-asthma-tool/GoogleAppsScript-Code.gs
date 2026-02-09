/**
 * 氣喘控制測驗（C-ACT/ACT）結果寫入 Google 試算表並寄送報告至病患 Email
 * 使用方式：
 * 1. 新增一個 Google 試算表
 * 2. 擴充功能 → Apps Script
 * 3. 將此檔案內容貼上，儲存
 * 4. 部署 → 新增部署 → 類型選「網路應用程式」
 *    - 執行身分：我
 *    - 存取權：任何人
 * 5. 部署後複製「網路應用程式 URL」，貼到 index.html 的 GOOGLE_SCRIPT_URL
 * 注意：寄信會使用「您登入 Apps Script 的 Google 帳號」的 Gmail，請確認該帳號可正常發信。
 */

function doGet(e) {
  var result = 'ok';
  try {
    var params = e.parameter;
    var action = params.action || '';
    var name = params.name || '';
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

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[0];
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        '時間', '姓名', '年齡區間', '測驗類型', '分數', '控制等級', '等級代碼',
        '高雄氣溫°C', '高雄濕度%', '天氣建議', 'Email', '同意寄送報告及行銷'
      ]);
    }

    var ageLabel = ageGroup === 'child' ? '4～11歲（兒童）' : ageGroup === 'adult' ? '12歲以上（成人）' : ageGroup;
    sheet.appendRow([
      timestamp,
      name,
      ageLabel,
      testType,
      score,
      level,
      levelKey,
      temp,
      humidity,
      weatherSuggestion,
      email,
      consent === 'yes' ? '是' : ''
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

function buildReportHtml(name, testType, score, level, levelDesc, weatherSuggestion) {
  var levelColor = '#22c55e';
  if (level.indexOf('部分') >= 0 || level.indexOf('未完全') >= 0) levelColor = '#eab308';
  if (level.indexOf('不佳') >= 0) levelColor = '#ef4444';
  var suggestion = weatherSuggestion ? '<p style="margin:12px 0 0;padding:10px;background:#f0f9ff;border-radius:8px;font-size:14px;color:#0c4a6e;">💡 今日建議：' + weatherSuggestion + '</p>' : '';
  return '<div style="font-family: \'Noto Sans TC\', sans-serif; max-width: 560px; margin: 0 auto; padding: 24px; color: #1e293b;">' +
    '<div style="text-align: center; padding-bottom: 16px; border-bottom: 2px solid #e2e8f0;">' +
    '<h1 style="margin: 0; font-size: 1.25rem;">徐嘉賢診所—過敏性氣喘 | 胸腔專科診所</h1>' +
    '<p style="margin: 6px 0 0; font-size: 0.9rem; color: #64748b;">黑眼圈奶爸 Dr. 徐嘉賢醫師</p>' +
    '</div>' +
    '<p style="margin-top: 20px;">' + name + ' 您好：</p>' +
    '<p>以下是您的氣喘控制測驗結果，僅供參考，建議與醫師討論後續照護。</p>' +
    '<div style="text-align: center; padding: 24px; margin: 20px 0; border-radius: 16px; background: #f8fafc; border: 2px solid ' + levelColor + ';">' +
    '<p style="margin: 0 0 4px; font-size: 0.9rem; color: #64748b;">' + testType + ' 總分</p>' +
    '<p style="margin: 0; font-size: 2rem; font-weight: 700; color: ' + levelColor + ';">' + score + '</p>' +
    '<p style="margin: 8px 0 0; font-weight: 500;">' + level + '：' + levelDesc + '</p>' +
    '</div>' +
    suggestion +
    '<p style="margin-top: 24px;">預約掛號：<a href="https://08143.vision.com.tw/Register" style="color: #0d9488;">徐嘉賢診所 線上掛號</a></p>' +
    '<p style="margin-top: 24px; font-size: 0.85rem; color: #64748b;">本結果僅供參考，請與醫師討論後續照護計畫。© 徐嘉賢診所</p>' +
    '</div>';
}
