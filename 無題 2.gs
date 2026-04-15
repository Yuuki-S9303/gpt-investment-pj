/** ❶ WebhookをScript Propertiesに登録（reporting / logs の2本） */
function setWebhooks_multi(){
  const p = PropertiesService.getScriptProperties();
  // ▼新しいWebhook URLをここに直接入力してからGAS上で実行してください（このファイルはコミットしない）
  const REPORTING = 'YOUR_NEW_REPORTING_WEBHOOK_URL_HERE';
  const LOGS      = 'YOUR_NEW_LOGS_WEBHOOK_URL_HERE';

  p.setProperty('DISCORD_WEBHOOK_REPORTING', REPORTING);
  p.setProperty('DISCORD_WEBHOOK_LOGS',      LOGS);

  // 旧フォールバック（DISCORD_WEBHOOK_URL）を残すならコメントのまま
  // 完全移行するなら下行のコメントを外す
  // p.deleteProperty('DISCORD_WEBHOOK_URL');

  Logger.log('OK: reporting/logs を登録しました');
}

/** ❷ 現在のWebhook設定を確認（そのまま実行でOK） */
function debug_showWebhooks(){
  const sp = PropertiesService.getScriptProperties().getProperties();
  const cfg = (typeof CFG==='object' && CFG) ? CFG : {};
  const out = {
    DISCORD_WEBHOOK_REPORTING: sp.DISCORD_WEBHOOK_REPORTING || cfg.DISCORD_WEBHOOK_REPORTING || '(none)',
    DISCORD_WEBHOOK_LOGS     : sp.DISCORD_WEBHOOK_LOGS      || cfg.DISCORD_WEBHOOK_LOGS      || '(none)',
    DISCORD_WEBHOOK_URL      : sp.DISCORD_WEBHOOK_URL       || cfg.DISCORD_WEBHOOK           || '(fallback / none)'
  };
  Logger.log(JSON.stringify(out, null, 2));
}

/** ❸ 送信ルータ（reporting / logs / fallback） */
function postDiscordTo_(channel /* 'reporting' | 'logs' */, text){
  const sp = PropertiesService.getScriptProperties();
  const urlReporting = (sp.getProperty('DISCORD_WEBHOOK_REPORTING') || '').trim();
  const urlLogs      = (sp.getProperty('DISCORD_WEBHOOK_LOGS') || '').trim();
  const urlFallback  = (sp.getProperty('DISCORD_WEBHOOK_URL') || '').trim();

  var url = '';
  if (channel === 'reporting' && urlReporting) url = urlReporting;
  else if (channel === 'logs' && urlLogs)      url = urlLogs;
  else url = urlFallback;  // フォールバック（暫定）

  if (!url) throw new Error('Webhook未設定: ' + channel);

  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    muteHttpExceptions: true,
    payload: JSON.stringify({ content: String(text||'') })
  });
  const code = res.getResponseCode();
  if (code !== 204 && code !== 200) {
    throw new Error(`Discord response ${code}: ${res.getContentText().slice(0,200)}`);
  }
}

/** ❹ 動作テスト：reporting / logs にそれぞれPing */
function test_reporting_ping(){
  const txt = '🧪 reportingルート・テスト ' + new Date().toISOString();
  postDiscordTo_('reporting', txt);
}
function test_logs_ping(){
  const txt = '🧪 logsルート・テスト ' + new Date().toISOString();
  postDiscordTo_('logs', txt);
}

/** ❺（任意）旧1本URLを仮コピーして即テストできるユーティリティ */
function setWebhooks_quickClone(){
  const p = PropertiesService.getScriptProperties();
  const base = (p.getProperty('DISCORD_WEBHOOK_URL') || '').trim();
  if (!base) throw new Error('DISCORD_WEBHOOK_URL が未設定です。まず既存の1本を設定してください。');
  p.setProperty('DISCORD_WEBHOOK_REPORTING', base);
  p.setProperty('DISCORD_WEBHOOK_LOGS',      base);
  Logger.log('reporting/logs を一時的に既存URLへ複製しました');
}
