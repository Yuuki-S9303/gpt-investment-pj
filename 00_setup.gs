function setConfig(){
  const p = PropertiesService.getScriptProperties();

  // ★ 以下のURLを新しいものに書き換えてから、このファイルはGitHubにpushしないこと
  p.setProperty('DISCORD_WEBHOOK_URL',        'YOUR_WEBHOOK_URL_HERE');         // メイン（フォールバック用）
  p.setProperty('DISCORD_WEBHOOK_REPORTING',  'YOUR_REPORTING_WEBHOOK_HERE');   // レポーティング
  p.setProperty('DISCORD_WEBHOOK_LOGS',       'YOUR_LOGS_WEBHOOK_HERE');        // ログ
  p.setProperty('DISCORD_WEBHOOK_BORDERLINE', 'YOUR_BORDERLINE_WEBHOOK_HERE');  // ボーダーライン
  p.setProperty('DISCORD_WEBHOOK_REMIND',     'YOUR_REMIND_WEBHOOK_HERE');      // S株リマインド

  p.setProperty('SECURITY_TOKEN','tok_bA7m2QxK5eN4pH9tR3vS6yL1zW8dJ0fC2');
  Logger.log('All props saved');
}
