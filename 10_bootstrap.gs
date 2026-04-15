/** === Bootstrap (Clean) ======================================
 * 役割:
 *  - 共通ユーティリティ（CFG, SS/sh/nowJST, notifyDiscord）
 *  - 「約定取り込み」の30分トリガーを1本だけ用意
 * 余計なトリガー／ダミー通知は持たない
 * =========================================================== */

const CFG = {
  SPREADSHEET_ID: "1EBlWYTlCCQqlWfyNWXsb2VcNxKH6YCxyLL4OXg4_NAA",
  CLOCK_TZ: "Asia/Tokyo"
};

function SS(){ return SpreadsheetApp.openById(CFG.SPREADSHEET_ID); }
function sh(name){ return SS().getSheetByName(name) || SS().insertSheet(name); }
function nowJST(){ return Utilities.formatDate(new Date(), CFG.CLOCK_TZ, 'yyyy-MM-dd HH:mm:ss'); }

function notifyDiscord(content){
  const url = PropertiesService.getScriptProperties().getProperty('DISCORD_WEBHOOK_URL');
  if (!url) return;
  UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ content })
  });
}

/** === 30分トリガーをセット（job_mail_import_sbi だけ） === */
function setup_ingest_trigger(){
  // 既存トリガーを確認し、同じ関数の重複だけ排除
  const all = ScriptApp.getProjectTriggers();
  all.forEach(t=>{
    if (t.getHandlerFunction && t.getHandlerFunction() === 'job_mail_import_sbi'){
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('job_mail_import_sbi')
    .timeBased()
    .everyMinutes(30)
    .inTimezone(CFG.CLOCK_TZ)
    .create();
}
