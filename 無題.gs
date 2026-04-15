/** 互換ログ（log_→sh の橋渡し） */
function log_(level, msg){
  try{
    // level は 'INFO' / 'ERROR' / 'NOTICE' 想定
    const m = String(level||'').toUpperCase();
    const kind = (m==='ERROR') ? 'エラー' : (m==='INFO' ? '情報' : 'お知らせ');
    sh(kind, String(msg||''));
  }catch(e){
    Logger.log(`[${level}] ${msg}`);
  }
}

/* ================================================================
 * Discord Webhook Router（用途別：logs / daily / sold ほか）
 * - postDiscordTo_('logs', text) でログ用へ送信
 * - プロパティ名:
 *    DISCORD_WEBHOOK_LOGS   … ログ用チャンネル（必須）
 *    DISCORD_WEBHOOK_DAILY  … 日報（任意）
 *    DISCORD_WEBHOOK_SOLD   … 売却サマリ（任意）
 *    DISCORD_WEBHOOK_DEFAULT … 既定（任意）
 *    DISCORD_WEBHOOK_URL    … 後方互換（任意）
 * ================================================================ */
if (typeof _cfg_ !== 'function') {
  function _cfg_(){
    const sp = PropertiesService.getScriptProperties().getProperties();
    return Object.assign({}, (typeof CFG==='object' && CFG) ? CFG : {}, sp);
  }
}

function _discordWebhookFor_(route /* 'logs'|'daily'|'sold'|'default' */){
  const cfg = _cfg_();
  const k = String(route||'default').toUpperCase();

  const map = {
    'LOGS':    cfg.DISCORD_WEBHOOK_LOGS,
    'DAILY':   cfg.DISCORD_WEBHOOK_DAILY,
    'SOLD':    cfg.DISCORD_WEBHOOK_SOLD,
    'DEFAULT': cfg.DISCORD_WEBHOOK_DEFAULT
  };
  let url = (map[k]||'').trim();

  // 後方互換（単一設定）
  if (!url) url = (cfg.DISCORD_WEBHOOK_URL||'').trim();

  if (!url) throw new Error('Discord webhook URL not configured for route='+route);
  return url;
}

function postDiscordTo_(route, text){
  const url = _discordWebhookFor_(route);
  UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ content: String(text||'') })
  });
}

/* 初期設定（必要なら1回実行してプロパティに保存） */
function setConfig_DiscordRoutes(){
  const p = PropertiesService.getScriptProperties();

  // ★ここをあなたのURLに置き換え
  const REPORTING = 'https://discord.com/api/webhooks/xxxxxxxx/reporting'; // daily/sold
  const LOGS      = 'https://discord.com/api/webhooks/xxxxxxxx/logs';      // logs

  p.setProperty('DISCORD_WEBHOOK_DAILY',  REPORTING);
  p.setProperty('DISCORD_WEBHOOK_SOLD',   REPORTING);
  p.setProperty('DISCORD_WEBHOOK_LOGS',   LOGS);
  p.setProperty('DISCORD_WEBHOOK_DEFAULT', REPORTING); // 任意（既定）
  // p.setProperty('DISCORD_WEBHOOK_URL', REPORTING);  // 後方互換が必要な場合のみ

  Logger.log('Discord routes saved.');
}

/* 動作テスト（logs/daily/sold それぞれ） */
function job_testRoutes(){
  postDiscordTo_('logs',  '🧪 logs: test message');
  postDiscordTo_('daily', '🧪 daily: test message');
  postDiscordTo_('sold',  '🧪 sold: test message');
}


/* ===== 既存ユーティリティ ===== */

function debug_listTriggers(){
  const ts = ScriptApp.getProjectTriggers();
  if (!ts.length) return Logger.log('No triggers');
  const rows = ts.map(t=>({
    handler: t.getHandlerFunction(),
    event: t.getEventType(),
    source: t.getTriggerSource(),
  }));
  Logger.log(JSON.stringify(rows, null, 2));
}
function cleanup_messyTriggers(){
  const TARGETS = [
    'job_dailySummarySmart',
    'job_dailySummary',        // 旧名の可能性
    'job_dailySummarySmart_',  // 下線付きを直接呼んでる残党
    'notifyDailyReportToDiscord', // 直呼びの誤設定があるケース
  ];
  ScriptApp.getProjectTriggers().forEach(t=>{
    const fn = t.getHandlerFunction();
    if (TARGETS.includes(fn)) {
      ScriptApp.deleteTrigger(t);
    }
  });
  Logger.log('Deleted legacy/duplicate triggers for daily summary.');
}

/* （重複を整理）Discordテスト */
function _testDiscord(){
  try{
    // ログ用チャンネルに飛ばす動作確認
    postDiscordTo_('logs', '🔧 Discordテスト OK ' + new Date().toISOString());
  }catch(e){
    Logger.log('postDiscordTo_(logs) 失敗: ' + e);
  }
}


/* ===== daily summary の送信可否を診断（通知は logs へ） ===== */
function dailySummary_diag(){
  const tz = (typeof CFG!=='undefined' && CFG && CFG.CLOCK_TZ) ? CFG.CLOCK_TZ : 'Asia/Tokyo';
  const now = new Date();
  const today = Utilities.formatDate(now, tz, 'yyyy/MM/dd');

  let sent = false, rowsToday = [], lastExec = 0, reached1600 = false, quiet15m = false, reason = '';

  try{
    // 既送フラグ
    sent = !!hasSummarySentToday_?.();

    // Fills準備＆当日抽出
    ensureHeaders_('Fills', ['createdAt','side','date','code','name','price','qty','market','memo','execTime']);
    const sF = sh('Fills');
    const fv = sF.getDataRange().getValues();
    const fh = fv[0]; const F=(n)=>fh.indexOf(n);
    rowsToday = fv.length>1 ? fv.slice(1).filter(r => String(r[F('date')]||'') === today) : [];

    rowsToday.forEach(r=>{
      const et = String(r[F('execTime')]||'').trim();
      const base = et ? new Date(et) : new Date(String(r[F('createdAt')]||''));
      if (base && !isNaN(base.getTime())) lastExec = Math.max(lastExec, base.getTime());
    });

    // JST 16:00到達＆静寂15分
    const [Y,M,D] = today.split('/').map(n=>parseInt(n,10));
    const jst16UtcMs = Date.UTC(Y, M-1, D, 16-9, 0, 0);
    reached1600 = (lastExec >= jst16UtcMs);
    quiet15m    = ((Date.now() - lastExec) >= 15*60*1000);

    if (sent) reason = 'hasSummarySentToday_ = true（既に送信済み扱い）';
    else if (!fv || fv.length<=1) reason = 'Fills空';
    else if (rowsToday.length===0) reason = '今日のFillsなし';
    else if (!lastExec) reason = 'execTime/createdAtの時刻解析不可';
    else if (!reached1600 && !quiet15m) reason = '引け約定未着 or 静寂不足（15分未満）';
    else reason = '送信可能条件OK（本体を回せば送れる想定）';

  }catch(e){
    reason = '診断中エラー: ' + e;
  }

  const payload =
    '🩺 dailySummary 診断\n' +
    `today=${today}\n` +
    `sentFlag=${sent}\n` +
    `rowsToday=${rowsToday.length}\n` +
    `lastExecIso=${ lastExec ? new Date(lastExec).toISOString() : '(none)' }\n` +
    `reached1600=${reached1600}\n` +
    `quiet15m=${quiet15m}\n` +
    `→ reason=${reason}`;

  try{ postDiscordTo_('logs', payload); }catch(e){ Logger.log('診断のDiscord送信失敗: '+e); }
}


/* ===== 手動強制（日報本体は別ファイルで postDiscordTo_('daily') を使用想定） ===== */
function daily_summary_forceSend(){
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1500)) {
    try{ postDiscordTo_('logs','⚠️ daily_summary_forceSend: 並走中でスキップ'); }catch(e){}
    return;
  }
  try{
    fillsToLedger();
    job_1510_dailySummary();   // ← 日報側の送信は postDiscordTo_('daily', …) にしておく
    markSummarySent_();
    try{ postDiscordTo_('logs','✅ dailySummary（手動強制）送信完了'); }catch(e){}
  }catch(err){
    try{ postDiscordTo_('logs',`🛑 dailySummary（手動強制）失敗: ${err}`); }catch(e){}
    throw err;
  }finally{
    lock.releaseLock();
  }
}


/* ===== ログ用の “安全送信” 互換関数（既存呼び出しを温存） =====
 * 既存コードが notifyDiscordSafe_(text) を呼んでも logs へ飛ぶようにする
 */
function notifyDiscordSafe_(text){
  try{
    if (!text) return 0;
    postDiscordTo_('logs', String(text));
    sh('情報', '[DISCORD/logs] sent');
    return 204; // Discordは通常204 No Content
  }catch(e){
    sh('エラー', '[DISCORD/logs] failed: ' + e);
    return -1;
  }
}


/** === 未反映SELLを自動再処理（AUDIT_SELL_ORPHAN → Fills再投入） ================== */

const CFG_REPROC = {
  SS_ID: '1EBlWYTlCCQqlWfyNWXsb2VcNxKH6YCxyLL4OXg4_NAA',           // ★必須（例: '1AbC...xyz'）
  SHEET_AUDIT: 'AUDIT_SELL_ORPHAN',
  SHEET_FILLS: 'Fills',
  // AUDITのKey_A列（列記号 or 数字どちらでもOK）
  AUDIT_KEY_COL: 'J',                                   // 既定: J列（Key_A_NORM 推奨）
  // Fillsヘッダ名（列位置は自動検出）
  FILLS_KEY_HEADER: 'Key_A_NORM',
  FILLS_PROCESSED_HEADER: 'ProcessedAt',
  // 実行後に呼ぶ関数名（存在しなくてもOK）
  PIPE_FUNCS: ['fillsToLedger', 'runEquityUpdate', 'runKPIUpdate'],
};

function job_reprocess_orphanSELL(){
  const ss = openTargetSpreadsheet_(CFG_REPROC.SS_ID);
  const shAudit = ss.getSheetByName(CFG_REPROC.SHEET_AUDIT);
  const shFills = ss.getSheetByName(CFG_REPROC.SHEET_FILLS);
  if (!shAudit) throw new Error(`シートが見つかりません: ${CFG_REPROC.SHEET_AUDIT}`);
  if (!shFills) throw new Error(`シートが見つかりません: ${CFG_REPROC.SHEET_FILLS}`);

  const keyCol = normalizeCol_(CFG_REPROC.AUDIT_KEY_COL);
  const lastA = shAudit.getLastRow();
  if (lastA < 2) { Logger.log('未反映SELLなし（AUDIT空）'); return; }

  // AUDITのKey_Aを取得
  const orphanKeys = shAudit.getRange(2, keyCol, lastA-1, 1).getValues()
    .flat().filter(v => v !== '');

  if (orphanKeys.length === 0) { Logger.log('未反映SELLなし（キー0件）'); return; }
  const setKeys = new Set(orphanKeys);

  // Fillsのヘッダ行を読み列位置を特定
  const fillsLastCol = shFills.getLastColumn();
  const header = shFills.getRange(1,1,1,fillsLastCol).getValues()[0];
  const idxKey = header.findIndex(h => String(h).trim() === CFG_REPROC.FILLS_KEY_HEADER) + 1;
  const idxProc = header.findIndex(h => String(h).trim() === CFG_REPROC.FILLS_PROCESSED_HEADER) + 1;
  if (!idxKey) throw new Error(`Fillsに列がありません: ${CFG_REPROC.FILLS_KEY_HEADER}`);
  if (!idxProc) throw new Error(`Fillsに列がありません: ${CFG_REPROC.FILLS_PROCESSED_HEADER}`);

  // Fills全行のKey/ProcessedAtを読み込み→一致行のProcessedAtを空に
  const lr = shFills.getLastRow();
  if (lr < 2) { Logger.log('Fillsにデータ行なし'); return; }

  const keys = shFills.getRange(2, idxKey, lr-1, 1).getValues();
  const procs = shFills.getRange(2, idxProc, lr-1, 1).getValues();
  let touched = 0;
  for (let i=0; i<keys.length; i++){
    const k = keys[i][0];
    if (k && setKeys.has(k) && procs[i][0] !== '') {
      procs[i][0] = '';  // クリア
      touched++;
    }
  }
  if (touched > 0) {
    shFills.getRange(2, idxProc, lr-1, 1).setValues(procs);
  }
  Logger.log(`ProcessedAt クリア行数: ${touched}`);

  // パイプライン実行（存在すれば）
  CFG_REPROC.PIPE_FUNCS.forEach(fn => {
    if (typeof this[fn] === 'function') {
      try { this[fn](); Logger.log(`[OK] ${fn}`); }
      catch(e){ Logger.log(`[NG] ${fn}: ${e}`); }
    } else {
      Logger.log(`[SKIP] 関数が見つかりません: ${fn}`);
    }
  });
}

/** ---- helpers ---- */
function openTargetSpreadsheet_(ssId){
  // バインドされていれば active を優先、無ければ ID で開く
  const active = SpreadsheetApp.getActiveSpreadsheet && SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  if (!ssId) throw new Error('SS_IDが未設定です（非バインドのため必須）。');
  return SpreadsheetApp.openById(ssId);
}
function normalizeCol_(col){           // 'J' or 10 → 10
  if (typeof col === 'number') return col;
  const s = String(col).trim().toUpperCase();
  let n = 0;
  for (let i=0; i<s.length; i++){
    n = n*26 + (s.charCodeAt(i) - 64);
  }
  return n;
}

function debug_checkSheets(){
  const ss = _getSpreadsheet_();
  const names = ss.getSheets().map(s=>s.getName());
  Logger.log('Sheets: '+names.join(', '));
  // 存在確認
  sheet('KPI'); sheet('Equity'); sheet('Ledger'); sheet('Fills');
  Logger.log('All required sheets resolved.');
}
