// ===== Spreadsheet 取得ユーティリティ（未バインド対応） =====
function getSs_(){
  // 1) バインドされていればこれで取得
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;

  // 2) スクリプトプロパティに ID があればそれで開く
  var id = '';
  try {
    id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID') || '';
  } catch(e) {}
  if (!id) {
    // 3) Settings シートに SPREADSHEET_ID があればそれを使う（バインドされていない可能性もあるので openById）
    id = tryReadSheetIdFromSettings_();
  }
  if (!id) {
    throw new Error('スプレッドシート未バインド。SPREADSHEET_ID を Script Properties か Settings に設定してください。');
  }
  return SpreadsheetApp.openById(id);
}

function tryReadSheetIdFromSettings_(){
  try {
    // バインドされていないと ActiveSpreadsheet は null のため、openById できないので何もせず返す
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return '';
    var sh = ss.getSheetByName('Settings');
    if (!sh) return '';
    var last = sh.getLastRow();
    if (last < 2) return '';
    var vals = sh.getRange(2,1,last-1,2).getValues(); // [Key, Value]
    for (var i=0; i<vals.length; i++){
      var k = String(vals[i][0]||'').trim();
      var v = String(vals[i][1]||'').trim();
      if (k === 'SPREADSHEET_ID' && v) return v;
    }
  } catch(e) {}
  return '';
}

// ===== ヘルパー：安全な数値/日付変換 =====
function num_(v){
  if (v === null || v === '' || v === undefined) return 0;
  if (typeof v === 'number') return v;
  return Number(String(v).replace(/,/g,'').trim()) || 0;
}
function dateOrNull_(v){
  if (!v) return null;
  try { return new Date(v); } catch(e){ return null; }
}

// ===== メイン：QCE-Lite ラン =====
function qce_run(){
  const ss = getSs_();

  // --- シート存在チェック ---
  const shU = ss.getSheetByName('Universe');
  if (!shU) throw new Error('Universe シートが見つかりません。作成してヘッダーを入れてください。');

  const shS = ss.getSheetByName('Signals') || ss.insertSheet('Signals');
  const shW = ss.getSheetByName('Watch_OUT') || ss.insertSheet('Watch_OUT');

  // --- Universe 読み込み＆ヘッダー検証 ---
  const u = shU.getDataRange().getValues();
  if (u.length < 2) throw new Error('Universe にデータがありません。ヘッダー行の下に銘柄行を入れてください。');
  const header = u.shift().map(h => String(h||'').trim());
  const idx = Object.fromEntries(header.map((h,i)=>[h, i]));

  // 必須ヘッダーの存在チェック（足りなければ明示）
  ['Code','Name','Price','Turnover','Vol','Vol20','High52','LastCatalystDate'].forEach(k=>{
    if (!(k in idx)) throw new Error('Universe ヘッダー不足: ' + k);
  });
  // Optional: Change20D% は無ければ0扱い
  const hasChg20 = ('Change20D%' in idx);

  // --- しきい値読込（未設定でもデフォルトで動く） ---
  const p = (typeof getParamMap_ === 'function') ? getParamMap_() : {};
  const liqMin   = (typeof Pn_==='function') ? Pn_(p, 'MIN_LIQUIDITY_JPY',      5e8) : 5e8;
  const liqMinLo = (typeof Pn_==='function') ? Pn_(p, 'MIN_LIQUIDITY_JPY_LO',   3e8) : 3e8;
  const distMax  = (typeof Pn_==='function') ? Pn_(p, 'MAX_52W_HIGH_DIST_PCT',  5)   : 5;
  const chg20Min = (typeof Pn_==='function') ? Pn_(p, 'MIN_20D_CHANGE_PCT',     20)  : 20;
  const vrMin    = (typeof Pn_==='function') ? Pn_(p, 'MIN_VOL_RATIO',          2)   : 2;
  const catLook  = (typeof Pn_==='function') ? Pn_(p, 'CATALYST_LOOKBACK_D',    7)   : 7;
  const now = new Date();

  // --- 計算 ---
  const signals = [];
  const watch = [];
  for (let i=0; i<u.length; i++){
    const r = u[i];

    const code = r[idx.Code];
    // 変更前
    // const name = r[idx.Name];
    // 変更後（両対応）
    const name =
      (idx.Name !== undefined && r[idx.Name]) ||
      (idx['銘柄名'] !== undefined && r[idx['銘柄名']]) ||
      (idx['銘柄']   !== undefined && r[idx['銘柄']]) ||
      '';
    const price = num_(r[idx.Price]);
    const turnover = num_(r[idx.Turnover]);
    const vol = num_(r[idx.Vol]);
    const vol20 = num_(r[idx.Vol20]);
    const high52 = num_(r[idx.High52]);
    const lastCat = dateOrNull_(r[idx.LastCatalystDate]);
    const change20D = hasChg20 ? num_(r[idx['Change20D%']]) : 0;

    if (!code || !price) continue;

    const volRatio = vol20 > 0 ? (vol / vol20) : 0;
    const distHighPct = high52 > 0 ? ((high52 - price) / high52 * 100) : 999;
    const hasCatalyst = lastCat ? ((now - lastCat) / 86400000) <= catLook : false;

    const liqOK = (turnover >= liqMin) || (turnover >= liqMinLo); // 低流動日ゆるめ判定
    const momOK = (distHighPct <= distMax) || (change20D >= chg20Min);
    const volOK = (volRatio >= vrMin);

    let score = 0;
    if (liqOK)       score += 1.0;
    if (momOK)       score += 1.2;
    if (volOK)       score += 1.0;
    if (hasCatalyst) score += 1.3;
    // 52W高値に近いほど微加点（上限0.02 * distMax）
    if (isFinite(distHighPct)) score += Math.max(0, (distMax - distHighPct)) * 0.02;

    // バケット（初期値とフィルタ条件を統一）
    let bucket = '除外（ウォッチ中のみ）';
    if (liqOK && volOK && momOK && hasCatalyst) {
      bucket = '入るべき';
    } else if (liqOK && (momOK || volOK)) {
      bucket = '注視';
    }

    // Signals行
    signals.push([
      code, price, turnover, volRatio, distHighPct, change20D,
      hasCatalyst, liqOK, momOK, volOK, score, null
    ]);

    // --- 日本語の理由文を生成 ---
    const parts = [];
    parts.push(liqOK ? '流動性◎' : '流動性△');
    parts.push(momOK ? 'モメンタム◎' : 'モメンタム△');
    parts.push(volOK ? '出来高◎' : '出来高△');
    parts.push(hasCatalyst ? '材料あり' : '材料なし');
    const reasonJP = parts.join('・');

    // --- Watch_OUT に行追加 ---
    if (bucket !== '除外（ウォッチ中のみ）') {
      watch.push([null, code, name, reasonJP, bucket, 'BUY', null, null]);
    }
  }

  // --- Signals 出力（空でもヘッダーは敷く） ---
  const sigHeaders = ['Code','Price','Turnover','VolRatio','Dist52wHigh%','Change20D%','HasCatalyst','LiquidityOK','MomentumOK','VolumeOK','Score_QCE','Score_DTO'];
  shS.clearContents();
  shS.getRange(1,1,1,sigHeaders.length).setValues([sigHeaders]);
  if (signals.length) {
    shS.getRange(2,1,signals.length,signals[0].length).setValues(signals);
  }

  // --- Watch_OUT 出力（空でもヘッダーは敷く） ---
  const wHeaders = ['Rank','Code','Name','Reason','Bucket','Side','SuggestQty','Budget'];
  shW.clearContents();
  shW.getRange(1,1,1,wHeaders.length).setValues([wHeaders]);

  // フォールバック（全落ち時のデバッグ用・不要なら削除可）
  if (watch.length === 0 && signals.length > 0){
    // スコア上位3件だけ注視として出す
    const sorted = signals
      .map((s)=>({code:s[0], price:s[1], score:s[10]}))
      .sort((a,b)=> b.score - a.score)
      .slice(0, Math.min(3, signals.length));
    sorted.forEach((t,i)=>{
      watch.push([i+1, t.code, '—', 'フォールバック出力（しきい値要調整）', '注視', 'BUY', null, null]);
    });
  }

  if (watch.length) {
    watch.forEach((r,i)=> r[0] = i + 1);
    shW.getRange(2,1,watch.length,watch[0].length).setValues(watch);
  }
}
