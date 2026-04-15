/**** Fills → Ledger 統合パイプライン（本番クリーン版 / 同日BUY→SELL対応・logs送出対応） *****
 * 前提:
 *  - 共通ユーティリティ: sh()（なければ Logger.log 代用）, CFG（任意）
 *  - ルーター: postDiscordTo_('logs', text) が別ファイルにあればそれを使用
 *    → 無ければ本ファイル内で最小版を定義（ScriptPropertiesの DISCORD_WEBHOOK_LOGS or DISCORD_WEBHOOK_URL を参照）
 *  - Fillsヘッダー:
 *      Date | Side | Code | Name | Price | Qty | Account | OrderNo | ExecType | Source | InsertedAt | (ProcessedAt 任意)
 *  - Ledgerヘッダー（既存運用に準拠）:
 *      購入日 | 銘柄コード | 銘柄名 | 購入単価 | 株数 | 売却日 | 売却単価 | 実現損益 | 購入金額 | 1R(想定リスク) | R | 保有日数 | メモ
 *  - ロジック:
 *      ・平均取得単価方式（SELL後の残原価 = avg × 残株数）
 *      ・同日トランザクションは BUY → SELL の順で処理
 *      ・保有0→BUY でロット起点日を更新（Ledger保有行の「購入日」用）
 ******************************************************************************/

/* ================== ルーター（存在しなければ定義） ================== */
if (typeof _cfg_ !== 'function') {
  function _cfg_(){
    const sp = PropertiesService.getScriptProperties().getProperties();
    return Object.assign({}, (typeof CFG==='object' && CFG) ? CFG : {}, sp);
  }
}
if (typeof postDiscordTo_ !== 'function') {
  function _discordWebhookFor_(route){
    const cfg = _cfg_();
    const r   = String(route||'').toUpperCase();
    let url = '';
    if (r === 'LOGS')      url = (cfg.DISCORD_WEBHOOK_LOGS||'').trim();
    if (!url)              url = (cfg.DISCORD_WEBHOOK_URL||'').trim(); // 後方互換
    if (!url) throw new Error('Discord webhook (logs) not configured');
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
}

/* ================== 共通小物 ================== */
// リスク%（未設定なら 5%）
const RISK_PCT = (typeof CFG !== 'undefined' && CFG.RISK_PCT != null) ? Number(CFG.RISK_PCT) : 0.05;

function round0(x){return Math.round(Number(x)||0);}
function round2(x){return Math.round((Number(x)||0)*100)/100;}

function idxCaseInsensitive(headers, name){
  const target = String(name).toLowerCase();
  for (let i=0;i<headers.length;i++){
    if (String(headers[i]).toLowerCase() === target) return i;
  }
  return -1;
}

function ensureHeadersLocal(sheet, headers){
  const cur = sheet.getRange(1,1,1,Math.max(sheet.getLastColumn(), headers.length)||headers.length).getValues()[0];
  const nonEmpty = cur.filter(v=>String(v||'').trim()!=='').length;
  if (sheet.getLastRow()===0 || sheet.getLastColumn()===0 || nonEmpty===0){
    sheet.clearContents();
    sheet.getRange(1,1,1,headers.length).setValues([headers]);
    return;
  }
  if (cur.length < headers.length){
    const add = headers.slice(cur.length);
    sheet.getRange(1, cur.length+1, 1, add.length).setValues([add]);
  }
}

function ensureFillsProcessedAt(){
  const sF = sh('Fills');
  if (sF.getLastRow() === 0){
    sF.getRange(1,1,1,1).setValue('ProcessedAt');
    return;
  }
  const fh = sF.getRange(1,1,1,sF.getLastColumn()).getValues()[0];
  if (!fh.includes('ProcessedAt')){
    sF.getRange(1, fh.length+1).setValue('ProcessedAt');
  }
}

/* ================== パイプライン・オーケストレーター ================== */
/** 公開：Gmail取込完了後に呼ぶオーケストレーター（logsへ実行結果を通知） */
function jobPipelineAfterIngest(){
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10 * 1000)) {
    try{ postDiscordTo_('logs','⚠️ [PIPELINE] lock busy, skip'); }catch(e){}
    return;
  }
  try {
    // 60秒デバウンス
    const prop  = PropertiesService.getScriptProperties();
    const lastTs = Number(prop.getProperty('PIPELINE_LAST_TS') || 0);
    const now    = Date.now();
    if (now - lastTs < 60 * 1000) {
      try{ postDiscordTo_('logs','⏳ [PIPELINE] debounced (<60s)'); }catch(e){}
      return;
    }

    // 実処理
    const r1 = fillsToLedger();     // 期待: {processed, deleted}
    const r2 = runEquityUpdate();   // 期待: {rows}
    const r3 = runKPIUpdate();      // 期待: {updated:true}

    // デバウンス更新（通知有無に関わらず実行完了として更新）
    prop.setProperty('PIPELINE_LAST_TS', String(now));

    // 件数サマリ
    const p = Number(r1?.processed || 0);
    const d = Number(r1?.deleted   || 0);
    const rows = Number(r2?.rows || 0);
    const kpiUpdated = Boolean(r3?.updated);

    // ✅ 通知は processed or deleted が1以上のときだけ
    if (p + d >= 1) {
      const msg =
`✅ [PIPELINE] done
fillsToLedger: processed=${p} / deleted=${d}
runEquityUpdate: rows=${rows}
runKPIUpdate: updated=${kpiUpdated}`;
      try{ postDiscordTo_('logs', msg); }catch(e){}
    } else {
      // 変化なし → Discord通知は出さない（コンソールだけ残す）
      console.log(`[PIPELINE] no changes (processed=${p}, deleted=${d}) → skip notify`);
    }

    // 呼び出し側で使いたい場合の戻り値（既存呼び出しは無視してOK）
    return { processed:p, deleted:d, eqRows:rows, kpiUpdated };

  } catch (err){
    try{ postDiscordTo_('logs', '🛑 [PIPELINE] failed: ' + err); }catch(e){}
    throw err;
  } finally {
    lock.releaseLock();
  }
}


/* ================== Fills → Ledger ================== */
/** Fills→Ledger（平均取得単価方式）＋ 1R / R / 保有日数 対応（未処理のみ）
 *  戻り値: { processed:number, deleted:number }
 */
function fillsToLedger(){
  // Ledger 見出し
  const sL = sh('Ledger');
  const headers = ['購入日','銘柄コード','銘柄名','購入単価','株数','売却日','売却単価','実現損益','購入金額','1R(想定リスク)','R','保有日数','メモ'];
  ensureHeadersLocal(sL, headers);
  const lh = sL.getRange(1,1,1,headers.length).getValues()[0];
  const L  = (n)=>lh.indexOf(n);

  // 既存保有インデックス（コード→保有行）
  const rowsL = sL.getLastRow()>1 ? sL.getRange(2,1,sL.getLastRow()-1,lh.length).getValues() : [];
  const holdIdx = new Map();
  rowsL.forEach((r,i)=>{
    const code = String(r[L('銘柄コード')]||'').trim();
    const qty  = Number(r[L('株数')]||0);
    const sold = r[L('売却日')];
    if (code && qty>0 && !sold){
      holdIdx.set(code,{
        rowIndex: i+2,
        code,
        name: r[L('銘柄名')],
        qty,
        avg: Number(r[L('購入単価')]||0),
        buyDate: r[L('購入日')]||''
      });
    }
  });

  // Fills 読み込み＆ProcessedAt列を保証
  ensureFillsProcessedAt();
  const sF = sh('Fills');
  const fv = sF.getDataRange().getValues();
  if (fv.length <= 1){
    try{ postDiscordTo_('logs','ℹ️ [fillsToLedger] no fills'); }catch(e){}
    return { processed:0, deleted:0 };
  }

  const fh = fv[0];
  const F = (n)=>{
    const i = idxCaseInsensitive(fh, n);
    if (i>=0) return i;
    const alt = { 'date':'Date', 'side':'Side', 'code':'Code', 'name':'Name', 'price':'Price', 'qty':'Qty' };
    return idxCaseInsensitive(fh, alt[n]||n);
  };

  const cDate = F('date'), cSide=F('side'), cCode=F('code'), cName=F('name'), cPrice=F('price'), cQty=F('qty');
  const cInsertedAt = F('InsertedAt'); // 並び安定化に使用（無ければ -1）
  if ([cDate,cSide,cCode,cName,cPrice,cQty].some(x=>x<0)){
    throw new Error('Fillsヘッダー不足: 必須列(Date/Side/Code/Name/Price/Qty)が見つかりません。');
  }

  const cProcessedAt = fh.indexOf('ProcessedAt')>=0 ? fh.indexOf('ProcessedAt') : fh.length;
  const rowsF = fv.slice(1);
  const tz = (typeof CFG!=='undefined' && CFG.CLOCK_TZ) ? CFG.CLOCK_TZ : Session.getScriptTimeZone();

  const toDelete = new Set();        // 全決済で消す保有行
  const processedMarks = [];         // [rowNumber, timestamp]

  // 未処理だけ抽出 → 同日 BUY→SELL
  const unprocessed = [];
  for (let i=0;i<rowsF.length;i++){
    const r = rowsF[i];
    const already = r[cProcessedAt] && String(r[cProcessedAt]).trim() !== '';
    if (already) continue;
    unprocessed.push({ r, i });
  }
  unprocessed.sort((A,B)=>{
    const a=A.r, b=B.r;
    const ad = new Date(a[cDate]).getTime() || 0;
    const bd = new Date(b[cDate]).getTime() || 0;
    if (ad !== bd) return ad - bd;                             // 日付昇順
    const as = String(a[cSide]||'').toUpperCase();
    const bs = String(b[cSide]||'').toUpperCase();
    if (as !== bs) return as==='BUY' ? -1 : 1;                 // 同日は BUY→SELL
    const ai = (cInsertedAt>=0) ? (new Date(a[cInsertedAt]).getTime()||0) : 0;
    const bi = (cInsertedAt>=0) ? (new Date(b[cInsertedAt]).getTime()||0) : 0;
    if (ai !== bi) return ai - bi;                             // 挿入時刻
    return A.i - B.i;                                          // 行番号で安定化
  });

  for (const {r, i} of unprocessed){
    const side = String(r[cSide]||'').toUpperCase();
    const code = String(r[cCode]||'').trim();
    const name = r[cName]||'';
    const price= Number(r[cPrice]||0);
    const qty  = Number(r[cQty]||0);
    const rawDate = r[cDate];

    if(!code || !qty || !price){
      Logger.log('[fillsToLedger] skip invalid: row='+(i+2)+' code='+code+' price='+price+' qty='+qty);
      processedMarks.push([i+2, Utilities.formatDate(new Date(), tz, 'yyyy/MM/dd HH:mm:ss')]);
      continue;
    }
    const dateStr = Utilities.formatDate(new Date(rawDate), tz, 'yyyy/MM/dd');

    if (side === 'BUY'){
      const h = holdIdx.get(code);
      if (h){
        const newQty = h.qty + qty;
        const newAvg = (h.avg*h.qty + price*qty) / newQty;
        sL.getRange(h.rowIndex, L('購入単価')+1).setValue(round2(newAvg));
        sL.getRange(h.rowIndex, L('株数')+1).setValue(newQty);
        sL.getRange(h.rowIndex, L('購入金額')+1).setValue(round0(newAvg*newQty)); // 残原価=avg×残株数
        if (!h.buyDate) sL.getRange(h.rowIndex, L('購入日')+1).setValue(dateStr);
        // 1R（ポジション全体）
        const col1R = L('1R(想定リスク)') + 1;
        if (col1R > 0) sL.getRange(h.rowIndex, col1R).setValue(round2(newAvg * newQty * RISK_PCT));
        holdIdx.set(code, {...h, qty:newQty, avg:newAvg, buyDate:h.buyDate||dateStr});
      }else{
        const oneRPos = round2(price * qty * RISK_PCT);
        sL.appendRow([dateStr, code, name, price, qty, '', '', '', round0(price*qty), oneRPos, '', '', 'GAS-import']);
        holdIdx.set(code, {rowIndex:sL.getLastRow(), code, name, qty, avg:price, buyDate:dateStr});
      }

    } else if (side === 'SELL'){
      const h = holdIdx.get(code);
      if (!h){
        Logger.log('[fillsToLedger] SELL without position → skip: '+code);
        processedMarks.push([i+2, Utilities.formatDate(new Date(), tz, 'yyyy/MM/dd HH:mm:ss')]);
        continue;
      }

      const avg = h.avg;
      const oneRForSell = round2(avg * qty * RISK_PCT);
      const realized = round0((price - avg) * qty);

      const dBuy  = new Date(h.buyDate || dateStr);
      const dSell = new Date(dateStr);
      const holdDaysRaw = Math.round((dSell - dBuy) / (1000*60*60*24));
      const holdDays = Math.max(1, holdDaysRaw);

      const Rval = (oneRForSell ? (realized / oneRForSell) : '');

      sL.appendRow([
        h.buyDate || '',
        code,
        name || h.name || '',
        avg || '',
        -qty,
        dateStr,          // 売却日
        price,            // 売却単価
        realized,         // 実現損益
        round0(avg*qty),  // 買付金額（原価・売却分）
        oneRForSell,      // 1R(想定リスク) ※売却分
        Rval,             // R
        holdDays,         // 保有日数
        'GAS-sell'
      ]);

      const remain = h.qty - qty;
      if (remain <= 0){
        // フラット：保有行は削除対象、ロット起点日リセット
        toDelete.add(h.rowIndex);
        holdIdx.delete(code);
      }else{
        // 部分売却：残原価は avg × 残株数（平均法）
        sL.getRange(h.rowIndex, L('株数')+1).setValue(remain);
        sL.getRange(h.rowIndex, L('購入金額')+1).setValue(round0(avg*remain));
        const col1R = L('1R(想定リスク)') + 1;
        if (col1R > 0) sL.getRange(h.rowIndex, col1R).setValue(round2(avg * remain * RISK_PCT));
        holdIdx.set(code, {...h, qty:remain});
      }
    }

    // Fills ProcessedAt
    processedMarks.push([i+2, Utilities.formatDate(new Date(), tz, 'yyyy/MM/dd HH:mm:ss')]);
  }

  // Fills ProcessedAt 一括書き込み
  if (processedMarks.length){
    const colP = cProcessedAt + 1;
    const rng = sF.getRange(2, colP, sF.getLastRow()-1, 1);
    const colVals = rng.getValues();
    for (const [rowNum, ts] of processedMarks){
      colVals[rowNum-2][0] = ts;
    }
    rng.setValues(colVals);
  }

  // 保有行の削除（降順で）
  if (toDelete.size){
    Array.from(toDelete).sort((a,b)=>b-a).forEach(rn => sL.deleteRow(rn));
  }

  Logger.log('[fillsToLedger] processed='+processedMarks.length+' deleted='+toDelete.size);
  return { processed: processedMarks.length, deleted: toDelete.size };
}

/* ================== Equity 再集計 ================== */
/** Equity再集計（売却行のみ集計）
 *  戻り値: { rows:number }
 */
function runEquityUpdate(){
  const sL = sh('Ledger');
  const lastRow = sL.getLastRow();
  if (lastRow < 2){
    Logger.log('[runEquityUpdate] no ledger rows');
    return { rows: 0 };
  }

  const lh = sL.getRange(1,1,1,sL.getLastColumn()).getValues()[0];
  const L  = (n)=>lh.indexOf(n);
  const cSellDate = L('売却日')+1;
  const cRealized = L('実現損益')+1;

  const vals = sL.getRange(2,1,lastRow-1,sL.getLastColumn()).getValues();
  const dayPL = new Map();
  vals.forEach(r=>{
    const d = r[cSellDate-1];
    const pl = Number(r[cRealized-1]||0);
    if (!d) return;
    const key = Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), 'yyyy/MM/dd');
    dayPL.set(key, (dayPL.get(key)||0) + pl);
  });

  const days = Array.from(dayPL.keys()).sort();
  const startEq = (typeof CFG!=='undefined' && CFG.SET_STARTING_EQUITY!=null) ? Number(CFG.SET_STARTING_EQUITY) : 0;
  let cum = 0;
  const out = [['日付','実現損益','累積損益','残高']];
  days.forEach(d=>{
    const pl = Math.round(dayPL.get(d));
    cum += pl;
    out.push([d, pl, cum, startEq + cum]);
  });

  const sE = sh('Equity');
  sE.clearContents();
  sE.getRange(1,1,out.length,out[0].length).setValues(out);
  Logger.log('[runEquityUpdate] rows='+ (out.length-1));
  return { rows: out.length-1 };
}

/* ================== KPI 更新 ================== */
/** KPI更新（売却行ベース）
 *  戻り値: { updated:true }
 */
function runKPIUpdate(){
  const sL = sh('Ledger');
  const lastRow = sL.getLastRow();
  if (lastRow < 2){
    Logger.log('[runKPIUpdate] no ledger rows');
    return { updated:false };
  }

  const lh = sL.getRange(1,1,1,sL.getLastColumn()).getValues()[0];
  const L  = (n)=>lh.indexOf(n);

  const cSellDate = L('売却日')+1;
  const cRealized = L('実現損益')+1;
  const cR        = L('R')+1;
  const cDays     = L('保有日数')+1;

  const vals = sL.getRange(2,1,lastRow-1,sL.getLastColumn()).getValues();
  const sells = vals.filter(r=>r[cSellDate-1]);

  const reals = sells.map(r=>Number(r[cRealized-1]||0));
  const wins = reals.filter(x=>x>0).length;
  const losses = reals.filter(x=>x<0).length;
  const trades = wins + losses;

  const Rs = sells.map(r=>Number(r[cR-1])).filter(x=>!isNaN(x));
  const avgR = Rs.length ? Rs.reduce((a,b)=>a+b,0)/Rs.length : '';

  const sumWin = reals.filter(x=>x>0).reduce((a,b)=>a+b,0);
  const sumLossAbs = Math.abs(reals.filter(x=>x<0).reduce((a,b)=>a+b,0));
  const pfTrade = sumLossAbs>0 ? (sumWin / sumLossAbs) : '';

  // 日別PF（参考）
  const byDay = new Map();
  sells.forEach(r=>{
    const d = Utilities.formatDate(new Date(r[cSellDate-1]), Session.getScriptTimeZone(), 'yyyy/MM/dd');
    const pl = Number(r[cRealized-1]||0);
    byDay.set(d, (byDay.get(d)||0) + pl);
  });
  const dayVals = Array.from(byDay.values());
  const pfDay = (()=> {
    const pos = dayVals.filter(x=>x>0).reduce((a,b)=>a+b,0);
    const neg = Math.abs(dayVals.filter(x=>x<0).reduce((a,b)=>a+b,0));
    return neg>0 ? pos/neg : '';
  })();

  const days = sells.map(r=>Number(r[cDays-1]||0)).filter(x=>!isNaN(x)&&x>0);
  const avgDays = days.length ? (days.reduce((a,b)=>a+b,0)/days.length) : '';

  const out = [
    ['指標','値'],
    ['勝率(取引ベース)', trades ? wins/trades : 0],
    ['平均R',            avgR],
    ['PF(取引ベース)',   pfTrade],
    ['PF(日別)',         pfDay],
    ['平均保有日数',      avgDays],
  ];

  const sK = sh('KPI');
  sK.clearContents();
  sK.getRange(1,1,out.length,2).setValues(out);
  Logger.log('[runKPIUpdate] updated KPI');
  return { updated:true };
}
