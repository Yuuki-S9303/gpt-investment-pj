/* ======================= 売却サマリ通知（Ledger日本語ヘッダ対応） ======================= */
/* 使い方:
 * - job_notifySoldToday() を手動実行 or トリガー登録
 * - 週末は自動で「前営業日」にフォールバック（無効にしたい場合は該当ロジックを外してください）
 */

/* --- 営業日ヘルパ --- */
function _isWeekendJST_(d){
  const tz='Asia/Tokyo';
  const wd = Number(Utilities.formatDate(d, tz, 'u')); // 1=月..7=日
  return (wd===6 || wd===7);
}
function _prevBizDateJST_(d){
  const x = new Date(d.getTime());
  do { x.setDate(x.getDate()-1); } while(_isWeekendJST_(x));
  return x;
}

/* --- Discordポスト（プロジェクトのどれかが定義済み前提） --- */
function _postDiscord_(text){
  try{
    if (typeof discord_ === 'function') return discord_(text);
    if (typeof notifyDiscordSafe_ === 'function') return notifyDiscordSafe_(text);
    throw new Error('discord_ / notifyDiscordSafe_ が見つかりません');
  }catch(e){
    sh('エラー','Discord送信失敗: '+e);
    throw e;
  }
}

/* --- Ledgerから、指定日(yyyy/MM/dd)の売却行を抽出（日本語ヘッダ対応） --- */
function _collectSoldOn_(tz, ymd){
  const sL = sheet('Ledger');  // ←シート名が違う場合はここを変更
  if (!sL || sL.getLastRow()<=1) return [];

  const head = sL.getRange(1,1,1,sL.getLastColumn()).getValues()[0].map(x=>String(x||'').trim());
  const H = (name)=> head.indexOf(name);

  const ixBuyDate   = H('購入日');
  const ixCode      = H('銘柄コード');
  const ixName      = H('銘柄名');
  const ixBuyPx     = H('購入単価');
  const ixQty       = H('株数');
  const ixSellDate  = H('売却日');
  const ixSellPx    = H('売却単価');
  const ixPnL       = H('実現損益');
  const ixRisk1R    = H('1R(想定リスク)');
  const ixR         = H('R');
  const ixMemo      = H('メモ');

  const vals = sL.getRange(2,1, sL.getLastRow()-1, head.length).getValues();
  const out  = [];

  for (const r of vals){
    const sellYmd = _asYmdJST_(r[ixSellDate], tz);
    if (sellYmd !== ymd) continue;

    const qtyRaw = _toNumberSafe_(r[ixQty]);      // 売りはマイナス想定
    const qtyAbs = Math.abs(qtyRaw||0);
    if (qtyAbs===0) continue;                     // 数量0はスキップ

    const buyPx  = _toNumberSafe_(r[ixBuyPx]);
    const sellPx = _toNumberSafe_(r[ixSellPx]);
    const pnl    = _toNumberSafe_(r[ixPnL]);
    const rVal   = r[ixR];

    out.push({
      code: String(r[ixCode]||''),
      name: String(r[ixName]||''),
      buyDate: _asYmdJST_(r[ixBuyDate], tz),
      qty: qtyAbs,
      buyPx, sellPx, pnl,
      r: (rVal==null||rVal==='') ? '' : String(rVal),
      risk1R: _toNumberSafe_(r[ixRisk1R]),
      memo: String(r[ixMemo]||'')
    });
  }
  return out;
}

/* --- 同一銘柄の部分約定を合算（数量合計/損益合計、平均売却単価を再計算） --- */
function _groupSoldByCode_(rows){
  const map = new Map();
  for (const x of rows){
    const key = `${x.code}|${x.name}`;
    const cur = map.get(key) || {code:x.code, name:x.name, qty:0, pnl:0, sellAmount:0, sellQty:0, buyPx:x.buyPx, memo:[]};
    cur.qty += Number(x.qty)||0;
    cur.pnl += Number(x.pnl)||0;
    // 平均売却単価用（売却金額と数量から算出）
    if (isFinite(x.sellPx) && isFinite(x.qty)){
      cur.sellAmount += x.sellPx * x.qty;
      cur.sellQty    += x.qty;
    }
    if (x.memo) cur.memo.push(x.memo);
    map.set(key, cur);
  }
  // 整形
  const out = [];
  for (const v of map.values()){
    const avgSell = (v.sellQty>0) ? (v.sellAmount / v.sellQty) : NaN;
    out.push({
      code: v.code,
      name: v.name,
      qty: v.qty,
      avgSellPx: isFinite(avgSell)? avgSell : NaN,
      buyPx: v.buyPx,
      pnl: v.pnl,
      memo: v.memo.filter(Boolean).join(', ')
    });
  }
  // 損益額降順
  out.sort((a,b)=> b.pnl - a.pnl);
  return out;
}

/* --- レンダ: 売却サマリ本文（Discord想定） --- */
function _renderSoldTodayReport_(ymd, soldRows){
  const line = '━━━━━━━━━━━━━━';
  if (!soldRows || soldRows.length===0){
    return [
      `【売却サマリ】${ymd}`,
      line,
      `本日は売却なし`,
      line
    ].join('\n');
  }

  const grouped = _groupSoldByCode_(soldRows);
  const totalPnl = grouped.reduce((s,x)=>s+(Number(x.pnl)||0), 0);

  const list = grouped.map(x=>{
    const pxTxt = (isFinite(x.avgSellPx) && isFinite(x.buyPx))
      ? `@${x.avgSellPx.toFixed(2)} / 建値${x.buyPx.toFixed(2)}`
      : `建値${isFinite(x.buyPx)? x.buyPx.toFixed(2):'-'}`;
    const memo = x.memo ? ` | ${x.memo}` : '';
    return `- ${x.code} ${x.name}：${x.qty}株 | ${pxTxt} | 損益 ${_yenSigned(x.pnl)}${memo}`;
  }).join('\n');

  // TOP3 勝ち/負け
  const winners = grouped.filter(x=>x.pnl>0).slice(0,3)
                   .map(x=>`  • ${x.code} ${x.name} ${_yenSigned(x.pnl)}`).join('\n') || '  • なし';
  const losers  = grouped.filter(x=>x.pnl<0).slice(0,3)
                   .map(x=>`  • ${x.code} ${x.name} ${_yenSigned(x.pnl)}`).join('\n') || '  • なし';

  return [
    `【売却サマリ】${ymd}`,
    line,
    `件数：${grouped.length} / 合計損益：${_yenSigned(totalPnl)}`,
    ``,
    list,
    ``,
    `🏆 Winners (Top3)`,
    winners,
    ``,
    `💔 Losers (Top3)`,
    losers,
    line
  ].join('\n');
}

/* --- 実行関数：週末は前営業日にフォールバック --- */
function job_notifySoldToday(){
  const tz  = (typeof CFG!=='undefined' && CFG && CFG.CLOCK_TZ) ? CFG.CLOCK_TZ : 'Asia/Tokyo';
  const now = new Date();

  // ★ 土日は売却サマリもスキップ
  if (_isWeekendJST_(now)) {
    sh('売却通知', '[SOLD] weekend detected, skip sold summary');
    return;
  }

  const ymd = Utilities.formatDate(now, tz, 'yyyy/MM/dd');

  const rows = _collectSoldOn_(tz, ymd);
  const text = _renderSoldTodayReport_(ymd, rows);
  _postDiscord_(text);
  sh('売却通知', `ymd=${ymd} rows=${rows.length}`);
}


/* --- 過去日指定で通知したい時用（テスト/リカバリ） --- */
function job_notifySoldOn_(ymd){ // ymd = '2025/11/01' など
  const tz = (typeof CFG!=='undefined' && CFG && CFG.CLOCK_TZ) ? CFG.CLOCK_TZ : 'Asia/Tokyo';
  const rows = _collectSoldOn_(tz, ymd);
  const text = _renderSoldTodayReport_(ymd, rows);
  _postDiscord_(text);
  sh('売却通知', `ymd=${ymd} rows=${rows.length}`);
}
