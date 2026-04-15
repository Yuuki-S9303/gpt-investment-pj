/**** 売却銘柄の損益通知（Discord）— 建値=「購入単価」のみ ******/
function _cfg_(){
  const sp = PropertiesService.getScriptProperties().getProperties();
  return Object.assign({}, (typeof CFG==='object' && CFG) ? CFG : {}, sp);
}
function _getSpreadsheet_(){
  try{ const ss = SpreadsheetApp.getActiveSpreadsheet(); if (ss) return ss; }catch(e){}
  const cfg = _cfg_();
  for (const k of ['SS_ID','SPREADSHEET_ID','SHEET_ID','MASTER_SSID','SSID']){
    const id = (cfg[k]||'').trim(); if (!id) continue;
    try{ const ss = SpreadsheetApp.openById(id); if (ss) return ss; }catch(e){}
  }
  throw new Error('Spreadsheet ID not found. ScriptProperty に SS_ID 等を設定してください。');
}
if (typeof sheet !== 'function'){
  function sheet(name){
    const ss = _getSpreadsheet_();
    const s  = ss.getSheetByName(name);
    if (!s) throw new Error('Sheet not found: '+name);
    return s;
  }
}
if (typeof sh !== 'function'){
  function sh(tag, msg){ try{ Logger.log(tag+': '+msg); }catch(e){} }
}
if (typeof _toNumberSafe_ !== 'function'){
  function _toNumberSafe_(v){
    if (v==null || v==='') return 0;
    if (typeof v==='number') return v;
    const n = Number(String(v).replace(/,/g,'').replace(/%$/,'').trim());
    return isNaN(n) ? 0 : n;
  }
}
if (typeof _asYmdJST_ !== 'function'){
  function _asYmdJST_(v, tz){
    if (v==null || v==='') return '';
    if (v instanceof Date) return Utilities.formatDate(v, tz||'Asia/Tokyo','yyyy/MM/dd');
    const s = String(v).trim();
    if (/^\d{4}\/\d{2}\/\d{2}$/.test(s)) return s;
    const d = new Date(s);
    return isNaN(d.getTime()) ? '' : Utilities.formatDate(d, tz||'Asia/Tokyo','yyyy/MM/dd');
  }
}
if (typeof _yenSigned !== 'function'){
  function _yenSigned(n){
    const sign = (Number(n)>=0 ? '＋' : '－');
    return `${sign}${Math.abs(Math.round(Number(n)||0)).toLocaleString('ja-JP')}円`;
  }
}
function _postDiscord_(text){
  try{ if (typeof notifyDiscordSafe_==='function'){ notifyDiscordSafe_(text); return; } }catch(e){}
  const url = (_cfg_().DISCORD_WEBHOOK_URL||'').trim();
  if (!url) throw new Error('DISCORD_WEBHOOK_URL が未設定です（ScriptProperties）');
  UrlFetchApp.fetch(url, {method:'post',contentType:'application/json',payload:JSON.stringify({content:text})});
}

/* ==== ヘッダ解決（大小無視・部分一致あり） ==== */
function _findCol_(headArr, candidates){
  const H = headArr.map(h=>String(h||'').trim());
  const lower = H.map(h=>h.toLowerCase());
  // 完全一致（大小そのまま）
  for (const k of candidates){ const i = H.indexOf(k); if (i>=0) return i; }
  // 大小無視の完全一致
  for (const k of candidates){ const i = lower.indexOf(String(k).toLowerCase()); if (i>=0) return i; }
  // 部分一致（大小無視）
  for (let i=0;i<lower.length;i++){
    for (const k of candidates){
      const key = String(k).toLowerCase();
      if (key && lower[i].includes(key)) return i;
    }
  }
  return -1;
}

/* ==== 週末フォールバック ==== */
function _isWeekendJST_(d){
  const tz='Asia/Tokyo'; const wd = Number(Utilities.formatDate(d, tz, 'u')); // 1=Mon..7=Sun
  return (wd===6 || wd===7);
}
function _prevBizDateJST_(d){
  const x = new Date(d.getTime());
  do { x.setDate(x.getDate()-1); } while(_isWeekendJST_(x));
  return x;
}

/* ==== 今日の SELL 抽出（Side列が無くてもOK／売却日ベース） ==== */
function _collectSoldTodayRows_(tz, ymd){
  const sL = sheet('Ledger');
  if (!sL || sL.getLastRow()<=1) return [];

  const head = sL.getRange(1,1,1,sL.getLastColumn()).getValues()[0];

  const ixSellDate = _findCol_(head, ['売却日','SellDate','Sell Date','約定日(売)','Date','日付']);
  const ixSide     = _findCol_(head, ['Side','売買']);
  const ixCode     = _findCol_(head, ['銘柄コード','Code']);
  const ixName     = _findCol_(head, ['銘柄名','Name']);
  const ixQty      = _findCol_(head, ['株数','数量','Qty','Quantity']);
  const ixPxSell   = _findCol_(head, ['売却単価','約定単価(売)','SellPrice','Sell Price']);
  const ixPxAny    = _findCol_(head, ['約定単価','Price','約定価格']);
  const ixBuyPrice = _findCol_(head, ['購入単価']); // ★建値はコレだけ見る
  const ixPnL      = _findCol_(head, ['実現損益','PnL','Realized']);
  const ixR        = _findCol_(head, ['R']);
  const ixExec     = _findCol_(head, ['ExecType']);
  const ixMode     = _findCol_(head, ['Mode','モード']);

  const vals = sL.getRange(2,1, sL.getLastRow()-1, head.length).getValues();
  const out = [];
  for (const r of vals){
    // 「売却日が今日」 or （SideがSELL かつ 日付が今日）
    const dSell = _asYmdJST_( (ixSellDate>=0 ? r[ixSellDate] : r[0]) , tz);
    const side  = (ixSide>=0 ? String(r[ixSide]||'').toUpperCase() : '');
    const isTodaySell = (dSell===ymd) || (side==='SELL' && dSell===ymd);
    if (!isTodaySell) continue;

    const qtyRaw = (ixQty>=0 ? _toNumberSafe_(r[ixQty]) : 0);
    const qty    = Math.abs(qtyRaw); // マイナス表記を吸収

    let price = 0;
    if (ixPxSell>=0) price = _toNumberSafe_(r[ixPxSell]);
    if (!price && ixPxAny>=0) price = _toNumberSafe_(r[ixPxAny]);

    // ★建値＝購入単価のみ
    const entry = (ixBuyPrice>=0) ? _toNumberSafe_(r[ixBuyPrice]) : 0;

    let pnl = (ixPnL>=0 ? _toNumberSafe_(r[ixPnL]) : NaN);
    if (!isFinite(pnl)) pnl = (price - entry) * qty;

    out.push({
      code:String(ixCode>=0 ? r[ixCode] : ''),
      name:String(ixName>=0 ? r[ixName] : ''),
      qty, price, entry, pnl,
      r:(ixR>=0 && r[ixR]!=null && r[ixR]!=='') ? String(r[ixR]) : '',
      exec:String(ixExec>=0 ? (r[ixExec]||'') : ''),
      mode:String(ixMode>=0 ? (r[ixMode]||'') : '')
    });
  }
  return out;
}

/* ==== 文面生成 ==== */
function _renderSoldTodayReport_(ymd, rows){
  if (!rows.length){
    return [
      `【売却サマリ】${ymd}`,
      '━━━━━━━━━━━━━━',
      '本日は売却なし',
      '━━━━━━━━━━━━━━'
    ].join('\n');
  }
  const lines = [];
  let total = 0, win=0, loss=0;
  for (const x of rows){
    total += (x.pnl||0);
    if (x.pnl>0) win++; else if (x.pnl<0) loss++;
    const priceTxt = (x.price ? x.price.toLocaleString('ja-JP') : '-');
    const entryTxt = (x.entry ? x.entry.toLocaleString('ja-JP') : '-'); // ←購入単価ベース
    const rTxt = (x.r ? ` / ${x.r}R` : '');
    const ex   = (x.exec ? ` / ${x.exec}` : '');
    const md   = (x.mode ? ` / ${x.mode}` : '');
    lines.push(`${x.code} ${x.name}：${x.qty}株 | @${priceTxt} / 建値${entryTxt} | 損益 ${_yenSigned(x.pnl)}${rTxt}${ex}${md}`);
  }
  const topW = [...rows].filter(x=>x.pnl>0).sort((a,b)=>b.pnl-a.pnl).slice(0,3);
  const topL = [...rows].filter(x=>x.pnl<0).sort((a,b)=>a.pnl-b.pnl).slice(0,3);
  const winners = topW.length ? topW.map(x=>`  • ${x.code} ${x.name} ${_yenSigned(x.pnl)}`).join('\n') : '  • なし';
  const losers  = topL.length ? topL.map(x=>`  • ${x.code} ${x.name} ${_yenSigned(x.pnl)}`).join('\n') : '  • なし';

  return [
    `【売却サマリ】${ymd}`,
    '━━━━━━━━━━━━━━',
    `件数：${rows.length} / 合計損益：${_yenSigned(total)}`,
    '',
    ...lines,
    '',
    '🏆 Winners (Top3)',
    winners,
    '',
    '💔 Losers (Top3)',
    losers,
    '━━━━━━━━━━━━━━'
  ].join('\n');
}

/* ==== 本体（週末は前営業日に寄せる） ==== */
function job_notifySoldToday(){
  const tz  = (_cfg_().CLOCK_TZ||'Asia/Tokyo');
  const now = new Date();
  const target = _isWeekendJST_(now) ? _prevBizDateJST_(now) : now;
  const ymd = Utilities.formatDate(target, tz, 'yyyy/MM/dd');

  const rows = _collectSoldTodayRows_(tz, ymd);
  const text = _renderSoldTodayReport_(ymd, rows);
  postDiscordTo_('sold', text);
  sh('売却通知', `ymd=${ymd} rows=${rows.length} sent`);
}

/* ==== 任意日付テスト ==== */
function job_notifySoldOn_(dateStr /* 'yyyy/MM/dd' */){
  const tz  = (_cfg_().CLOCK_TZ||'Asia/Tokyo');
  const ymd = dateStr || Utilities.formatDate(new Date(), tz, 'yyyy/MM/dd');
  const rows = _collectSoldTodayRows_(tz, ymd);
  const text = _renderSoldTodayReport_(ymd, rows);
  postDiscordTo_('sold', text);
  sh('売却通知TEST', `ymd=${ymd} rows=${rows.length}`);
}
