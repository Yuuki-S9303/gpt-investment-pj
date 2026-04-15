/************************************************************
 * Daily Summary / Discord Reporter (統合・安定版)
 * - Active Spreadsheet 非依存（SS_ID などのScriptProperty可）
 * - KPI抽出を _collectKpi_STRONG_ 1本に統一
 * - sh() と sheet() の誤用修正
 * - 重複定義削除＆SyntaxError解消
 ************************************************************/

/* ====== 基本ユーティリティ ====== */
function _cfg_(){
  const sp = PropertiesService.getScriptProperties().getProperties();
  const out = Object.assign({}, (typeof CFG==='object' && CFG) ? CFG : {}, sp);
  return out;
}

function _getSpreadsheet_(){
  // 1) Active（コンテナバインド時）
  try{
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) return ss;
  }catch(e){}

  // 2) Script Properties / CFG
  const cfg = _cfg_();
  const candKeys = ['SS_ID','SPREADSHEET_ID','SHEET_ID','MASTER_SSID','SSID'];
  for (const k of candKeys){
    const id = (cfg[k] || '').trim?.() || '';
    if (!id) continue;
    try{
      const ss = SpreadsheetApp.openById(id);
      if (ss) return ss;
    }catch(e){}
  }
  throw new Error('Spreadsheet ID not found. Set Script Property "SS_ID" (or SPREADSHEET_ID / MASTER_SSID).');
}

/** 厳密指定のシート取得（見つからなければ例外） */
if (typeof sheet !== 'function') {
  function sheet(name){
    const ss = _getSpreadsheet_();
    const s  = ss.getSheetByName(name);
    if (!s) throw new Error('Sheet not found: '+name);
    return s;
  }
}

/** ログ */
if (typeof sh !== 'function') {
  function sh(tag, msg){ try{ Logger.log(`${tag}: ${msg}`); }catch(e){} }
}

/** ゆるふわ名前マッチ（大小/前後空白無視） */
function _sheetLoose_(name){
  const ss = _getSpreadsheet_();
  const want = String(name||'').trim().toLowerCase();
  for (const s of ss.getSheets()){
    const nm = String(s.getName()||'').trim().toLowerCase();
    if (nm === want) return s;
  }
  return null;
}

/* ====== 数値/日付ユーティリティ ====== */
if (typeof _toNumberSafe_ !== 'function') {
  function _toNumberSafe_(v){
    if (v==null || v==='') return 0;
    if (typeof v === 'number') return v;
    const s = String(v).replace(/,/g,'').replace(/%$/,'').trim();
    const n = Number(s);
    return isNaN(n) ? 0 : n;
  }
}
if (typeof _asYmdJST_ !== 'function') {
  function _asYmdJST_(v, tz){
    if (v==null || v==='') return '';
    if (v instanceof Date) return Utilities.formatDate(v, tz||'Asia/Tokyo', 'yyyy/MM/dd');
    const s = String(v).trim();
    if (/^\d{4}\/\d{2}\/\d{2}$/.test(s)) return s;
    const d = new Date(s);
    return isNaN(d.getTime()) ? '' : Utilities.formatDate(d, tz||'Asia/Tokyo', 'yyyy/MM/dd');
  }
}
if (typeof _numOr !== 'function') {
  function _numOr(v, d=0){
    if (v==null) return d;
    const s0 = String(v).trim();
    if (!s0 || s0.toUpperCase()==='N/A') return d;
    const s  = s0.replace(/,/g,'').replace(/%$/,'');
    const n  = Number(s);
    return isNaN(n) ? d : n;
  }
} else {
  const _numOr_old = _numOr;
  _numOr = function(v, d=0){
    if (v!=null && typeof v!=='number'){
      const s = String(v).trim().replace(/,/g,'').replace(/%$/,'');
      const n = Number(s);
      if (!isNaN(n)) return n;
    }
    return _numOr_old(v, d);
  }
}
if (typeof _pctText01 !== 'function') {
  function _pctText01(v, digits=2){ const n=_numOr(v,0); return (n*100).toFixed(digits)+'%'; }
}
if (typeof _yenSigned !== 'function') {
  function _yenSigned(n){
    const sign = (Number(n)>=0 ? '＋' : '－');
    return `${sign}${Math.abs(Math.round(Number(n)||0)).toLocaleString('ja-JP')}円`;
  }
}

/** 表示用の柔軟パーサ */
function _parseNumFlex_(v){
  if (v==null || v==='') return NaN;
  if (typeof v === 'number') return v;
  const s = String(v).trim().replace(/,/g,'').replace(/%$/,'');
  const n = Number(s);
  return isNaN(n) ? NaN : n;
}
function _numTextOrDash_(v, digits=3){
  const n = _parseNumFlex_(v);
  return isNaN(n) ? '—' : n.toFixed(digits);
}
function _pctTextFlex_(v, digits=2){
  const n = _parseNumFlex_(v);
  if (isNaN(n)) return '—';
  const val = (n<=1 ? n*100 : n);
  return val.toFixed(digits) + '%';
}

/* ====== 集計: Fills / Equity / Ledger / KPI ====== */
function _collectFillsToday_(tz, ymd){
  try{
    const sF = sheet('Fills');
    if (!sF || sF.getLastRow()<=1) return { total:0, buy:0, sell:0 };
    const vals = sF.getDataRange().getValues();
    const h = vals[0].map(v=>String(v||'').toLowerCase());
    const ixDate = h.indexOf('date');
    const ixSide = h.indexOf('side');
    let total=0, buy=0, sell=0;
    for (let i=1;i<vals.length;i++){
      const r = vals[i];
      const d = _asYmdJST_(r[ixDate], tz);
      if (d === ymd){
        total++;
        const s = String(r[ixSide]||'').toUpperCase();
        if (s==='BUY')  buy++;
        if (s==='SELL') sell++;
      }
    }
    return { total, buy, sell };
  }catch(e){
    sh('警告','Fills集計失敗: '+e);
    return { total:0, buy:0, sell:0 };
  }
}

// ★ これで既存の _collectEquityToday_ を置き換え
function _collectEquityToday_(tz, ymd){
  try{
    const sE = sheet('Equity');
    if (!sE || sE.getLastRow()<=1) return { dailyPL:0, cumPL:0 };

    const v  = sE.getDataRange().getValues();
    const h  = v[0].map(x=>String(x||'').trim());
    const idx = (cands)=>{ for (const c of cands){ const i=h.indexOf(c); if (i>=0) return i; } return -1; };
    const ixDate = idx(['日付','date','Date']);
    const ixDaily= idx(['実現損益','dailypl','実現']);
    const ixCum  = idx(['累積損益','cumpl','累積']);

    if (ixDate<0 || ixCum<0) return { dailyPL:0, cumPL:0 };

    let dailyPL = 0;
    let cumPL   = 0;
    let foundToday = false;
    let latestCumUpTo = NaN;

    for (let i=1;i<v.length;i++){
      const r = v[i];
      const d = _asYmdJST_(r[ixDate], tz);
      const cumVal = _toNumberSafe_(r[ixCum]);

      // 「指定日以下で最新」の累計を常に更新
      if (d && cumVal!=null && d <= ymd){
        latestCumUpTo = cumVal;
      }

      // 当日行があれば通常通りそれを採用
      if (d === ymd){
        foundToday = true;
        dailyPL = (ixDaily>=0 ? _toNumberSafe_(r[ixDaily]) : 0);
        cumPL   = cumVal;
        break; // 当日が見つかったら終了
      }
    }

    // 当日が無ければ累積=直近日の累計、日次=0
    if (!foundToday){
      cumPL = (isNaN(latestCumUpTo) ? 0 : latestCumUpTo);
      dailyPL = 0;
    }

    return { dailyPL, cumPL };
  }catch(e){
    sh('警告','Equity読み取り失敗: '+e);
    return { dailyPL:0, cumPL:0 };
  }
}


function _collectWinLossFromLedger_(tz){
  try{
    const sL = sheet('Ledger');
    if (!sL || sL.getLastRow()<=1) return null;
    const lh = sL.getRange(1,1,1,sL.getLastColumn()).getValues()[0];
    const L  = (n)=>lh.indexOf(n);
    const rows = sL.getRange(2,1, sL.getLastRow()-1, lh.length).getValues();
    let win=0, loss=0, trades=0;
    for (const r of rows){
      const realized = _toNumberSafe_(r[L('実現損益')]);
      const sellDate = r[L('売却日')];
      if (!sellDate || realized===0) continue;
      trades++;
      if (realized>0) win++;
      else if (realized<0) loss++;
    }
    return { win, loss, trades };
  }catch(e){
    sh('警告','Ledger勝敗集計失敗: '+e);
    return null;
  }
}

// ★ 差し替え：日本語の買/売・購入日/売却日にも対応
// ★ Ledgerから「今日の新規エントリー」と HardStop/TimeStop/DTO を集計する（売却サマリ準拠）
function _collectExecAndTopFromLedgerToday_(tz, ymd){
  const sL = sheet('Ledger');
  if (!sL || sL.getLastRow() <= 1) {
    return {
      hardstop_count: 0,
      timestop_count: 0,
      dto_count: 0,
      winners: [],      // Winners/Losers は売却サマリ側で処理するのでここは空でOK
      losers: [],
      entries: []       // 👈 ここに新規エントリーを詰める
    };
  }

  const head = sL.getRange(1, 1, 1, sL.getLastColumn()).getValues()[0].map(x => String(x || '').trim());
  const H = (name) => head.indexOf(name);

  // 売却サマリと同じヘッダ定義
  const ixBuyDate   = H('購入日');
  const ixCode      = H('銘柄コード');
  const ixName      = H('銘柄名');
  const ixBuyPx     = H('購入単価');
  const ixQty       = H('株数');
  const ixSellDate  = H('売却日');
  const ixPnL       = H('実現損益');
  const ixMemo      = H('メモ');
  // もし Mode 列をあとで増やすならここに追加してもOK
  // const ixMode      = H('モード');

  const vals = sL.getRange(2, 1, sL.getLastRow() - 1, head.length).getValues();

  // ===== 1) 今日の売却行（HardStop / TimeStop / DTO カウント用） =====
  const soldToday = [];
  for (const r of vals) {
    const sellYmd = _asYmdJST_(r[ixSellDate], tz);
    if (sellYmd === ymd) soldToday.push(r);
  }

  const classifyMemo = (memoRaw) => {
    const memo = String(memoRaw || '').toUpperCase();
    const isHard = /HARDSTOP|ﾊｰﾄﾞｽﾄｯﾌﾟ|-3%/.test(memo);
    const isTime = /TIMESTOP|時間|引け/.test(memo) && !/DTO/.test(memo);
    const isDTO  = /DTO/.test(memo) && /EOD|引け|CLOSE/.test(memo);
    return { isHard, isTime, isDTO };
  };

  let hardstop_count = 0;
  let timestop_count = 0;
  let dto_count      = 0;

  for (const r of soldToday) {
    const memoFlags = classifyMemo(ixMemo >= 0 ? r[ixMemo] : '');
    if (memoFlags.isHard) hardstop_count++;
    if (memoFlags.isTime) timestop_count++;
    if (memoFlags.isDTO)  dto_count++;
  }

  // ===== 2) 今日の新規エントリー（購入日=ymd & 株数>0） =====
  const entries = [];
  for (const r of vals) {
    const buyYmd = _asYmdJST_(r[ixBuyDate], tz);
    if (buyYmd !== ymd) continue;

    const qty = _toNumberSafe_(r[ixQty]);
    if (!qty || qty <= 0) continue; // マイナスは売り、0はスキップ

    const code  = String(r[ixCode] || '');
    const name  = String(r[ixName] || '');
    const buyPx = _toNumberSafe_(r[ixBuyPx]);
    const entry = isFinite(buyPx) ? buyPx : NaN;
    const mode  = 'NORMAL';  // 必要ならあとで "DTO" 等を入れられるように拡張
    const border = isFinite(entry)
      ? `S:${(entry * 0.97).toFixed(2)} / T:+5%→${(entry * 1.05).toFixed(2)}`
      : '-';

    entries.push({
      code,
      name,
      entry: isFinite(entry) ? entry.toFixed(2) : '-',
      mode,
      border
    });
  }

  return {
    hardstop_count,
    timestop_count,
    dto_count,
    winners: [],   // ここは売却サマリでやる
    losers: [],
    entries
  };
}



// ★ 差し替え版：保有シート名の候補を複数試す
// ★ Ledgerから「売却日が空欄の行＝保有中」を集計して、保有継続（抜粋）用の配列を返す
function _buildHoldingsExcerpt_(){
  try{
    const tz  = (typeof CFG!=='undefined' && CFG && CFG.CLOCK_TZ) ? CFG.CLOCK_TZ : 'Asia/Tokyo';

    // 1) Ledger から保有中ポジションを集計（売却日が空欄、株数>0）
    const sL = sheet('Ledger');
    if (!sL || sL.getLastRow() <= 1) return [];

    const head = sL.getRange(1,1,1,sL.getLastColumn()).getValues()[0].map(x=>String(x||'').trim());
    const H = (keys)=> {
      const arr = Array.isArray(keys) ? keys : [keys];
      for (const k of arr){
        const i = head.indexOf(k);
        if (i >= 0) return i;
      }
      return -1;
    };

    const idxCode    = H(['銘柄コード','Code']);
    const idxName    = H(['銘柄名','Name']);
    const idxBuyPx   = H(['購入単価','取得単価','AvgEntry','平均取得単価']);
    const idxQty     = H(['株数','数量','Qty']);
    const idxSellDt  = H(['売却日','SellDate']);

    // 必須ヘッダが無い場合は何もしない
    if (idxCode < 0 || idxBuyPx < 0 || idxQty < 0 || idxSellDt < 0) {
      sh('警告', 'Ledgerヘッダ不足のため _buildHoldingsExcerpt_ スキップ');
      return [];
    }

    const vals = sL.getRange(2,1, sL.getLastRow()-1, head.length).getValues();

    // key = code|name でグルーピング（増し玉があっても平均単価を出せるように一応集計）
    const map = new Map();
    for (const r of vals){
      const sellDate = r[idxSellDt];
      if (sellDate) continue; // 売却済み → 保有対象外

      const qty = _toNumberSafe_(r[idxQty]);
      if (!qty || qty <= 0) continue; // 0株や売り行はスキップ

      const code  = String(r[idxCode] || '').trim();
      const name  = String(r[idxName] || '').trim();
      const buyPx = _toNumberSafe_(r[idxBuyPx]);

      if (!code) continue;

      const key = `${code}|${name}`;
      const cur = map.get(key) || {code, name, qtySum:0, entryAmount:0};
      const q   = isFinite(qty) ? qty : 0;
      const e   = isFinite(buyPx) ? buyPx : NaN;

      cur.qtySum += q;
      if (isFinite(e)) cur.entryAmount += e * q;

      map.set(key, cur);
    }

    if (!map.size) return []; // 保有なし

    // 2) 可能なら Positions / Holdings シートから「現在値」を拾う（任意）
    const lastMap = new Map();
    try{
      const cfg = (typeof CFG!=='undefined' && CFG) ? CFG : {};
      const ss  = _getSpreadsheet_();

      const candNames = [];
      if (cfg.HOLDINGS_SHEET) candNames.push(cfg.HOLDINGS_SHEET);
      candNames.push('Positions','Holdings','HoldingsView','保有一覧');

      let sPos = null;
      for (const nm of candNames){
        const t = _sheetLoose_(nm);
        if (t) { sPos = t; break; }
      }

      if (sPos && sPos.getLastRow() > 1){
        const h2 = sPos.getRange(1,1,1,sPos.getLastColumn()).getValues()[0].map(x=>String(x||'').trim());
        const H2 = (keys)=> {
          const arr = Array.isArray(keys) ? keys : [keys];
          for (const k of arr){
            const i = h2.indexOf(k);
            if (i >= 0) return i;
          }
          return -1;
        };
        const idxPCode = H2(['Code','銘柄コード']);
        const idxLast  = H2(['Last','現在値','Price','現在']);

        if (idxPCode >= 0 && idxLast >= 0){
          const v2 = sPos.getRange(2,1, sPos.getLastRow()-1, h2.length).getValues();
          for (const r2 of v2){
            const c = String(r2[idxPCode] || '').trim();
            if (!c) continue;
            const last = _toNumberSafe_(r2[idxLast]);
            if (isFinite(last)) lastMap.set(c, last);
          }
        }
      }
    }catch(e){
      sh('警告','Positions lookup failed in _buildHoldingsExcerpt_: '+e);
    }

    // 3) 出力用配列に変換（entry/last/pct/stop/tp をセット）
    const out = [];
    for (const v of map.values()){
      const entry = (v.qtySum > 0 && isFinite(v.entryAmount))
        ? (v.entryAmount / v.qtySum)
        : NaN;

      const last  = lastMap.has(v.code) ? _toNumberSafe_(lastMap.get(v.code)) : NaN;
      const pct   = (isFinite(entry) && isFinite(last) && entry > 0)
        ? (((last/entry)-1)*100).toFixed(2)
        : '-';

      out.push({
        code:  v.code,
        entry: isFinite(entry) ? entry.toFixed(2) : '-',
        last:  isFinite(last)  ? last.toFixed(2)  : '-',
        pct,
        stop:  isFinite(entry) ? (entry*0.97).toFixed(2) : '-',
        tp:    isFinite(entry) ? (entry*1.05).toFixed(2) : '-'
      });
    }

    // 4) 含み益率降順でソートして上位3件だけ抜粋
    out.sort((a,b)=>{
      const pa = parseFloat(a.pct);
      const pb = parseFloat(b.pct);
      if (isNaN(pa) && isNaN(pb)) return a.code.localeCompare(b.code);
      if (isNaN(pa)) return 1;
      if (isNaN(pb)) return -1;
      return pb - pa;
    });

    return out.slice(0, 3);
  }catch(e){
    sh('警告','_buildHoldingsExcerpt_ failed: '+e);
    return [];
  }
}




/* ====== KPI 読み（この1本に統一） ====== */
function _collectKpi_STRONG_(){
  const LABELS = [
    '勝率(取引ベース)','平均R','PF(取引ベース)','PF(日別)','平均保有日数',
    '最大ドローダウン','最大ドローダウン%','取引数','勝ち数','負け数'
  ];
  const out = {}; LABELS.forEach(k=>out[k]='N/A');

  const s = sheet('KPI');
  if (!s || s.getLastRow()<=1) { Logger.log('KPI_EMPTY'); return out; }

  const v = s.getDataRange().getValues();

  // A列=指標, B列=値 を辞書化
  const dict = new Map();
  for (let i=1; i<v.length; i++){
    const k = String(v[i][0] ?? '').trim();
    const val = (v[i][1]==null ? '' : String(v[i][1]));
    if (k) dict.set(k, val);
  }

  // 完全一致
  for (const lab of LABELS){
    if (dict.has(lab)) out[lab] = dict.get(lab);
  }

  // 含む一致（表記ゆれ救済）
  const CONTAINS = [
    ['勝率(取引ベース)', /勝率/],
    ['平均R',            /平均\s*R/i],
    ['PF(取引ベース)',    /PF.*(取引|トレード)/i],
    ['PF(日別)',          /PF.*日別/i],
    ['平均保有日数',       /平均.*保有.*日数/],
    ['最大ドローダウン',   /(最大|ＭＡＸ).*ドロ|MDD/i],
    ['最大ドローダウン%',  /(最大|ＭＡＸ).*ドロ.*%|MDD%/i],
    ['取引数',            /取引数|トレード数|回数/],
    ['勝ち数',            /勝ち数|勝ち|Win/i],
    ['負け数',            /負け数|負け|Loss/i],
  ];
  for (const [disp, pat] of CONTAINS){
    if (out[disp] !== 'N/A') continue;
    for (const [k,vv] of dict.entries()){
      if (pat.test(k)) { out[disp] = vv; break; }
    }
  }

  Logger.log('KPI_DICT: '+JSON.stringify(Object.fromEntries(dict)));
  Logger.log('KPI_OUT : '+JSON.stringify(out));
  return out;
}

/**
 * Ledgerから「今日の売却 Top Winners / Losers」のテキストを作成
 * ※売却サマリ用の _collectSoldOn_ + _groupSoldByCode_ を流用
 */
function _buildWinnersLosersTextFromLedger_(tz, ymd){
  try {
    // 売却サマリと同じロジックで「今日の売却行」を取得
    const rows = _collectSoldOn_(tz, ymd);
    if (!rows || !rows.length) {
      return { winners: '―', losers: '―' };
    }

    // 同一銘柄を合算（損益ベースでソート）
    const grouped = _groupSoldByCode_(rows);

    // 売却サマリとほぼ同じフォーマットでTOP3を生成
    const winners = grouped
      .filter(x => x.pnl > 0)
      .slice(0, 3)
      .map(x => `- ${x.code} ${x.name} ${_yenSigned(x.pnl)}`)
      .join('\n') || '―';

    const losers = grouped
      .filter(x => x.pnl < 0)
      .slice(0, 3)
      .map(x => `- ${x.code} ${x.name} ${_yenSigned(x.pnl)}`)
      .join('\n') || '―';

    return { winners, losers };
  } catch (e) {
    sh('警告', '_buildWinnersLosersTextFromLedger_ failed: ' + e);
    return { winners: '―', losers: '―' };
  }
}


/* ====== レンダラ（PLUS：本文組み立て） ====== */
function _renderDailyReport_PLUS_(ymd, fills, eq, kpi, wl, ext, holdings, soldWl){
  // ガード
  kpi = kpi || {};
  wl  = wl  || {};
  ext = ext || {hardstop_count:0, timestop_count:0, dto_count:0, winners:[], losers:[], entries:[]};
  holdings = holdings || [];

  // KPI数値
  const winRateNum = _numOr(kpi['勝率(取引ベース)'], 0); // 0〜1想定
  const avgRNum    = _numOr(kpi['平均R'], 0);
  const pfTradeNum = _numOr(kpi['PF(取引ベース)'], 0);
  const pfDailyNum = _numOr(kpi['PF(日別)'], 0);
  const avgHoldNum = _numOr(kpi['平均保有日数'], 0);
  let   maxDDAbs   = _numOr(kpi['最大ドローダウン'], 0);
  let   maxDDPct   = _numOr(kpi['最大ドローダウン%'], 0);
  let   nTrades    = _numOr(kpi['取引数'], 0);
  const kpiWin     = _numOr(kpi['勝ち数'], 0);
  const kpiLoss    = _numOr(kpi['負け数'], 0);

  // 勝敗フォールバック
  if (!nTrades && wl.trades!=null) nTrades = wl.trades;
  const winDisp  = (kpi['勝ち数']!=='N/A') ? kpiWin  : (wl.win  ?? 0);
  const lossDisp = (kpi['負け数']!=='N/A') ? kpiLoss : (wl.loss ?? 0);

  // MDD% 補完（Equityか累積損益の系列から計算）
  let needRecalcDD = !(isFinite(maxDDPct) && maxDDPct>=0 && maxDDPct<=1);
  if (needRecalcDD || (maxDDAbs===0 && maxDDPct===0)) {
    try{
      const sE = sheet('Equity');
      const v  = (sE && sE.getLastRow()>1) ? sE.getDataRange().getValues() : [];
      if (v.length>1){
        const h = v[0].map(x=>String(x||'').trim());
        const idxEquity = ['残高','Equity','equity'].map(c=>h.indexOf(c)).find(i=>i>=0);
        const idxCumPL  = ['累積損益','cumpl','累積'].map(c=>h.indexOf(c)).find(i=>i>=0);
        let series = [];
        if (idxEquity!=null && idxEquity>=0){
          series = v.slice(1).map(r=>Number(String(r[idxEquity]).replace(/,/g,''))).filter(n=>!isNaN(n));
        } else if (idxCumPL!=null && idxCumPL>=0){
          series = v.slice(1).map(r=>Number(String(r[idxCumPL]).replace(/,/g,''))).filter(n=>!isNaN(n));
        }
        if (series.length){
          let peak=-1e99, mddAbs=0, mddPct=0;
          for (const x of series){
            if (x>peak) peak=x;
            const ddAbs= peak-x;
            const ddPct= (peak>0) ? Math.max(0, Math.min(1, ddAbs/peak)) : 0;
            if (ddAbs>mddAbs) mddAbs=ddAbs;
            if (ddPct>mddPct) mddPct = ddPct;
          }
          maxDDAbs=mddAbs; maxDDPct=mddPct;
        }
      }
    }catch(e){}
  }

  // 表示整形
  const header = `【日報】${ymd}`;
  const line   = '━━━━━━━━━━━━━━';
  const winRateTxt = _pctText01(winRateNum);
  const avgRTxt    = avgRNum.toFixed(3);
  const pfTradeTxt = pfTradeNum.toFixed(3);
  const pfDailyTxt = pfDailyNum.toFixed(3);
  const avgHoldTxt = avgHoldNum.toFixed(2);
  const maxDDTxt   = _yenSigned(-maxDDAbs).replace('＋','－');
  const tradesTxt  = String(nTrades);

    soldWl = soldWl || {};

  // 売却サマリ側の winners / losers が渡ってきていればそれを優先
  const winnersTxt = soldWl.winners
    || ((ext.winners && ext.winners.length)
        ? ext.winners.map(w => `- ${w.code} ${w.name} ${w.pl}（${w.r}R）理由:${w.reason}`).join('\n')
        : '―');

  const losersTxt = soldWl.losers
    || ((ext.losers && ext.losers.length)
        ? ext.losers.map(l => `- ${l.code} ${l.name} ${l.pl}（${l.r}R）理由:${l.reason}`).join('\n')
        : '―');

  const entriesTxt = (ext.entries&&ext.entries.length)
    ? ext.entries.map(e=>`- ${e.code} ${e.name} @${e.entry}（${e.mode}）| 次ボーダー:${e.border}`).join('\n') : '―';
  const holdingsTxt = (holdings&&holdings.length)
    ? holdings.map(h=>`- ${h.code} 建値${h.entry} / 現在${h.last}（${h.pct}%）| S:${h.stop} / T:${h.tp}`).join('\n') : '―';
  const execLine = `HardStop：${ext.hardstop_count} / TimeStop：${ext.timestop_count} / DTO：${ext.dto_count}`;

  return [
    `${header}`,
    `${line}`,
    `📊 本日取引サマリ`,
    `本日損益：${_yenSigned(eq.dailyPL)}`,
    `取引回数：${fills.total}（BUY:${fills.buy} / SELL:${fills.sell}）`,
    `累積損益：${_yenSigned(eq.cumPL)}`,
    execLine,
    ``,
    `📈 KPI`,
    `勝率(取引ベース)：${winRateTxt}`,
    `平均R：${avgRTxt}`,
    `PF(取引ベース)：${pfTradeTxt}`,
    `PF(日別)：${pfDailyTxt}`,
    `平均保有日数：${avgHoldTxt}`,
    `最大ドローダウン：${maxDDTxt}`,
    `取引数：${tradesTxt}（勝ち${winDisp} / 負け${lossDisp}）`,
    ``,
    `🏆 Top Winners`,
    winnersTxt,
    ``,
    `💔 Top Losers`,
    losersTxt,
    ``,
    `🆕 新規エントリー`,
    entriesTxt,
    ``,
    `📌 保有継続（抜粋）`,
    holdingsTxt,
    `${line}`
  ].join('\n');
}

/* ====== エントリーポイント ====== */
function job_1510_dailySummary(){
  const tz  = (typeof CFG!=='undefined' && CFG && CFG.CLOCK_TZ) ? CFG.CLOCK_TZ : 'Asia/Tokyo';

  // ★ 土日は日報スキップ
  const now = new Date();
  if (typeof _isWeekendJST_ === 'function' && _isWeekendJST_(now)) {
    sh('情報', '[DAILY] weekend detected, skip daily summary');
    return;
  }

  const ymd = Utilities.formatDate(now, tz, 'yyyy/MM/dd');


  // 0) 参照シート存在チェック（ログ出力のみ）
  try{
    const ss = _getSpreadsheet_();
    const names = ss.getSheets().map(s=>s.getName());
    sh('情報','Sheets: '+names.join(', '));
    // 最低限
    ['Fills','Equity','Ledger','KPI'].forEach(n=>{
      const s = _sheetLoose_(n);
      if (!s) throw new Error('Required sheet missing: '+n);
    });
    sh('情報','All required sheets resolved.');
  }catch(e){
    sh('エラー','Sheet resolve failed: '+e);
    throw e;
  }

  // 1) Fills（当日）
  const fills = _collectFillsToday_(tz, ymd);

  // 2) Equity（当日損益/累積）
  const eq    = _collectEquityToday_(tz, ymd);

  // 3) KPI（指標,値）
  const kpi   = _collectKpi_STRONG_();

  // 4) Ledger 勝敗（通算）
  const wl    = _collectWinLossFromLedger_(tz) || {win:0,loss:0,trades:0};

  // 5) Ledger当日 执行種別/トップ/新規
  const ext   = _collectExecAndTopFromLedgerToday_(tz, ymd);

    // 6) 保有抜粋（任意）
  const holdings = _buildHoldingsExcerpt_();

  // 6.5) 売却サマリの Winners / Losers を流用
  const soldWl = _buildWinnersLosersTextFromLedger_(tz, ymd);

  // 7) 描画（soldWl を追加で渡す）
  const text  = _renderDailyReport_PLUS_(ymd, fills, eq, kpi, wl, ext, holdings, soldWl);

  // 8) Discord送信
  try {
    if (typeof discord_ === 'function') discord_(text);
    else postDiscordTo_('daily', text);
  } catch(e){

  }
}
