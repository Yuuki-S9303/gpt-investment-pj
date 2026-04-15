/** =====================================================================
 *  JPX全銘柄 → （Yahoo or GOOGLEFINANCE） → Universe更新 → QCE/DTOふるい出し
 *  （GAS / スプレッドシート直付け or スタンドアロン対応）
 *  v1.2：Yahooが401/403なら自動で GOOGLEFINANCE にフォールバック
 * =====================================================================*/

var TZ = 'Asia/Tokyo';
var SPREADSHEET_ID = ''; // ← 任意：既存ブックに出したい場合はIDを入れる

function getSS_() {
  if (SPREADSHEET_ID) { try { return SpreadsheetApp.openById(SPREADSHEET_ID); } catch(e){} }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;
  var created = SpreadsheetApp.create('Universe_Workbook_' + Utilities.formatDate(new Date(), TZ, 'yyyyMMdd_HHmmss'));
  Logger.log('新規スプレッドシートを作成: ' + created.getUrl());
  return created;
}

/* === 外部取得設定 === */
var YF_QUOTE_ENDPOINT = 'https://query1.finance.yahoo.com/v7/finance/quote';
var YF_CHART_ENDPOINT = 'https://query1.finance.yahoo.com/v8/finance/chart/';
var YF_BATCH_SIZE = 120;
var YF_SLEEP_MS   = 300;
var CHART_TOP_N   = 60;

/* === JPX一覧 === */
var JPX_LIST_PAGE = 'https://www.jpx.co.jp/markets/statistics-equities/misc/01.html';
var JPX_XLS_REGEX = /https?:\/\/[^"' ]+\/att\/[^"' ]+\.(?:xls|xlsx)/ig;
var MARKET_ALLOW  = ['プライム','スタンダード','グロース'];
var MARKET_DENY   = ['ETF','ETN','REIT','インフラ','PRO','プロ','ベンチャーファンド','投資法人','投資証券'];

/* === スクリーニング基準 === */
var LIQ_MIN_DEFAULT = 500000000;
var LIQ_MIN_RELAXED = 300000000;
var USE_RELAXED_DAY = false;
var VOL_RATIO_MIN   = 2.0;
var HI_BAND_PCT     = -0.05;
var CHG20D_MIN      = 0.20;

/* === エントリ：Universe更新 === */
function job_update_universe_all(){
  var xlsUrl = findLatestJpxExcelUrl_();
  if (!xlsUrl) throw new Error('JPXの上場銘柄一覧Excelが見つかりません。');

  var sheetId = importJpxExcelToSheet_(xlsUrl);
  var codes   = extractStockCodesFromJpxSheet_(sheetId); // 4桁コード配列
  var ok = updateUniverseFromYahooDynamic_(codes);       // Yahooで試す

  if (!ok) { // 401/403/空 → GOOGLEFINANCEに切替
    Logger.log('Yahooが利用不可のため、GOOGLEFINANCEでUniverseを構築します。');
    updateUniverseFromGoogleFinance_(codes);
  }
  Logger.log('Universe更新完了（対象コード数）: ' + codes.length + '銘柄');
}

/* ------ JPX：Excel URL抽出（多段フォールバック） ------ */
function findLatestJpxExcelUrl_(){
  var pages = [
    'https://www.jpx.co.jp/markets/statistics-equities/misc/01.html',
    'https://www.jpx.co.jp/english/markets/statistics-equities/misc/01.html'
  ];
  for (var i=0;i<pages.length;i++){
    try{
      var html = UrlFetchApp.fetch(pages[i], {muteHttpExceptions:true}).getContentText('UTF-8');
      var links = html.match(JPX_XLS_REGEX) || [];
      if (links.length) return links[links.length-1];
    }catch(e){}
  }
  var fallbacks = [
    'https://www.jpx.co.jp/markets/statistics-equities/misc/tvdivq0000001vg2-att/data_j.xlsx',
    'https://www.jpx.co.jp/markets/statistics-equities/misc/tvdivq0000001vg2-att/data_j.xls'
  ];
  for (var j=0;j<fallbacks.length;j++){
    try{
      var res = UrlFetchApp.fetch(fallbacks[j], {muteHttpExceptions:true, method:'get'});
      if (res.getResponseCode() === 200) return fallbacks[j];
    }catch(e){}
  }
  return null;
}

/* ------ JPX：Excel→Googleスプレッドシート変換（Drive API v2 必須） ------ */
function importJpxExcelToSheet_(xlsUrl){
  var res = UrlFetchApp.fetch(xlsUrl, {muteHttpExceptions:true});
  if (res.getResponseCode() !== 200) throw new Error('JPX Excelダウンロード失敗: ' + res.getResponseCode());
  var blob = res.getBlob().setName('jpx_list.xls');
  var file = Drive.Files.insert(
    { title: 'JPX_上場銘柄一覧_取込_' + Utilities.formatDate(new Date(), TZ, 'yyyyMMdd_HHmmss'),
      mimeType: 'application/vnd.google-apps.spreadsheet' },
    blob
  );
  return file.id;
}

/* ------ JPX：コード抽出 ------ */
function extractStockCodesFromJpxSheet_(sheetId){
  var ss = SpreadsheetApp.openById(sheetId);
  var sh = ss.getSheets()[0];
  var v  = sh.getDataRange().getValues();
  if (v.length < 2) return [];
  var header  = v[0].map(String);
  var idxCode = findColIndex_(header, ['コード','銘柄コード','証券コード','Code']);
  var idxMkt  = findColIndex_(header, ['市場','市場・商品区分','Market']);
  var idxType = findColIndex_(header, ['33業種区分','17業種区分','種類','分類'], true);

  var out = [];
  for (var r=1;r<v.length;r++){
    var code   = String(v[r][idxCode] || '').trim();
    var market = String(v[r][idxMkt]  || '').trim();
    var type   = idxType>=0 ? String(v[r][idxType] || '').trim() : '';
    if (!/^\d{4}$/.test(code)) continue;
    if (!containsAny_(market, MARKET_ALLOW)) continue;
    if (containsAny_(market, MARKET_DENY) || containsAny_(type, MARKET_DENY)) continue;
    out.push(code);
  }
  return Array.from ? Array.from(new Set(out)).sort() : uniqueSort_(out);
}
function containsAny_(text, words){ for (var i=0;i<words.length;i++){ if (text.indexOf(words[i])>=0) return true; } return false; }
function uniqueSort_(arr){ var m={},o=[]; for (var i=0;i<arr.length;i++){ if (!m[arr[i]]){m[arr[i]]=1;o.push(arr[i]);} } o.sort(); return o; }
function findColIndex_(header, keys, optional){
  for (var i=0;i<header.length;i++){ var h=String(header[i]||''); for (var k=0;k<keys.length;k++){ if (h.indexOf(keys[k])>=0) return i; } }
  return optional ? -1 : -1;
}

/* ------ YahooでUniverse作成（戻り値：true=成功/false=ダメ） ------ */
function updateUniverseFromYahooDynamic_(codeList){
  var ss = getSS_();
  var sh = ss.getSheetByName('Universe') || ss.insertSheet('Universe');
  var headers = ['code','symbol','name','market','sector','close','prevClose','chg','chgPct','volume','avgVol20d(approx)','turnover','marketCap','52wHigh','52wLow','distTo52wHigh','distTo52wLow','updatedAtJST'];
  sh.clearContents(); sh.appendRow(headers);

  var rows = fetchYahooQuoteBulk_(codeList);
  if (!rows || rows.length===0) return false;

  sh.getRange(2,1,rows.length,headers.length).setValues(rows);
  sh.setFrozenRows(1); sh.autoResizeColumns(1, headers.length);
  return true;
}

function fetchYahooQuoteBulk_(codes){
  var symbols = codes.map(function(c){ return c + '.T'; });
  var chunks = []; for (var i=0;i<symbols.length;i+=YF_BATCH_SIZE){ chunks.push(symbols.slice(i,i+YF_BATCH_SIZE)); }
  var out = [];
  var jstNow = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');

  var gotAny = false;
  for (var ci=0; ci<chunks.length; ci++){
    var syms = chunks[ci];
    var url = YF_QUOTE_ENDPOINT + '?symbols=' + encodeURIComponent(syms.join(','));
    var res = UrlFetchApp.fetch(url, { muteHttpExceptions:true });
    var code = res.getResponseCode();
    if (code !== 200){
      Logger.log('Yahoo quote error: ' + code + ' ' + res.getContentText().slice(0,200));
      if (code===401 || code===403) return []; // フォールバックへ
      Utilities.sleep(YF_SLEEP_MS);
      continue;
    }
    var json = JSON.parse(res.getContentText());
    var quotes = (json && json.quoteResponse && json.quoteResponse.result) ? json.quoteResponse.result : [];
    if (quotes.length) gotAny = true;

    for (var q=0; q<quotes.length; q++){
      var r = quotes[q];
      var symbol = r.symbol || '';
      var code4  = symbol.replace('.T','');
      var name   = r.shortName || r.longName || '';
      var market = r.fullExchangeName || r.exchange || 'TSE';
      var sector = r.sector || '';
      var close  = num_(r.regularMarketPrice);
      var prevC  = num_(r.regularMarketPreviousClose);
      var chg    = (isNum_(close)&&isNum_(prevC)) ? close-prevC : '';
      var chgPct = (isNum_(chg)&&isNum_(prevC)&&prevC!==0) ? chg/prevC : '';
      var vol    = num_(r.regularMarketVolume);
      var avg20  = num_(r.averageDailyVolume10Day) || num_(r.averageDailyVolume3Month) || '';
      var tnov   = (isNum_(close)&&isNum_(vol)) ? Math.round(close*vol) : '';
      var mcap   = num_(r.marketCap);
      var hi52   = num_(r.fiftyTwoWeekHigh);
      var lo52   = num_(r.fiftyTwoWeekLow);
      var dHi    = (isNum_(hi52)&&isNum_(close)&&hi52!==0) ? (close/hi52 - 1) : '';
      var dLo    = (isNum_(lo52)&&isNum_(close)&&lo52!==0) ? (close/lo52 - 1) : '';
      out.push([code4, symbol, name, market, sector, close, prevC, chg, chgPct, vol, avg20, tnov, mcap, hi52, lo52, dHi, dLo, jstNow]);
    }
    Utilities.sleep(YF_SLEEP_MS);
  }
  return gotAny ? out : [];
}
function num_(v){ return (typeof v === 'number' && !isNaN(v)) ? v : ''; }
function isNum_(v){ return typeof v === 'number' && !isNaN(v); }

/* ------ GOOGLEFINANCE フォールバック（数式でUniverseを構築） ------ */
// ------ GOOGLEFINANCE フォールバック（東京は TYO:コード / バッチ＋flush） ------
// ------ GOOGLEFINANCE フォールバック（値はsetValues、式はsetFormulas） ------
function updateUniverseFromGoogleFinance_(codeList){
  var ss = getSS_();
  var sh = ss.getSheetByName('Universe') || ss.insertSheet('Universe');
  sh.clearContents();

  var headers = ['code','symbol','name','market','sector','close','prevClose','chg','chgPct','volume','avgVol20d(approx)','turnover','marketCap','52wHigh','52wLow','distTo52wHigh','distTo52wLow','updatedAtJST'];
  sh.getRange(1,1,1,headers.length).setValues([headers]);

  var nowJST = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss');

  var BATCH = 200;
  for (var start=0; start<codeList.length; start+=BATCH){
    var end = Math.min(start+BATCH, codeList.length);
    var rowsValues = [];    // A,B,R列（code, symbol, updatedAtJST）
    var rowsFormulas = [];  // C〜Q列（name〜distTo52wLow）

    for (var i=start; i<end; i++){
      var code  = codeList[i];
      var symGF = 'TYO:' + code;   // GOOGLEFINANCE 用
      var symYF = code + '.T';     // 記録用

      var rowIndex = (i - start) + 2 + start; // シート上の行番号

      // --- 値（A,B,R列）は setValues で入れる ---
      rowsValues.push([ code, symYF, nowJST ]);

      // --- 数式（C〜Q列）は setFormulas で入れる ---
      var fName   = '=IFERROR(GOOGLEFINANCE("'+symGF+'","name"),)';
      var fPrice  = '=IFERROR(GOOGLEFINANCE("'+symGF+'","price"),)';
      var fPrev   = '=IFERROR(GOOGLEFINANCE("'+symGF+'","closeyest"),)';
      var fVol    = '=IFERROR(GOOGLEFINANCE("'+symGF+'","volume"),)';
      var fHi52   = '=IFERROR(GOOGLEFINANCE("'+symGF+'","high52"),)';
      var fLo52   = '=IFERROR(GOOGLEFINANCE("'+symGF+'","low52"),)';
      var fMcap   = '=IFERROR(GOOGLEFINANCE("'+symGF+'","marketcap"),)';
      var fAvgVol = '=IFERROR(AVERAGE(QUERY(GOOGLEFINANCE("'+symGF+'","volume",TODAY()-45,TODAY(),"DAILY"),"select Col2 where Col2 is not null",0)),)';

      var fChg    = '=IF(OR($F'+rowIndex+'="", $G'+rowIndex+'=""), , $F'+rowIndex+'-$G'+rowIndex+')';
      var fChgPct = '=IF(OR($H'+rowIndex+'="", $G'+rowIndex+'="", $G'+rowIndex+'=0), , $H'+rowIndex+'/$G'+rowIndex+')';
      var fTnov   = '=IF(OR($F'+rowIndex+'="", $J'+rowIndex+'=""), , $F'+rowIndex+'*$J'+rowIndex+')';
      var fDistHi = '=IF(OR($F'+rowIndex+'="", $N'+rowIndex+'="", $N'+rowIndex+'=0), , $F'+rowIndex+'/$N'+rowIndex+'-1)';
      var fDistLo = '=IF(OR($F'+rowIndex+'="", $O'+rowIndex+'="", $O'+rowIndex+'=0), , $F'+rowIndex+'/$O'+rowIndex+'-1)';

      // C〜Q（15列）: [name, market, sector, close, prevClose, chg, chgPct, volume, avgVol20d, turnover, marketCap, 52wHigh, 52wLow, distHi, distLo]
      rowsFormulas.push([
        fName, '', '',              // name / market / sector（market/sectorは空のまま）
        fPrice, fPrev, fChg, fChgPct,
        fVol, fAvgVol, fTnov,
        fMcap, fHi52, fLo52, fDistHi, fDistLo
      ]);
    }

    // A,B列（code,symbol）
    sh.getRange(2+start, 1, rowsValues.length, 2).setValues(rowsValues.map(function(r){ return [r[0], r[1]]; }));
    // C〜Q列（数式）
    sh.getRange(2+start, 3, rowsFormulas.length, 15).setFormulas(rowsFormulas);
    // R列（updatedAtJST）
    sh.getRange(2+start, 18, rowsValues.length, 1).setValues(rowsValues.map(function(r){ return [r[2]]; }));

    SpreadsheetApp.flush();
    Utilities.sleep(1200);
  }

  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, headers.length);
}




/* === Catalystテンプレ（任意） === */
function job_create_catalyst_template(){
  var ss = getSS_();
  var sh = ss.getSheetByName('Catalyst') || ss.insertSheet('Catalyst');
  sh.clearContents();
  sh.getRange(1,1,1,3).setValues([['code','eventDate(yyyy-MM-dd)','note']]);
  var today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  sh.getRange(2,1,1,3).setValues([['7011', today, '決算・説明会など']]);
  sh.setFrozenRows(1); sh.autoResizeColumns(1,3);
}

/* === スクリーニング（QCE/DTO） === */
function job_screen_all(){
  var universe = readUniverse_();
  var catalyst = loadCatalystMap_();

  var liqMin = USE_RELAXED_DAY ? LIQ_MIN_RELAXED : LIQ_MIN_DEFAULT;
  var prelim = [];
  for (var i=0;i<universe.length;i++){
    var r = universe[i];
    var hasLiq = (r.turnover || 0) >= liqMin;
    var volOK  = ratio_(r.volume, r.avgVol20d) >= VOL_RATIO_MIN;
    var nearHi = ((r.distHi != null ? r.distHi : -1) >= HI_BAND_PCT);
    if (hasLiq && volOK && (nearHi || true)) prelim.push(r);
  }

  // 20Dは Yahooのchartがブロックされる可能性があるため、ここは一旦スキップ/将来差し替え可
  var withChg20d = prelim.map(function(x){ x.chg20d = null; return x; });

  var rowsQCE = [];
  for (var j=0;j<withChg20d.length;j++){
    var rr = withChg20d[j];
    var hasCatalyst = !!catalyst[rr.code];

    if (!hasCatalyst){
      rowsQCE.push(makeRow_(rr, '除外（ウォッチ中のみ）', 'カタリスト無'));
      continue;
    }
    var nearHi2   = ((rr.distHi != null ? rr.distHi : -1) >= HI_BAND_PCT);
    var volOK2    = ratio_(rr.volume, rr.avgVol20d) >= VOL_RATIO_MIN;

    if (nearHi2 && volOK2){
      rowsQCE.push(makeRow_(rr, '入るべき', reason_(rr, '高値圏・出来高◎')));
      continue;
    }
    rowsQCE.push(makeRow_(rr, '注視（要:高値圏/出来高）', reason_(rr)));
  }

  var TOP_N_DTO = 120;
  var dtoPool = withChg20d
    .filter(function(x){ return !!catalyst[x.code]; })
    .sort(function(a,b){ return (b.turnover||0)-(a.turnover||0); })
    .slice(0, TOP_N_DTO);

  var rowsDTO = [];
  for (var k=0;k<dtoPool.length;k++){
    var d = dtoPool[k];
    var ok = (ratio_(d.volume, d.avgVol20d) >= VOL_RATIO_MIN) &&
             ( ((d.distHi != null ? d.distHi : -1) >= HI_BAND_PCT) );
    rowsDTO.push(makeRow_(d, ok ? '入るべき' : '注視', 'DTO基準（流動性上位・高値圏・材料）'));
  }

  writeSheet_('QCE_OUT', rowsQCE);
  writeSheet_('DTO_OUT', rowsDTO);
  Logger.log('QCE_OUT=' + rowsQCE.length + ', DTO_OUT=' + rowsDTO.length);
}

/* === Universe読取/出力ユーティリティ === */
function readUniverse_(){
  var ss = getSS_();
  var sh = ss.getSheetByName('Universe');
  if (!sh) throw new Error('Universeシートがありません。先に job_update_universe_all() を実行してください。');

  var v = sh.getDataRange().getValues();
  var idx = index_(v[0]);
  var out = [];
  for (var r=1; r<v.length; r++){
    var row = v[r];
    out.push({
      code: str_(row[idx.code]),
      name: str_(row[idx.name]),
      close: num0_(row[idx.close]),
      volume: num0_(row[idx.volume]),
      avgVol20d: num0_(row[idx.avgVol20d]),
      turnover: num0_(row[idx.turnover]),
      marketCap: num0_(row[idx.marketCap]),
      hi52: num0_(row[idx.hi52]),
      lo52: num0_(row[idx.lo52]),
      distHi: num0_(row[idx.distHi])
    });
  }
  return out;
}
function index_(hdr){
  var m = {};
  for (var i=0;i<hdr.length;i++){
    var k = String(hdr[i]).toLowerCase();
    if (k==='code') m.code=i;
    if (k==='name') m.name=i;
    if (k==='close') m.close=i;
    if (k==='volume') m.volume=i;
    if (k==='avgvol20d(approx)') m.avgVol20d=i;
    if (k==='turnover(=close*vol)' || k==='turnover') m.turnover=i;
    if (k==='marketcap') m.marketCap=i;
    if (k==='52whigh') m.hi52=i;
    if (k==='52wlow')  m.lo52=i;
    if (k==='distto52whigh') m.distHi=i;
  }
  return m;
}
function writeSheet_(name, rows){
  var ss = getSS_();
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  var headers = ['Code','Name','Bucket','Side','Reason','Close','Turnover','VolRatio','distTo52wHigh','chg20d'];
  sh.clearContents(); sh.appendRow(headers);
  if (rows.length){ sh.getRange(2,1,rows.length,headers.length).setValues(rows); }
  sh.setFrozenRows(1); sh.autoResizeColumns(1, headers.length);
}

/* === Catalyst & 20D（簡易） === */
function loadCatalystMap_(){
  var ss = getSS_();
  var sh = ss.getSheetByName('Catalyst');
  if (!sh) return {};
  var v = sh.getDataRange().getValues();
  var today = new Date(), map = {};
  for (var r=1;r<v.length;r++){
    var code = String(v[r][0]||'').trim();
    var dateStr = String(v[r][1]||'').trim();
    if (!/^\d{4}$/.test(code) || !dateStr) continue;
    var d = new Date(dateStr + 'T00:00:00+09:00');
    var diff = (today - d) / (1000*60*60*24);
    if (diff <= 10) map[code] = true;
  }
  return map;
}

/* === スコア出力補助 === */
function makeRow_(r, bucket, reason){
  var volRatio = ratio_(r.volume, r.avgVol20d);
  return [ r.code, r.name, bucket, (bucket==='除外（ウォッチ中のみ）' ? '—' : 'BUY'), reason, r.close, r.turnover, volRatio, r.distHi, r.chg20d ];
}
function reason_(r, extra){
  var parts = [];
  if ((r.distHi != null ? r.distHi : -1) >= HI_BAND_PCT) parts.push('高値圏');
  if (ratio_(r.volume, r.avgVol20d) >= VOL_RATIO_MIN) parts.push('出来高◎');
  if (extra) parts.push(extra);
  return parts.join('・') || '条件ギリ届かず';
}
function ratio_(a,b){ return (typeof a==='number' && typeof b==='number' && b>0) ? a/b : 0; }
function num0_(v){ return (typeof v==='number' && !isNaN(v)) ? v : 0; }
function str_(v){ return (v==null) ? '' : String(v); }

/* === トリガー（任意） === */
function setupTriggers(){
  ScriptApp.newTrigger('job_update_universe_all').timeBased().atHour(16).nearMinute(10).everyDays(1).create();
  ScriptApp.newTrigger('job_screen_all').timeBased().atHour(16).nearMinute(20).everyDays(1).create();
}
