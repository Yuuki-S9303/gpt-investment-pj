/** 稼働ウィンドウ：平日 9:00–18:00(JST) だけ動かす */
function shouldRunNow_(){
  const tz = (typeof CFG!=='undefined' && CFG.CLOCK_TZ) || 'Asia/Tokyo';
  const now = new Date();
  const h   = Number(Utilities.formatDate(now, tz, 'H'));     // 0-23
  const dow = Number(Utilities.formatDate(now, tz, 'u'));     // 1=Mon ... 7=Sun
  return (dow >= 1 && dow <= 5) && (h >= 9 && h < 18);
}


/** 一連の更新：Gmail取込 → Fills→Ledger → Equity → KPI */
function jobAllInOne(){
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30 * 1000)) { 
    console.log('[ALL-IN-ONE] lock busy, skip'); 
    return; 
  }
  try {
    if (!shouldRunNow_()){ 
      console.log('[ALL-IN-ONE] outside window'); 
      return; 
    }

    // ① Gmail→Fills
    let ingestNote = '';
    try {
      const ingestRes = jobMailImportSBI();
      if (ingestRes && typeof ingestRes === 'object') {
        ingestNote = `\nmailIngest: ${JSON.stringify(ingestRes)}`;
      }
    } catch (e) {
      console.warn('[ALL-IN-ONE] jobMailImportSBI failed but continue:', e);
      ingestNote = `\nmailIngest: failed=${String(e)}`;
    }

    // ② Fills→Ledger→Equity→KPI
    const stats = jobPipelineAfterIngest();

    if (stats) {
      console.log('[ALL-IN-ONE] stats=', JSON.stringify(stats));
      notifyPipelineIfChanged_(stats, ingestNote);
    } else {
      console.log('[ALL-IN-ONE] pipeline skipped (debounced/locked inside)');
    }

    // ②.5 RiskGuard（BanList/RiskState 更新）
    // stats の有無に関わらず実行：Gmail取込が0件でも、保有状況/連敗判定は更新したい
    try {
      rgRunAll();
      console.log('[ALL-IN-ONE] RiskGuard updated (BanList/RiskState)');
    } catch (e) {
      console.warn('[ALL-IN-ONE] rgRunAll failed:', e);
    }

    // ★ ③ 新規BUY＋ボーダー通知：statsの有無に関わらず実行
    try {
      const sent = notifyNewBuyBorders();
      console.log(`[ALL-IN-ONE] buy-border notified: ${sent}`);
    } catch (e) {
      console.warn('[ALL-IN-ONE] notifyNewBuyBorders failed:', e);
    }

    console.log('[ALL-IN-ONE] done');
  } finally {
    lock.releaseLock();
  }
}




/** === パイプライン通知（fillsToLedgerが動いた時だけ送る） === */
function notifyPipelineIfChanged_(stats, extraLog=""){
  // stats 期待形: { processed:number, deleted:number, eqRows?:number, kpiUpdated?:boolean }
  const p = Number(stats?.processed || 0);
  const d = Number(stats?.deleted   || 0);
  if ((p + d) < 1) {
    // 変化なし → 通知スキップ
    console.log('[PIPELINE] no changes, notification skipped');
    return;
  }
  const msg =
`✅ [PIPELINE] done
fillsToLedger: processed=${p} / deleted=${d}
runEquityUpdate: rows=${Number(stats?.eqRows || 0)}
runKPIUpdate: updated=${Boolean(stats?.kpiUpdated)}
${extraLog||""}`;
  notifyDiscordSafe_(msg);  // 既存の安全通知関数
}


/** 今すぐ一連を実行（時間帯ガードなし・手動実行用） */
function jobAllInOneNow(){
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30 * 1000)) { 
    console.log('[ALL-IN-ONE NOW] lock busy, skip'); 
    return; 
  }
  try {
    // Gmail→Fills
    jobMailImportSBI();

    // Fills→Ledger→Equity→KPI
    const stats = jobPipelineAfterIngest();

    // ②.5 RiskGuard（BanList/RiskState 更新）
    // 手動実行でも、statsの有無に関わらず更新してOK
    try {
      rgRunAll();
      console.log('[ALL-IN-ONE NOW] RiskGuard updated (BanList/RiskState)');
    } catch (e) {
      console.warn('[ALL-IN-ONE NOW] rgRunAll failed:', e);
    }

    // 手動実行時も、fillsToLedger が動いたらボーダー通知
    if (stats) {
      const p = Number(stats.processed || 0);
      const d = Number(stats.deleted   || 0);
      if ((p + d) >= 1) {
        try {
          const sent = notifyNewBuyBorders_();
          console.log(`[ALL-IN-ONE NOW] buy-border notified: ${sent}`);
        } catch (e) {
          console.warn('[ALL-IN-ONE NOW] notifyNewBuyBorders_ failed:', e);
        }
      } else {
        console.log('[ALL-IN-ONE NOW] no fillsToLedger changes, skip buy-border notify');
      }
    }

    console.log('[ALL-IN-ONE NOW] done');
  } finally {
    lock.releaseLock();
  }
}

