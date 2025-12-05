/**
 * AddTargetKeyword.gs
 * 統合データシートにtarget_keyword列を追加し、ターゲットKW分析シートからデータを連携
 */

/**
 * 統合データシートにtarget_keyword列を追加
 */
function addTargetKeywordToIntegratedSheet() {
  console.log('=== target_keyword列追加開始 ===');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const integratedSheet = ss.getSheetByName('統合データ');
  const targetKWSheet = ss.getSheetByName('ターゲットKW分析');
  
  if (!integratedSheet || !targetKWSheet) {
    console.error('シートが見つかりません');
    return;
  }
  
  // 統合データシートのヘッダーを取得
  const intHeaders = integratedSheet.getRange(1, 1, 1, integratedSheet.getLastColumn()).getValues()[0];
  
  // target_keyword列が既に存在するか確認
  let targetKwColIdx = intHeaders.indexOf('target_keyword');
  
  if (targetKwColIdx === -1) {
    // 列を追加（page_titleの後、つまり列Cに挿入）
    const insertPosition = 3; // C列
    integratedSheet.insertColumnBefore(insertPosition);
    integratedSheet.getRange(1, insertPosition).setValue('target_keyword');
    targetKwColIdx = insertPosition - 1; // 0-indexed
    console.log('target_keyword列を追加しました（列C）');
  } else {
    console.log('target_keyword列は既に存在します（列' + (targetKwColIdx + 1) + '）');
  }
  
  // ターゲットKW分析シートからデータを取得
  const targetKWData = targetKWSheet.getDataRange().getValues();
  const targetKWHeaders = targetKWData[0];
  
  const kwPageUrlIdx = targetKWHeaders.indexOf('page_url');
  const kwTargetKwIdx = targetKWHeaders.indexOf('target_keyword');
  const kwSearchVolumeIdx = targetKWHeaders.indexOf('search_volume');
  
  // ページURL → ターゲットKW のマップを作成（検索ボリューム最大のものを選択）
  const kwMap = new Map();
  
  for (let i = 1; i < targetKWData.length; i++) {
    const row = targetKWData[i];
    const pageUrl = normalizeUrlForMatch(row[kwPageUrlIdx]);
    const targetKw = row[kwTargetKwIdx] || '';
    const searchVolume = row[kwSearchVolumeIdx] || 0;
    
    if (!pageUrl || !targetKw) continue;
    
    // 既存のエントリより検索ボリュームが大きい場合は上書き
    if (!kwMap.has(pageUrl) || kwMap.get(pageUrl).searchVolume < searchVolume) {
      kwMap.set(pageUrl, {
        targetKeyword: targetKw,
        searchVolume: searchVolume
      });
    }
  }
  
  console.log(`ターゲットKWマップ構築: ${kwMap.size}件`);
  
  // 統合データシートのデータを取得（ヘッダー更新後）
  const intData = integratedSheet.getDataRange().getValues();
  const updatedHeaders = intData[0];
  const pageUrlIdx = updatedHeaders.indexOf('page_url');
  const newTargetKwColIdx = updatedHeaders.indexOf('target_keyword');
  
  let matchedCount = 0;
  let unmatchedCount = 0;
  
  // 各ページにターゲットKWを設定
  for (let i = 1; i < intData.length; i++) {
    const pageUrl = normalizeUrlForMatch(intData[i][pageUrlIdx]);
    
    if (!pageUrl) continue;
    
    const kwData = kwMap.get(pageUrl);
    
    if (kwData) {
      integratedSheet.getRange(i + 1, newTargetKwColIdx + 1).setValue(kwData.targetKeyword);
      matchedCount++;
    } else {
      unmatchedCount++;
      console.log(`マッチなし: ${pageUrl}`);
    }
  }
  
  console.log('=== target_keyword列追加完了 ===');
  console.log(`マッチ成功: ${matchedCount}件`);
  console.log(`マッチなし: ${unmatchedCount}件`);
  
  return {
    matched: matchedCount,
    unmatched: unmatchedCount
  };
}

/**
 * URLを正規化（マッチング用）
 */
function normalizeUrlForMatch(url) {
  if (!url) return '';
  
  let normalized = url.toString().toLowerCase();
  
  // フルURLからパスのみ抽出
  if (normalized.includes('://')) {
    try {
      const urlObj = new URL(normalized);
      normalized = urlObj.pathname;
    } catch (e) {
      // パース失敗時はそのまま
    }
  }
  
  // 末尾スラッシュを除去
  normalized = normalized.replace(/\/$/, '');
  
  // 先頭スラッシュを確保
  if (!normalized.startsWith('/')) {
    normalized = '/' + normalized;
  }
  
  return normalized;
}

/**
 * target_keyword追加後に5軸スコアリングを再実行
 */
function addTargetKeywordAndRescore() {
  // 1. target_keyword列を追加
  const result = addTargetKeywordToIntegratedSheet();
  
  if (result.matched === 0) {
    console.error('ターゲットKWのマッチングが0件です。URLの形式を確認してください。');
    return;
  }
  
  // 2. 5軸スコアリングを再実行
  console.log('\n--- 5軸スコアリング再実行 ---');
  recalculateAllScoresV2();
}
