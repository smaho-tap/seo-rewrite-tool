/**
 * ============================================================================
 * Day 11-12: ページURL自動取得・DA連携
 * ============================================================================
 * ターゲットKW分析からページURLを取得し、競合分析シートに連携
 */

/**
 * ターゲットKW分析からページURLを取得して競合分析シートに書き込み
 */
function linkPageUrlsToCompetitorAnalysis() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetKWSheet = ss.getSheetByName('ターゲットKW分析');
  const competitorSheet = ss.getSheetByName('競合分析');
  
  if (!targetKWSheet) {
    throw new Error('ターゲットKW分析シートが見つかりません');
  }
  
  if (!competitorSheet) {
    throw new Error('競合分析シートが見つかりません');
  }
  
  Logger.log('=== ページURL自動取得開始 ===');
  Logger.log('');
  
  // ターゲットKW分析からデータ取得（A列: keyword, B列: page_url）
  const targetKWData = targetKWSheet.getDataRange().getValues();
  const targetKWMap = {};
  
  for (let i = 1; i < targetKWData.length; i++) {
    const keyword = targetKWData[i][0]; // A列
    const pageUrl = targetKWData[i][1]; // B列
    
    if (keyword && pageUrl) {
      targetKWMap[keyword] = pageUrl;
    }
  }
  
  Logger.log(`ターゲットKW分析: ${Object.keys(targetKWMap).length} 件取得`);
  Logger.log('');
  
  // 競合分析シートのデータ取得
  const competitorLastRow = competitorSheet.getLastRow();
  if (competitorLastRow < 2) {
    Logger.log('⚠ 競合分析シートにデータがありません');
    return;
  }
  
  const competitorData = competitorSheet.getRange(2, 1, competitorLastRow - 1, 3).getValues();
  
  let matchCount = 0;
  let noMatchCount = 0;
  
  const pageUrls = [];
  
  competitorData.forEach((row, index) => {
    const rowIndex = index + 2;
    const keyword = row[1]; // B列
    
    if (targetKWMap[keyword]) {
      pageUrls.push([targetKWMap[keyword]]);
      matchCount++;
      Logger.log(`✓ [行${rowIndex}] ${keyword} → ${targetKWMap[keyword]}`);
    } else {
      pageUrls.push(['']);
      noMatchCount++;
      Logger.log(`✗ [行${rowIndex}] ${keyword} → マッチなし`);
    }
  });
  
  // C列（page_url）に一括書き込み
  competitorSheet.getRange(2, 3, pageUrls.length, 1).setValues(pageUrls);
  
  Logger.log('');
  Logger.log('=== ページURL書き込み完了 ===');
  Logger.log(`マッチ: ${matchCount} 件`);
  Logger.log(`マッチなし: ${noMatchCount} 件`);
  
  return { matchCount, noMatchCount };
}

/**
 * 自社サイトのDAを更新（ページURLからDA取得）
 */
function updateOwnSiteDA() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const competitorSheet = ss.getSheetByName('競合分析');
  
  if (!competitorSheet) {
    throw new Error('競合分析シートが見つかりません');
  }
  
  const integratedSheet = ss.getSheetByName('統合データ');
  let ownDomain = 'smaho-tap.com';
  
  if (integratedSheet && integratedSheet.getLastRow() > 1) {
    const firstUrl = integratedSheet.getRange(2, 2).getValue();
    if (firstUrl && firstUrl.indexOf('http') === 0) {
      ownDomain = extractDomain(firstUrl);
    }
  }
  
  Logger.log('=== 自社サイトDA更新開始 ===');
  Logger.log(`自社ドメイン: ${ownDomain}`);
  Logger.log('');
  
  const lastRow = competitorSheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('⚠ データがありません');
    return;
  }
  
  const pageUrls = competitorSheet.getRange(2, 3, lastRow - 1, 1).getValues();
  
  const absoluteUrls = pageUrls.map(row => {
    let url = row[0];
    if (url && typeof url === 'string') {
      if (url.indexOf('http') !== 0) {
        url = 'https://' + ownDomain + url;
      }
      return url;
    }
    return '';
  }).filter(url => url);
  
  if (absoluteUrls.length === 0) {
    Logger.log('⚠ ページURLがありません');
    return;
  }
  
  Logger.log(`ページURL数: ${absoluteUrls.length}`);
  Logger.log('');
  
  const daMap = fetchDAWithSmartCaching(absoluteUrls);
  
  Logger.log('');
  Logger.log('自社サイトDAをE列に書き込み中...');
  Logger.log('');
  
  let updateCount = 0;
  
  pageUrls.forEach((row, index) => {
    const rowIndex = index + 2;
    let pageUrl = row[0];
    
    if (pageUrl && typeof pageUrl === 'string') {
      if (pageUrl.indexOf('http') !== 0) {
        pageUrl = 'https://' + ownDomain + pageUrl;
      }
      
      const domain = extractDomain(pageUrl);
      const daData = daMap[domain];
      
      if (daData) {
        competitorSheet.getRange(rowIndex, 5).setValue(daData.da);
        updateCount++;
        Logger.log(`✓ [行${rowIndex}] ${pageUrl} → DA: ${daData.da}`);
      } else {
        Logger.log(`✗ [行${rowIndex}] ${pageUrl} → DAなし`);
      }
    }
  });
  
  Logger.log('');
  Logger.log('=== 自社サイトDA更新完了 ===');
  Logger.log(`更新数: ${updateCount} 件`);
  
  return updateCount;
}

/**
 * 完全な競合分析実行（URL取得 → DA取得 → スコア算出）
 * @param {number} startRow - 開始行（デフォルト: 2）
 * @param {number} endRow - 終了行（デフォルト: null = 最終行）
 */
function runFullCompetitorAnalysis(startRow = 2, endRow = null) {
  Logger.log('=================================');
  Logger.log('完全な競合分析実行');
  Logger.log('=================================');
  Logger.log('');
  
  // ステップ1: ページURL取得
  Logger.log('【ステップ1】ページURL自動取得');
  Logger.log('');
  const urlResult = linkPageUrlsToCompetitorAnalysis();
  Logger.log('');
  
  if (urlResult.matchCount === 0) {
    Logger.log('⚠ ページURLが1件もマッチしませんでした');
    return;
  }
  
  // ステップ2: 自社サイトDA取得
  Logger.log('【ステップ2】自社サイトDA取得');
  Logger.log('');
  const daCount = updateOwnSiteDA();
  Logger.log('');
  
  if (daCount === 0) {
    Logger.log('⚠ 自社サイトDAが1件も取得できませんでした');
    return;
  }
  
  // ステップ3: 勝算度スコア算出
  Logger.log('【ステップ3】勝算度スコア算出');
  Logger.log('');
  updateWinnableScores(startRow, endRow);
  
  Logger.log('');
  Logger.log('=================================');
  Logger.log('完全な競合分析完了');
  Logger.log('=================================');
  Logger.log(`ページURL: ${urlResult.matchCount} 件`);
  Logger.log(`自社DA: ${daCount} 件`);
  Logger.log('');
  Logger.log('競合分析シートを確認してください');
}

/**
 * テスト実行: 最初の5行で完全な競合分析
 */
function testFullCompetitorAnalysis() {
  Logger.log('=== 完全な競合分析テスト（5行） ===');
  Logger.log('');
  
  runFullCompetitorAnalysis(2, 6);
  
  Logger.log('');
  Logger.log('=== テスト完了 ===');
}

/**
 * 全行で完全な競合分析実行
 */
function runFullCompetitorAnalysisForAll() {
  Logger.log('=== 全行で完全な競合分析実行 ===');
  Logger.log('');
  
  runFullCompetitorAnalysis();
  
  Logger.log('');
  Logger.log('=== 全行完了 ===');
}
