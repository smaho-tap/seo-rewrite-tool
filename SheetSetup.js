/**
 * SEOリライト支援ツール - SheetSetup.gs
 * Day 7-8: データ構造拡張・KW管理・イベント分析・GTM分析
 * 
 * 実装内容:
 * - 5つの新規シート作成
 * - クエリ単位スコアリング
 * - ターゲットKW分析
 * - KW自動スクリーニング（条件付き保護実装）
 * - データ品質診断
 * 
 * バージョン: 2.0
 * 最終更新: 2025-11-25
 */

// ===================================================================
// Day 7-8.1: 新規シート作成（5シート）
// ===================================================================

/**
 * 5つの新規シートを一括作成
 */
function createNewSheets() {
  Logger.log('===== 新規シート作成開始 =====\n');
  
  createQueryAnalysisSheet();
  createTargetKeywordSheet();
  createKeywordRemovalSheet();
  createEventAnalysisSheet();
  createGTMAnalysisSheet();
  
  Logger.log('\n===== 5つの新規シート作成完了 =====');
}

/**
 * クエリ分析シート作成（12列）
 */
function createQueryAnalysisSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = 'クエリ分析';
  
  // 既存シートがあれば削除
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    ss.deleteSheet(sheet);
  }
  
  // 新規シート作成
  sheet = ss.insertSheet(sheetName);
  
  // ヘッダー行
  const headers = [
    'query_id',
    'page_url',
    'query',
    'position',
    'clicks',
    'impressions',
    'ctr',
    'query_score',
    'cv_proximity',
    'target_kw_match',
    'improvement_potential',
    'last_updated'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // ヘッダー行の書式設定
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  
  // 列幅設定
  sheet.setColumnWidth(1, 200); // query_id
  sheet.setColumnWidth(2, 250); // page_url
  sheet.setColumnWidth(3, 300); // query
  sheet.setColumnWidth(4, 80);  // position
  sheet.setColumnWidth(5, 80);  // clicks
  sheet.setColumnWidth(6, 100); // impressions
  sheet.setColumnWidth(7, 80);  // ctr
  sheet.setColumnWidth(8, 100); // query_score
  sheet.setColumnWidth(9, 120); // cv_proximity
  sheet.setColumnWidth(10, 120); // target_kw_match
  sheet.setColumnWidth(11, 130); // improvement_potential
  sheet.setColumnWidth(12, 150); // last_updated
  
  // データ検証ルール（cv_proximity）
  const cvProximityRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['高', '中', '低'], true)
    .build();
  sheet.getRange(2, 9, 500, 1).setDataValidation(cvProximityRule);
  
  // 行を固定
  sheet.setFrozenRows(1);
  
  Logger.log('✅ クエリ分析シート作成完了');
}

/**
 * ターゲットKW分析シート作成（18列）
 */
function createTargetKeywordSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = 'ターゲットKW分析';
  
  // 既存シートがあれば削除
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    ss.deleteSheet(sheet);
  }
  
  // 新規シート作成
  sheet = ss.insertSheet(sheetName);
  
  // ヘッダー行（18列）
  const headers = [
    'keyword_id',
    'page_url',
    'target_keyword',
    'gyron_position',
    'gsc_position',
    'gsc_clicks',
    'gsc_impressions',
    'gsc_ctr',
    'search_volume',
    'competition_level',
    'kw_score',
    'performance_score',
    'search_volume_score',
    'strategic_value_score',
    'removal_score',
    'status',
    'notes',
    'last_updated'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // ヘッダー行の書式設定
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#34a853');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  
  // 列幅設定
  sheet.setColumnWidth(1, 200); // keyword_id
  sheet.setColumnWidth(2, 250); // page_url
  sheet.setColumnWidth(3, 250); // target_keyword
  sheet.setColumnWidth(4, 100); // gyron_position
  sheet.setColumnWidth(5, 100); // gsc_position
  sheet.setColumnWidth(6, 100); // gsc_clicks
  sheet.setColumnWidth(7, 120); // gsc_impressions
  sheet.setColumnWidth(8, 100); // gsc_ctr
  sheet.setColumnWidth(9, 120); // search_volume
  sheet.setColumnWidth(10, 120); // competition_level
  sheet.setColumnWidth(11, 100); // kw_score
  sheet.setColumnWidth(12, 120); // performance_score
  sheet.setColumnWidth(13, 130); // search_volume_score
  sheet.setColumnWidth(14, 140); // strategic_value_score
  sheet.setColumnWidth(15, 120); // removal_score
  sheet.setColumnWidth(16, 120); // status
  sheet.setColumnWidth(17, 300); // notes
  sheet.setColumnWidth(18, 150); // last_updated
  
  // データ検証ルール（competition_level）
  const competitionRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['易', '中', '難', '激戦'], true)
    .build();
  sheet.getRange(2, 10, 500, 1).setDataValidation(competitionRule);
  
  // データ検証ルール（status）
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['維持', '最優先改善', '要改善', '除外候補'], true)
    .build();
  sheet.getRange(2, 16, 500, 1).setDataValidation(statusRule);
  
  // 行を固定
  sheet.setFrozenRows(1);
  
  Logger.log('✅ ターゲットKW分析シート作成完了');
}

/**
 * KW除外候補シート作成（8列）
 */
function createKeywordRemovalSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = 'KW除外候補';
  
  // 既存シートがあれば削除
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    ss.deleteSheet(sheet);
  }
  
  // 新規シート作成
  sheet = ss.insertSheet(sheetName);
  
  // ヘッダー行
  const headers = [
    'keyword_id',
    'target_keyword',
    'page_url',
    'removal_score',
    'removal_reasons',
    'user_decision',
    'decision_date',
    'notes'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // ヘッダー行の書式設定
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#ea4335');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  
  // 列幅設定
  sheet.setColumnWidth(1, 200); // keyword_id
  sheet.setColumnWidth(2, 250); // target_keyword
  sheet.setColumnWidth(3, 250); // page_url
  sheet.setColumnWidth(4, 120); // removal_score
  sheet.setColumnWidth(5, 400); // removal_reasons
  sheet.setColumnWidth(6, 120); // user_decision
  sheet.setColumnWidth(7, 120); // decision_date
  sheet.setColumnWidth(8, 300); // notes
  
  // データ検証ルール（user_decision）
  const decisionRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['除外', '維持', '保留'], true)
    .build();
  sheet.getRange(2, 6, 500, 1).setDataValidation(decisionRule);
  
  // 行を固定
  sheet.setFrozenRows(1);
  
  Logger.log('✅ KW除外候補シート作成完了');
}

/**
 * イベント分析シート作成（8列）
 */
function createEventAnalysisSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = 'イベント分析';
  
  // 既存シートがあれば削除
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    ss.deleteSheet(sheet);
  }
  
  // 新規シート作成
  sheet = ss.insertSheet(sheetName);
  
  // ヘッダー行
  const headers = [
    'event_id',
    'event_name',
    'event_category',
    'event_count',
    'cv_contribution',
    'importance',
    'enabled',
    'last_updated'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // ヘッダー行の書式設定
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#fbbc04');
  headerRange.setFontColor('#000000');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  
  // 列幅設定
  sheet.setColumnWidth(1, 200); // event_id
  sheet.setColumnWidth(2, 250); // event_name
  sheet.setColumnWidth(3, 150); // event_category
  sheet.setColumnWidth(4, 120); // event_count
  sheet.setColumnWidth(5, 130); // cv_contribution
  sheet.setColumnWidth(6, 120); // importance
  sheet.setColumnWidth(7, 100); // enabled
  sheet.setColumnWidth(8, 150); // last_updated
  
  // データ検証ルール（event_category）
  const categoryRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['CV', 'エンゲージメント', 'その他'], true)
    .build();
  sheet.getRange(2, 3, 500, 1).setDataValidation(categoryRule);
  
  // データ検証ルール（importance）
  const importanceRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['高', '中', '低'], true)
    .build();
  sheet.getRange(2, 6, 500, 1).setDataValidation(importanceRule);
  
  // データ検証ルール（enabled）
  const enabledRule = SpreadsheetApp.newDataValidation()
    .requireCheckbox()
    .build();
  sheet.getRange(2, 7, 500, 1).setDataValidation(enabledRule);
  
  // 行を固定
  sheet.setFrozenRows(1);
  
  Logger.log('✅ イベント分析シート作成完了');
}

/**
 * GTM分析シート作成（9列）
 */
function createGTMAnalysisSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = 'GTM分析';
  
  // 既存シートがあれば削除
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    ss.deleteSheet(sheet);
  }
  
  // 新規シート作成
  sheet = ss.insertSheet(sheetName);
  
  // ヘッダー行
  const headers = [
    'tag_id',
    'tag_name',
    'tag_type',
    'trigger_name',
    'firing_count',
    'is_necessary',
    'removal_reason',
    'status',
    'last_updated'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // ヘッダー行の書式設定
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#9c27b0');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  
  // 列幅設定
  sheet.setColumnWidth(1, 150); // tag_id
  sheet.setColumnWidth(2, 250); // tag_name
  sheet.setColumnWidth(3, 150); // tag_type
  sheet.setColumnWidth(4, 200); // trigger_name
  sheet.setColumnWidth(5, 120); // firing_count
  sheet.setColumnWidth(6, 120); // is_necessary
  sheet.setColumnWidth(7, 300); // removal_reason
  sheet.setColumnWidth(8, 120); // status
  sheet.setColumnWidth(9, 150); // last_updated
  
  // データ検証ルール（is_necessary）
  const necessaryRule = SpreadsheetApp.newDataValidation()
    .requireCheckbox()
    .build();
  sheet.getRange(2, 6, 500, 1).setDataValidation(necessaryRule);
  
  // データ検証ルール（status）
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['維持', '削除推奨', '保留'], true)
    .build();
  sheet.getRange(2, 8, 500, 1).setDataValidation(statusRule);
  
  // 行を固定
  sheet.setFrozenRows(1);
  
  Logger.log('✅ GTM分析シート作成完了');
}

/**
 * シート再作成用ヘルパー関数
 */
function recreateTargetKeywordSheet() {
  Logger.log('ターゲットKW分析シートを再作成します...');
  createTargetKeywordSheet();
  Logger.log('✅ 再作成完了');
}

// ===================================================================
// Day 7-8.2: クエリ単位スコアリング実装
// ===================================================================

/**
 * クエリ分析実行（テスト版: 上位10ページのみ）
 */
function analyzeQueries(limitPages = null) {
  Logger.log('===== クエリ分析開始 =====\n');
  
  // GSC_RAWからデータ取得
  const gscData = getGSCRawData();
  Logger.log(`GSC_RAW取得: ${gscData.length}行`);
  
  // ターゲットKW取得
  const targetKeywords = getTargetKeywords();
  Logger.log(`ターゲットKW取得: ${targetKeywords.length}件`);
  
  // ページごとにクエリをグループ化
  const pageGroups = groupByPage(gscData);
  const pageUrls = Object.keys(pageGroups);
  Logger.log(`ユニークページ数: ${pageUrls.length}`);
  
  // テスト用: limitPagesが指定されていれば制限
  const processPages = limitPages ? pageUrls.slice(0, limitPages) : pageUrls;
  Logger.log(`処理対象: ${processPages.length}ページ`);
  
  const allQueries = [];
  let processedCount = 0;
  
  for (const pageUrl of processPages) {
    const queries = pageGroups[pageUrl];
    
    // 上位20-30クエリを抽出
    const topQueries = getTopQueries(queries, 30);
    
    // 各クエリをスコアリング
    for (const query of topQueries) {
      const scored = scoreQuery(query, pageUrl, targetKeywords);
      allQueries.push(scored);
    }
    
    processedCount++;
    if (processedCount % 10 === 0) {
      Logger.log(`進捗: ${processedCount}/${processPages.length}`);
    }
  }
  
  Logger.log(`\n総クエリ数: ${allQueries.length}件`);
  
  // クエリ分析シートに書き込み
  writeQueryAnalysisData(allQueries);
  
  Logger.log('\n===== クエリ分析完了 =====');
  return allQueries;
}

/**
 * GSC_RAWシートからデータ取得
 */
function getGSCRawData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('GSC_RAW');
  
  if (!sheet) {
    Logger.log('⚠️ GSC_RAWシートが見つかりません');
    return [];
  }
  
  const data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) {
    return [];
  }
  
  // ヘッダー行を除く
  const dataRows = data.slice(1);
  
  return dataRows.map(row => ({
    date: row[0],
    pageUrl: normalizeUrl(row[1] || ''),
    query: (row[2] || '').trim(),
    position: parseFloat(row[3]) || 0,
    clicks: parseInt(row[4]) || 0,
    impressions: parseInt(row[5]) || 0,
    ctr: parseFloat(row[6]) || 0
  })).filter(row => row.pageUrl && row.query);
}

/**
 * GyronSEO_RAWからターゲットキーワード一覧を取得（修正版）
 */
function getTargetKeywords() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gyronSheet = ss.getSheetByName('GyronSEO_RAW');
  
  if (!gyronSheet) {
    Logger.log('⚠️ GyronSEO_RAWシートが見つかりません');
    return [];
  }
  
  const data = gyronSheet.getDataRange().getValues();
  
  if (data.length <= 1) {
    Logger.log('⚠️ GyronSEO_RAWにデータがありません');
    return [];
  }
  
  // ヘッダー行を除く
  const dataRows = data.slice(1);
  
  // 各行をオブジェクト化（正しい列マッピング）
  const keywords = dataRows
    .filter(row => row[0] && row[0].trim() !== '') // A列（keyword）が空でない
    .map(row => ({
      keyword: (row[0] || '').trim(),              // A列: keyword
      pageUrl: normalizeUrl(row[1] || ''),         // B列: url
      position: parseFloat(row[4]) || 101,         // E列: latest_position
      position7dAgo: parseFloat(row[5]) || 101,    // F列: position_7d_ago
      position30dAgo: parseFloat(row[6]) || 101,   // G列: position_30d_ago
      position90dAgo: parseFloat(row[7]) || 101,   // H列: position_90d_ago
      trend: row[8] || '--',                       // I列: position_trend
      searchVolume: 0  // GyronSEO_RAWにはこの列がない（Day 11-12で追加予定）
    }));
  
  Logger.log(`ターゲットKW取得: ${keywords.length}件`);
  return keywords;
}

/**
 * ページごとにクエリをグループ化
 */
function groupByPage(gscData) {
  const groups = {};
  
  for (const row of gscData) {
    if (!groups[row.pageUrl]) {
      groups[row.pageUrl] = [];
    }
    groups[row.pageUrl].push(row);
  }
  
  return groups;
}

/**
 * 上位N件のクエリを抽出（表示回数降順）
 */
function getTopQueries(queries, limit) {
  return queries
    .sort((a, b) => b.impressions - a.impressions)
    .slice(0, limit);
}

/**
 * クエリをスコアリング
 */
function scoreQuery(query, pageUrl, targetKeywords) {
  // CTRギャップスコア
  const ctrGapScore = calculateCTRGapScore(query.position, query.ctr);
  
  // 改善余地スコア
  const improvementScore = calculateImprovementScore(query.position, query.impressions);
  
  // CV近接度
  const cvProximity = calculateCVProximity(query.query);
  
  // ターゲットKW一致チェック
  const targetKWMatch = checkTargetKeywordMatch(query.query, pageUrl, targetKeywords);
  
  // クエリスコア（総合）
  const queryScore = Math.round(
    ctrGapScore * 0.4 + 
    improvementScore * 0.4 + 
    (cvProximity === '高' ? 20 : cvProximity === '中' ? 10 : 0)
  );
  
  return {
    queryId: `q_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`,
    pageUrl: query.pageUrl,
    query: query.query,
    position: query.position,
    clicks: query.clicks,
    impressions: query.impressions,
    ctr: query.ctr,
    queryScore: queryScore,
    cvProximity: cvProximity,
    targetKwMatch: targetKWMatch,
    improvementPotential: improvementScore,
    lastUpdated: new Date()
  };
}

/**
 * CTRギャップスコア計算
 */
function calculateCTRGapScore(position, actualCtr) {
  // 順位ごとの期待CTR（業界平均）
  const expectedCtrMap = {
    1: 0.30, 2: 0.15, 3: 0.10, 4: 0.07, 5: 0.05,
    6: 0.04, 7: 0.03, 8: 0.025, 9: 0.02, 10: 0.015
  };
  
  const expectedCtr = expectedCtrMap[Math.round(position)] || 0.01;
  const gap = Math.max(0, expectedCtr - actualCtr / 100);
  
  // ギャップが大きいほど高スコア
  return Math.min(100, gap * 500);
}

/**
 * 改善余地スコア計算
 */
function calculateImprovementScore(position, impressions) {
  let positionScore = 0;
  
  if (position >= 4 && position <= 7) {
    positionScore = 100; // 最も改善余地あり
  } else if (position >= 8 && position <= 10) {
    positionScore = 80;
  } else if (position >= 11 && position <= 20) {
    positionScore = 50;
  } else if (position >= 2 && position <= 3) {
    positionScore = 70; // 1位を狙える
  } else {
    positionScore = 20;
  }
  
  // 表示回数による重み付け
  const impressionWeight = Math.min(1, impressions / 1000);
  
  return Math.round(positionScore * (0.5 + impressionWeight * 0.5));
}

/**
 * CV近接度判定
 */
function calculateCVProximity(query) {
  const highCvWords = ['おすすめ', '比較', 'ランキング', '評判', '口コミ', '選び方', '安い', '激安', '最安値'];
  const midCvWords = ['メリット', 'デメリット', '違い', '方法', '手順'];
  
  const lowerQuery = query.toLowerCase();
  
  if (highCvWords.some(word => lowerQuery.includes(word))) {
    return '高';
  } else if (midCvWords.some(word => lowerQuery.includes(word))) {
    return '中';
  } else {
    return '低';
  }
}

/**
 * ターゲットKW一致チェック
 */
function checkTargetKeywordMatch(query, pageUrl, targetKeywords) {
  const matchingKW = targetKeywords.find(kw => 
    kw.pageUrl === pageUrl && query.includes(kw.keyword)
  );
  
  return matchingKW ? true : false;
}

/**
 * クエリ分析データをシートに書き込み
 */
function writeQueryAnalysisData(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('クエリ分析');
  
  if (!sheet) {
    Logger.log('⚠️ クエリ分析シートが見つかりません');
    return;
  }
  
  if (data.length === 0) {
    Logger.log('書き込むデータがありません');
    return;
  }
  
  // 既存データをクリア（ヘッダー行以外）
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }
  
  // データを2次元配列に変換
  const values = data.map(row => [
    row.queryId,
    row.pageUrl,
    row.query,
    row.position,
    row.clicks,
    row.impressions,
    row.ctr,
    row.queryScore,
    row.cvProximity,
    row.targetKwMatch,
    row.improvementPotential,
    row.lastUpdated
  ]);
  
  // 書き込み
  sheet.getRange(2, 1, values.length, values[0].length).setValues(values);
  
  Logger.log(`クエリ分析シートに${values.length}行を書き込みました`);
}

/**
 * URL正規化関数
 */
function normalizeUrl(url) {
  if (!url) return '';
  
  // プロトコルとドメインを削除してパスのみ取得
  let normalized = url
    .replace(/^https?:\/\/[^\/]+/, '')
    .replace(/\/$/, ''); // 末尾のスラッシュを削除
  
  // 空の場合は'/'を返す
  return normalized || '/';
}

/**
 * テスト関数: クエリ分析実行（上位10ページのみ）
 */
function testAnalyzeQueries() {
  Logger.log('===== クエリ分析 テスト開始 =====\n');
  
  const result = analyzeQueries(10); // 上位10ページのみ
  
  Logger.log(`\n書き込み完了: ${result.length}件`);
  Logger.log('\n===== テスト完了 =====');
}

// ===================================================================
// Day 7-8.3: ターゲットKW分析実装
// ===================================================================

/**
 * ターゲットKW分析実行
 */
function analyzeTargetKeywords() {
  Logger.log('===== ターゲットKW分析開始 =====\n');
  
  // GyronSEOターゲットKW取得
  const targetKeywords = getTargetKeywords();
  Logger.log(`ターゲットKW: ${targetKeywords.length}件`);
  
  // GSC_RAWからデータ取得
  const gscData = getGSCRawData();
  Logger.log(`GSC_RAW: ${gscData.length}行`);
  
  // ページ+クエリでGSCデータをインデックス化
  const gscIndex = {};
  for (const row of gscData) {
    const key = `${row.pageUrl}|${row.query}`;
    gscIndex[key] = row;
  }
  
  const results = [];
  let matchCount = 0;
  
  for (const kw of targetKeywords) {
    // GSCデータと照合
    const gscMatch = Object.values(gscIndex).find(gsc => 
      gsc.pageUrl === kw.pageUrl && gsc.query.includes(kw.keyword)
    );
    
    if (gscMatch) {
      matchCount++;
    }
    
    // スコアリング
    const performanceScore = calculatePerformanceScore(kw, gscMatch);
    const searchVolumeScore = calculateSearchVolumeScore(kw.searchVolume);
    const strategicValueScore = calculateStrategicValueScore(kw, gscMatch);
    
    const kwScore = Math.round(
      performanceScore * 0.4 +
      searchVolumeScore * 0.3 +
      strategicValueScore * 0.3
    );
    
    // 除外スコア計算
    const removalScore = calculateRemovalScore(kw, gscMatch);
    
    // ステータス判定
    let status = '維持';
    if (removalScore >= 100) {
      status = '除外候補';
    } else if (kwScore >= 80) {
      status = '最優先改善';
    } else if (kwScore >= 60) {
      status = '要改善';
    }
    
    results.push({
      keywordId: `kw_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`,
      pageUrl: kw.pageUrl,
      targetKeyword: kw.keyword,
      gyronPosition: kw.position,
      gscPosition: gscMatch ? gscMatch.position : '',
      gscClicks: gscMatch ? gscMatch.clicks : 0,
      gscImpressions: gscMatch ? gscMatch.impressions : 0,
      gscCtr: gscMatch ? gscMatch.ctr : 0,
      searchVolume: kw.searchVolume,
      competitionLevel: '中', // Day 11-12で更新予定
      kwScore: kwScore,
      performanceScore: performanceScore,
      searchVolumeScore: searchVolumeScore,
      strategicValueScore: strategicValueScore,
      removalScore: removalScore,
      status: status,
      notes: generateNotes(kw, gscMatch, status),
      lastUpdated: new Date()
    });
  }
  
  Logger.log(`\nGSCマッチ: ${matchCount}件 (${(matchCount/targetKeywords.length*100).toFixed(1)}%)`);
  
  // ターゲットKW分析シートに書き込み
  writeTargetKeywordAnalysisData(results);
  
  // スコア分布を集計
  const highScore = results.filter(r => r.kwScore >= 80).length;
  const removalCandidates = results.filter(r => r.removalScore >= 100).length;
  
  Logger.log(`\n高スコア(80+): ${highScore}件`);
  Logger.log(`除外候補(100+): ${removalCandidates}件`);
  
  // 上位5件を表示
  const top5 = results
    .sort((a, b) => b.kwScore - a.kwScore)
    .slice(0, 5);
  
  Logger.log(`\n===== 高スコア上位5件 =====`);
  for (let i = 0; i < top5.length; i++) {
    const kw = top5[i];
    Logger.log(`${i + 1}. ${kw.targetKeyword} (${kw.pageUrl})`);
    Logger.log(`   KWスコア: ${kw.kwScore}, 順位: ${kw.gyronPosition}, 検索Vol: ${kw.searchVolume}`);
    Logger.log(`   GSCマッチ: ${kw.gscPosition ? 'true' : 'false'}, 表示: ${kw.gscImpressions}, CTR: ${kw.gscCtr.toFixed(2)}%`);
    Logger.log(`   ステータス: ${kw.status}, メモ: ${kw.notes}`);
  }
  
  Logger.log('\n===== ターゲットKW分析完了 =====');
  return results;
}

/**
 * パフォーマンススコア計算
 */
function calculatePerformanceScore(kw, gscMatch) {
  let score = 0;
  
  // 順位スコア
  if (kw.position <= 3) {
    score += 40; // 上位キープ
  } else if (kw.position <= 10) {
    score += 30;
  } else if (kw.position <= 20) {
    score += 20;
  } else {
    score += 10;
  }
  
  // CTRスコア（GSCデータがあれば）
  if (gscMatch) {
    if (gscMatch.ctr > 5) {
      score += 30;
    } else if (gscMatch.ctr > 2) {
      score += 20;
    } else {
      score += 10;
    }
  }
  
  // 表示回数スコア
  if (gscMatch && gscMatch.impressions > 1000) {
    score += 30;
  } else if (gscMatch && gscMatch.impressions > 100) {
    score += 20;
  } else {
    score += 10;
  }
  
  return Math.min(100, score);
}

/**
 * 検索ボリュームスコア計算
 */
function calculateSearchVolumeScore(volume) {
  if (volume >= 1000) return 100;
  if (volume >= 500) return 70;
  if (volume >= 100) return 40;
  if (volume >= 10) return 20;
  return 0; // Day 11-12でデータ追加予定
}

/**
 * 戦略的価値スコア計算
 */
function calculateStrategicValueScore(kw, gscMatch) {
  let score = 50; // 基本点
  
  // トレンドボーナス
  if (kw.trend === '↑') {
    score += 30;
  } else if (kw.trend === '→') {
    score += 10;
  }
  
  // GSCマッチボーナス
  if (gscMatch) {
    score += 20;
  }
  
  return Math.min(100, score);
}

/**
 * 除外スコア計算
 */
function calculateRemovalScore(kw, gscMatch) {
  let score = 0;
  
  // 基準1: 順位30位以下が継続
  if (kw.position >= 30) {
    score += 80;
  }
  
  // 基準2: 月間検索ボリューム10未満
  if (kw.searchVolume < 10) {
    score += 70;
  }
  
  // 基準3: 実際の表示回数が5未満
  if (gscMatch && gscMatch.impressions < 5) {
    score += 70;
  } else if (!gscMatch) {
    score += 70; // GSCにデータなし = 検索されていない
  }
  
  // 基準4: GyronとGSCで大きなズレ（20位以上の差）
  if (gscMatch && Math.abs(kw.position - gscMatch.position) >= 20) {
    score += 40;
  }
  
  // 基準5: 下降トレンド
  if (kw.trend === '↓') {
    score += 60;
  }
  
  return score;
}

/**
 * メモ生成
 */
function generateNotes(kw, gscMatch, status) {
  const notes = [];
  
  if (kw.position <= 3) {
    notes.push('上位キープ');
  }
  
  if (kw.searchVolume < 10) {
    notes.push('検索需要極小');
  }
  
  if (!gscMatch || (gscMatch && gscMatch.impressions < 5)) {
    notes.push('実際の検索なし');
  }
  
  if (status === '除外候補') {
    notes.push('⚠️除外推奨');
  }
  
  return notes.join(', ');
}

/**
 * ターゲットKW分析データをシートに書き込み（18列対応）
 */
function writeTargetKeywordAnalysisData(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('ターゲットKW分析');
  
  if (!sheet) {
    Logger.log('⚠️ ターゲットKW分析シートが見つかりません');
    return;
  }
  
  if (data.length === 0) {
    Logger.log('書き込むデータがありません');
    return;
  }
  
  // 既存データをクリア（ヘッダー行以外）
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }
  
  // データを2次元配列に変換（18列）
  const values = data.map(row => [
    row.keywordId,
    row.pageUrl,
    row.targetKeyword,
    row.gyronPosition,
    row.gscPosition || '',
    row.gscClicks,
    row.gscImpressions,
    row.gscCtr,
    row.searchVolume,
    row.competitionLevel,
    row.kwScore,
    row.performanceScore,
    row.searchVolumeScore,
    row.strategicValueScore,
    row.removalScore,
    row.status,
    row.notes,
    row.lastUpdated
  ]);
  
  // 書き込み
  sheet.getRange(2, 1, values.length, values[0].length).setValues(values);
  
  // 書式設定
  // H列（CTR）: パーセント表示
  sheet.getRange(2, 8, values.length, 1).setNumberFormat('0.00"%"');
  
  // K-O列（スコア5列）: 数値書式
  sheet.getRange(2, 11, values.length, 5).setNumberFormat('0');
  
  // P列（status）に条件付き書式
  applyConditionalFormattingToStatus(sheet, values.length);
  
  Logger.log(`ターゲットKW分析シートに${values.length}行を書き込みました`);
}

/**
 * status列に条件付き書式を適用
 */
function applyConditionalFormattingToStatus(sheet, dataRows) {
  const range = sheet.getRange(2, 16, dataRows, 1); // P列（status）
  
  // 既存の条件付き書式をクリア
  range.clearFormat();
  
  const rules = sheet.getConditionalFormatRules();
  
  // 除外候補: 赤
  const rule1 = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('除外候補')
    .setBackground('#f4cccc')
    .setRanges([range])
    .build();
  
  // 最優先改善: オレンジ
  const rule2 = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('最優先改善')
    .setBackground('#fce5cd')
    .setRanges([range])
    .build();
  
  // 要改善: 黄色
  const rule3 = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('要改善')
    .setBackground('#fff2cc')
    .setRanges([range])
    .build();
  
  // 維持: 緑
  const rule4 = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('維持')
    .setBackground('#d9ead3')
    .setRanges([range])
    .build();
  
  rules.push(rule1, rule2, rule3, rule4);
  sheet.setConditionalFormatRules(rules);
}

/**
 * テスト関数: ターゲットKW分析実行
 */
function testAnalyzeTargetKeywords() {
  Logger.log('===== ターゲットKW分析 テスト開始 =====\n');
  
  const result = analyzeTargetKeywords();
  
  Logger.log(`\n書き込み完了: ${result.length}件`);
  Logger.log('\n===== テスト完了 =====');
}

// ===================================================================
// Day 7-8.4: KW自動スクリーニング実装（条件付き保護）
// ===================================================================

/**
 * KW自動スクリーニング（Day 7-8.4）修正版
 * ターゲットKW分析シートから除外候補を自動抽出
 * - 1ページ1KWの場合は条件付き保護
 * - removalScore < 150: 完全保護
 * - removalScore >= 150: 除外候補に含めるが⚠️マーク表示
 */
function screenKeywordsForRemoval() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetKWSheet = ss.getSheetByName('ターゲットKW分析');
  const removalSheet = ss.getSheetByName('KW除外候補');
  
  if (!targetKWSheet) {
    Logger.log('⚠️ ターゲットKW分析シートが見つかりません');
    return;
  }
  
  if (!removalSheet) {
    Logger.log('⚠️ KW除外候補シートが見つかりません');
    return;
  }
  
  // ターゲットKW分析シートからデータ取得
  const data = targetKWSheet.getDataRange().getValues();
  
  if (data.length <= 1) {
    Logger.log('⚠️ ターゲットKW分析シートにデータがありません');
    return;
  }
  
  // ステップ1: ページごとのKW数をカウント
  const kwCountByPage = {};
  for (let i = 1; i < data.length; i++) {
    const pageUrl = data[i][1]; // B列: page_url
    kwCountByPage[pageUrl] = (kwCountByPage[pageUrl] || 0) + 1;
  }
  
  Logger.log(`\nページごとのKW数カウント完了: ${Object.keys(kwCountByPage).length}ページ`);
  
  // ステップ2: 除外候補を抽出
  const removalCandidates = [];
  let protectedCount = 0; // 保護されたKW数
  let warningCount = 0;   // ⚠️マーク付きKW数
  
  // ヘッダー行を除いてループ
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    const keywordId = row[0];
    const pageUrl = row[1];
    const targetKeyword = row[2];
    const gyronPosition = parseFloat(row[3]) || 101;
    const gscPosition = parseFloat(row[4]) || 101;
    const gscClicks = parseFloat(row[5]) || 0;
    const gscImpressions = parseFloat(row[6]) || 0;
    const searchVolume = parseFloat(row[8]) || 0;
    const removalScore = parseFloat(row[14]) || 0;
    const status = row[15] || '';
    
    // 除外候補（removalScore >= 100）のみ処理
    if (removalScore >= 100) {
      const kwCount = kwCountByPage[pageUrl] || 0;
      const isOnlyKW = (kwCount === 1);
      
      // 【条件付き保護】1ページ1KW + removalScore < 150 → 完全保護
      if (isOnlyKW && removalScore < 150) {
        protectedCount++;
        Logger.log(`🛡️ 保護: ${targetKeyword} (唯一のKW、スコア${removalScore}点 < 150点)`);
        continue; // 除外候補に含めない
      }
      
      // 除外理由を特定
      const reasons = [];
      
      // ⚠️マーク: 唯一のKWで重度の除外候補
      if (isOnlyKW && removalScore >= 150) {
        reasons.push('⚠️唯一のKW（要慎重判断）');
        warningCount++;
      }
      
      // 基準1: 順位30位以下が継続
      if (gyronPosition >= 30) {
        reasons.push('順位30位以下が継続（改善の見込み薄）');
      }
      
      // 基準2: 月間検索ボリューム10未満
      if (searchVolume < 10) {
        reasons.push('月間検索ボリューム10未満（需要なし）');
      }
      
      // 基準3: 実際の表示回数が5未満
      if (gscImpressions < 5) {
        reasons.push('GSC表示回数5未満（実際の検索なし）');
      }
      
      // 基準4: GyronとGSCで大きなズレ（20位以上の差）
      if (Math.abs(gyronPosition - gscPosition) >= 20) {
        reasons.push('GyronとGSCで20位以上の差（計測ミス疑い）');
      }
      
      // notesの作成
      let notes = `順位: ${gyronPosition}位, 表示: ${gscImpressions}回, 検索Vol: ${searchVolume}`;
      
      // ⚠️マーク: notesにも追記
      if (isOnlyKW && removalScore >= 150) {
        notes = `⚠️このページの唯一のターゲットKW | ${notes}`;
      }
      
      removalCandidates.push({
        keywordId: keywordId,
        targetKeyword: targetKeyword,
        pageUrl: pageUrl,
        removalScore: removalScore,
        removalReasons: reasons.join(', '),
        userDecision: '', // ユーザーが後で判断
        decisionDate: '',
        notes: notes,
        isOnlyKW: isOnlyKW && removalScore >= 150 // ⚠️マーク判定用
      });
    }
  }
  
  Logger.log(`\n===== KW自動スクリーニング結果 =====`);
  Logger.log(`総キーワード数: ${data.length - 1}件`);
  Logger.log(`除外候補: ${removalCandidates.length}件`);
  Logger.log(`  └ 通常の除外候補: ${removalCandidates.length - warningCount}件`);
  Logger.log(`  └ ⚠️要慎重判断（唯一のKW）: ${warningCount}件`);
  Logger.log(`保護されたKW: ${protectedCount}件（1ページ1KW + スコア<150点）`);
  
  if (removalCandidates.length === 0) {
    Logger.log('✅ 除外候補はありませんでした');
    return;
  }
  
  // KW除外候補シートに書き込み
  writeKeywordRemovalData(removalCandidates);
  
  Logger.log(`\n===== 除外理由トップ5 =====`);
  for (let i = 0; i < Math.min(5, removalCandidates.length); i++) {
    const candidate = removalCandidates[i];
    Logger.log(`${i + 1}. ${candidate.targetKeyword}`);
    Logger.log(`   除外スコア: ${candidate.removalScore}点`);
    Logger.log(`   理由: ${candidate.removalReasons}`);
    Logger.log(`   ${candidate.notes}`);
  }
  
  Logger.log(`\n✅ KW除外候補シートに${removalCandidates.length}件を書き込みました`);
}

/**
 * KW除外候補データをシートに書き込み（修正版）
 */
function writeKeywordRemovalData(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('KW除外候補');
  
  if (!sheet) {
    Logger.log('⚠️ KW除外候補シートが見つかりません');
    return;
  }
  
  if (data.length === 0) {
    Logger.log('書き込むデータがありません');
    return;
  }
  
  // 既存データをクリア（ヘッダー行以外）
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
    sheet.clearConditionalFormatRules(); // 既存の条件付き書式もクリア
  }
  
  // データを2次元配列に変換
  const values = data.map(row => [
    row.keywordId,
    row.targetKeyword,
    row.pageUrl,
    row.removalScore,
    row.removalReasons,
    row.userDecision,
    row.decisionDate,
    row.notes
  ]);
  
  // 書き込み
  sheet.getRange(2, 1, values.length, values[0].length).setValues(values);
  
  // 書式設定
  // D列（removalScore）に数値書式
  sheet.getRange(2, 4, values.length, 1).setNumberFormat('0');
  
  // 条件付き書式を適用
  applyConditionalFormattingToRemovalSheet(sheet, data, values.length);
  
  Logger.log(`KW除外候補シートに${values.length}行を書き込みました`);
}

/**
 * KW除外候補シートに条件付き書式を適用（修正版）
 * - removalScoreによる色分け
 * - ⚠️マーク（唯一のKW）の強調表示
 */
function applyConditionalFormattingToRemovalSheet(sheet, data, dataRows) {
  const rules = [];
  
  // ルール1-3: removalScore（D列）の色分け
  const scoreRange = sheet.getRange(2, 4, dataRows, 1);
  
  // 150点以上: 濃い赤
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(150)
    .setBackground('#ea4335')
    .setFontColor('#ffffff')
    .setRanges([scoreRange])
    .build());
  
  // 120-149点: 赤
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(120, 149)
    .setBackground('#f4cccc')
    .setRanges([scoreRange])
    .build());
  
  // 100-119点: オレンジ
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(100, 119)
    .setBackground('#fce5cd')
    .setRanges([scoreRange])
    .build());
  
  // ルール4: ⚠️マーク（唯一のKW）の強調表示
  // removal_reasons列（E列）に⚠️が含まれる場合、黄色背景
  const reasonsRange = sheet.getRange(2, 5, dataRows, 1);
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('⚠️')
    .setBackground('#fff2cc') // 黄色
    .setBold(true)
    .setRanges([reasonsRange])
    .build());
  
  // ルール5: notes列（H列）に⚠️が含まれる場合も黄色背景
  const notesRange = sheet.getRange(2, 8, dataRows, 1);
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('⚠️')
    .setBackground('#fff2cc') // 黄色
    .setBold(true)
    .setRanges([notesRange])
    .build());
  
  sheet.setConditionalFormatRules(rules);
  
  Logger.log('条件付き書式を適用しました');
}

/**
 * テスト関数: KW自動スクリーニング実行
 */
function testScreenKeywordsForRemoval() {
  Logger.log('===== KW自動スクリーニング テスト開始 =====\n');
  
  screenKeywordsForRemoval();
  
  Logger.log('\n===== テスト完了 =====');
}

// ===================================================================
// データ品質診断関数
// ===================================================================

/**
 * GyronSEO_RAWデータの品質診断
 */
function checkGyronSEOData() {
  Logger.log('===== GyronSEO_RAW データ品質診断 =====\n');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('GyronSEO_RAW');
  
  if (!sheet) {
    Logger.log('⚠️ GyronSEO_RAWシートが見つかりません');
    return;
  }
  
  const data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) {
    Logger.log('⚠️ GyronSEO_RAWにデータがありません');
    return;
  }
  
  const dataRows = data.slice(1);
  
  Logger.log(`総キーワード数: ${dataRows.length}件\n`);
  
  // 順位分布
  const positionDist = {
    '1-10': 0,
    '11-30': 0,
    '31-50': 0,
    '51-100': 0,
    '101': 0
  };
  
  // 検索ボリューム分布
  const volumeDist = {
    '0': 0,
    '1-10': 0,
    '11-100': 0,
    '101-500': 0,
    '501+': 0
  };
  
  // トレンド分布
  const trendDist = {
    '↑': 0,
    '→': 0,
    '↓': 0,
    '--': 0
  };
  
  for (const row of dataRows) {
    const position = parseFloat(row[4]) || 101; // E列
    const volume = parseFloat(row[8]) || 0;     // I列（存在しない可能性）
    const trend = row[8] || '--';               // I列
    
    // 順位分布
    if (position <= 10) positionDist['1-10']++;
    else if (position <= 30) positionDist['11-30']++;
    else if (position <= 50) positionDist['31-50']++;
    else if (position <= 100) positionDist['51-100']++;
    else positionDist['101']++;
    
    // 検索ボリューム分布
    if (volume === 0) volumeDist['0']++;
    else if (volume <= 10) volumeDist['1-10']++;
    else if (volume <= 100) volumeDist['11-100']++;
    else if (volume <= 500) volumeDist['101-500']++;
    else volumeDist['501+']++;
    
    // トレンド分布
    if (trend === '↑') trendDist['↑']++;
    else if (trend === '→') trendDist['→']++;
    else if (trend === '↓') trendDist['↓']++;
    else trendDist['--']++;
  }
  
  Logger.log('【順位分布】');
  for (const [range, count] of Object.entries(positionDist)) {
    const pct = (count / dataRows.length * 100).toFixed(1);
    Logger.log(`  ${range}位: ${count}件 (${pct}%)`);
  }
  
  Logger.log('\n【検索ボリューム分布】');
  for (const [range, count] of Object.entries(volumeDist)) {
    const pct = (count / dataRows.length * 100).toFixed(1);
    Logger.log(`  ${range}: ${count}件 (${pct}%)`);
  }
  
  Logger.log('\n【トレンド分布】');
  for (const [symbol, count] of Object.entries(trendDist)) {
    const pct = (count / dataRows.length * 100).toFixed(1);
    let label = symbol;
    if (symbol === '↑') label = '上昇(↑)';
    else if (symbol === '→') label = '横ばい(→)';
    else if (symbol === '↓') label = '下降(↓)';
    else if (symbol === '--') label = '不明(--)';
    Logger.log(`  ${label}: ${count}件 (${pct}%)`);
  }
  
  // 警告判定
  Logger.log('\n【診断結果】');
  
  if (positionDist['101'] / dataRows.length > 0.5) {
    Logger.log(`⚠️ 警告: 圏外(101位)が${(positionDist['101'] / dataRows.length * 100).toFixed(1)}%と非常に多い`);
    Logger.log('   → 除外候補が多くなる主要因');
  }
  
  if (volumeDist['0'] / dataRows.length > 0.5) {
    Logger.log(`⚠️ 警告: 検索ボリューム0が${(volumeDist['0'] / dataRows.length * 100).toFixed(1)}%と非常に多い`);
    Logger.log('   → 除外候補が多くなる要因');
  }
  
  Logger.log('\n===== 診断完了 =====');
}

/**
 * デバッグ用: KW除外候補シートを強制クリア
 */
function forceClearRemovalSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('KW除外候補');
  
  if (!sheet) {
    Logger.log('⚠️ KW除外候補シートが見つかりません');
    return;
  }
  
  // 既存データを完全クリア（ヘッダー行以外）
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
    Logger.log(`✅ ${lastRow - 1}行をクリアしました`);
  } else {
    Logger.log('ℹ️ クリアするデータがありません');
  }
  
  // 条件付き書式もクリア
  sheet.clearConditionalFormatRules();
  Logger.log('✅ 条件付き書式をクリアしました');
  
  Logger.log('\n===== 強制クリア完了 =====');
  Logger.log('次に testScreenKeywordsForRemoval() を実行してください');
}