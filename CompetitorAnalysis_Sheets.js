/**
 * ============================================================================
 * Day 11-12: 競合分析シート作成（修正版）
 * ============================================================================
 * 競合分析シートとDA履歴シート（スマートキャッシング用）を作成
 */

/**
 * 競合分析シートを作成
 */
function createCompetitorAnalysisSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 既存シートチェック
  let sheet = ss.getSheetByName('競合分析');
  if (sheet) {
    Logger.log('⚠ 競合分析シートは既に存在します');
    ss.deleteSheet(sheet);
    Logger.log('既存シートを削除しました');
  }
  
  // 新規シート作成
  sheet = ss.insertSheet('競合分析');
  Logger.log('✓ 競合分析シートを作成しました');
  
  // ヘッダー行作成
  const headers = [
    // 基本情報（4列）
    'analysis_id',
    'target_keyword',
    'page_url',
    'analysis_date',
    
    // 自社サイト（2列）
    'own_site_da',
    'own_site_current_rank',
    
    // 検索結果Top10（20列）
    'rank_1_url', 'rank_1_da',
    'rank_2_url', 'rank_2_da',
    'rank_3_url', 'rank_3_da',
    'rank_4_url', 'rank_4_da',
    'rank_5_url', 'rank_5_da',
    'rank_6_url', 'rank_6_da',
    'rank_7_url', 'rank_7_da',
    'rank_8_url', 'rank_8_da',
    'rank_9_url', 'rank_9_da',
    'rank_10_url', 'rank_10_da',
    
    // 競合分析結果（4列）
    'avg_da_top10',
    'weaker_sites_count',
    'weaker_sites_highest_rank',
    'da_distribution_type',
    
    // 勝算度評価（5列）
    'winnable_score',
    'target_rank',
    'competition_level',
    'rewrite_roi',
    'last_updated'
  ];
  
  // ヘッダー設定
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4a86e8');
  headerRange.setFontColor('#ffffff');
  
  // 列幅自動調整
  for (let i = 1; i <= headers.length; i++) {
    sheet.autoResizeColumn(i);
  }
  
  // 最小列幅設定
  sheet.setColumnWidth(1, 150);  // analysis_id
  sheet.setColumnWidth(2, 200);  // target_keyword
  sheet.setColumnWidth(3, 250);  // page_url
  sheet.setColumnWidth(4, 120);  // analysis_date
  
  // URL列の幅設定
  for (let i = 7; i <= 25; i += 2) {
    sheet.setColumnWidth(i, 250);  // rank_X_url
  }
  
  // データ検証・書式設定
  setupCompetitorAnalysisFormatting(sheet);
  
  Logger.log('✓ ヘッダー設定完了（27列）');
  Logger.log('✓ 列幅調整完了');
  Logger.log('✓ 書式設定完了');
}

/**
 * 競合分析シートの書式設定
 */
function setupCompetitorAnalysisFormatting(sheet) {
  // 日付列（D列）
  const dateRange = sheet.getRange('D2:D1000');
  dateRange.setNumberFormat('yyyy-mm-dd');
  
  // DA列（E列、H列、J列、...）
  const daColumns = [5, 8, 10, 12, 14, 16, 18, 20, 22, 24, 26, 27];
  daColumns.forEach(col => {
    const range = sheet.getRange(2, col, 999);
    range.setNumberFormat('0');
  });
  
  // 順位列（F列）
  const rankRange = sheet.getRange('F2:F1000');
  rankRange.setNumberFormat('0');
  
  // weaker_sites_count列（AB列 = 28列目）
  const weakerCountRange = sheet.getRange(2, 28, 999);
  weakerCountRange.setNumberFormat('0');
  
  // weaker_sites_highest_rank列（AC列 = 29列目）
  const weakerRankRange = sheet.getRange(2, 29, 999);
  weakerRankRange.setNumberFormat('0');
  
  // winnable_score列（AE列 = 31列目）
  const winnableRange = sheet.getRange(2, 31, 999);
  winnableRange.setNumberFormat('0');
  
  // target_rank列（AF列 = 32列目）
  const targetRankRange = sheet.getRange(2, 32, 999);
  targetRankRange.setNumberFormat('0');
  
  // last_updated列（AI列 = 35列目）
  const lastUpdatedRange = sheet.getRange(2, 35, 999);
  lastUpdatedRange.setNumberFormat('yyyy-mm-dd hh:mm:ss');
  
  // 条件付き書式設定
  setupConditionalFormatting(sheet);
  
  Logger.log('✓ 書式設定完了');
}

/**
 * 条件付き書式設定（修正版）
 */
function setupConditionalFormatting(sheet) {
  const rules = [];
  
  // 勝算度スコアの条件付き書式（AE列 = 31列目）
  const winnableRange = sheet.getRange(2, 31, 999);
  
  rules.push(
    // 90-100点: 濃い緑
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(90)
      .setBackground('#00ff00')
      .setFontColor('#000000')
      .setRanges([winnableRange])
      .build(),
    
    // 60-89点: 緑
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(60, 89)
      .setBackground('#93c47d')
      .setFontColor('#000000')
      .setRanges([winnableRange])
      .build(),
    
    // 40-59点: 黄色
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(40, 59)
      .setBackground('#ffd966')
      .setFontColor('#000000')
      .setRanges([winnableRange])
      .build(),
    
    // 1-39点: オレンジ
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(1, 39)
      .setBackground('#f6b26b')
      .setFontColor('#000000')
      .setRanges([winnableRange])
      .build(),
    
    // 0点: 赤
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberEqualTo(0)
      .setBackground('#e06666')
      .setFontColor('#ffffff')
      .setRanges([winnableRange])
      .build()
  );
  
  // 競合レベルの条件付き書式（AG列 = 33列目）
  const levelRange = sheet.getRange(2, 33, 999);
  
  rules.push(
    // 易: 緑
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('易')
      .setBackground('#93c47d')
      .setFontColor('#000000')
      .setRanges([levelRange])
      .build(),
    
    // 中: 黄色
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('中')
      .setBackground('#ffd966')
      .setFontColor('#000000')
      .setRanges([levelRange])
      .build(),
    
    // 難: オレンジ
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('難')
      .setBackground('#f6b26b')
      .setFontColor('#000000')
      .setRanges([levelRange])
      .build(),
    
    // 激戦: 赤
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('激戦')
      .setBackground('#e06666')
      .setFontColor('#ffffff')
      .setRanges([levelRange])
      .build()
  );
  
  // すべてのルールを一度に設定
  sheet.setConditionalFormatRules(rules);
  Logger.log('✓ 条件付き書式設定完了');
}

/**
 * DA履歴シートを作成（スマートキャッシング用）
 */
function createDAHistorySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 既存シートチェック
  let sheet = ss.getSheetByName('DA履歴');
  if (sheet) {
    Logger.log('⚠ DA履歴シートは既に存在します');
    ss.deleteSheet(sheet);
    Logger.log('既存シートを削除しました');
  }
  
  // 新規シート作成
  sheet = ss.insertSheet('DA履歴');
  Logger.log('✓ DA履歴シートを作成しました');
  
  // ヘッダー行作成
  const headers = [
    'domain',
    'da',
    'pa',
    'last_updated',
    'cache_until'
  ];
  
  // ヘッダー設定
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#6aa84f');
  headerRange.setFontColor('#ffffff');
  
  // 列幅設定
  sheet.setColumnWidth(1, 250);  // domain
  sheet.setColumnWidth(2, 80);   // da
  sheet.setColumnWidth(3, 80);   // pa
  sheet.setColumnWidth(4, 180);  // last_updated
  sheet.setColumnWidth(5, 180);  // cache_until
  
  // データ検証・書式設定
  setupDAHistoryFormatting(sheet);
  
  Logger.log('✓ ヘッダー設定完了（5列）');
  Logger.log('✓ 列幅調整完了');
  Logger.log('✓ 書式設定完了');
}

/**
 * DA履歴シートの書式設定
 */
function setupDAHistoryFormatting(sheet) {
  // DA列（B列）
  const daRange = sheet.getRange('B2:B10000');
  daRange.setNumberFormat('0');
  
  // PA列（C列）
  const paRange = sheet.getRange('C2:C10000');
  paRange.setNumberFormat('0');
  
  // 日付列（D列、E列）
  const dateRange1 = sheet.getRange('D2:D10000');
  dateRange1.setNumberFormat('yyyy-mm-dd hh:mm:ss');
  
  const dateRange2 = sheet.getRange('E2:E10000');
  dateRange2.setNumberFormat('yyyy-mm-dd hh:mm:ss');
  
  // 条件付き書式（キャッシュ有効期限）は後で実装
  // Apps Scriptの日付比較の制限により、シンプルな実装に変更
  
  Logger.log('✓ DA履歴シートの書式設定完了');
}

/**
 * 両方のシートを一括作成（便利関数）
 */
function createCompetitorAnalysisSheets() {
  Logger.log('=== 競合分析シート作成開始 ===');
  Logger.log('');
  
  Logger.log('1. 競合分析シート作成中...');
  createCompetitorAnalysisSheet();
  Logger.log('');
  
  Logger.log('2. DA履歴シート作成中...');
  createDAHistorySheet();
  Logger.log('');
  
  Logger.log('=== 完了 ===');
  Logger.log('✓ 競合分析シート（27列）作成完了');
  Logger.log('✓ DA履歴シート（5列）作成完了');
  Logger.log('');
  Logger.log('次のステップ: DataForSEO検索結果取得関数を実装');
}

/**
 * シート作成確認（テスト用）
 */
function checkCompetitorSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const competitorSheet = ss.getSheetByName('競合分析');
  const daHistorySheet = ss.getSheetByName('DA履歴');
  
  Logger.log('=== シート確認 ===');
  Logger.log(`競合分析シート: ${competitorSheet ? '✓ 存在' : '✗ 未作成'}`);
  Logger.log(`DA履歴シート: ${daHistorySheet ? '✓ 存在' : '✗ 未作成'}`);
  
  if (competitorSheet) {
    Logger.log(`競合分析シート列数: ${competitorSheet.getLastColumn()}`);
    Logger.log(`期待値: 27列`);
  }
  
  if (daHistorySheet) {
    Logger.log(`DA履歴シート列数: ${daHistorySheet.getLastColumn()}`);
    Logger.log(`期待値: 5列`);
  }
}