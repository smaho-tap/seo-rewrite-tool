/**
 * より詳細なデバッグ - セルの型と値を確認
 */
function debugCompetitorDataDetailed() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('競合分析');
  
  Logger.log('=== シート詳細確認 ===');
  Logger.log('シート名: ' + sheet.getName());
  Logger.log('最終行: ' + sheet.getLastRow());
  Logger.log('最終列: ' + sheet.getLastColumn());
  
  // ヘッダー行を確認
  Logger.log('\n=== ヘッダー行（1行目）確認 ===');
  const headerB = sheet.getRange(1, 2).getValue(); // B1
  const headerE = sheet.getRange(1, 5).getValue(); // E1
  const headerH = sheet.getRange(1, 8).getValue(); // H1
  
  Logger.log('B1 (target_keyword): ' + headerB);
  Logger.log('E1 (own_site_da): ' + headerE);
  Logger.log('H1 (rank_1_da): ' + headerH);
  
  // データ行を確認（2行目と3行目）
  Logger.log('\n=== 2行目のデータ ===');
  Logger.log('B2: ' + sheet.getRange(2, 2).getValue() + ' (型: ' + typeof sheet.getRange(2, 2).getValue() + ')');
  Logger.log('E2: ' + sheet.getRange(2, 5).getValue() + ' (型: ' + typeof sheet.getRange(2, 5).getValue() + ')');
  Logger.log('H2: ' + sheet.getRange(2, 8).getValue() + ' (型: ' + typeof sheet.getRange(2, 8).getValue() + ')');
  Logger.log('J2: ' + sheet.getRange(2, 10).getValue() + ' (型: ' + typeof sheet.getRange(2, 10).getValue() + ')');
  
  Logger.log('\n=== 3行目のデータ ===');
  Logger.log('B3: ' + sheet.getRange(3, 2).getValue() + ' (型: ' + typeof sheet.getRange(3, 2).getValue() + ')');
  Logger.log('E3: ' + sheet.getRange(3, 5).getValue() + ' (型: ' + typeof sheet.getRange(3, 5).getValue() + ')');
  Logger.log('H3: ' + sheet.getRange(3, 8).getValue() + ' (型: ' + typeof sheet.getRange(3, 8).getValue() + ')');
  Logger.log('J3: ' + sheet.getRange(3, 10).getValue() + ' (型: ' + typeof sheet.getRange(3, 10).getValue() + ')');
  
  // 実際にデータがある行を探す
  Logger.log('\n=== データ行探索 ===');
  for (let row = 1; row <= sheet.getLastRow(); row++) {
    const keyword = sheet.getRange(row, 2).getValue();
    if (keyword && keyword.toString().includes('amazon')) {
      Logger.log('✓ データ発見: ' + row + '行目');
      Logger.log('  キーワード: ' + keyword);
      Logger.log('  自社DA (E列): ' + sheet.getRange(row, 5).getValue());
      Logger.log('  rank_1_da (H列): ' + sheet.getRange(row, 8).getValue());
      Logger.log('  rank_2_da (J列): ' + sheet.getRange(row, 10).getValue());
    }
  }
}

/**
 * fetchSearchResults()の返り値を確認
 */
function debugFetchSearchResults() {
  const keyword = 'amazon で iphone を 買う メリット';
  
  Logger.log('=== fetchSearchResults() デバッグ ===');
  Logger.log('キーワード: ' + keyword);
  
  const results = fetchSearchResults(keyword);
  
  Logger.log('\n返り値の型: ' + typeof results);
  Logger.log('返り値がnull?: ' + (results === null));
  Logger.log('返り値がundefined?: ' + (results === undefined));
  
  if (results) {
    Logger.log('配列?: ' + Array.isArray(results));
    Logger.log('length: ' + (results.length || 'lengthプロパティなし'));
    
    if (results.length > 0) {
      Logger.log('\n最初の要素:');
      Logger.log(JSON.stringify(results[0], null, 2));
    }
  } else {
    Logger.log('⚠ resultsがnullまたはundefinedです');
  }
}

/**
 * fetchSearchResults()の返り値の構造を詳しく確認
 */
function debugSearchResultsStructure() {
  const keyword = 'amazon で iphone を 買う メリット';
  
  Logger.log('=== fetchSearchResults() 構造確認 ===');
  
  const results = fetchSearchResults(keyword);
  
  Logger.log('\n返り値の型: ' + typeof results);
  Logger.log('配列?: ' + Array.isArray(results));
  
  if (results && typeof results === 'object') {
    Logger.log('\nオブジェクトのキー:');
    const keys = Object.keys(results);
    Logger.log(keys);
    
    // 各キーの値の型を確認
    for (let key of keys) {
      Logger.log('\n' + key + ':');
      Logger.log('  型: ' + typeof results[key]);
      
      if (Array.isArray(results[key])) {
        Logger.log('  配列の長さ: ' + results[key].length);
        
        if (results[key].length > 0) {
          Logger.log('  最初の要素:');
          Logger.log('  ' + JSON.stringify(results[key][0], null, 2).substring(0, 500));
        }
      }
    }
  }
  
  Logger.log('\n=== 全体構造（最初の1000文字） ===');
  Logger.log(JSON.stringify(results, null, 2).substring(0, 1000));
}

/**
 * extractDomain()のテスト
 */
function testExtractDomain() {
  const testUrl = '/purchase-amazon-iphone';
  
  Logger.log('=== extractDomain() テスト ===');
  Logger.log('入力URL: ' + testUrl);
  
  const domain = extractDomain(testUrl);
  
  Logger.log('抽出されたドメイン: "' + domain + '"');
  Logger.log('長さ: ' + (domain ? domain.length : 0));
  Logger.log('型: ' + typeof domain);
  
  // 正しい例
  const testUrl2 = 'https://smaho-tap.com/purchase-amazon-iphone';
  const domain2 = extractDomain(testUrl2);
  
  Logger.log('\n正しい例:');
  Logger.log('入力URL: ' + testUrl2);
  Logger.log('抽出されたドメイン: "' + domain2 + '"');
}

/**
 * getDAFromCache()の返り値を確認
 */
function testGetDAFromCache() {
  const domain = 'smaho-tap.com';
  
  Logger.log('=== getDAFromCache() テスト ===');
  Logger.log('ドメイン: ' + domain);
  
  const cachedData = getDAFromCache(domain);
  
  Logger.log('\n返り値:');
  Logger.log('型: ' + typeof cachedData);
  Logger.log('値: ' + JSON.stringify(cachedData, null, 2));
  
  if (cachedData && typeof cachedData === 'object') {
    Logger.log('\nオブジェクトのキー:');
    Logger.log(Object.keys(cachedData));
    
    if (cachedData.da !== undefined) {
      Logger.log('\nDA値: ' + cachedData.da);
    }
  }
}