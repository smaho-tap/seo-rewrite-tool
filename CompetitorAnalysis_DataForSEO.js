/**
 * ============================================================================
 * CompetitorAnalysis_DataForSEO.gs v2.0
 * ============================================================================
 * DataForSEO APIを使用して、キーワードの検索結果（上位10サイト）を取得
 * 
 * v2.0更新内容（Day 16）:
 * - AI Overview（AIO）データ取得対応
 * - expand_ai_overview: true 追加（追加コストなし）
 * - AIOデータを含むレスポンス構造に変更
 */

/**
 * DataForSEO APIで検索結果を取得（AIO対応版）
 * 
 * @param {string} keyword - 検索キーワード
 * @return {Object} 検索結果オブジェクト（AIOデータ含む）
 */
function fetchSearchResults(keyword) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const login = scriptProperties.getProperty('DATAFORSEO_LOGIN');
  const password = scriptProperties.getProperty('DATAFORSEO_PASSWORD');
  
  if (!login || !password) {
    throw new Error('DataForSEO認証情報が設定されていません');
  }
  
  const url = 'https://api.dataforseo.com/v3/serp/google/organic/live/advanced';
  
  const requestBody = [{
    "keyword": keyword,
    "location_code": 2392,  // 日本
    "language_code": "ja",
    "device": "desktop",
    "depth": 10,            // 上位10件
    "expand_ai_overview": true  // ★ AIOデータ取得（追加コストなし）
  }];
  
  const options = {
    method: 'POST',
    headers: {
      'Authorization': 'Basic ' + Utilities.base64Encode(login + ':' + password),
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(requestBody),
    muteHttpExceptions: true
  };
  
  // リトライロジック
  const maxRetries = 3;
  let lastError = null;
  
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      Logger.log(`[${keyword}] DataForSEO API呼び出し中... (試行 ${attempt}/${maxRetries})`);
      
      const response = UrlFetchApp.fetch(url, options);
      const responseCode = response.getResponseCode();
      
      if (responseCode === 200) {
        const data = JSON.parse(response.getContentText());
        
        // エラーチェック
        if (data.status_code === 20000) {
          const items = data.tasks[0].result[0].items || [];
          
          // AIOデータを抽出
          const aioData = extractAIOFromItems(items, keyword);
          
          // オーガニック結果のみを抽出（AI Overview、広告等を除外）
          const organicResults = items.filter(item => {
            return item.type === 'organic';
          });
          
          Logger.log(`✓ [${keyword}] 全結果: ${items.length}件、オーガニック: ${organicResults.length}件、AIO: ${aioData.hasAIO ? 'あり' : 'なし'}`);
          
          return {
            keyword: keyword,
            results: organicResults.slice(0, 10).map((item, index) => ({
              rank: index + 1,
              url: item.url || '',
              domain: extractDomain(item.url || ''),
              title: item.title || '',
              description: item.description || ''
            })),
            aio: aioData,  // ★ AIOデータを追加
            cost: 0.006,
            timestamp: new Date()
          };
        } else {
          throw new Error(`DataForSEO APIエラー: ${data.status_message}`);
        }
      } else if (responseCode === 429) {
        // レート制限
        const waitTime = Math.pow(2, attempt) * 1000;
        Logger.log(`⚠ [${keyword}] レート制限（429）、${waitTime/1000}秒待機...`);
        Utilities.sleep(waitTime);
        continue;
      } else {
        throw new Error(`HTTP ${responseCode}: ${response.getContentText()}`);
      }
    } catch (error) {
      lastError = error;
      Logger.log(`✗ [${keyword}] 試行 ${attempt}/${maxRetries} 失敗: ${error.message}`);
      
      if (attempt < maxRetries) {
        const waitTime = Math.pow(2, attempt) * 1000;
        Logger.log(`${waitTime/1000}秒待機後に再試行...`);
        Utilities.sleep(waitTime);
      }
    }
  }
  
  throw new Error(`[${keyword}] DataForSEO API取得失敗（${maxRetries}回試行）: ${lastError.message}`);
}

/**
 * APIレスポンスからAIOデータを抽出
 * 
 * @param {Array} items - APIレスポンスのitems配列
 * @param {string} keyword - キーワード（ログ用）
 * @return {Object} AIOデータオブジェクト
 */
function extractAIOFromItems(items, keyword) {
  // AIOアイテムを検索
  const aioItem = items.find(item => item.type === 'ai_overview');
  
  if (!aioItem) {
    return {
      hasAIO: false,
      references: [],
      totalReferences: 0,
      aioPosition: null,
      isAsynchronous: false
    };
  }
  
  // 全ての引用元を収集
  const allReferences = [];
  
  // items配列内のreferencesを収集（本文内の引用）
  if (aioItem.items && Array.isArray(aioItem.items)) {
    aioItem.items.forEach((element, elementIndex) => {
      if (element.references && Array.isArray(element.references)) {
        element.references.forEach((ref, refIndex) => {
          allReferences.push({
            position: allReferences.length + 1,
            source: ref.source || '',
            domain: ref.domain || '',
            url: ref.url || '',
            title: ref.title || '',
            text: ref.text || '',
            location: 'inline',  // 本文内
            elementIndex: elementIndex
          });
        });
      }
    });
  }
  
  // AIO下部のreferences（「もっと見る」セクション）
  if (aioItem.references && Array.isArray(aioItem.references)) {
    aioItem.references.forEach((ref, refIndex) => {
      // 重複チェック（URLベース）
      const isDuplicate = allReferences.some(r => r.url === ref.url);
      if (!isDuplicate) {
        allReferences.push({
          position: allReferences.length + 1,
          source: ref.source || '',
          domain: ref.domain || '',
          url: ref.url || '',
          title: ref.title || '',
          text: ref.text || '',
          location: 'footer'  // 下部の「もっと見る」
        });
      }
    });
  }
  
  Logger.log(`✓ [${keyword}] AIO検出: 引用元 ${allReferences.length}件`);
  
  return {
    hasAIO: true,
    references: allReferences,
    totalReferences: allReferences.length,
    aioPosition: aioItem.rank_absolute || 1,
    isAsynchronous: aioItem.asynchronous_ai_overview || false
  };
}

/**
 * URLからドメインを抽出
 * 
 * @param {string} url - URL
 * @return {string} ドメイン
 */
function extractDomain(url) {
  try {
    if (!url) return '';
    const match = url.match(/^https?:\/\/([^\/]+)/);
    return match ? match[1] : '';
  } catch (error) {
    return '';
  }
}

/**
 * 複数キーワードの検索結果を一括取得（AIO対応版）
 * 
 * @param {Array} keywords - キーワード配列
 * @return {Array} 検索結果配列（AIOデータ含む）
 */
function fetchMultipleSearchResults(keywords) {
  const results = [];
  let successCount = 0;
  let failCount = 0;
  let totalCost = 0;
  let aioCount = 0;
  
  Logger.log(`=== ${keywords.length} キーワードの検索結果取得開始（AIO対応） ===`);
  
  for (let i = 0; i < keywords.length; i++) {
    const keyword = keywords[i];
    
    try {
      const result = fetchSearchResults(keyword);
      results.push(result);
      successCount++;
      totalCost += result.cost;
      
      if (result.aio && result.aio.hasAIO) {
        aioCount++;
      }
      
      Logger.log(`[${i + 1}/${keywords.length}] ✓ ${keyword} - 成功 (AIO: ${result.aio.hasAIO ? 'あり' : 'なし'})`);
      
      // レート制限対策（1秒待機）
      if (i < keywords.length - 1) {
        Utilities.sleep(1000);
      }
    } catch (error) {
      Logger.log(`[${i + 1}/${keywords.length}] ✗ ${keyword} - 失敗: ${error.message}`);
      failCount++;
      
      results.push({
        keyword: keyword,
        results: [],
        aio: { hasAIO: false, references: [], totalReferences: 0 },
        error: error.message,
        timestamp: new Date()
      });
    }
  }
  
  Logger.log('');
  Logger.log('=== 取得結果サマリー ===');
  Logger.log(`成功: ${successCount} / ${keywords.length}`);
  Logger.log(`失敗: ${failCount} / ${keywords.length}`);
  Logger.log(`AIOあり: ${aioCount} / ${successCount}`);
  Logger.log(`推定コスト: $${totalCost.toFixed(3)}`);
  
  return results;
}

/**
 * テスト実行: AIO対応版
 */
function testDataForSEOFetchWithAIO() {
  Logger.log('=== DataForSEO検索結果取得テスト（AIO対応版） ===');
  Logger.log('');
  
  // テスト用キーワード
  const testKeywords = [
    'iphone 保険',
    'iphone 安く買う方法'
  ];
  
  Logger.log(`テストキーワード: ${testKeywords.join(', ')}`);
  Logger.log(`推定コスト: $${(testKeywords.length * 0.006).toFixed(3)}`);
  Logger.log('');
  
  const results = fetchMultipleSearchResults(testKeywords);
  
  // 結果をログ出力
  Logger.log('');
  Logger.log('=== 取得結果詳細 ===');
  
  results.forEach(result => {
    if (result.error) {
      Logger.log(`✗ ${result.keyword}: エラー - ${result.error}`);
    } else {
      Logger.log(`✓ ${result.keyword}:`);
      Logger.log(`  オーガニック結果: ${result.results.length}件`);
      Logger.log(`  AIO: ${result.aio.hasAIO ? 'あり' : 'なし'}`);
      
      if (result.aio.hasAIO) {
        Logger.log(`  AIO引用元: ${result.aio.totalReferences}件`);
        result.aio.references.slice(0, 3).forEach((ref, i) => {
          Logger.log(`    ${i + 1}. ${ref.domain} - ${ref.title.substring(0, 40)}...`);
        });
      }
      
      Logger.log(`  上位3サイト:`);
      result.results.slice(0, 3).forEach(item => {
        Logger.log(`    ${item.rank}位: ${item.domain}`);
      });
    }
  });
  
  return results;
}

// ============================================================================
// 以下、既存の関数（変更なし）
// ============================================================================

/**
 * ターゲットKW分析シートから主要キーワードを取得
 */
function getTopKeywords(limit = 5) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('ターゲットKW分析');
  
  if (!sheet) {
    throw new Error('ターゲットKW分析シートが見つかりません');
  }
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const keywordColIndex = headers.indexOf('ターゲットキーワード');
  
  if (keywordColIndex === -1) {
    throw new Error('ターゲットキーワード列が見つかりません');
  }
  
  const keywords = [];
  for (let i = 1; i < Math.min(data.length, limit + 1); i++) {
    const keyword = data[i][keywordColIndex];
    if (keyword) {
      keywords.push(keyword);
    }
  }
  
  Logger.log(`✓ ターゲットKW分析シートから ${keywords.length} 件のキーワードを取得`);
  
  return keywords;
}

/**
 * 検索結果を競合分析シートに書き込み（1キーワード）
 */
function writeSearchResultToSheet(searchResult, pageUrl = '') {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('競合分析');
  
  if (!sheet) {
    throw new Error('競合分析シートが見つかりません');
  }
  
  const lastRow = sheet.getLastRow();
  const newRow = lastRow + 1;
  
  const analysisId = `${searchResult.keyword}_${Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd')}`;
  
  const rowData = [
    analysisId,
    searchResult.keyword,
    pageUrl,
    new Date(),
    '',
    ''
  ];
  
  for (let i = 0; i < 10; i++) {
    if (i < searchResult.results.length) {
      const result = searchResult.results[i];
      rowData.push(result.url);
      rowData.push('');
    } else {
      rowData.push('');
      rowData.push('');
    }
  }
  
  rowData.push('');  // avg_da_top10
  rowData.push('');  // weaker_sites_count
  rowData.push('');  // weaker_sites_highest_rank
  rowData.push('');  // da_distribution_type
  rowData.push('');  // winnable_score
  rowData.push('');  // target_rank
  rowData.push('');  // competition_level
  rowData.push('');  // rewrite_roi
  rowData.push(new Date());  // last_updated
  
  sheet.getRange(newRow, 1, 1, rowData.length).setValues([rowData]);
  
  Logger.log(`✓ [${searchResult.keyword}] 競合分析シートに書き込み完了（行 ${newRow}）`);
}

/**
 * 複数検索結果を競合分析シートに一括書き込み
 */
function writeMultipleSearchResultsToSheet(searchResults) {
  Logger.log(`=== ${searchResults.length} 件の検索結果を競合分析シートに書き込み ===`);
  
  const pageUrls = getPageUrlsForKeywords(searchResults.map(r => r.keyword));
  
  let successCount = 0;
  let failCount = 0;
  
  searchResults.forEach((result, index) => {
    try {
      if (!result.error) {
        const pageUrl = pageUrls[result.keyword] || '';
        writeSearchResultToSheet(result, pageUrl);
        successCount++;
      } else {
        Logger.log(`✗ [${result.keyword}] エラーのためスキップ: ${result.error}`);
        failCount++;
      }
    } catch (error) {
      Logger.log(`✗ [${result.keyword}] 書き込みエラー: ${error.message}`);
      failCount++;
    }
  });
  
  Logger.log('');
  Logger.log('=== 書き込み結果 ===');
  Logger.log(`成功: ${successCount} / ${searchResults.length}`);
  Logger.log(`失敗: ${failCount} / ${searchResults.length}`);
}

/**
 * ターゲットKW分析シートからキーワードに対応するページURLを取得
 */
function getPageUrlsForKeywords(keywords) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('ターゲットKW分析');
  
  if (!sheet) {
    Logger.log('⚠ ターゲットKW分析シートが見つかりません');
    return {};
  }
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const keywordColIndex = headers.indexOf('ターゲットキーワード');
  const urlColIndex = headers.indexOf('ページURL');
  
  if (keywordColIndex === -1 || urlColIndex === -1) {
    Logger.log('⚠ 必要な列が見つかりません');
    return {};
  }
  
  const pageUrls = {};
  
  for (let i = 1; i < data.length; i++) {
    const keyword = data[i][keywordColIndex];
    const url = data[i][urlColIndex];
    
    if (keyword && keywords.includes(keyword)) {
      pageUrls[keyword] = url;
    }
  }
  
  Logger.log(`✓ ${Object.keys(pageUrls).length} 件のページURLを取得`);
  
  return pageUrls;
}

/**
 * 検索結果を競合分析シートの指定行に書き込み
 */
function writeSearchResultToSheetRow(sheet, row, results) {
  if (!sheet) {
    throw new Error('シートが指定されていません');
  }
  
  if (!results || !Array.isArray(results)) {
    throw new Error('検索結果が配列ではありません');
  }
  
  Logger.log(`検索結果を ${row} 行目に書き込み中... (${results.length} 件)`);
  
  const startCol = 7;
  
  for (let i = 0; i < 10; i++) {
    const urlCol = startCol + (i * 2);
    const daCol = urlCol + 1;
    
    if (i < results.length && results[i]) {
      const result = results[i];
      sheet.getRange(row, urlCol).setValue(result.url || '');
      sheet.getRange(row, daCol).setValue('');
    } else {
      sheet.getRange(row, urlCol).setValue('');
      sheet.getRange(row, daCol).setValue('');
    }
  }
  
  Logger.log(`✓ ${row} 行目に検索結果を書き込み完了`);
}