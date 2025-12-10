/**
 * CompetitorAnalysis_Auto.gs v3.0 完全版
 * 競合分析バッチ処理・DA自動再取得・AIO順位追跡統合
 * 
 * 【更新履歴】
 * - v1.0: バッチ処理実装
 * - v2.0: トリガー分割実装（タイムアウト対策）
 * - v3.0: AIO順位追跡を統合★NEW
 * 
 * 【トリガー設定】
 * 1. runBatch1 (月曜 5:00) - KW 1-50 + AIO記録
 * 2. runBatch2 (月曜 5:10) - KW 51-100 + AIO記録
 * 3. runBatch3 (月曜 5:15) - KW 101-150 + AIO記録
 * 4. runBatch4 (月曜 5:20) - KW 151-200 + AIO記録
 * 5. runBatch5 (月曜 5:25) - KW 201-213 + AIO記録
 * 6. weeklyDARetry (月曜 5:35) - DA未取得再取得 + AIOサマリー
 * 
 * @version 3.0
 * @lastUpdated 2025-12-02
 */

// ============================================================
// 設定
// ============================================================

/**
 * AIO順位追跡を有効化するフラグ
 * false にするとAIO処理をスキップ（コスト削減）
 */
const ENABLE_AIO_TRACKING = true;

// ============================================================
// テスト関数
// ============================================================

function testAutoCompetitorAnalysis() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetKWSheet = ss.getSheetByName('ターゲットKW分析');
  const competitorSheet = ss.getSheetByName('競合分析');
  
  Logger.log('=== 自動競合分析テスト（1件） ===');
  Logger.log('');
  
  Logger.log('【ステップ1】既存データクリア');
  const lastRow = competitorSheet.getLastRow();
  if (lastRow > 1) {
    competitorSheet.deleteRows(2, lastRow - 1);
    Logger.log('✓ ' + (lastRow - 1) + ' 行削除');
  }
  Logger.log('');
  
  Logger.log('【ステップ2】ターゲットKW分析から1件取得');
  const targetData = targetKWSheet.getRange(2, 1, 1, 3).getValues();
  
  const keywords = [];
  const pageUrls = [];
  
  targetData.forEach(function(row, index) {
    const keyword = row[2];
    const pageUrl = row[1];
    
    if (keyword && pageUrl) {
      keywords.push(keyword);
      pageUrls.push(pageUrl);
      Logger.log('✓ [' + (index + 1) + '] ' + keyword);
      Logger.log('  URL: ' + pageUrl);
    }
  });
  
  Logger.log('');
  Logger.log('取得: ' + keywords.length + ' 件');
  Logger.log('');
  
  Logger.log('【ステップ3】競合分析シートに書き込み');
  const now = new Date();
  const rows = keywords.map(function(keyword, index) {
    return [
      keyword + '_' + Utilities.formatDate(now, 'Asia/Tokyo', 'yyyyMMdd'),
      keyword,
      pageUrls[index],
      now
    ];
  });
  
  competitorSheet.getRange(2, 1, rows.length, 4).setValues(rows);
  Logger.log('✓ ' + rows.length + ' 行書き込み完了');
  Logger.log('');
  
  Logger.log('【ステップ4】DataForSEO検索結果取得');
  keywords.forEach(function(keyword, index) {
    const rowIndex = index + 2;
    Logger.log('[' + keyword + '] 検索中...');
    
    try {
      const result = fetchSearchResults(keyword);
      
      if (result && Array.isArray(result) && result.length > 0) {
        writeSearchResultToSheet({ keyword: keyword, organic_results: result }, rowIndex);
        Logger.log('✓ 成功: ' + result.length + ' 件');
      } else if (result && typeof result === 'object') {
        writeSearchResultToSheet(result, rowIndex);
        Logger.log('✓ 成功');
      }
      
      Utilities.sleep(1000);
    } catch (error) {
      Logger.log('✗ エラー: ' + error.message);
    }
  });
  
  Logger.log('');
  
  Logger.log('【ステップ5】自社DA取得');
  const ownDomain = 'smaho-tap.com';
  const absoluteUrls = pageUrls.map(function(url) {
    if (url.indexOf('http') !== 0) {
      return 'https://' + ownDomain + url;
    }
    return url;
  });
  
  Logger.log('絶対URL: ' + absoluteUrls.join(', '));
  
  const daMap = fetchDAWithSmartCaching(absoluteUrls);
  
  Logger.log('DAマップ件数: ' + Object.keys(daMap).length);
  
  absoluteUrls.forEach(function(url, index) {
    const rowIndex = index + 2;
    const domain = extractDomain(url);
    
    Logger.log('[デバッグ] URL: ' + url + ', Domain: ' + domain);
    
    var daData = daMap[domain];
    if (!daData && domain.endsWith('/')) {
      daData = daMap[domain.slice(0, -1)];
    } else if (!daData) {
      daData = daMap[domain + '/'];
    }
    
    if (daData) {
      competitorSheet.getRange(rowIndex, 5).setValue(daData.da);
      Logger.log('✓ [行' + rowIndex + '] DA: ' + daData.da + ' を書き込み');
    } else {
      Logger.log('✗ [行' + rowIndex + '] DAなし');
    }
  });
  
  Logger.log('');
  
  Logger.log('【ステップ6】競合DA取得');
  
  Utilities.sleep(5000);
  
  const dataForDA = competitorSheet.getRange(2, 1, 2, 27).getValues();
  
  Logger.log('[デバッグ] 取得した行数: ' + dataForDA.length);
  Logger.log('[デバッグ] 行2のG列: ' + dataForDA[0][6]);
  Logger.log('[デバッグ] 行3のG列: ' + dataForDA[1][6]);
  
  const competitorUrls = [];
  
  dataForDA.forEach(function(row, rowIdx) {
    for (var i = 6; i <= 24; i += 2) {
      const url = row[i];
      if (url && typeof url === 'string' && url.length > 0) {
        competitorUrls.push(url);
        Logger.log('[デバッグ] 行' + (rowIdx + 2) + ' 競合URL追加: ' + url);
      }
    }
  });
  
  Logger.log('競合URL数: ' + competitorUrls.length);
  
  if (competitorUrls.length > 0) {
    const competitorDAMap = fetchDAWithSmartCaching(competitorUrls);
    
    Logger.log('競合DAマップ件数: ' + Object.keys(competitorDAMap).length);
    
    dataForDA.forEach(function(row, index) {
      const rowIndex = index + 2;
      
      for (var i = 0; i < 10; i++) {
        const urlColIndex = 6 + (i * 2);
        const daColNumber = 8 + (i * 2);
        const url = row[urlColIndex];
        
        if (url && typeof url === 'string' && url.length > 0) {
          const domain = extractDomain(url);
          
          var daData = competitorDAMap[domain];
          if (!daData && domain.endsWith('/')) {
            daData = competitorDAMap[domain.slice(0, -1)];
          } else if (!daData) {
            daData = competitorDAMap[domain + '/'];
          }
          
          if (daData) {
            competitorSheet.getRange(rowIndex, daColNumber).setValue(daData.da);
            Logger.log('✓ [行' + rowIndex + ', 列' + daColNumber + '] ' + domain + ' → DA: ' + daData.da);
          } else {
            Logger.log('✗ [行' + rowIndex + ', 列' + daColNumber + '] ' + domain + ' → DAなし');
          }
        }
      }
    });
    
    Logger.log('✓ 競合DA書き込み完了');
  } else {
    Logger.log('⚠ 競合URLが見つかりませんでした');
  }
  
  Logger.log('');
  
  Logger.log('【ステップ7】勝算度スコア算出');
  updateWinnableScores(2, 3);
  
  Logger.log('');
  Logger.log('=== テスト完了 ===');
  Logger.log('競合分析シートを確認してください');
}

/**
 * 統合テスト（修正版v2）
 * 1キーワードの完全な競合分析フローをテスト
 */
function testAutoCompetitorAnalysisFixed() {
  Logger.log('=== 競合分析統合テスト開始（修正版v2） ===\n');
  
  try {
    // ステップ1: ターゲットKW分析シートから1件取得
    Logger.log('【ステップ1】ターゲットKW取得');
    const targetKWSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ターゲットKW分析');
    if (!targetKWSheet) {
      throw new Error('ターゲットKW分析シートが見つかりません');
    }
    
    const keyword = targetKWSheet.getRange(2, 3).getValue(); // C2: target_keyword
    const pageUrl = targetKWSheet.getRange(2, 2).getValue(); // B2: page_url
    
    if (!keyword || !pageUrl) {
      throw new Error('ターゲットKWまたはページURLが取得できません');
    }
    
    Logger.log('✓ キーワード: ' + keyword);
    Logger.log('✓ ページURL: ' + pageUrl);
    
    // ステップ2: 競合分析シートを準備
    Logger.log('\n【ステップ2】競合分析シート準備');
    const competitorSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('競合分析');
    if (!competitorSheet) {
      throw new Error('競合分析シートが見つかりません');
    }
    
    Logger.log('✓ 現在の最終行: ' + competitorSheet.getLastRow());
    
    // ステップ3: 分析IDと基本情報を書き込み
    Logger.log('\n【ステップ3】基本情報書き込み');
    const row = 2; // データ行
    const analysisId = 'CA_' + new Date().getTime();
    const analysisDate = Utilities.formatDate(new Date(), 'JST', 'yyyy-MM-dd');
    
    competitorSheet.getRange(row, 1).setValue(analysisId);        // A: analysis_id
    competitorSheet.getRange(row, 2).setValue(keyword);           // B: target_keyword
    competitorSheet.getRange(row, 3).setValue(pageUrl);           // C: page_url
    competitorSheet.getRange(row, 4).setValue(analysisDate);      // D: analysis_date
    
    Logger.log('✓ 基本情報を書き込みました');
    
    // ステップ4: DataForSEO検索結果取得
    Logger.log('\n【ステップ4】DataForSEO検索結果取得');
    const searchData = fetchSearchResults(keyword);    // オブジェクトを取得
    const searchResults = searchData.results;          // resultsプロパティから配列を取り出す
    Logger.log('✓ [' + keyword + '] 検索結果 ' + searchResults.length + ' 件取得成功');
    
    // 検索結果を競合分析シートに書き込み（修正版関数を使用）
    writeSearchResultToSheetRow(competitorSheet, row, searchResults);
    Logger.log('✓ 検索結果を書き込みました');
    
   // ステップ5: 自社DA取得
    Logger.log('\n【ステップ5】自社DA取得');
    
    // 自社ドメインを直接指定
    const ownDomain = 'smaho-tap.com';
    
    // キャッシュから取得を試みる
    let ownDA;
    const cachedData = getDAFromCache(ownDomain);
    
    if (cachedData === null) {
      // キャッシュになければMoz APIで取得
      Logger.log(`[${ownDomain}] Moz APIでDA取得中...`);
      const daResults = fetchDomainAuthority([ownDomain]);
      ownDA = daResults[0] || 0;
      
      // キャッシュに保存
      if (ownDA > 0) {
        saveDAToCache(ownDomain, ownDA);
      }
    } else {
      // キャッシュから取得（オブジェクトからDA値を取り出す）
      ownDA = cachedData.da;
      Logger.log(`✓ [${ownDomain}] キャッシュヒット（DA: ${ownDA}）`);
    }
    
    competitorSheet.getRange(row, 5).setValue(ownDA); // E列: own_site_da
    Logger.log('✓ [行' + row + '] DA: ' + ownDA + ' を書き込み');
    
    // ステップ6: 競合DA取得
    Logger.log('\n【ステップ6】競合DA取得');
    updateCompetitorDA(row, row);  // startRow, endRowを両方指定（2行目のみ）
    Logger.log('✓ 競合DA書き込み完了');
    
    // ステップ7: 勝算度スコア算出
    Logger.log('\n【ステップ7】勝算度スコア算出');
    updateWinnableScores(row, row);  // startRow, endRowを両方指定（2行目のみ）
    
    // 結果確認
    const winnableScore = competitorSheet.getRange(row, 34).getValue(); // AH列（winnable_score）
    const competitorLevel = competitorSheet.getRange(row, 35).getValue(); // AI列（competitor_level）
    
    Logger.log('✓ [行' + row + '] ' + keyword + ': 勝算度' + winnableScore + '点（' + competitorLevel + '）');
    
    Logger.log('\n=== 統合テスト完了 ===');
    Logger.log('✓ 全ステップが正常に完了しました');
    
  } catch (error) {
    Logger.log('❌ エラー発生: ' + error.message);
    Logger.log(error.stack);
  }
}

// ============================================================
// 全213KW競合分析実行（手動用）
// ============================================================

/**
 * 全213KWの競合分析を実行（フェーズ6）
 * 
 * 実行時間: 約15-20分
 * コスト: DataForSEO $1.28 + Moz API 約200-300 rows
 */
function runFullCompetitorAnalysis() {
  Logger.log('=== 全213KW競合分析開始 ===');
  Logger.log('推定時間: 15-20分');
  Logger.log('推定コスト: DataForSEO $1.28 + Moz API 200-300 rows');
  Logger.log('AIO追跡: ' + (ENABLE_AIO_TRACKING ? '有効' : '無効'));
  Logger.log('');
  
  const startTime = new Date();
  
  try {
    // ステップ1: ターゲットKW分析から全件取得
    Logger.log('【ステップ1】ターゲットKW分析から全件取得');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const targetKWSheet = ss.getSheetByName('ターゲットKW分析');
    
    if (!targetKWSheet) {
      throw new Error('ターゲットKW分析シートが見つかりません');
    }
    
    const data = targetKWSheet.getDataRange().getValues();
    const headers = data[0];
    const keywordColIndex = headers.indexOf('target_keyword');
    const pageUrlColIndex = headers.indexOf('page_url');
    
    if (keywordColIndex === -1 || pageUrlColIndex === -1) {
      throw new Error('必要な列が見つかりません');
    }
    
    const keywords = [];
    const pageUrls = [];
    
    for (let i = 1; i < data.length; i++) {
      const keyword = data[i][keywordColIndex];
      const pageUrl = data[i][pageUrlColIndex];
      
      if (keyword && pageUrl) {
        keywords.push(keyword);
        pageUrls.push(pageUrl);
      }
    }
    
    Logger.log(`✓ ${keywords.length} 件のキーワードを取得`);
    Logger.log('');
    
    // ステップ2: 競合分析シートをクリア
    Logger.log('【ステップ2】競合分析シートをクリア');
    const competitorSheet = ss.getSheetByName('競合分析');
    
    if (!competitorSheet) {
      throw new Error('競合分析シートが見つかりません');
    }
    
    const lastRow = competitorSheet.getLastRow();
    if (lastRow > 1) {
      competitorSheet.deleteRows(2, lastRow - 1);
      Logger.log(`✓ ${lastRow - 1} 行削除`);
    }
    Logger.log('');
    
    // ステップ3: 基本情報を一括書き込み
    Logger.log('【ステップ3】基本情報を一括書き込み');
    const now = new Date();
    const analysisDate = Utilities.formatDate(now, 'JST', 'yyyy-MM-dd');
    
    const basicInfoRows = keywords.map((keyword, index) => {
      return [
        `CA_${now.getTime()}_${index}`,  // A: analysis_id
        keyword,                          // B: target_keyword
        pageUrls[index],                  // C: page_url
        analysisDate                      // D: analysis_date
      ];
    });
    
    competitorSheet.getRange(2, 1, basicInfoRows.length, 4).setValues(basicInfoRows);
    Logger.log(`✓ ${basicInfoRows.length} 行の基本情報を書き込み`);
    Logger.log('');
    
    // ステップ4: DataForSEO検索結果を一括取得
    Logger.log('【ステップ4】DataForSEO検索結果を一括取得');
    Logger.log(`${keywords.length} キーワードを処理中...`);
    Logger.log('');
    
    let successCount = 0;
    let failCount = 0;
    const allSearchResults = []; // AIO用に結果を保存★
    
    for (let i = 0; i < keywords.length; i++) {
      const keyword = keywords[i];
      const row = i + 2;
      
      try {
        const searchData = fetchSearchResults(keyword);
        const searchResults = searchData.results;
        
        writeSearchResultToSheetRow(competitorSheet, row, searchResults);
        
        // AIO用に結果を保存★
        allSearchResults.push(searchData);
        
        successCount++;
        
        if ((i + 1) % 10 === 0) {
          Logger.log(`進捗: ${i + 1} / ${keywords.length} (${Math.round((i + 1) / keywords.length * 100)}%)`);
        }
        
        // API制限対策（1秒待機）
        Utilities.sleep(1000);
        
      } catch (error) {
        Logger.log(`✗ [${keyword}] 検索結果取得失敗: ${error.message}`);
        failCount++;
      }
    }
    
    Logger.log('');
    Logger.log('=== DataForSEO検索結果取得完了 ===');
    Logger.log(`成功: ${successCount} / ${keywords.length}`);
    Logger.log(`失敗: ${failCount} / ${keywords.length}`);
    Logger.log('');
    
    // ステップ5: 自社DA取得（キャッシュ利用）
    Logger.log('【ステップ5】自社DA取得');
    const ownDomain = 'smaho-tap.com';
    
    let ownDA;
    const cachedData = getDAFromCache(ownDomain);
    
    if (cachedData === null) {
      Logger.log(`[${ownDomain}] Moz APIでDA取得中...`);
      const daResults = fetchDomainAuthority([ownDomain]);
      ownDA = daResults[0] || 0;
      
      if (ownDA > 0) {
        saveDAToCache([{ domain: ownDomain, da: ownDA, pa: 0 }]);
      }
    } else {
      ownDA = cachedData.da;
      Logger.log(`✓ [${ownDomain}] キャッシュヒット（DA: ${ownDA}）`);
    }
    
    // 全行のE列に自社DAを書き込み
    const ownDAColumn = Array(keywords.length).fill([ownDA]);
    competitorSheet.getRange(2, 5, keywords.length, 1).setValues(ownDAColumn);
    Logger.log(`✓ ${keywords.length} 行に自社DA（${ownDA}）を書き込み`);
    Logger.log('');
    
    // ステップ6: 競合DA取得（全URL一括処理）
    Logger.log('【ステップ6】競合DA取得');
    updateCompetitorDA(2, 2 + keywords.length - 1);
    Logger.log('✓ 競合DA取得完了');
    Logger.log('');
    
    // ステップ7: 勝算度スコア算出
    Logger.log('【ステップ7】勝算度スコア算出');
    updateWinnableScores(2, 2 + keywords.length - 1);
    Logger.log('✓ 勝算度スコア算出完了');
    Logger.log('');
    
    // ステップ8: AIO順位追跡★NEW
    if (ENABLE_AIO_TRACKING && allSearchResults.length > 0) {
      Logger.log('【ステップ8】AIO順位追跡');
      processAIOInBatch(allSearchResults);
      Logger.log('');
    }
    
    const endTime = new Date();
    const duration = (endTime - startTime) / 1000 / 60; // 分
    
    Logger.log('=== 全213KW競合分析完了 ===');
    Logger.log(`処理時間: ${duration.toFixed(1)}分`);
    Logger.log(`成功: ${successCount} / ${keywords.length}`);
    Logger.log(`推定コスト: DataForSEO $${(successCount * 0.006).toFixed(2)}`);
    Logger.log('');
    Logger.log('競合分析シートを確認してください');
    
  } catch (error) {
    Logger.log('❌ エラー発生: ' + error.message);
    Logger.log(error.stack);
  }
}

// ============================================================
// バッチ処理（6分制限対策）
// ============================================================

/**
 * バッチ処理版：全213KWを50件ずつ処理（6分制限対策）
 * ★v3.0: AIO順位追跡を統合
 * 
 * @param {number} batchNumber - バッチ番号（1, 2, 3, 4, 5）
 */
function runCompetitorAnalysisBatch(batchNumber) {
  const batchSize = 50;
  const startIndex = (batchNumber - 1) * batchSize;
  const endIndex = Math.min(startIndex + batchSize - 1, 212); // 0-based index
  
  Logger.log(`=== バッチ${batchNumber}: ${startIndex + 1}〜${endIndex + 1}件目を処理 ===`);
  Logger.log('AIO追跡: ' + (ENABLE_AIO_TRACKING ? '有効' : '無効'));
  Logger.log('');
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const targetKWSheet = ss.getSheetByName('ターゲットKW分析');
    const competitorSheet = ss.getSheetByName('競合分析');
    
    // データ取得
    const data = targetKWSheet.getDataRange().getValues();
    const headers = data[0];
    const keywordColIndex = headers.indexOf('target_keyword');
    const pageUrlColIndex = headers.indexOf('page_url');
    
    const keywords = [];
    const pageUrls = [];
    
    for (let i = startIndex + 1; i <= endIndex + 1 && i < data.length; i++) {
      const keyword = data[i][keywordColIndex];
      const pageUrl = data[i][pageUrlColIndex];
      
      if (keyword && pageUrl) {
        keywords.push(keyword);
        pageUrls.push(pageUrl);
      }
    }
    
    Logger.log(`✓ ${keywords.length} 件のキーワードを取得`);
    Logger.log('');
    
    // 基本情報書き込み（初回バッチのみシートクリア）
    if (batchNumber === 1) {
      Logger.log('【初回】競合分析シートをクリア');
      const lastRow = competitorSheet.getLastRow();
      if (lastRow > 1) {
        competitorSheet.deleteRows(2, lastRow - 1);
      }
    }
    
    const now = new Date();
    const analysisDate = Utilities.formatDate(now, 'JST', 'yyyy-MM-dd');
    const startRow = batchNumber === 1 ? 2 : competitorSheet.getLastRow() + 1;
    
    const basicInfoRows = keywords.map((keyword, index) => {
      return [
        `CA_${now.getTime()}_${startIndex + index}`,
        keyword,
        pageUrls[index],
        analysisDate
      ];
    });
    
    competitorSheet.getRange(startRow, 1, basicInfoRows.length, 4).setValues(basicInfoRows);
    Logger.log(`✓ ${basicInfoRows.length} 行の基本情報を書き込み（行${startRow}〜）`);
    Logger.log('');
    
    // DataForSEO検索結果取得
    Logger.log('【DataForSEO】検索結果取得中...');
    let successCount = 0;
    const allSearchResults = []; // AIO用に結果を保存★
    
    for (let i = 0; i < keywords.length; i++) {
      const keyword = keywords[i];
      const row = startRow + i;
      
      try {
        const searchData = fetchSearchResults(keyword);
        const searchResults = searchData.results;
        
        writeSearchResultToSheetRow(competitorSheet, row, searchResults);
        
        // AIO用に結果を保存★
        allSearchResults.push(searchData);
        
        successCount++;
        
        if ((i + 1) % 10 === 0) {
          Logger.log(`進捗: ${i + 1} / ${keywords.length}`);
        }
        
        Utilities.sleep(1000);
        
      } catch (error) {
        Logger.log(`✗ [${keyword}] エラー: ${error.message}`);
      }
    }
    
    Logger.log(`✓ DataForSEO完了: ${successCount} / ${keywords.length}`);
    Logger.log('');
    
    // 自社DA
    Logger.log('【自社DA】取得中...');
    const ownDomain = 'smaho-tap.com';
    let ownDA;
    const cachedData = getDAFromCache(ownDomain);
    
    if (cachedData === null) {
      const daResults = fetchDomainAuthority([ownDomain]);
      ownDA = daResults[0] || 0;
      if (ownDA > 0) {
        saveDAToCache([{ domain: ownDomain, da: ownDA, pa: 0 }]);
      }
    } else {
      ownDA = cachedData.da;
    }
    
    const ownDAColumn = Array(keywords.length).fill([ownDA]);
    competitorSheet.getRange(startRow, 5, keywords.length, 1).setValues(ownDAColumn);
    Logger.log(`✓ 自社DA（${ownDA}）書き込み完了`);
    Logger.log('');
    
    // 競合DA
    Logger.log('【競合DA】取得中...');
    updateCompetitorDA(startRow, startRow + keywords.length - 1);
    Logger.log('✓ 競合DA完了');
    Logger.log('');
    
    // 勝算度スコア
    Logger.log('【勝算度スコア】算出中...');
    updateWinnableScores(startRow, startRow + keywords.length - 1);
    Logger.log('✓ 勝算度スコア完了');
    Logger.log('');
    
    // AIO順位追跡★NEW
    if (ENABLE_AIO_TRACKING && allSearchResults.length > 0) {
      Logger.log('【AIO順位追跡】処理中...');
      processAIOInBatch(allSearchResults);
      Logger.log('');
    }
    
    Logger.log(`=== バッチ${batchNumber}完了 ===`);
    Logger.log(`処理件数: ${keywords.length}`);
    
    // 次のバッチの案内
    if (endIndex < 212) {
      Logger.log('');
      Logger.log(`次のバッチ: runCompetitorAnalysisBatch(${batchNumber + 1}) を実行してください`);
    } else {
      Logger.log('');
      Logger.log('🎉 全バッチ完了！競合分析シートを確認してください');
    }
    
  } catch (error) {
    Logger.log(`❌ バッチ${batchNumber}でエラー: ${error.message}`);
    Logger.log(error.stack);
  }
}

/**
 * 全5バッチを順次実行（手動実行用）
 */
function runAllBatches() {
  for (let i = 1; i <= 5; i++) {
    Logger.log(`\n========== バッチ${i}/5 ==========\n`);
    runCompetitorAnalysisBatch(i);
    
    if (i < 5) {
      Logger.log('\n次のバッチまで30秒待機...\n');
      Utilities.sleep(30000);
    }
  }
  
  Logger.log('\n🎉 全213KW競合分析完了！');
  
  // 最終処理（AIOサマリー）
  if (ENABLE_AIO_TRACKING) {
    Logger.log('\n=== AIOサマリーレポート生成 ===');
    generateWeeklyAIOReport();
  }
}

/**
 * バッチ実行用の個別関数（トリガー用）
 */
function runBatch1() {
  runCompetitorAnalysisBatch(1);
}

function runBatch2() {
  runCompetitorAnalysisBatch(2);
}

function runBatch3() {
  runCompetitorAnalysisBatch(3);
}

function runBatch4() {
  runCompetitorAnalysisBatch(4);
}

function runBatch5() {
  runCompetitorAnalysisBatch(5);
}

// ============================================================
// AIO順位追跡統合 ★Day 16追加 v3.0
// ============================================================

/**
 * バッチ処理でAIO順位も記録する
 * 
 * @param {Array} searchResults - fetchSearchResultsの結果配列
 */
function processAIOInBatch(searchResults) {
  if (!ENABLE_AIO_TRACKING) {
    Logger.log('AIO追跡は無効です');
    return;
  }
  
  // AIOTracking.gsの関数が存在するか確認
  if (typeof processAIOForMultipleKeywords !== 'function') {
    Logger.log('⚠️ AIOTracking.gsが読み込まれていません');
    Logger.log('  AIOTracking.gsをApps Scriptに追加してください');
    return;
  }
  
  try {
    const aioResults = processAIOForMultipleKeywords(searchResults);
    
    const aioDisplayed = aioResults.filter(r => r.hasAIO).length;
    const aioWithOwnSite = aioResults.filter(r => r.ownSiteFound).length;
    
    Logger.log('✓ AIO処理完了');
    Logger.log('  AIO表示: ' + aioDisplayed + '/' + aioResults.length + '件');
    Logger.log('  自社引用: ' + aioWithOwnSite + '件');
    
    // TOP3にいるキーワードを表示
    const top3Keywords = aioResults.filter(r => r.ownSiteFound && r.ownSitePosition <= 3);
    if (top3Keywords.length > 0) {
      Logger.log('  TOP3引用キーワード:');
      top3Keywords.forEach(r => {
        Logger.log(`    - ${r.keyword} (${r.ownSitePosition}位)`);
      });
    }
    
  } catch (e) {
    Logger.log('✗ AIO処理エラー: ' + e.message);
  }
}

/**
 * 週次最終処理でAIOサマリーレポートを生成
 */
function generateWeeklyAIOReport() {
  if (!ENABLE_AIO_TRACKING) {
    Logger.log('AIO追跡は無効です');
    return;
  }
  
  // AIOTracking.gsの関数が存在するか確認
  if (typeof generateAIOSummaryReport !== 'function') {
    Logger.log('⚠️ AIOTracking.gsが読み込まれていません');
    return;
  }
  
  try {
    const report = generateAIOSummaryReport();
    
    Logger.log('✓ AIOサマリー生成完了');
    Logger.log('');
    Logger.log('【AIOサマリー】');
    Logger.log('  総キーワード: ' + report.summary.totalKeywords);
    Logger.log('  AIO表示: ' + report.summary.aioDisplayed + '件');
    Logger.log('  自社引用: ' + report.summary.ownSiteInAIO + '件');
    Logger.log('  TOP3引用: ' + report.summary.aioTop3 + '件');
    
    if (report.summary.improved > 0) {
      Logger.log('  ↑ 改善: ' + report.summary.improved + '件');
    }
    if (report.summary.declined > 0) {
      Logger.log('  ↓ 悪化: ' + report.summary.declined + '件');
    }
    if (report.summary.newAppearance > 0) {
      Logger.log('  ★ 新規引用: ' + report.summary.newAppearance + '件');
    }
    if (report.summary.lostAppearance > 0) {
      Logger.log('  ✗ 引用消失: ' + report.summary.lostAppearance + '件');
    }
    
    // 自社引用キーワード一覧
    if (report.details && report.details.keywordsInAIO && report.details.keywordsInAIO.length > 0) {
      Logger.log('');
      Logger.log('【自社引用キーワード TOP10】');
      const top10 = report.details.keywordsInAIO.slice(0, 10);
      top10.forEach((item, index) => {
        Logger.log(`  ${index + 1}. ${item.keyword} (${item.position}位)`);
      });
    }
    
  } catch (e) {
    Logger.log('✗ AIOサマリーエラー: ' + e.message);
  }
}

// ============================================================
// DA未取得再取得関数
// ============================================================

/**
 * DA未取得のURLを抽出して再取得
 */
function retryMissingDAs() {
  Logger.log('=== DA未取得URL再取得開始 ===');
  Logger.log('');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('競合分析');
  
  if (!sheet) {
    throw new Error('競合分析シートが見つかりません');
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('データがありません');
    return;
  }
  
  // 全データを取得
  const data = sheet.getRange(2, 1, lastRow - 1, 36).getValues();
  
  const missingDAs = [];
  
  // DA未取得のURLを抽出
  data.forEach((row, index) => {
    const rowIndex = index + 2;
    
    // G, I, K, M, O, Q, S, U, W, Y列（URL）
    // H, J, L, N, P, R, T, V, X, Z列（DA）
    for (let i = 0; i < 10; i++) {
      const urlColIndex = 6 + (i * 2);  // 6, 8, 10...
      const daColIndex = 7 + (i * 2);   // 7, 9, 11...
      
      const url = row[urlColIndex];
      const da = row[daColIndex];
      
      if (url && (!da || da === 0 || da === '')) {
        const domain = extractDomain(url);
        if (domain) {
          missingDAs.push({
            row: rowIndex,
            col: daColIndex + 1,  // シート列番号（1-based）
            url: url,
            domain: domain
          });
        }
      }
    }
  });
  
  Logger.log(`DA未取得URL: ${missingDAs.length}件`);
  Logger.log('');
  
  if (missingDAs.length === 0) {
    Logger.log('✓ 全てのURLでDA取得済み');
    return;
  }
  
  // 重複を除外
  const uniqueDomains = [...new Set(missingDAs.map(item => item.domain))];
  Logger.log(`ユニークドメイン: ${uniqueDomains.length}件`);
  Logger.log('');
  
  // Moz APIで再取得
  Logger.log('Moz APIで再取得中...');
  
  try {
    const daResults = fetchDomainAuthority(uniqueDomains);
    
    Logger.log(`✓ ${daResults.length}件のDA取得成功`);
    Logger.log('');
    
    // 結果をマップ化（末尾スラッシュを削除）
    const daMap = {};
    daResults.forEach(result => {
      const domain = result.domain.replace(/\/$/, '');  // 末尾スラッシュを削除
      daMap[domain] = result.da;
      Logger.log(`[マップ化] "${result.domain}" → "${domain}" (DA: ${result.da})`);
    });
    
    // シートに書き込み
    let updateCount = 0;
    missingDAs.forEach(item => {
      const da = daMap[item.domain];
      if (da !== undefined && da > 0) {
        sheet.getRange(item.row, item.col).setValue(da);
        updateCount++;
        Logger.log(`✓ [行${item.row}] ${item.domain} → DA: ${da}`);
      } else {
        Logger.log(`✗ [行${item.row}] ${item.domain} → DA取得失敗`);
      }
    });
    
    Logger.log('');
    Logger.log('=== 再取得完了 ===');
    Logger.log(`更新数: ${updateCount} / ${missingDAs.length}`);
    
    // キャッシュに保存
    if (daResults.length > 0) {
      saveDAToCache(daResults);
      Logger.log('✓ DA履歴シートに保存');
    }
    
  } catch (error) {
    Logger.log(`❌ エラー: ${error.message}`);
    Logger.log(error.stack);
  }
}

/**
 * DA未取得URLを一覧表示（実行前確認用）
 */
function listMissingDAs() {
  Logger.log('=== DA未取得URL一覧 ===');
  Logger.log('');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('競合分析');
  
  if (!sheet) {
    throw new Error('競合分析シートが見つかりません');
  }
  
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(2, 1, lastRow - 1, 36).getValues();
  
  const missingDAs = [];
  
  data.forEach((row, index) => {
    const rowIndex = index + 2;
    const keyword = row[1];  // B列
    
    for (let i = 0; i < 10; i++) {
      const urlColIndex = 6 + (i * 2);
      const daColIndex = 7 + (i * 2);
      
      const url = row[urlColIndex];
      const da = row[daColIndex];
      
      if (url && (!da || da === 0 || da === '')) {
        const domain = extractDomain(url);
        if (domain) {
          missingDAs.push({
            keyword: keyword,
            rank: i + 1,
            domain: domain,
            url: url
          });
        }
      }
    }
  });
  
  Logger.log(`合計: ${missingDAs.length}件`);
  Logger.log('');
  
  // ドメインごとにグループ化
  const domainCount = {};
  missingDAs.forEach(item => {
    domainCount[item.domain] = (domainCount[item.domain] || 0) + 1;
  });
  
  Logger.log('ドメイン別件数:');
  Object.entries(domainCount).forEach(([domain, count]) => {
    Logger.log(`  ${domain}: ${count}件`);
  });
  
  Logger.log('');
  Logger.log('retryMissingDAs() を実行して再取得してください');
}

/**
 * DA未取得を再取得（未取得件数を返すバージョン）
 */
function retryMissingDAsWithCount() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('競合分析');
  
  if (!sheet) {
    return { updated: 0, missingCount: 0, retriedCount: 0 };
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return { updated: 0, missingCount: 0, retriedCount: 0 };
  }
  
  const data = sheet.getRange(2, 1, lastRow - 1, 36).getValues();
  const missingDAs = [];
  
  // DA未取得のURLを抽出
  data.forEach((row, index) => {
    const rowIndex = index + 2;
    
    for (let i = 0; i < 10; i++) {
      const urlColIndex = 6 + (i * 2);
      const daColIndex = 7 + (i * 2);
      
      const url = row[urlColIndex];
      const da = row[daColIndex];
      
      if (url && (!da || da === 0 || da === '')) {
        const domain = extractDomain(url);
        if (domain) {
          missingDAs.push({
            row: rowIndex,
            col: daColIndex + 1,
            url: url,
            domain: domain
          });
        }
      }
    }
  });
  
  if (missingDAs.length === 0) {
    return { updated: 0, missingCount: 0, retriedCount: 0 };
  }
  
  // 重複を除外
  const uniqueDomains = [...new Set(missingDAs.map(item => item.domain))];
  
  // Moz APIで再取得（最大50件）
  const domainsToFetch = uniqueDomains.slice(0, 50);
  
  try {
    const daResults = fetchDomainAuthority(domainsToFetch);
    
    // 結果をマップ化（末尾スラッシュを削除）
    const daMap = {};
    daResults.forEach(result => {
      const domain = result.domain.replace(/\/$/, '');
      daMap[domain] = result.da;
    });
    
    // シートに書き込み
    let updateCount = 0;
    missingDAs.forEach(item => {
      const da = daMap[item.domain];
      if (da !== undefined && da > 0) {
        sheet.getRange(item.row, item.col).setValue(da);
        updateCount++;
      }
    });
    
    // キャッシュに保存
    if (daResults.length > 0) {
      saveDAToCache(daResults);
    }
    
    return {
      updated: updateCount,
      missingCount: missingDAs.length - updateCount,
      retriedCount: updateCount
    };
    
  } catch (error) {
    Logger.log(`DA再取得エラー: ${error.message}`);
    return { updated: 0, missingCount: missingDAs.length, retriedCount: 0 };
  }
}

// ============================================================
// 週次自動実行関数
// ============================================================

/**
 * DA未取得の自動再取得 + AIOサマリー（週次トリガー用）★v3.0更新
 * トリガー設定: 月曜 5:35
 * 
 * 【処理内容】
 * - 競合分析シートでDA未取得（空白または0）の行を検出
 * - Moz APIで再取得
 * - 最大15回リトライ
 * - 5分経過で安全に中断
 * - AIOサマリーレポート生成★NEW
 * 
 * @returns {void}
 */
function weeklyDARetry() {
  const startTime = new Date();
  const SAFE_EXECUTION_TIME = 240;  // 4分（AIOサマリー用に1分確保）
  const MAX_DA_RETRY = 15;
  
  Logger.log('===========================================');
  Logger.log('=== 週次DA自動再取得開始（v3.0 AIO対応）===');
  Logger.log('===========================================');
  Logger.log('開始時刻: ' + startTime.toLocaleString('ja-JP'));
  Logger.log('AIO追跡: ' + (ENABLE_AIO_TRACKING ? '有効' : '無効'));
  
  let totalRetried = 0;
  let lastMissingCount = -1;
  let retryCount = 0;
  
  for (let retry = 1; retry <= MAX_DA_RETRY; retry++) {
    // 4分経過チェック（AIOサマリー用に1分確保）
    const elapsed = (new Date() - startTime) / 1000;
    if (elapsed > SAFE_EXECUTION_TIME) {
      Logger.log('⚠️ 4分経過のため中断（次週のトリガーで継続）');
      break;
    }
    
    // DA未取得を再取得
    const result = retryMissingDAsWithCount();
    retryCount++;
    
    Logger.log(`リトライ ${retry}: DA未取得 ${result.missingCount}件, 再取得 ${result.retriedCount}件`);
    
    // DA未取得が0件になったら終了
    if (result.missingCount === 0) {
      Logger.log('✅ DA未取得が0件になりました！');
      break;
    }
    
    // 進捗がない場合は終了（無限ループ防止）
    if (result.missingCount === lastMissingCount) {
      Logger.log('⚠️ 進捗なし（取得不可能なドメインの可能性）、終了します');
      break;
    }
    
    totalRetried += result.retriedCount || 0;
    lastMissingCount = result.missingCount;
    
    // API制限対策の待機
    Utilities.sleep(2000);
  }
  
  Logger.log('');
  Logger.log('=== DA再取得完了 ===');
  Logger.log(`リトライ回数: ${retryCount}回`);
  Logger.log(`再取得成功数: ${totalRetried}件`);
  Logger.log(`残りDA未取得: ${lastMissingCount}件`);
  
  // AIOサマリーレポート生成★NEW
  if (ENABLE_AIO_TRACKING) {
    Logger.log('');
    Logger.log('===========================================');
    Logger.log('=== AIOサマリーレポート生成 ===');
    Logger.log('===========================================');
    generateWeeklyAIOReport();
  }
  
  const endTime = new Date();
  const duration = ((endTime - startTime) / 1000).toFixed(1);
  
  Logger.log('');
  Logger.log('===========================================');
  Logger.log('=== 週次最終処理完了（v3.0）===');
  Logger.log('===========================================');
  Logger.log(`処理時間: ${duration}秒`);
}

// ============================================================
// トリガー設定関数
// ============================================================

/**
 * 週次競合分析トリガーを設定（6分割版 v3.0）
 * 
 * 【設定されるトリガー】
 * 1. runBatch1 - 月曜 5:00
 * 2. runBatch2 - 月曜 5:10
 * 3. runBatch3 - 月曜 5:15
 * 4. runBatch4 - 月曜 5:20
 * 5. runBatch5 - 月曜 5:25
 * 6. weeklyDARetry - 月曜 5:35（AIOサマリー含む）
 * 
 * 【使い方】
 * 1. Apps Scriptエディタでこの関数を実行
 * 2. 既存の競合分析トリガーは自動削除される
 * 3. 6つの新しいトリガーが設定される
 */
function setupWeeklyCompetitorTriggers() {
  Logger.log('===========================================');
  Logger.log('=== 週次競合分析トリガー設定（v3.0）===');
  Logger.log('===========================================');
  
  // Step 1: 既存の競合分析トリガーを削除
  Logger.log('');
  Logger.log('【Step 1】既存トリガーの削除');
  const triggers = ScriptApp.getProjectTriggers();
  let deletedCount = 0;
  
  for (let i = 0; i < triggers.length; i++) {
    const trigger = triggers[i];
    const funcName = trigger.getHandlerFunction();
    
    // 競合分析関連のトリガーを削除
    if (funcName.indexOf('Batch') !== -1 || 
        funcName.indexOf('weeklyCompetitor') !== -1 ||
        funcName === 'weeklyDARetry') {
      ScriptApp.deleteTrigger(trigger);
      Logger.log('  削除: ' + funcName);
      deletedCount++;
    }
  }
  Logger.log('  削除完了: ' + deletedCount + '件');
  
  // Step 2: 新しいトリガーを設定
  Logger.log('');
  Logger.log('【Step 2】新規トリガーの設定');
  
  const triggerConfigs = [
    { func: 'runBatch1', hour: 5, minute: 0 },
    { func: 'runBatch2', hour: 5, minute: 10 },
    { func: 'runBatch3', hour: 5, minute: 15 },
    { func: 'runBatch4', hour: 5, minute: 20 },
    { func: 'runBatch5', hour: 5, minute: 25 },
    { func: 'weeklyDARetry', hour: 5, minute: 35 }
  ];
  
  for (let j = 0; j < triggerConfigs.length; j++) {
    const config = triggerConfigs[j];
    
    ScriptApp.newTrigger(config.func)
      .timeBased()
      .onWeekDay(ScriptApp.WeekDay.MONDAY)
      .atHour(config.hour)
      .nearMinute(config.minute)
      .create();
    
    const timeStr = config.hour + ':' + (config.minute < 10 ? '0' : '') + config.minute;
    Logger.log('  作成: ' + config.func + ' (月曜 ' + timeStr + ')');
  }
  
  Logger.log('');
  Logger.log('===========================================');
  Logger.log('✅ 6つのトリガーを設定しました（v3.0 AIO対応）');
  Logger.log('===========================================');
  Logger.log('');
  Logger.log('【トリガー一覧】');
  Logger.log('1. runBatch1     - 月曜 5:00  (KW 1-50 + AIO)');
  Logger.log('2. runBatch2     - 月曜 5:10  (KW 51-100 + AIO)');
  Logger.log('3. runBatch3     - 月曜 5:15  (KW 101-150 + AIO)');
  Logger.log('4. runBatch4     - 月曜 5:20  (KW 151-200 + AIO)');
  Logger.log('5. runBatch5     - 月曜 5:25  (KW 201-213 + AIO)');
  Logger.log('6. weeklyDARetry - 月曜 5:35  (DA再取得 + AIOサマリー)');
}

/**
 * 現在のトリガー一覧を表示
 */
function listAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  
  Logger.log('===========================================');
  Logger.log('=== 現在のトリガー一覧 ===');
  Logger.log('===========================================');
  
  if (triggers.length === 0) {
    Logger.log('トリガーは設定されていません');
    return;
  }
  
  for (let i = 0; i < triggers.length; i++) {
    const trigger = triggers[i];
    const funcName = trigger.getHandlerFunction();
    const triggerType = trigger.getEventType();
    
    Logger.log((i + 1) + '. ' + funcName + ' (' + triggerType + ')');
  }
  
  Logger.log('-------------------------------------------');
  Logger.log('合計: ' + triggers.length + '件');
}

/**
 * 競合分析関連のトリガーのみ削除
 */
function deleteCompetitorTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  let deletedCount = 0;
  
  Logger.log('=== 競合分析トリガーを削除 ===');
  
  for (let i = 0; i < triggers.length; i++) {
    const trigger = triggers[i];
    const funcName = trigger.getHandlerFunction();
    
    if (funcName.indexOf('Batch') !== -1 || 
        funcName.indexOf('weeklyCompetitor') !== -1 ||
        funcName === 'weeklyDARetry') {
      ScriptApp.deleteTrigger(trigger);
      Logger.log('削除: ' + funcName);
      deletedCount++;
    }
  }
  
  Logger.log('削除完了: ' + deletedCount + '件');
}

// ============================================================
// 旧トリガー設定関数（互換性のため残す）
// ============================================================

/**
 * 週次自動更新トリガーを設定（DA自動再取得版）
 * @deprecated setupWeeklyCompetitorTriggers() を使用してください
 */
function setupWeeklyCompetitorAnalysisTrigger() {
  Logger.log('⚠️ この関数は非推奨です。');
  Logger.log('代わりに setupWeeklyCompetitorTriggers() を実行してください。');
  Logger.log('');
  Logger.log('setupWeeklyCompetitorTriggers() を自動実行します...');
  Logger.log('');
  
  setupWeeklyCompetitorTriggers();
}

/**
 * トリガー一覧を表示（旧版）
 */
function listCompetitorAnalysisTriggers() {
  Logger.log('=== トリガー一覧 ===');
  Logger.log('');
  
  const triggers = ScriptApp.getProjectTriggers();
  let found = false;
  
  triggers.forEach(trigger => {
    const funcName = trigger.getHandlerFunction();
    if (funcName === 'runFullCompetitorAnalysis' ||
        funcName.indexOf('Batch') !== -1 ||
        funcName.indexOf('weeklyCompetitor') !== -1 ||
        funcName === 'weeklyDARetry') {
      found = true;
      const triggerSource = trigger.getTriggerSource();
      const eventType = trigger.getEventType();
      
      Logger.log('関数: ' + funcName);
      Logger.log('種類: ' + triggerSource);
      Logger.log('イベント: ' + eventType);
      Logger.log('');
    }
  });
  
  if (!found) {
    Logger.log('⚠ 競合分析トリガーが設定されていません');
    Logger.log('setupWeeklyCompetitorTriggers() を実行してください');
  }
  
  Logger.log('=== トリガー一覧終了 ===');
}

// ============================================================
// 旧週次自動実行関数（非推奨）
// ============================================================

/**
 * 週次自動実行用（バッチ処理 + DA自動再取得）
 * @deprecated 6分制限でタイムアウトするため、6トリガー分割版を使用
 * 
 * 【代替手順】
 * setupWeeklyCompetitorTriggers() を実行して6つのトリガーを設定
 */
function weeklyCompetitorAnalysisWithDARetry() {
  Logger.log('⚠️ この関数は6分制限でタイムアウトする可能性があります。');
  Logger.log('');
  Logger.log('【推奨】setupWeeklyCompetitorTriggers() を実行して');
  Logger.log('6つのトリガーに分割してください。');
  Logger.log('');
  Logger.log('今回は処理を続行します...');
  Logger.log('');
  
  Logger.log('=== 週次競合分析開始（DA自動再取得版） ===');
  Logger.log('実行時刻: ' + new Date());
  Logger.log('');
  
  const startTime = new Date();
  
  try {
    // バッチ1-5を順次実行
    for (let i = 1; i <= 5; i++) {
      Logger.log(`\n========== バッチ${i}/5 ==========\n`);
      runCompetitorAnalysisBatch(i);
      
      if (i < 5) {
        Logger.log('次のバッチまで5秒待機...');
        Utilities.sleep(5000);
      }
    }
    
    Logger.log('');
    Logger.log('=== 全バッチ完了 ===');
    Logger.log('');
    
    // DA未取得を自動再取得（最大15回）
    Logger.log('=== DA自動再取得開始 ===');
    
    let totalUpdated = 0;
    
    for (let retry = 1; retry <= 15; retry++) {
      Logger.log(`\n--- DA再取得 ${retry}/15 ---\n`);
      
      const result = retryMissingDAsWithCount();
      totalUpdated += result.updated;
      
      Logger.log(`更新: ${result.updated}件`);
      Logger.log(`残り未取得: ${result.missingCount}件`);
      
      if (result.missingCount === 0) {
        Logger.log('');
        Logger.log('✓ 全てのDA取得完了');
        break;
      }
      
      if (retry < 15 && result.missingCount > 0) {
        Logger.log('5秒待機...');
        Utilities.sleep(5000);
      }
    }
    
    // AIOサマリーレポート生成★NEW
    if (ENABLE_AIO_TRACKING) {
      Logger.log('');
      Logger.log('=== AIOサマリーレポート生成 ===');
      generateWeeklyAIOReport();
    }
    
    const endTime = new Date();
    const duration = (endTime - startTime) / 1000 / 60; // 分
    
    Logger.log('');
    Logger.log('=== 週次競合分析完了 ===');
    Logger.log(`実行時間: ${duration.toFixed(1)}分`);
    Logger.log(`DA更新合計: ${totalUpdated}件`);
    Logger.log('');
    
  } catch (error) {
    Logger.log('');
    Logger.log('❌ エラー発生: ' + error.message);
    Logger.log(error.stack);
  }
}

// ============================================================
// デバッグ・テスト関数
// ============================================================

/**
 * ターゲットKW分析シートの列名を確認（デバッグ）
 */
function debugTargetKWSheetHeaders() {
  Logger.log('=== ターゲットKW分析シート 列名確認 ===');
  Logger.log('');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('ターゲットKW分析');
  
  if (!sheet) {
    Logger.log('❌ ターゲットKW分析シートが見つかりません');
    return;
  }
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  Logger.log('列数: ' + headers.length);
  Logger.log('');
  Logger.log('列名一覧:');
  
  headers.forEach((header, index) => {
    Logger.log(`列${index + 1}: "${header}"`);
  });
  
  Logger.log('');
  Logger.log('=== 確認完了 ===');
}

/**
 * Moz APIの返り値形式を確認（デバッグ）
 */
function debugMozResponse() {
  Logger.log('=== Moz API返り値デバッグ ===');
  Logger.log('');
  
  // テスト用ドメイン
  const testDomains = [
    'news.yahoo.co.jp',
    'gadgenect.jp',
    'smaho-tap.com'
  ];
  
  Logger.log('テストドメイン: ' + testDomains.join(', '));
  Logger.log('');
  
  try {
    const results = fetchDomainAuthority(testDomains);
    
    Logger.log('取得件数: ' + results.length);
    Logger.log('');
    
    results.forEach(result => {
      Logger.log('--- 結果 ---');
      Logger.log('domain: "' + result.domain + '"');
      Logger.log('型: ' + typeof result.domain);
      Logger.log('長さ: ' + result.domain.length);
      Logger.log('DA: ' + result.da);
      Logger.log('PA: ' + result.pa);
      Logger.log('');
    });
    
    // extractDomain()との比較
    Logger.log('=== extractDomain()との比較 ===');
    Logger.log('');
    
    const testUrls = [
      'https://news.yahoo.co.jp/articles/xxx',
      'https://gadgenect.jp/xxx',
      'https://smaho-tap.com/purchase-amazon-iphone'
    ];
    
    testUrls.forEach(url => {
      const extracted = extractDomain(url);
      Logger.log('URL: ' + url);
      Logger.log('抽出: "' + extracted + '"');
      Logger.log('');
    });
    
  } catch (error) {
    Logger.log('❌ エラー: ' + error.message);
  }
}

/**
 * weeklyDARetry のテスト実行（手動）
 */
function testWeeklyDARetry() {
  Logger.log('=== weeklyDARetry テスト実行 ===');
  weeklyDARetry();
}

/**
 * バッチ処理のテスト（バッチ1のみ）
 */
function testBatch1() {
  Logger.log('=== バッチ1 テスト実行 ===');
  runBatch1();
}

/**
 * DA未取得件数をカウント
 */
function countMissingDAs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('競合分析');
  
  if (!sheet) {
    Logger.log('競合分析シートが見つかりません');
    return;
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('データがありません');
    return;
  }
  
  const data = sheet.getRange(2, 1, lastRow - 1, 36).getValues();
  
  let missingCount = 0;
  let totalUrls = 0;
  
  data.forEach(row => {
    for (let i = 0; i < 10; i++) {
      const urlColIndex = 6 + (i * 2);
      const daColIndex = 7 + (i * 2);
      
      const url = row[urlColIndex];
      const da = row[daColIndex];
      
      if (url && url !== '') {
        totalUrls++;
        if (!da || da === '' || da === 0) {
          missingCount++;
        }
      }
    }
  });
  
  Logger.log('=== DA未取得カウント ===');
  Logger.log('総URL数: ' + totalUrls);
  Logger.log('DA取得済み: ' + (totalUrls - missingCount));
  Logger.log('DA未取得: ' + missingCount);
  Logger.log('取得率: ' + ((totalUrls - missingCount) / totalUrls * 100).toFixed(1) + '%');
}

// ============================================================
// AIO統合テスト関数 ★v3.0追加
// ============================================================

/**
 * AIO統合テスト（1キーワード）
 */
function testAIOIntegrationSingle() {
  Logger.log('=== AIO統合テスト（1キーワード） ===');
  Logger.log('');
  
  const testKeyword = 'iphone 画面がセピア色になる';
  
  Logger.log('テストキーワード: ' + testKeyword);
  Logger.log('AIO追跡: ' + (ENABLE_AIO_TRACKING ? '有効' : '無効'));
  Logger.log('');
  
  // AIOTracking.gsの関数確認
  if (typeof processAIOForMultipleKeywords !== 'function') {
    Logger.log('❌ AIOTracking.gsが読み込まれていません');
    Logger.log('  AIOTracking.gsをApps Scriptに追加してください');
    return;
  }
  
  if (typeof getOwnDomain !== 'function') {
    Logger.log('❌ getOwnDomain関数が見つかりません');
    return;
  }
  
  Logger.log('自社ドメイン: ' + getOwnDomain());
  Logger.log('');
  
  // 検索結果取得
  Logger.log('【ステップ1】検索結果取得');
  try {
    const searchData = fetchSearchResults(testKeyword);
    
    Logger.log('✓ 検索結果取得成功');
    Logger.log('  オーガニック結果: ' + (searchData.results ? searchData.results.length : 0) + '件');
    Logger.log('  AIO: ' + (searchData.aio && searchData.aio.hasAIO ? 'あり（' + searchData.aio.totalReferences + '件引用）' : 'なし'));
    Logger.log('');
    
    // AIO処理
    Logger.log('【ステップ2】AIO順位処理');
    const aioResults = processAIOForMultipleKeywords([searchData]);
    
    if (aioResults.length > 0) {
      const result = aioResults[0];
      Logger.log('✓ AIO処理成功');
      Logger.log('  AIO表示: ' + (result.hasAIO ? 'あり' : 'なし'));
      Logger.log('  自社引用: ' + (result.ownSiteFound ? result.ownSitePosition + '位' : 'なし'));
      
      if (result.ownSiteUrl) {
        Logger.log('  引用URL: ' + result.ownSiteUrl);
      }
    }
    
    Logger.log('');
    Logger.log('=== テスト完了 ===');
    Logger.log('AIO順位履歴シートを確認してください');
    
  } catch (error) {
    Logger.log('❌ エラー: ' + error.message);
    Logger.log(error.stack);
  }
}

/**
 * AIO統合テスト（複数キーワード）
 */
function testAIOIntegrationMultiple() {
  Logger.log('=== AIO統合テスト（複数キーワード） ===');
  Logger.log('');
  
  const testKeywords = [
    'iphone 画面がセピア色になる',
    'iphone 保険',
    'スマホ 画面 修理'
  ];
  
  Logger.log('テストキーワード: ' + testKeywords.length + '件');
  Logger.log('AIO追跡: ' + (ENABLE_AIO_TRACKING ? '有効' : '無効'));
  Logger.log('');
  
  // AIOTracking.gsの関数確認
  if (typeof fetchMultipleSearchResults !== 'function') {
    Logger.log('❌ fetchMultipleSearchResults関数が見つかりません');
    Logger.log('  CompetitorAnalysis_DataForSEO.gsを確認してください');
    return;
  }
  
  if (typeof processAIOForMultipleKeywords !== 'function') {
    Logger.log('❌ AIOTracking.gsが読み込まれていません');
    return;
  }
  
  Logger.log('自社ドメイン: ' + getOwnDomain());
  Logger.log('');
  
  // 検索結果取得
  Logger.log('【ステップ1】検索結果取得');
  const searchResults = fetchMultipleSearchResults(testKeywords);
  
  const successCount = searchResults.filter(r => r.results && r.results.length > 0).length;
  const aioCount = searchResults.filter(r => r.aio && r.aio.hasAIO).length;
  
  Logger.log('✓ 検索結果: ' + successCount + '/' + testKeywords.length + '成功');
  Logger.log('  AIOあり: ' + aioCount + '件');
  Logger.log('');
  
  // AIO処理
  Logger.log('【ステップ2】AIO順位処理');
  const aioResults = processAIOForMultipleKeywords(searchResults);
  
  Logger.log('');
  Logger.log('【結果サマリー】');
  
  let aioDisplayed = 0;
  let ownSiteInAIO = 0;
  
 for (let i = 0; i < aioResults.length; i++) {
    const result = aioResults[i];
    Logger.log('');
    Logger.log(result.keyword + ':');
    Logger.log('  AIO: ' + (result.hasAIO ? 'あり' : 'なし'));
    
    if (result.hasAIO) {
      aioDisplayed++;
      Logger.log('  自社引用: ' + (result.ownSiteFound ? result.ownSitePosition + '位' : 'なし'));
      
      if (result.ownSiteFound) {
        ownSiteInAIO++;
      }
    }
  }
  
  Logger.log('');
  Logger.log('【統計】');
  Logger.log('  AIO表示: ' + aioDisplayed + '/' + testKeywords.length);
  Logger.log('  自社引用: ' + ownSiteInAIO + '/' + aioDisplayed);
  Logger.log('');
  Logger.log('=== テスト完了 ===');
  Logger.log('AIO順位履歴シートを確認してください');
}

/**
 * AIOサマリーレポートのテスト
 */
function testAIOSummaryReport() {
  Logger.log('=== AIOサマリーレポートテスト ===');
  Logger.log('');
  
  if (typeof generateAIOSummaryReport !== 'function') {
    Logger.log('❌ AIOTracking.gsが読み込まれていません');
    return;
  }
  
  generateWeeklyAIOReport();
  
  Logger.log('');
  Logger.log('=== テスト完了 ===');
}

/**
 * AIO追跡の有効/無効を切り替え（デバッグ用）
 */
function toggleAIOTracking() {
  Logger.log('=== AIO追跡設定 ===');
  Logger.log('');
  Logger.log('現在の設定: ENABLE_AIO_TRACKING = ' + ENABLE_AIO_TRACKING);
  Logger.log('');
  Logger.log('設定を変更するには、コード内の ENABLE_AIO_TRACKING を編集してください:');
  Logger.log('  const ENABLE_AIO_TRACKING = true;  // AIO追跡有効');
  Logger.log('  const ENABLE_AIO_TRACKING = false; // AIO追跡無効');
}