/**
 * ============================================================================
 * Day 11-12: ターゲットKW分析連携・完全自動化
 * ============================================================================
 * ターゲットKW分析の既存キーワードで競合分析を実行
 */

/**
 * 競合分析シートの既存データをクリア（ヘッダーは残す）
 */
function clearCompetitorAnalysisData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('競合分析');
  
  if (!sheet) {
    throw new Error('競合分析シートが見つかりません');
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1);
    Logger.log(`✓ 既存データ削除: ${lastRow - 1} 行`);
  } else {
    Logger.log('既存データなし');
  }
}

/**
 * ターゲットKW分析から指定件数のキーワードをインポート
 * @param {number} count - 取得件数（デフォルト: 5）
 * @return {Array} インポートしたキーワードとページURL
 */
function importKeywordsFromTargetKW(count = 5) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetKWSheet = ss.getSheetByName('ターゲットKW分析');
  const competitorSheet = ss.getSheetByName('競合分析');
  
  if (!targetKWSheet) {
    throw new Error('ターゲットKW分析シートが見つかりません');
  }
  
  if (!competitorSheet) {
    throw new Error('競合分析シートが見つかりません');
  }
  
  Logger.log(`=== ターゲットKW分析から${count}件インポート ===`);
  Logger.log('');
  
  // ターゲットKW分析からデータ取得（A列: keyword_id, B列: page_url, C列: target_keyword）
  const targetKWData = targetKWSheet.getDataRange().getValues();
  
  const importData = [];
  const now = new Date();
  const dateStr = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyyMMdd');
  
  for (let i = 1; i < Math.min(count + 1, targetKWData.length); i++) {
    const keywordId = targetKWData[i][0]; // A列
    const pageUrl = targetKWData[i][1];   // B列
    const keyword = targetKWData[i][2];   // C列
    
    if (keyword && pageUrl) {
      const analysisId = `${keyword}_${dateStr}`;
      
      importData.push({
        analysis_id: analysisId,
        target_keyword: keyword,
        page_url: pageUrl,
        analysis_date: now
      });
      
      Logger.log(`✓ [${i}] ${keyword} → ${pageUrl}`);
    }
  }
  
  Logger.log('');
  Logger.log(`インポート件数: ${importData.length} 件`);
  
  // 競合分析シートに書き込み
  if (importData.length > 0) {
    const rows = importData.map(item => [
      item.analysis_id,      // A列
      item.target_keyword,   // B列
      item.page_url,         // C列
      item.analysis_date     // D列
    ]);
    
    competitorSheet.getRange(2, 1, rows.length, 4).setValues(rows);
    Logger.log('✓ 競合分析シートに書き込み完了');
  }
  
  Logger.log('');
  
  return importData;
}

/**
 * ターゲットKWベースで完全な競合分析を実行
 * @param {number} keywordCount - 分析するキーワード数（デフォルト: 5）
 */
function runCompetitorAnalysisForTargetKW(keywordCount = 5) {
  Logger.log('===========================================');
  Logger.log('ターゲットKWベース完全競合分析');
  Logger.log('===========================================');
  Logger.log('');
  
  // ステップ1: 既存データクリア
  Logger.log('【ステップ1】既存データクリア');
  Logger.log('');
  clearCompetitorAnalysisData();
  Logger.log('');
  
  // ステップ2: ターゲットKW分析からインポート
  Logger.log('【ステップ2】ターゲットKW分析からインポート');
  Logger.log('');
  const importData = importKeywordsFromTargetKW(keywordCount);
  Logger.log('');
  
  if (importData.length === 0) {
    Logger.log('⚠ インポートするデータがありません');
    return;
  }
  
  // ステップ3: DataForSEO検索結果取得
  Logger.log('【ステップ3】DataForSEO検索結果取得');
  Logger.log('');
  
  const keywords = importData.map(item => item.target_keyword);
  
  // キーワードごとに検索結果を取得して書き込み
  let successCount = 0;
  let failCount = 0;
  
  keywords.forEach((keyword, index) => {
    try {
      const rowIndex = index + 2;
      Logger.log(`[${keyword}] DataForSEO API呼び出し中... (試行 1/3、${index + 1}/${keywords.length}件)`);
      
      const result = fetchSearchResults(keyword);
      
      // デバッグログ追加
      Logger.log(`[DEBUG] result型: ${typeof result}`);
      Logger.log(`[DEBUG] result配列判定: ${Array.isArray(result)}`);
      Logger.log(`[DEBUG] result内容: ${JSON.stringify(result).substring(0, 200)}`);
      
      // fetchSearchResults()は配列を返すので、配列チェックに修正
      if (result && Array.isArray(result) && result.length > 0) {
        // writeSearchResultToSheet()はオブジェクト形式を期待するのでラップ
        writeSearchResultToSheet({ keyword: keyword, organic_results: result }, rowIndex);
        successCount++;
        Logger.log(`[${index + 1}/${keywords.length}] ✓ ${keyword} - 成功`);
      } else if (result && typeof result === 'object' && !Array.isArray(result)) {
        // オブジェクト形式の場合
        writeSearchResultToSheet(result, rowIndex);
        successCount++;
        Logger.log(`[${index + 1}/${keywords.length}] ✓ ${keyword} - 成功（オブジェクト形式）`);
      } else {
        failCount++;
        Logger.log(`[${index + 1}/${keywords.length}] ✗ ${keyword} - 失敗（結果なし）`);
      }
      
      // レート制限対策: 1秒待機
      if (index < keywords.length - 1) {
        Utilities.sleep(1000);
      }
    } catch (error) {
      failCount++;
      Logger.log(`[${index + 1}/${keywords.length}] ✗ ${keyword} - エラー: ${error.message}`);
    }
  });
  
  Logger.log('');
  Logger.log('=== 取得結果サマリー ===');
  Logger.log(`成功: ${successCount} / ${keywords.length}`);
  Logger.log(`失敗: ${failCount} / ${keywords.length}`);
  Logger.log(`推定コスト: $${(successCount * 0.006).toFixed(3)}`);
  Logger.log('');
  
  if (successCount === 0) {
    Logger.log('⚠ 検索結果が1件も取得できませんでした');
    return;
  }
  
  Logger.log('✓ 検索結果書き込み完了');
  Logger.log('');
  
  // ステップ4: 自社サイトDA取得
  Logger.log('【ステップ4】自社サイトDA取得');
  Logger.log('');
  const daCount = updateOwnSiteDA();
  Logger.log('');
  
  if (daCount === 0) {
    Logger.log('⚠ 自社サイトDAが取得できませんでした');
    return;
  }
  
  // ステップ5: 競合サイトDA取得（既に取得済みのはず）
  Logger.log('【ステップ5】競合サイトDA取得');
  Logger.log('');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('競合分析');
  const data = sheet.getRange(2, 1, keywordCount, 27).getValues();
  
  // 競合URLを収集
  const competitorUrls = [];
  data.forEach(row => {
    for (let i = 6; i <= 24; i += 2) {
      const url = row[i];
      if (url) {
        competitorUrls.push(url);
      }
    }
  });
  
  Logger.log(`競合URL数: ${competitorUrls.length}`);
  
  // DAを取得（スマートキャッシング）
  const daMap = fetchDAWithSmartCaching(competitorUrls);
  
  Logger.log('');
  Logger.log('競合サイトDAをシートに書き込み中...');
  Logger.log('');
  
  // 競合DAを書き込み
  data.forEach((row, index) => {
    const rowIndex = index + 2;
    
    for (let i = 0; i < 10; i++) {
      const urlIndex = 6 + (i * 2);
      const daColNumber = 8 + (i * 2);
      const url = row[urlIndex];
      
      if (url) {
        const domain = extractDomain(url);
        const daData = daMap[domain];
        
        if (daData) {
          sheet.getRange(rowIndex, daColNumber).setValue(daData.da);
        }
      }
    }
  });
  
  Logger.log('✓ 競合DA書き込み完了');
  Logger.log('');
  
  // ステップ6: 勝算度スコア算出
  Logger.log('【ステップ6】勝算度スコア算出');
  Logger.log('');
  updateWinnableScores(2, keywordCount + 1);
  
  Logger.log('');
  Logger.log('===========================================');
  Logger.log('ターゲットKWベース完全競合分析完了');
  Logger.log('===========================================');
  Logger.log(`処理キーワード数: ${keywordCount} 件`);
  Logger.log('');
  Logger.log('競合分析シートを確認してください');
}

/**
 * テスト実行: ターゲットKW分析の最初の5件で競合分析
 */
function testTargetKWCompetitorAnalysis() {
  Logger.log('=== ターゲットKWベース競合分析テスト（5件） ===');
  Logger.log('');
  
  runCompetitorAnalysisForTargetKW(5);
  
  Logger.log('');
  Logger.log('=== テスト完了 ===');
}

/**
 * 本番実行: ターゲットKW分析の全213件で競合分析（フェーズ6用）
 */
function runFullTargetKWCompetitorAnalysis() {
  Logger.log('=== 全213件の競合分析実行 ===');
  Logger.log('');
  Logger.log('⚠ 注意: この処理には約30-40分かかります');
  Logger.log('⚠ DataForSEO API: 約$1.28');
  Logger.log('⚠ Moz API: 約800-900 rows使用');
  Logger.log('');
  
  // 確認メッセージ
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '全213件の競合分析実行',
    '約30-40分かかります。実行しますか？',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    runCompetitorAnalysisForTargetKW(213);
    
    Logger.log('');
    Logger.log('=== 全213件完了 ===');
  } else {
    Logger.log('キャンセルされました');
  }
}