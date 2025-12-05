/**
 * AIO分析チャット・週次レポート統合
 * Day 16実装
 * 
 * 【追加場所】
 * 1. WebApp.gs に isAIOAnalysisRequest, handleAIOAnalysisChat 等を追加
 * 2. handleChatMessage の競合分析判定の後に AIO判定を追加
 * 3. Scoring.gs の generateWeeklyReport を更新
 * 
 * 作成日: 2025/12/02
 */

// ========================================
// WebApp.gs に追加するコード
// ========================================

/**
 * メッセージがAIO分析リクエストかどうかを判定
 * @param {string} message - ユーザーメッセージ
 * @return {boolean} AIO分析リクエストかどうか
 */
function isAIOAnalysisRequest(message) {
  var aioKeywords = [
    'aio', 'AI Overview', 'AIオーバービュー', 
    'AI概要', 'AI引用', 'AIO対策', 'AIO順位',
    'AIに引用', 'AIで表示', 'AI表示'
  ];
  
  var lowerMessage = message.toLowerCase();
  
  for (var i = 0; i < aioKeywords.length; i++) {
    if (lowerMessage.includes(aioKeywords[i].toLowerCase())) {
      return true;
    }
  }
  
  return false;
}

/**
 * AIO分析リクエストを処理
 * @param {string} message - ユーザーメッセージ
 * @return {string} レスポンス
 */
function handleAIOAnalysisChat(message) {
  try {
    Logger.log('AIO分析リクエスト処理開始');
    
    // ページURL抽出
    var pageUrl = extractPageUrl(message);
    
    // キーワード抽出
    var keyword = extractKeywordFromMessage(message);
    
    Logger.log('抽出URL: ' + pageUrl);
    Logger.log('抽出キーワード: ' + keyword);
    
    var response = '';
    
    // 1. AIO順位履歴データを取得
    var aioHistory = getAIOHistoryForChat(pageUrl, keyword);
    
    // 2. AIO順位履歴がある場合
    if (aioHistory && aioHistory.length > 0) {
      response += formatAIOHistoryForChat(aioHistory);
      response += '\n---\n\n';
    }
    
    // 3. 特定ページが指定されている場合、AIO最適化提案を生成
    if (pageUrl) {
      response += '## 📝 AIO最適化提案\n\n';
      try {
        // Scoring.gs の generateAIOSuggestion を呼び出し
        var suggestion = generateAIOSuggestion(pageUrl);
        response += suggestion;
      } catch (e) {
        Logger.log('AIO提案生成エラー: ' + e.message);
        response += 'AIO提案の生成中にエラーが発生しました: ' + e.message;
      }
    } else if (keyword) {
      // キーワード指定の場合、そのキーワードのAIO状況を表示
      response += formatAIOKeywordAnalysis(keyword, aioHistory);
    } else {
      // 何も指定がない場合、AIO全体サマリーを表示
      response += formatAIOOverallSummary();
    }
    
    return response;
    
  } catch (error) {
    Logger.log('AIO分析エラー: ' + error.message);
    return 'AIO分析中にエラーが発生しました: ' + error.message;
  }
}

/**
 * メッセージからキーワードを抽出
 * @param {string} message - ユーザーメッセージ
 * @return {string|null} 抽出されたキーワード
 */
function extractKeywordFromMessage(message) {
  // 「〜のAIO」「〜でAIO」「〜はAIO」パターン
  var patterns = [
    /「(.+?)」.*(?:aio|AI)/i,
    /(.+?)(?:の|で|は).*(?:aio|AI)/i
  ];
  
  for (var i = 0; i < patterns.length; i++) {
    var match = message.match(patterns[i]);
    if (match && match[1]) {
      var keyword = match[1].trim();
      // 不要な文字を除去
      keyword = keyword.replace(/^(の|で|は|を|が|に)/, '').trim();
      if (keyword.length > 0 && keyword.length < 50) {
        return keyword;
      }
    }
  }
  
  return null;
}

/**
 * AIO順位履歴をチャット用に取得
 * @param {string} pageUrl - ページURL（オプション）
 * @param {string} keyword - キーワード（オプション）
 * @return {Array} AIO履歴データ
 */
function getAIOHistoryForChat(pageUrl, keyword) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('AIO順位履歴');
    
    if (!sheet) {
      Logger.log('AIO順位履歴シートが見つかりません');
      return [];
    }
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    
    var keywordIdx = headers.indexOf('keyword');
    var hasAioIdx = headers.indexOf('has_aio');
    var ownSiteIdx = headers.indexOf('own_site_in_aio');
    var positionIdx = headers.indexOf('aio_reference_position');
    var urlIdx = headers.indexOf('referenced_url');
    var checkDateIdx = headers.indexOf('check_date');
    var changeIdx = headers.indexOf('aio_position_change');
    
    var results = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var rowKeyword = row[keywordIdx] || '';
      var rowUrl = row[urlIdx] || '';
      
      // フィルタリング
      var match = false;
      if (keyword && rowKeyword.toLowerCase().includes(keyword.toLowerCase())) {
        match = true;
      }
      if (pageUrl && rowUrl.toLowerCase().includes(pageUrl.toLowerCase())) {
        match = true;
      }
      if (!keyword && !pageUrl) {
        match = true; // 全件
      }
      
      if (match) {
        results.push({
          keyword: rowKeyword,
          hasAIO: row[hasAioIdx],
          ownSiteInAIO: row[ownSiteIdx],
          position: row[positionIdx],
          url: rowUrl,
          checkDate: row[checkDateIdx],
          change: row[changeIdx]
        });
      }
    }
    
    // 日付の新しい順にソート
    results.sort(function(a, b) {
      return new Date(b.checkDate) - new Date(a.checkDate);
    });
    
    Logger.log('AIO履歴取得: ' + results.length + '件');
    return results;
    
  } catch (error) {
    Logger.log('AIO履歴取得エラー: ' + error.message);
    return [];
  }
}

/**
 * AIO履歴をチャット用にフォーマット
 * @param {Array} history - AIO履歴データ
 * @return {string} フォーマットされた文字列
 */
function formatAIOHistoryForChat(history) {
  var response = '## 📊 AIO順位履歴\n\n';
  
  // 自社引用ありのキーワードをピックアップ
  var ownSiteKeywords = history.filter(function(h) { return h.ownSiteInAIO === true; });
  
  if (ownSiteKeywords.length > 0) {
    response += '### ⭐ 自社サイトがAIOに引用されているキーワード\n\n';
    response += '| キーワード | 引用順位 | 変動 | 確認日 |\n';
    response += '|------------|----------|------|--------|\n';
    
    var shown = {};
    for (var i = 0; i < ownSiteKeywords.length && Object.keys(shown).length < 5; i++) {
      var h = ownSiteKeywords[i];
      if (!shown[h.keyword]) {
        var dateStr = '';
        if (h.checkDate) {
          var d = new Date(h.checkDate);
          dateStr = (d.getMonth() + 1) + '/' + d.getDate();
        }
        response += '| ' + h.keyword + ' | ' + h.position + '位 | ' + (h.change || '-') + ' | ' + dateStr + ' |\n';
        shown[h.keyword] = true;
      }
    }
    response += '\n';
  }
  
  // AIOあり（自社引用なし）のキーワード
  var aioKeywords = history.filter(function(h) { return h.hasAIO === true && h.ownSiteInAIO !== true; });
  
  if (aioKeywords.length > 0) {
    response += '### 📋 AIO表示あり（自社引用なし）\n\n';
    
    var shown = {};
    var count = 0;
    for (var i = 0; i < aioKeywords.length && count < 5; i++) {
      var h = aioKeywords[i];
      if (!shown[h.keyword]) {
        response += '- ' + h.keyword + '\n';
        shown[h.keyword] = true;
        count++;
      }
    }
    response += '\n';
  }
  
  return response;
}

/**
 * キーワード指定のAIO分析をフォーマット
 * @param {string} keyword - キーワード
 * @param {Array} history - AIO履歴データ
 * @return {string} フォーマットされた文字列
 */
function formatAIOKeywordAnalysis(keyword, history) {
  var response = '## 🔍 「' + keyword + '」のAIO分析\n\n';
  
  var keywordHistory = history.filter(function(h) {
    return h.keyword.toLowerCase().includes(keyword.toLowerCase());
  });
  
  if (keywordHistory.length === 0) {
    response += 'このキーワードのAIO履歴データはまだありません。\n\n';
    response += '週次の競合分析実行後にデータが蓄積されます。';
    return response;
  }
  
  var latest = keywordHistory[0];
  
  response += '**AIO表示**: ' + (latest.hasAIO ? 'あり ✅' : 'なし') + '\n';
  response += '**自社引用**: ' + (latest.ownSiteInAIO ? latest.position + '位 ⭐' : 'なし') + '\n';
  
  if (latest.change) {
    response += '**順位変動**: ' + latest.change + '\n';
  }
  
  if (latest.url) {
    response += '**引用URL**: ' + latest.url + '\n';
  }
  
  response += '\n### 💡 AIO対策のポイント\n\n';
  
  if (latest.ownSiteInAIO) {
    response += '✅ 自社サイトがAIOに引用されています！\n';
    response += '- 現在の順位を維持・向上させましょう\n';
    response += '- コンテンツの鮮度を保つことが重要です\n';
  } else if (latest.hasAIO) {
    response += '⚠️ AIOが表示されていますが、自社サイトは引用されていません。\n\n';
    response += '**改善案**:\n';
    response += '1. 質問に対する明確な回答を冒頭に配置\n';
    response += '2. FAQスキーマの実装を検討\n';
    response += '3. 簡潔で読みやすい文章構造に改善\n';
  } else {
    response += 'このキーワードではAIOが表示されていません。\n';
    response += '通常のSEO対策に注力しましょう。';
  }
  
  return response;
}

/**
 * AIO全体サマリーをフォーマット
 * @return {string} フォーマットされた文字列
 */
function formatAIOOverallSummary() {
  var response = '## 📈 AIO分析サマリー\n\n';
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('AIO順位履歴');
    
    if (!sheet || sheet.getLastRow() <= 1) {
      response += 'AIO順位履歴データがまだありません。\n\n';
      response += '週次の競合分析を実行すると、AIOデータが蓄積されます。';
      return response;
    }
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    
    var keywordIdx = headers.indexOf('keyword');
    var hasAioIdx = headers.indexOf('has_aio');
    var ownSiteIdx = headers.indexOf('own_site_in_aio');
    
    // ユニークキーワードを集計
    var keywordStats = {};
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var kw = row[keywordIdx];
      
      if (!keywordStats[kw]) {
        keywordStats[kw] = {
          hasAIO: row[hasAioIdx],
          ownSiteInAIO: row[ownSiteIdx]
        };
      }
    }
    
    var totalKeywords = Object.keys(keywordStats).length;
    var aioKeywords = 0;
    var ownSiteKeywords = 0;
    
    for (var kw in keywordStats) {
      if (keywordStats[kw].hasAIO) aioKeywords++;
      if (keywordStats[kw].ownSiteInAIO) ownSiteKeywords++;
    }
    
    response += '### 📊 統計情報\n\n';
    response += '- 追跡キーワード数: ' + totalKeywords + '件\n';
    response += '- AIO表示あり: ' + aioKeywords + '件 (' + Math.round(aioKeywords / totalKeywords * 100) + '%)\n';
    response += '- 自社引用あり: ' + ownSiteKeywords + '件 (' + Math.round(ownSiteKeywords / totalKeywords * 100) + '%)\n\n';
    
    response += '### 🎯 使い方\n\n';
    response += '特定のキーワードやページのAIO状況を確認するには:\n';
    response += '- 「iPhone 保険 のAIO状況は？」\n';
    response += '- 「/iphonerepair-screen のAIO対策を提案して」\n';
    
  } catch (error) {
    response += 'データ取得中にエラーが発生しました: ' + error.message;
  }
  
  return response;
}


// ========================================
// handleChatMessage への追加コード
// ========================================

/*
 * handleChatMessage 関数内、競合分析判定の後に以下を追加:
 * 
 * // ========================================
 * // 優先度0.5: AIO分析リクエスト（Day 16追加）
 * // ========================================
 * if (isAIOAnalysisRequest(userMessage)) {
 *   Logger.log('AIO分析リクエストを検出');
 *   return handleAIOAnalysisChat(userMessage);
 * }
 */


// ========================================
// Scoring.gs に追加するコード（週次レポート用）
// ========================================

/**
 * 週次レポート生成（AIO情報追加版）
 * ※既存の generateWeeklyReport を置き換え
 */
function generateWeeklyReportWithAIO(topPages) {
  var now = new Date();
  var dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy/MM/dd');
  
  var report = '【SEOリライト支援ツール】週次レポート\n';
  report += 'レポート日時: ' + dateStr + '\n\n';
  report += '=== 今週リライトすべきページ TOP10 ===\n\n';
  
  topPages.forEach(function(page, index) {
    report += (index + 1) + '位: ' + page.url + '\n';
    report += '   スコア: ' + page.score + '点\n\n';
  });
  
  // AIOサマリーを追加（Day 16）
  report += '\n=== AIO（AI Overview）状況 ===\n\n';
  report += getAIOSummaryForReport();
  
  report += '\n詳細は統合データシートをご確認ください。\n';
  
  return report;
}

/**
 * 週次レポート用AIOサマリーを取得
 * @return {string} AIOサマリー文字列
 */
function getAIOSummaryForReport() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('AIO順位履歴');
    
    if (!sheet || sheet.getLastRow() <= 1) {
      return 'AIOデータなし（次回競合分析実行後に表示）\n';
    }
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    
    var keywordIdx = headers.indexOf('keyword');
    var hasAioIdx = headers.indexOf('has_aio');
    var ownSiteIdx = headers.indexOf('own_site_in_aio');
    var positionIdx = headers.indexOf('aio_reference_position');
    var checkDateIdx = headers.indexOf('check_date');
    
    // 過去7日間のデータをフィルタ
    var oneWeekAgo = new Date();
    oneWeekAgo.setDate(oneWeekAgo.getDate() - 7);
    
    var recentData = [];
    for (var i = 1; i < data.length; i++) {
      var checkDate = new Date(data[i][checkDateIdx]);
      if (checkDate >= oneWeekAgo) {
        recentData.push({
          keyword: data[i][keywordIdx],
          hasAIO: data[i][hasAioIdx],
          ownSiteInAIO: data[i][ownSiteIdx],
          position: data[i][positionIdx]
        });
      }
    }
    
    if (recentData.length === 0) {
      return '過去7日間のAIOデータなし\n';
    }
    
    // 集計（キーワード単位で最新のみ）
    var keywordStats = {};
    for (var i = 0; i < recentData.length; i++) {
      var d = recentData[i];
      if (!keywordStats[d.keyword]) {
        keywordStats[d.keyword] = d;
      }
    }
    
    var totalKeywords = Object.keys(keywordStats).length;
    var aioKeywords = 0;
    var ownSiteKeywords = [];
    
    for (var kw in keywordStats) {
      if (keywordStats[kw].hasAIO) aioKeywords++;
      if (keywordStats[kw].ownSiteInAIO) {
        ownSiteKeywords.push({
          keyword: kw,
          position: keywordStats[kw].position
        });
      }
    }
    
    var summary = '';
    summary += '追跡KW数: ' + totalKeywords + '件\n';
    summary += 'AIO表示: ' + aioKeywords + '件 (' + Math.round(aioKeywords / totalKeywords * 100) + '%)\n';
    summary += '自社引用: ' + ownSiteKeywords.length + '件\n\n';
    
    if (ownSiteKeywords.length > 0) {
      summary += '【自社引用キーワード】\n';
      // 上位5件を表示（順位順）
      ownSiteKeywords.sort(function(a, b) { return a.position - b.position; });
      for (var i = 0; i < Math.min(ownSiteKeywords.length, 5); i++) {
        var kw = ownSiteKeywords[i];
        summary += '- ' + kw.keyword + ' (' + kw.position + '位)\n';
      }
    }
    
    return summary;
    
  } catch (error) {
    Logger.log('AIOサマリー取得エラー: ' + error.message);
    return 'AIOサマリー取得エラー\n';
  }
}


// ========================================
// テスト関数
// ========================================

/**
 * AIO分析チャットのテスト
 */
function testAIOAnalysisChat() {
  Logger.log('=== AIO分析チャットテスト ===');
  
  // テスト1: AIOリクエスト判定
  var testMessages = [
    'AIOの状況を教えて',
    'iPhone 保険 のAIO対策を提案して',
    '/iphonerepair-screen のAIO順位は？',
    'AI Overviewに引用されているキーワードは？',
    '今日の天気は？'
  ];
  
  Logger.log('\n--- リクエスト判定テスト ---');
  for (var i = 0; i < testMessages.length; i++) {
    var msg = testMessages[i];
    var isAIO = isAIOAnalysisRequest(msg);
    Logger.log('「' + msg + '」→ AIO分析: ' + isAIO);
  }
  
  // テスト2: AIO全体サマリー
  Logger.log('\n--- AIOサマリーテスト ---');
  var summary = formatAIOOverallSummary();
  Logger.log(summary);
  
  // テスト3: 週次レポート用サマリー
  Logger.log('\n--- 週次レポート用サマリーテスト ---');
  var reportSummary = getAIOSummaryForReport();
  Logger.log(reportSummary);
  
  Logger.log('\n=== テスト完了 ===');
}

/**
 * handleChatMessageからAIO分析が呼ばれるかのテスト
 */
function testAIOChatIntegration() {
  Logger.log('=== AIOチャット統合テスト ===');
  
  var testMessage = 'AIOの状況を教えて';
  
  // isAIOAnalysisRequest が true を返すか確認
  var isAIO = isAIOAnalysisRequest(testMessage);
  Logger.log('リクエスト判定: ' + isAIO);
  
  if (isAIO) {
    var result = handleAIOAnalysisChat(testMessage);
    Logger.log('\n--- 結果 ---');
    Logger.log(result);
  }
  
  Logger.log('\n=== テスト完了 ===');
}
