/**
 * SEOリライト支援ツール - WebApp.gs
 * Day 15完全版: 競合分析チャット機能追加
 * 最終更新: 2025年12月1日
 */

// ========================================
// Webアプリ基本設定
// ========================================

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('SEOリライト支援ツール')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ========================================
// メインチャット処理（Day 15更新）
// ========================================

function handleChatMessage(userMessage) {
  try {
    Logger.log('=== チャットメッセージ処理開始 ===');
    Logger.log('ユーザーメッセージ: ' + userMessage);
    
    if (!userMessage || userMessage.trim() === '') {
      return 'メッセージが空です';
    }
    
    // ========================================
    // 優先度0: 競合分析リクエスト（Day 15追加）
    // ========================================
    if (isCompetitorAnalysisRequest(userMessage)) {
      Logger.log('競合分析リクエストを検出');
      return handleCompetitorAnalysisChat(userMessage, null);
    }

     // ========================================
    // 優先度0.5: AIO分析リクエスト（Day 16追加）
    // ========================================
    if (isAIOAnalysisRequest(userMessage)) {
      Logger.log('AIO分析リクエストを検出');
      return handleAIOAnalysisChat(userMessage);
    }
    
    // ========================================
    // 優先度0.6: GSCズレ分析リクエスト（Day 22追加）
    // ========================================
    if (isGSCGapAnalysisRequest(userMessage)) {
      Logger.log('GSCズレ分析リクエストを検出');
      return handleGSCGapAnalysisFromChat(userMessage);
    }
    
    // 意図分析
    var intentData = analyzeIntent(userMessage);
    Logger.log('意図: ' + intentData.intent);
    
    var data = null;
    var contextPrompt = userMessage + '\n\n';
    
    // リライト提案クエリの場合
    if (intentData.intent === 'rewrite_suggestion_query') {
      Logger.log('リライト提案モード');
      
      var pageUrl = intentData.pageUrl;
      
      if (pageUrl) {
        // 特定ページの提案
        Logger.log('特定ページの提案: ' + pageUrl);
        var result = generateRewriteSuggestions(pageUrl);
        
        if (result.success) {
          var suggestionResponse = result.suggestion;
          // 推奨KW情報を追加
          var kwInfo = getRecommendedKeywordForChat(pageUrl);
          if (kwInfo) {
            suggestionResponse += kwInfo;
          }
          return suggestionResponse;
        } else {
          return 'ページが見つかりませんでした: ' + pageUrl;
        }
      } else {
        // 優先度上位ページの提案
        Logger.log('優先度上位ページの提案');
        var topPages = getTopPriorityPagesFiltered(5);
        
        if (topPages.length === 0) {
          return '対象ページが見つかりませんでした。スコアリングを実行してください。';
        }
        
        // 上位1ページの詳細提案を生成
        var result = generateRewriteSuggestions(topPages[0].url);
        
        var response = '【リライト優先度上位5ページ】\n\n';
        
        for (var i = 0; i < topPages.length; i++) {
          var page = topPages[i];
          response += (i + 1) + '. ' + page.url + '\n';
          response += '   総合スコア: ' + page.totalScore + '点\n';
          response += '   (機会損失: ' + page.opportunityScore + ' / パフォーマンス: ' + page.performanceScore + ' / ビジネス: ' + page.businessImpactScore + ')\n\n';
        }
        
        response += '---\n\n';
        response += '【1位ページの詳細リライト提案】\n\n';
        
        if (result.success) {
          response += result.suggestion;
          // 1位ページの推奨KW情報を追加
          var kwInfo = getRecommendedKeywordForChat(topPages[0].url);
          if (kwInfo) {
            response += kwInfo;
          }
        } else {
          response += '提案生成中にエラーが発生しました。';
        }
        
        return response;
      }
    }
    
    // リライト効果測定の場合
    if (intentData.intent === 'rewrite_effect_query') {
      Logger.log('リライト効果測定モード');
      
      var rewriteDate = intentData.rewriteDate;
      var comparisonDays = intentData.comparisonDays;
      var pageUrl = intentData.pageUrl;
      
      Logger.log('リライト日: ' + rewriteDate);
      Logger.log('比較期間: ' + comparisonDays + '日間');
      if (pageUrl) Logger.log('対象URL: ' + pageUrl);
      
      // Before/After期間を計算
      var periods = calculateBeforeAfterPeriods(rewriteDate, comparisonDays);
      
      Logger.log('Before期間: ' + periods.before.start + ' 〜 ' + periods.before.end);
      Logger.log('After期間: ' + periods.after.start + ' 〜 ' + periods.after.end);
      
      // データ取得
      var beforeData = fetchIntegratedDataForDateRange(
        periods.before.start, 
        periods.before.end, 
        pageUrl
      );
      
      var afterData = fetchIntegratedDataForDateRange(
        periods.after.start, 
        periods.after.end, 
        pageUrl
      );
      
      Logger.log('Beforeデータ: ' + beforeData.length + '件');
      Logger.log('Afterデータ: ' + afterData.length + '件');
      
      // Before/After統計計算
      var beforeStats = calculatePeriodStats(beforeData);
      var afterStats = calculatePeriodStats(afterData);
      
      // コンテキストプロンプト構築
      contextPrompt = buildBeforeAfterPrompt(
        userMessage,
        rewriteDate,
        comparisonDays,
        beforeStats,
        afterStats,
        pageUrl
      );
      
    } else if (intentData.intent === 'date_range_query') {
      // 期間指定クエリの場合
      Logger.log('期間指定クエリモード');
      
      var startDate = intentData.startDate;
      var endDate = intentData.endDate;
      var subIntent = intentData.subIntent || 'general';
      
      Logger.log('期間: ' + startDate + ' 〜 ' + endDate);
      Logger.log('サブ意図: ' + subIntent);
      
      // データ取得
      data = fetchIntegratedDataForDateRange(startDate, endDate);
      
      Logger.log('データ取得完了: ' + (data ? data.length : 0) + '件');
      
      // サブ意図に応じてデータを整形
      var tempIntentData = {
        intent: subIntent,
        needsData: true,
        dataType: 'integrated'
      };
      
      // 期間情報を含めたメッセージ
      var messageWithPeriod = userMessage + '\n\n【分析期間】' + startDate + ' 〜 ' + endDate + '\n\n';
      
      contextPrompt = buildContextPrompt(messageWithPeriod, tempIntentData, data);
      
    } else if (intentData.needsData) {
      // 通常のデータ参照クエリ
      Logger.log('通常のデータ参照モード');
      
      if (intentData.dataType === 'top_pages') {
        data = getTopPages(10);
      } else {
        data = getIntegratedData();
      }
      
      Logger.log('データ取得完了: ' + (data ? data.length : 0) + '件');
      
      contextPrompt = buildContextPrompt(userMessage, intentData, data);
    }
    
    var systemPrompt = buildSystemPrompt();
    
    Logger.log('プロンプト長: ' + contextPrompt.length + '文字');
    
    var response = callClaudeAPI(contextPrompt, systemPrompt);
    
    Logger.log('Claude応答取得成功');
    Logger.log('=== 処理完了 ===');
    
    return response;
    
  } catch (error) {
    Logger.log('=== エラー発生 ===');
    Logger.log('エラー: ' + error.message);
    Logger.log('スタック: ' + error.stack);
    
    return 'エラーが発生しました: ' + error.message;
  }
}

// ========================================
// 意図分析（Day 5修正版）
// ========================================

function analyzeIntent(userMessage) {
  var message = userMessage.toLowerCase();
  
  // =============================================
  // 優先度0: リライト提案クエリ（最優先）
  // =============================================
  
  var isSuggestionQuery = message.includes('リライト') && 
                          (message.includes('提案') || message.includes('おすすめ') || 
                           message.includes('候補') || message.includes('すべき'));
  
  var isTopPagesQuery = message.includes('優先') || message.includes('上位') ||
                        message.includes('どのページ') || message.includes('どの記事');
  
  if (isSuggestionQuery || (isTopPagesQuery && message.includes('リライト'))) {
    var specifiedUrl = extractPageUrl(userMessage);
    return {
      intent: 'rewrite_suggestion_query',
      needsData: true,
      dataType: 'suggestion',
      pageUrl: specifiedUrl
    };
  }
  
  // =============================================
  // 優先度1: リライト効果測定クエリ
  // =============================================
  
  var hasRewriteKeyword = message.includes('リライト') && 
                           (message.includes('効果') || message.includes('測定') || 
                            message.includes('比較') || message.includes('前後'));
  
  if (hasRewriteKeyword) {
    var rewriteDate = extractRewriteDate(userMessage);
    var comparisonDays = extractComparisonDays(userMessage);
    var pageUrl = extractPageUrl(userMessage);
    
    return {
      intent: 'rewrite_effect_query',
      needsData: true,
      dataType: 'date_range',
      rewriteDate: rewriteDate,
      comparisonDays: comparisonDays,
      pageUrl: pageUrl
    };
  }
  
  // =============================================
  // 優先度2: 期間指定の通常クエリ
  // =============================================
  
  var hasDateRange = message.includes('過去') || message.includes('先週') || 
                      message.includes('先月') || message.includes('今月');
  
  if (hasDateRange) {
    var dateRange = extractDateRange(userMessage);
    
    if (dateRange) {
      // 質問の種類を判定
      var queryIntent = 'general';
      
      if (message.includes('直帰率') || message.includes('離脱率')) {
        queryIntent = 'bounce_query';
      } else if (message.includes('ページビュー') || message.includes('pv') || 
                 (message.includes('アクセス') && message.includes('多い'))) {
        queryIntent = 'traffic_query';
      } else if (message.includes('ctr') || message.includes('クリック率')) {
        queryIntent = 'ctr_query';
      } else if (message.includes('順位') && (message.includes('低い') || message.includes('悪い'))) {
        queryIntent = 'ranking_query';
      } else if (message.includes('改善') || message.includes('どのページ')) {
        queryIntent = 'improvement_query';
      } else if (message.includes('トップ') || message.includes('ベスト') || 
                 (message.includes('多い') && message.includes('上位'))) {
        queryIntent = 'traffic_query';
      }
      
      return {
        intent: 'date_range_query',
        needsData: true,
        dataType: 'date_range',
        startDate: dateRange.start,
        endDate: dateRange.end,
        subIntent: queryIntent
      };
    }
  }
  
  // =============================================
  // 優先度3: 通常のデータ参照クエリ（期間指定なし）
  // =============================================
  
  if (message.includes('ページビュー') || message.includes('pv') || 
      (message.includes('アクセス') && message.includes('多い')) ||
      (message.includes('訪問') && message.includes('多い'))) {
    return { intent: 'traffic_query', needsData: true, dataType: 'integrated' };
  }
  
  if (message.includes('ctr') || message.includes('クリック率') || 
      (message.includes('クリック') && (message.includes('低い') || message.includes('少ない')))) {
    return { intent: 'ctr_query', needsData: true, dataType: 'integrated' };
  }
  
  if (message.includes('直帰率') || message.includes('離脱率') || 
      (message.includes('直帰') && !message.includes('検索'))) {
    return { intent: 'bounce_query', needsData: true, dataType: 'integrated' };
  }
  
  if (message.includes('改善すべき') || message.includes('優先度')) {
    return { intent: 'improvement_query', needsData: true, dataType: 'integrated' };
  }
  
  if ((message.includes('順位') && (message.includes('低い') || message.includes('悪い'))) ||
      (message.includes('ランキング') && message.includes('下位'))) {
    return { intent: 'ranking_query', needsData: true, dataType: 'integrated' };
  }
  
  if ((message.includes('トップ') || message.includes('ベスト') || 
       message.includes('人気') || (message.includes('多い') && message.includes('上位'))) && 
      !message.includes('順位') && !message.includes('ランキング')) {
    return { intent: 'top_pages_query', needsData: true, dataType: 'top_pages' };
  }
  
  return { intent: 'general_seo', needsData: false, dataType: null };
}

// ========================================
// 日付抽出関数
// ========================================

/**
 * リライト日を抽出
 */
function extractRewriteDate(userMessage) {
  var today = new Date();
  
  // 「5月15日」形式
  var match = userMessage.match(/(\d{1,2})月(\d{1,2})日/);
  if (match) {
    var month = parseInt(match[1]);
    var day = parseInt(match[2]);
    var year = today.getFullYear();
    
    // 未来の日付の場合は前年
    var date = new Date(year, month - 1, day);
    if (date > today) {
      year--;
      date = new Date(year, month - 1, day);
    }
    
    return formatDate(date);
  }
  
  // 「2025-05-15」形式
  match = userMessage.match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (match) {
    return match[0];
  }
  
  // デフォルト: 30日前
  var defaultDate = new Date();
  defaultDate.setDate(defaultDate.getDate() - 30);
  return formatDate(defaultDate);
}

/**
 * 比較期間（日数）を抽出
 */
function extractComparisonDays(userMessage) {
  // 「7日間」「30日間」形式
  var match = userMessage.match(/(\d+)日間/);
  if (match) {
    return parseInt(match[1]);
  }
  
  // 「1週間」形式
  if (userMessage.includes('1週間') || userMessage.includes('一週間')) {
    return 7;
  }
  
  // 「2週間」形式
  if (userMessage.includes('2週間') || userMessage.includes('二週間')) {
    return 14;
  }
  
  // 「1ヶ月」「1か月」形式
  if (userMessage.includes('1ヶ月') || userMessage.includes('1か月') || 
      userMessage.includes('一ヶ月') || userMessage.includes('一か月')) {
    return 30;
  }
  
  // 「2ヶ月」形式
  if (userMessage.includes('2ヶ月') || userMessage.includes('2か月') || 
      userMessage.includes('二ヶ月') || userMessage.includes('二か月')) {
    return 60;
  }
  
  // デフォルト: 7日間
  return 7;
}

/**
 * ページURLを抽出（改善版 - 3パターン対応）
 * 対応形式:
 * 1. 完全URL: https://smaho-tap.com/iphonerepair-battery-replacement-makes-sense
 * 2. パス形式: /iphonerepair-battery-replacement-makes-sense
 * 3. スラッグ形式: iphonerepair-battery-replacement-makes-sense
 */
function extractPageUrl(userMessage) {
  Logger.log('=== extractPageUrl開始 ===');
  Logger.log('入力メッセージ: ' + userMessage);
  
  var extractedPath = null;
  
  // =============================================
  // パターン1: 完全URL（https://smaho-tap.com/xxx）
  // =============================================
  var fullUrlMatch = userMessage.match(/https?:\/\/[^\/\s]+\/([a-zA-Z0-9\-_\/]+)/);
  if (fullUrlMatch) {
    extractedPath = '/' + fullUrlMatch[1].replace(/\/$/, ''); // 末尾スラッシュ除去
    Logger.log('完全URLから抽出: ' + extractedPath);
    return extractedPath;
  }
  
  // =============================================
  // パターン2: パス形式（/xxx）
  // =============================================
  var pathMatch = userMessage.match(/\/([a-zA-Z0-9\-_]+(?:\/[a-zA-Z0-9\-_]+)*)/);
  if (pathMatch) {
    extractedPath = '/' + pathMatch[1].replace(/\/$/, '');
    Logger.log('パス形式から抽出: ' + extractedPath);
    return extractedPath;
  }
  
  // =============================================
  // パターン3: スラッグ形式（/なし）→ 統合データから検索
  // =============================================
  // 英字で始まり、英数字・ハイフン・アンダースコアで構成される6文字以上の文字列
  var slugMatch = userMessage.match(/\b([a-zA-Z][a-zA-Z0-9\-_]{5,})\b/);
  if (slugMatch) {
    var potentialSlug = slugMatch[1].toLowerCase();
    Logger.log('スラッグ候補: ' + potentialSlug);
    
    // 除外ワード（一般的な単語を誤検出しないように）
    var excludeWords = ['リライト', '提案して', 'ください', 'おねがい', 'analysis', 'please', 'rewrite'];
    var isExcluded = excludeWords.some(function(word) {
      return potentialSlug.indexOf(word.toLowerCase()) !== -1;
    });
    
    if (!isExcluded) {
      // 統合データから該当するURLを検索
      try {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var sheet = ss.getSheetByName('統合データ');
        
        if (sheet) {
          var data = sheet.getDataRange().getValues();
          var headers = data[0];
          var urlIdx = headers.indexOf('page_url');
          
          if (urlIdx !== -1) {
            // 完全一致を優先
            for (var i = 1; i < data.length; i++) {
              var rowUrl = (data[i][urlIdx] || '').toString();
              var normalizedRowUrl = rowUrl.toLowerCase().replace(/^\//, '').replace(/\/$/, '');
              
              if (normalizedRowUrl === potentialSlug) {
                extractedPath = rowUrl.startsWith('/') ? rowUrl : '/' + rowUrl;
                Logger.log('スラッグ完全一致: ' + potentialSlug + ' → ' + extractedPath);
                return extractedPath;
              }
            }
            
            // 部分一致（スラッグがURLに含まれる場合）
            for (var i = 1; i < data.length; i++) {
              var rowUrl = (data[i][urlIdx] || '').toString();
              var normalizedRowUrl = rowUrl.toLowerCase();
              
              if (normalizedRowUrl.indexOf(potentialSlug) !== -1) {
                extractedPath = rowUrl.startsWith('/') ? rowUrl : '/' + rowUrl;
                Logger.log('スラッグ部分一致: ' + potentialSlug + ' → ' + extractedPath);
                return extractedPath;
              }
            }
          }
        }
      } catch (error) {
        Logger.log('URL検索エラー: ' + error.message);
      }
    }
  }
  
  Logger.log('URL抽出失敗: null');
  return null;
}

/**
 * 期間範囲を抽出
 */
function extractDateRange(userMessage) {
  var today = new Date();
  var startDate, endDate;
  
  // 「過去7日間」「過去30日間」
  var match = userMessage.match(/過去(\d+)日間/);
  if (match) {
    var days = parseInt(match[1]);
    endDate = new Date();
    endDate.setDate(endDate.getDate() - 1); // 昨日まで
    startDate = new Date(endDate);
    startDate.setDate(startDate.getDate() - days + 1);
    
    return {
      start: formatDate(startDate),
      end: formatDate(endDate)
    };
  }
  
  // 「先週」
  if (userMessage.includes('先週')) {
    var lastWeekEnd = new Date(today);
    lastWeekEnd.setDate(lastWeekEnd.getDate() - today.getDay()); // 今週の日曜
    lastWeekEnd.setDate(lastWeekEnd.getDate() - 1); // 先週の土曜
    
    var lastWeekStart = new Date(lastWeekEnd);
    lastWeekStart.setDate(lastWeekStart.getDate() - 6); // 先週の日曜
    
    return {
      start: formatDate(lastWeekStart),
      end: formatDate(lastWeekEnd)
    };
  }
  
  // 「先月」
  if (userMessage.includes('先月')) {
    var lastMonthEnd = new Date(today.getFullYear(), today.getMonth(), 0);
    var lastMonthStart = new Date(today.getFullYear(), today.getMonth() - 1, 1);
    
    return {
      start: formatDate(lastMonthStart),
      end: formatDate(lastMonthEnd)
    };
  }
  
  // 「今月」
  if (userMessage.includes('今月')) {
    var thisMonthStart = new Date(today.getFullYear(), today.getMonth(), 1);
    var thisMonthEnd = new Date();
    thisMonthEnd.setDate(thisMonthEnd.getDate() - 1);
    
    return {
      start: formatDate(thisMonthStart),
      end: formatDate(thisMonthEnd)
    };
  }
  
  return null;
}

/**
 * Before/After期間を計算
 */
function calculateBeforeAfterPeriods(rewriteDate, days) {
  var rewrite = new Date(rewriteDate);
  
  // Before期間
  var beforeEnd = new Date(rewrite);
  beforeEnd.setDate(beforeEnd.getDate() - 1);
  
  var beforeStart = new Date(beforeEnd);
  beforeStart.setDate(beforeStart.getDate() - days + 1);
  
  // After期間
  var afterStart = new Date(rewrite);
  afterStart.setDate(afterStart.getDate() + 1);
  
  var afterEnd = new Date(afterStart);
  afterEnd.setDate(afterEnd.getDate() + days - 1);
  
  return {
    before: {
      start: formatDate(beforeStart),
      end: formatDate(beforeEnd)
    },
    after: {
      start: formatDate(afterStart),
      end: formatDate(afterEnd)
    }
  };
}

/**
 * 日付をYYYY-MM-DD形式に変換
 */
function formatDate(date) {
  var year = date.getFullYear();
  var month = String(date.getMonth() + 1).padStart(2, '0');
  var day = String(date.getDate()).padStart(2, '0');
  return year + '-' + month + '-' + day;
}

// ========================================
// Before/Afterプロンプト構築
// ========================================

function buildBeforeAfterPrompt(userMessage, rewriteDate, days, beforeStats, afterStats, pageUrl) {
  var prompt = userMessage + '\n\n';
  
  prompt += '【リライト効果測定】\n';
  prompt += 'リライト実施日: ' + rewriteDate + '\n';
  prompt += '比較期間: 前後' + days + '日間\n';
  if (pageUrl) {
    prompt += '対象ページ: ' + pageUrl + '\n';
  }
  prompt += '\n';
  
  prompt += '【Before（リライト前' + days + '日間）】\n';
  prompt += '- 総ページビュー: ' + beforeStats.total_page_views + '\n';
  prompt += '- 平均ページビュー/日: ' + beforeStats.avg_page_views.toFixed(1) + '\n';
  prompt += '- 平均直帰率: ' + beforeStats.avg_bounce_rate.toFixed(1) + '%\n';
  prompt += '- 平均セッション時間: ' + Math.round(beforeStats.avg_session_duration) + '秒\n';
  prompt += '- 総クリック数: ' + beforeStats.total_clicks + '\n';
  prompt += '- 総表示回数: ' + beforeStats.total_impressions + '\n';
  prompt += '- 平均検索順位: ' + beforeStats.avg_position.toFixed(1) + '位\n';
  prompt += '- 平均CTR: ' + beforeStats.avg_ctr.toFixed(2) + '%\n';
  prompt += '\n';
  
  prompt += '【After（リライト後' + days + '日間）】\n';
  prompt += '- 総ページビュー: ' + afterStats.total_page_views + '\n';
  prompt += '- 平均ページビュー/日: ' + afterStats.avg_page_views.toFixed(1) + '\n';
  prompt += '- 平均直帰率: ' + afterStats.avg_bounce_rate.toFixed(1) + '%\n';
  prompt += '- 平均セッション時間: ' + Math.round(afterStats.avg_session_duration) + '秒\n';
  prompt += '- 総クリック数: ' + afterStats.total_clicks + '\n';
  prompt += '- 総表示回数: ' + afterStats.total_impressions + '\n';
  prompt += '- 平均検索順位: ' + afterStats.avg_position.toFixed(1) + '位\n';
  prompt += '- 平均CTR: ' + afterStats.avg_ctr.toFixed(2) + '%\n';
  prompt += '\n';
  
  prompt += '【変化率】\n';
  
  // PV変化
  var pvChange = calculateChangeRate(beforeStats.avg_page_views, afterStats.avg_page_views);
  prompt += '- ページビュー: ' + pvChange.text + '\n';
  
  // 直帰率変化
  var bounceChange = calculateChangeRate(beforeStats.avg_bounce_rate, afterStats.avg_bounce_rate);
  prompt += '- 直帰率: ' + bounceChange.text + '\n';
  
  // セッション時間変化
  var durationChange = calculateChangeRate(beforeStats.avg_session_duration, afterStats.avg_session_duration);
  prompt += '- セッション時間: ' + durationChange.text + '\n';
  
  // クリック数変化
  var clicksChange = calculateChangeRate(beforeStats.total_clicks, afterStats.total_clicks);
  prompt += '- クリック数: ' + clicksChange.text + '\n';
  
  // 順位変化
  var positionDiff = beforeStats.avg_position - afterStats.avg_position;
  if (positionDiff > 0) {
    prompt += '- 検索順位: ' + positionDiff.toFixed(1) + '位上昇 ✅\n';
  } else if (positionDiff < 0) {
    prompt += '- 検索順位: ' + Math.abs(positionDiff).toFixed(1) + '位下降 ⚠️\n';
  } else {
    prompt += '- 検索順位: 変化なし\n';
  }
  
  // CTR変化
  var ctrChange = calculateChangeRate(beforeStats.avg_ctr, afterStats.avg_ctr);
  prompt += '- CTR: ' + ctrChange.text + '\n';
  
  Logger.log('=== Before/Afterプロンプト（先頭800文字） ===');
  Logger.log(prompt.substring(0, 800));
  Logger.log('=== プロンプト終了 ===');
  
  return prompt;
}

/**
 * 変化率を計算
 */
function calculateChangeRate(before, after) {
  if (before === 0) {
    if (after > 0) {
      return { rate: 100, text: '新規データ ✅' };
    } else {
      return { rate: 0, text: 'データなし' };
    }
  }
  
  var rate = ((after - before) / before) * 100;
  var direction = rate > 0 ? '増加' : '減少';
  var emoji = rate > 0 ? ' ✅' : ' ⚠️';
  
  if (Math.abs(rate) < 1) {
    return { rate: rate, text: 'ほぼ変化なし' };
  }
  
  return {
    rate: rate,
    text: Math.abs(rate).toFixed(1) + '% ' + direction + emoji
  };
}

// ========================================
// プロンプト構築
// ========================================

function buildSystemPrompt() {
  var prompt = 'あなたはSEO専門家のアシスタントです。\n\n' +
    '【役割】\n' +
    '- GA4とGoogle Search Consoleのデータに基づいた具体的な提案\n' +
    '- データから問題点を発見し、改善施策を提案\n\n' +
    '【回答スタイル】\n' +
    '1. 提供されたデータを必ず参照する\n' +
    '2. URLを明記して具体的なページを特定する\n' +
    '3. 数値データを活用した客観的な分析\n' +
    '4. 実行可能な改善施策を提案\n' +
    '5. 簡潔で分かりやすい表現\n\n' +
    '【重要】\n' +
    '- 参考データに記載されているURLを必ず回答に含めてください\n' +
    '- ページタイトルだけでなく、URLも明記してください';
  
  return prompt;
}

function buildContextPrompt(userMessage, intentData, data) {
  var contextPrompt = userMessage + '\n\n';
  
  if (!intentData.needsData || !data || data.length === 0) {
    Logger.log('データなしでプロンプト構築');
    return userMessage;
  }
  
  contextPrompt += '【参考データ】\n\n';
  
  var formattedData = '';
  
  switch (intentData.intent) {
    case 'ranking_query':
      formattedData = formatRankingData(data);
      break;
    case 'traffic_query':
      formattedData = formatTrafficData(data);
      break;
    case 'ctr_query':
      formattedData = formatCTRData(data);
      break;
    case 'bounce_query':
      formattedData = formatBounceData(data);
      break;
    case 'improvement_query':
      formattedData = formatImprovementData(data);
      break;
    case 'top_pages_query':
      formattedData = formatTopPagesData(data);
      break;
    case 'general':
      formattedData = formatTopPagesData(data);
      break;
    default:
      formattedData = formatTopPagesData(data);
  }
  
  contextPrompt += formattedData;
  
  Logger.log('=== 構築されたプロンプト（先頭800文字） ===');
  Logger.log(contextPrompt.substring(0, 800));
  Logger.log('=== プロンプト終了 ===');
  
  return contextPrompt;
}

// ========================================
// データ整形
// ========================================

function formatRankingData(data) {
  var formatted = '【検索順位データ】順位が低い（数字が大きい）ページ 上位5件\n\n';
  
  var sortedData = data
    .filter(function(row) { return row.avg_position && parseFloat(row.avg_position) > 0; })
    .sort(function(a, b) { return parseFloat(b.avg_position) - parseFloat(a.avg_position); });
  
  sortedData.slice(0, 5).forEach(function(row, index) {
    formatted += (index + 1) + '. ' + (row.page_title || 'タイトルなし') + '\n';
    formatted += '   URL: ' + (row.page_url || 'URLなし') + '\n';
    formatted += '   順位: ' + parseFloat(row.avg_position).toFixed(1) + '位\n';
    formatted += '   表示: ' + (row.total_impressions_30d || row.impressions || 0) + '回\n';
    formatted += '   CTR: ' + (row.avg_ctr ? (parseFloat(row.avg_ctr) * 100).toFixed(2) + '%' : (row.ctr ? (parseFloat(row.ctr) * 100).toFixed(2) + '%' : '-')) + '\n\n';
  });
  
  return formatted;
}

function formatTrafficData(data) {
  var formatted = '【ページビューデータ】PVが多い順 上位5件\n\n';
  
  var sortedData = data
    .filter(function(row) { 
      return (row.avg_page_views_30d && parseFloat(row.avg_page_views_30d) > 0) || 
             (row.page_views && parseFloat(row.page_views) > 0); 
    })
    .sort(function(a, b) { 
      var pvA = parseFloat(a.avg_page_views_30d || a.page_views || 0);
      var pvB = parseFloat(b.avg_page_views_30d || b.page_views || 0);
      return pvB - pvA;
    });
  
  sortedData.slice(0, 5).forEach(function(row, index) {
    var pv = row.avg_page_views_30d || row.page_views || 0;
    formatted += (index + 1) + '. ' + (row.page_title || 'タイトルなし') + '\n';
    formatted += '   URL: ' + (row.page_url || 'URLなし') + '\n';
    formatted += '   PV: ' + Math.round(parseFloat(pv)) + '\n';
    formatted += '   直帰率: ' + (row.bounce_rate ? parseFloat(row.bounce_rate).toFixed(1) + '%' : '-') + '\n';
    formatted += '   セッション時間: ' + (row.avg_session_duration ? Math.round(parseFloat(row.avg_session_duration)) + '秒' : '-') + '\n\n';
  });
  
  return formatted;
}

function formatCTRData(data) {
  var formatted = '【CTRデータ】CTRが低い順 上位5件\n\n';
  
  var sortedData = data
    .filter(function(row) { 
      var ctr = row.avg_ctr || row.ctr || 0;
      var impressions = row.total_impressions_30d || row.impressions || 0;
      return parseFloat(ctr) > 0 && impressions > 100; 
    })
    .sort(function(a, b) { 
      var ctrA = parseFloat(a.avg_ctr || a.ctr || 0);
      var ctrB = parseFloat(b.avg_ctr || b.ctr || 0);
      return ctrA - ctrB;
    });
  
  sortedData.slice(0, 5).forEach(function(row, index) {
    var ctr = row.avg_ctr || row.ctr || 0;
    var impressions = row.total_impressions_30d || row.impressions || 0;
    formatted += (index + 1) + '. ' + (row.page_title || 'タイトルなし') + '\n';
    formatted += '   URL: ' + (row.page_url || 'URLなし') + '\n';
    formatted += '   CTR: ' + (parseFloat(ctr) * 100).toFixed(2) + '%\n';
    formatted += '   順位: ' + (row.avg_position ? parseFloat(row.avg_position).toFixed(1) + '位' : '-') + '\n';
    formatted += '   表示: ' + impressions + '回\n\n';
  });
  
  return formatted;
}

function formatBounceData(data) {
  var formatted = '【直帰率データ】直帰率が高い順 上位5件\n\n';
  
  var sortedData = data
    .filter(function(row) { return row.bounce_rate && parseFloat(row.bounce_rate) > 0; })
    .sort(function(a, b) { return parseFloat(b.bounce_rate) - parseFloat(a.bounce_rate); });
  
  sortedData.slice(0, 5).forEach(function(row, index) {
    var pv = row.avg_page_views_30d || row.page_views || 0;
    formatted += (index + 1) + '. ' + (row.page_title || 'タイトルなし') + '\n';
    formatted += '   URL: ' + (row.page_url || 'URLなし') + '\n';
    formatted += '   直帰率: ' + parseFloat(row.bounce_rate).toFixed(1) + '%\n';
    formatted += '   PV: ' + Math.round(parseFloat(pv)) + '\n';
    formatted += '   セッション時間: ' + (row.avg_session_duration ? Math.round(parseFloat(row.avg_session_duration)) + '秒' : '-') + '\n\n';
  });
  
  return formatted;
}

function formatImprovementData(data) {
  var formatted = '【改善候補ページ】4-10位で表示回数が多い 上位5件\n\n';
  
  var candidates = data.filter(function(row) {
    var position = parseFloat(row.avg_position);
    var impressions = parseFloat(row.total_impressions_30d || row.impressions || 0);
    return position >= 4 && position <= 10 && impressions > 100;
  }).sort(function(a, b) { 
    var impA = parseFloat(a.total_impressions_30d || a.impressions || 0);
    var impB = parseFloat(b.total_impressions_30d || b.impressions || 0);
    return impB - impA;
  });
  
  if (candidates.length === 0) {
    return '※4-10位で表示回数が多いページは現在ありません\n';
  }
  
  candidates.slice(0, 5).forEach(function(row, index) {
    var impressions = row.total_impressions_30d || row.impressions || 0;
    var ctr = row.avg_ctr || row.ctr || 0;
    formatted += (index + 1) + '. ' + (row.page_title || 'タイトルなし') + '\n';
    formatted += '   URL: ' + (row.page_url || 'URLなし') + '\n';
    formatted += '   順位: ' + parseFloat(row.avg_position).toFixed(1) + '位\n';
    formatted += '   表示: ' + impressions + '回\n';
    formatted += '   CTR: ' + (parseFloat(ctr) * 100).toFixed(2) + '%\n\n';
  });
  
  return formatted;
}

function formatTopPagesData(data) {
  var formatted = '【上位パフォーマンスページ】PV順 上位5件\n\n';
  
  var sortedData = data
    .filter(function(row) { 
      return (row.avg_page_views_30d && parseFloat(row.avg_page_views_30d) > 0) || 
             (row.page_views && parseFloat(row.page_views) > 0); 
    })
    .sort(function(a, b) { 
      var pvA = parseFloat(a.avg_page_views_30d || a.page_views || 0);
      var pvB = parseFloat(b.avg_page_views_30d || b.page_views || 0);
      return pvB - pvA;
    });
  
  sortedData.slice(0, 5).forEach(function(row, index) {
    var pv = row.avg_page_views_30d || row.page_views || 0;
    formatted += (index + 1) + '. ' + (row.page_title || 'タイトルなし') + '\n';
    formatted += '   URL: ' + (row.page_url || 'URLなし') + '\n';
    formatted += '   PV: ' + Math.round(parseFloat(pv)) + '\n';
    formatted += '   順位: ' + (row.avg_position ? parseFloat(row.avg_position).toFixed(1) + '位' : '-') + '\n';
    formatted += '   直帰率: ' + (row.bounce_rate ? parseFloat(row.bounce_rate).toFixed(1) + '%' : '-') + '\n\n';
  });
  
  return formatted;
}

// ========================================
// データ取得
// ========================================

function getIntegratedData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('統合データ');
    
    if (!sheet) {
      throw new Error('統合データシートが見つかりません');
    }
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var rows = data.slice(1);
    
    var result = rows.map(function(row) {
      var obj = {};
      headers.forEach(function(header, index) {
        obj[header] = row[index];
      });
      return obj;
    });
    
    Logger.log('統合データ取得: ' + result.length + '件');
    
    if (result.length > 0) {
      Logger.log('サンプルデータのキー: ' + Object.keys(result[0]).join(', '));
    }
    
    return result;
    
  } catch (error) {
    Logger.log('統合データ取得エラー: ' + error.message);
    throw error;
  }
}

function getTopPages(n) {
  n = n || 10;
  
  try {
    var allData = getIntegratedData();
    
    var sorted = allData
      .filter(function(row) { return row['avg_page_views_30d'] && parseFloat(row['avg_page_views_30d']) > 0; })
      .sort(function(a, b) {
        var pvA = parseFloat(a['avg_page_views_30d']);
        var pvB = parseFloat(b['avg_page_views_30d']);
        return pvB - pvA;
      });
    
    var topPages = sorted.slice(0, n);
    
    Logger.log('上位' + n + 'ページ取得成功');
    return topPages;
    
  } catch (error) {
    Logger.log('上位ページ取得エラー: ' + error.message);
    throw error;
  }
}

// ========================================
// テスト関数
// ========================================

function testRewriteSuggestion() {
  var testMessage = "リライト提案して";
  
  Logger.log('=== リライト提案テスト ===');
  var result = handleChatMessage(testMessage);
  
  Logger.log('=== 結果 ===');
  Logger.log(result);
  
  return result;
}

function testRewriteEffect() {
  var testMessage = "10月15日にリライトした効果を前後7日間で測定して";
  
  Logger.log('=== リライト効果測定テスト ===');
  var result = handleChatMessage(testMessage);
  
  Logger.log('=== 結果 ===');
  Logger.log(result);
  
  return result;
}

function testDateRange() {
  var testMessage = "過去7日間で直帰率が高いページは？";
  
  Logger.log('=== 期間指定テスト ===');
  var result = handleChatMessage(testMessage);
  
  Logger.log('=== 結果 ===');
  Logger.log(result);
  
  return result;
}

// ============================================
// 競合分析チャット連携（Day 15追加）
// ============================================

/**
 * メッセージが競合分析リクエストかどうかを判定
 * @param {string} message - ユーザーメッセージ
 * @return {boolean} 競合分析リクエストかどうか
 */
function isCompetitorAnalysisRequest(message) {
  var competitorKeywords = [
    '競合', '上位サイト', '比較', 'ライバル', 
    '1位', '2位', '3位', 'トップ', '検索結果',
    'DA', 'ドメインオーソリティ', '勝算'
  ];
  
  var urlPattern = /https?:\/\/[^\s]+/;
  
  // URLが含まれていて分析を依頼している場合
  if (urlPattern.test(message) && (message.includes('分析') || message.includes('比較') || message.includes('調べ'))) {
    return true;
  }
  
  // 競合関連キーワードが含まれている場合
  for (var i = 0; i < competitorKeywords.length; i++) {
    if (message.includes(competitorKeywords[i])) {
      return true;
    }
  }
  
  return false;
}

/**
 * 競合分析リクエストを処理
 * @param {string} message - ユーザーメッセージ
 * @param {string} currentPageUrl - 現在選択中のページURL（あれば）
 * @return {string} レスポンス
 */
function handleCompetitorAnalysisChat(message, currentPageUrl) {
  try {
    // 意図分析
    var intent = analyzeCompetitorIntent(message);
    
    Logger.log('競合分析意図: ' + intent.type + ', アクション: ' + intent.action);
    
    var response = '';
    
    switch (intent.type) {
      case 'url_direct':
        // URL直接指定の場合
        response = handleUrlDirectAnalysis(intent.urls);
        break;
        
      case 'keyword_competitor':
        // キーワード指定の場合
        response = handleKeywordCompetitorChat(intent.keyword);
        break;
        
      case 'diff_analysis':
        // 差分分析の場合
        if (currentPageUrl) {
          response = handleDiffAnalysisChat(currentPageUrl);
        } else {
          response = '差分分析を行うには、まず分析対象のページを指定してください。\n\n例：「/iphone-insurance の競合と比較して」';
        }
        break;
        
      default:
        // 一般的な競合分析の質問
        response = handleGeneralCompetitorQuestion(message);
        break;
    }
    
    return response;
    
  } catch (error) {
    Logger.log('競合分析エラー: ' + error.message);
    return '競合分析中にエラーが発生しました: ' + error.message + '\n\nもう一度お試しください。';
  }
}

/**
 * URL直接指定の分析
 */
function handleUrlDirectAnalysis(urls) {
  if (!urls || urls.length === 0) {
    return 'URLが検出できませんでした。分析したいURLを指定してください。';
  }
  
  // 複数URLの場合は比較分析
  if (urls.length > 1) {
    var result = compareMultipleCompetitors(urls);
    return formatMultipleCompetitorResult(result);
  }
  
  // 単一URLの場合
  var content = fetchCompetitorContent(urls[0]);
  if (!content.success) {
    return 'URL「' + urls[0] + '」のコンテンツを取得できませんでした。\n\n理由: ' + content.error;
  }
  
  return formatSingleSiteAnalysis(content);
}

/**
 * キーワード指定の競合分析
 */
function handleKeywordCompetitorChat(keyword) {
  if (!keyword) {
    return 'キーワードが検出できませんでした。\n\n例：「iPhone 保険で上位サイトを分析して」';
  }
  
  // 競合分析シートから検索
  var result = findCompetitorAnalysis(keyword);
  
  if (!result.found) {
    return '「' + keyword + '」の競合分析データが見つかりませんでした。\n\n' +
           '競合分析シートにデータがない可能性があります。\n' +
           '週次の競合分析実行後に再度お試しください。';
  }
  
  return formatKeywordCompetitorResult(result);
}

/**
 * 差分分析
 */
function handleDiffAnalysisChat(pageUrl) {
  // ページに関連するキーワードを取得
  var keywords = findKeywordsForPage(pageUrl);
  
  if (!keywords || keywords.length === 0) {
    return '「' + pageUrl + '」に関連するキーワードが見つかりませんでした。';
  }
  
  // 最初のキーワードで競合分析
  var competitorResult = findCompetitorAnalysis(keywords[0]);
  
  if (!competitorResult.found) {
    return '競合分析データが見つかりませんでした。';
  }
  
  // 上位3サイトのコンテンツを取得
  var topUrls = [];
  for (var i = 0; i < competitorResult.topSites.length; i++) {
    var site = competitorResult.topSites[i];
    if (!site.isOwnSite && topUrls.length < 3) {
      topUrls.push(site.url);
    }
  }
  
  if (topUrls.length === 0) {
    return '比較対象の競合サイトが見つかりませんでした。';
  }
  
  var competitorContents = compareMultipleCompetitors(topUrls);
  
  // 自社コンテンツを取得
  var ownUrl = 'https://smaho-tap.com' + pageUrl;
  var ownContent = fetchCompetitorContent(ownUrl);
  
  if (!ownContent.success) {
    return '自社ページのコンテンツを取得できませんでした。';
  }
  
  // 差分分析
  var diff = analyzeDifference(ownContent, competitorContents.results);
  
  return formatDiffAnalysisResult(ownContent, competitorContents, diff, competitorResult);
}

/**
 * 一般的な競合分析の質問
 */
function handleGeneralCompetitorQuestion(message) {
  return '競合分析について、以下のような質問ができます：\n\n' +
         '📊 **キーワード指定**\n' +
         '「iPhone 保険 で上位サイトを分析して」\n\n' +
         '🔗 **URL直接指定**\n' +
         '「https://example.com を分析して」\n\n' +
         '📈 **差分分析**\n' +
         '「/iphone-insurance の競合と比較して」\n\n' +
         '何を分析しますか？';
}

// ============================================
// 結果フォーマット関数
// ============================================

/**
 * 単一サイト分析結果をフォーマット
 */
function formatSingleSiteAnalysis(content) {
  var response = '## 📊 サイト分析結果\n\n';
  response += '**URL**: ' + content.url + '\n';
  response += '**タイトル**: ' + content.title + '\n\n';
  
  response += '### 📝 コンテンツ概要\n';
  response += '- 文字数: ' + content.wordCount.toLocaleString() + '文字\n';
  response += '- 画像数: ' + content.imageCount + '枚\n';
  response += '- H2見出し数: ' + content.h2Count + '個\n';
  response += '- H3見出し数: ' + content.h3Count + '個\n\n';
  
  response += '### 🎯 コンテンツ機能\n';
  response += '- 目次: ' + (content.hasToc ? 'あり ✅' : 'なし') + '\n';
  response += '- FAQ: ' + (content.hasFaq ? 'あり ✅' : 'なし') + '\n';
  response += '- 動画: ' + (content.hasVideo ? 'あり ✅' : 'なし') + '\n\n';
  
  if (content.h2s && content.h2s.length > 0) {
    response += '### 📑 見出し構成（H2）\n';
    var maxH2 = Math.min(content.h2s.length, 10);
    for (var i = 0; i < maxH2; i++) {
      response += (i + 1) + '. ' + content.h2s[i] + '\n';
    }
    if (content.h2s.length > 10) {
      response += '...他' + (content.h2s.length - 10) + '件\n';
    }
  }
  
  return response;
}

/**
 * キーワード競合分析結果をフォーマット
 */
function formatKeywordCompetitorResult(result) {
  var response = '## 🏆 競合分析結果\n\n';
  response += '**キーワード**: ' + result.keyword + '\n';
  response += '**勝算度**: ' + result.winnableScore + '点\n';
  response += '**競合レベル**: ' + result.competitorLevel + '\n';
  response += '**自社順位**: ' + result.ownSiteRank + '位\n';
  response += '**自社DA**: ' + result.ownSiteDA + '\n\n';
  
  // 鮮度警告
  if (result.freshness === 'stale') {
    response += '⚠️ データが30日以上前のものです。最新データの取得をお勧めします。\n\n';
  }
  
  response += '### 📊 上位10サイト\n';
  response += '| 順位 | サイト | DA |\n';
  response += '|------|--------|----|\n';
  
  for (var i = 0; i < result.topSites.length; i++) {
    var site = result.topSites[i];
    var ownMark = site.isOwnSite ? ' ⭐自社' : '';
    var domain = extractDomainFromUrl(site.url);
    response += '| ' + site.rank + '位 | ' + domain + ownMark + ' | ' + site.da + ' |\n';
  }
  
  response += '\n### 💡 分析コメント\n';
  
  if (result.winnableScore >= 80) {
    response += '✅ **チャンス大！** 勝算度が高く、積極的にリライトすべきキーワードです。';
  } else if (result.winnableScore >= 50) {
    response += '🔶 **勝算あり** 適切な施策で上位表示が狙えます。';
  } else {
    response += '⚠️ **競合が強い** 長期的な戦略が必要です。';
  }
  
  return response;
}

/**
 * 複数サイト比較結果をフォーマット
 */
function formatMultipleCompetitorResult(result) {
  if (!result.success) {
    return 'サイト情報の取得に失敗しました。';
  }
  
  var response = '## 📊 複数サイト比較結果\n\n';
  response += '取得成功: ' + result.totalFetched + 'サイト / エラー: ' + result.totalErrors + 'サイト\n\n';
  
  response += '### 📈 統計情報\n';
  response += '- 平均文字数: ' + result.stats.avgWordCount.toLocaleString() + '文字\n';
  response += '- 平均画像数: ' + result.stats.avgImageCount + '枚\n';
  response += '- 平均H2数: ' + result.stats.avgH2Count + '個\n\n';
  
  response += '### 📝 各サイト詳細\n';
  for (var i = 0; i < result.results.length; i++) {
    var site = result.results[i];
    response += '\n**' + (i + 1) + '. ' + site.title + '**\n';
    response += '- 文字数: ' + site.wordCount.toLocaleString() + '文字\n';
    response += '- 画像: ' + site.imageCount + '枚 / H2: ' + site.h2Count + '個\n';
  }
  
  return response;
}

/**
 * 差分分析結果をフォーマット
 */
function formatDiffAnalysisResult(ownContent, competitorContents, diff, competitorResult) {
  var response = '## 🔍 差分分析結果\n\n';
  response += '**キーワード**: ' + competitorResult.keyword + '\n';
  response += '**自社順位**: ' + competitorResult.ownSiteRank + '位\n\n';
  
  response += '### 📊 数値比較\n';
  response += '| 項目 | 自社 | 競合平均 | 差分 |\n';
  response += '|------|------|----------|------|\n';
  
  var avgWordCount = competitorContents.stats.avgWordCount;
  var wordDiff = ownContent.wordCount - avgWordCount;
  var wordDiffStr = wordDiff >= 0 ? '+' + wordDiff : '' + wordDiff;
  response += '| 文字数 | ' + ownContent.wordCount.toLocaleString() + ' | ' + avgWordCount.toLocaleString() + ' | ' + wordDiffStr + ' |\n';
  
  var avgImageCount = competitorContents.stats.avgImageCount;
  var imageDiff = ownContent.imageCount - avgImageCount;
  var imageDiffStr = imageDiff >= 0 ? '+' + imageDiff : '' + imageDiff;
  response += '| 画像数 | ' + ownContent.imageCount + ' | ' + avgImageCount + ' | ' + imageDiffStr + ' |\n';
  
  response += '\n### ⚠️ 不足している要素\n';
  
  if (diff.missingFeatures && diff.missingFeatures.length > 0) {
    for (var i = 0; i < diff.missingFeatures.length; i++) {
      response += '- ' + diff.missingFeatures[i] + '\n';
    }
  } else {
    response += '特になし ✅\n';
  }
  
  response += '\n### 💡 改善提案\n';
  
  if (wordDiff < -1000) {
    response += '- 📝 コンテンツ量が競合より少ないです。' + Math.abs(wordDiff).toLocaleString() + '文字程度の追加を検討してください。\n';
  }
  
  if (imageDiff < -5) {
    response += '- 🖼️ 画像が競合より少ないです。' + Math.abs(imageDiff) + '枚程度の追加を検討してください。\n';
  }
  
  if (!ownContent.hasFaq && diff.missingFeatures && diff.missingFeatures.indexOf('FAQ') !== -1) {
    response += '- ❓ FAQセクションの追加を検討してください。\n';
  }
  
  return response;
}

/**
 * URLからドメインを抽出（表示用）
 */
function extractDomainFromUrl(url) {
  try {
    var match = url.match(/^(?:https?:\/\/)?(?:www\.)?([^\/]+)/);
    return match ? match[1] : url;
  } catch (e) {
    return url;
  }
}

/**
 * 会話履歴から現在のページURLを抽出
 */
function extractPageUrlFromHistory(conversationHistory) {
  if (!conversationHistory || conversationHistory.length === 0) {
    return null;
  }
  
  // 最新のメッセージから遡ってページURLを探す
  for (var i = conversationHistory.length - 1; i >= 0; i--) {
    var msg = conversationHistory[i];
    if (msg.content) {
      var urlMatch = msg.content.match(/\/[a-z0-9\-]+(?:\/[a-z0-9\-]+)*/i);
      if (urlMatch) {
        return urlMatch[0];
      }
    }
  }
  
  return null;
}

// ============================================
// 競合分析チャット統合テスト
// ============================================

function testWebAppCompetitorIntegration() {
  Logger.log('=== WebApp競合分析統合テスト ===');
  
  // テスト1: 競合分析リクエスト判定
  var testMessages = [
    'iPhone amazon で買う で競合分析して',
    'https://example.com を分析して',
    '上位サイトと比較したい',
    '今日の天気は？',
    'リライト提案して'
  ];
  
  Logger.log('\n--- リクエスト判定テスト ---');
  for (var i = 0; i < testMessages.length; i++) {
    var msg = testMessages[i];
    var isCompetitor = isCompetitorAnalysisRequest(msg);
    Logger.log('「' + msg + '」→ 競合分析: ' + isCompetitor);
  }
  
  // テスト2: handleChatMessageから競合分析が呼ばれるか
  Logger.log('\n--- handleChatMessage統合テスト ---');
  var competitorMessage = 'iPhone amazon で買う で競合分析して';
  
  try {
    var result = handleChatMessage(competitorMessage);
    Logger.log('結果（先頭500文字）:');
    Logger.log(result.substring(0, 500));
    Logger.log('\n✅ 統合テスト成功！');
  } catch (error) {
    Logger.log('❌ エラー: ' + error.message);
  }
}

// ============================================
// doPost() に追加する分岐処理
// ============================================

/*
既存のdoPost関数内に以下の分岐を追加してください:

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    
    // ... 既存の分岐 ...
    
    // タスク管理API（新規追加）
    if (action === 'registerTask') {
      return handleRegisterTask(data);
    }
    
    if (action === 'completeTask') {
      return handleCompleteTask(data);
    }
    
    if (action === 'getPageTasks') {
      return handleGetPageTasks(data);
    }
    
    if (action === 'getPendingTasks') {
      return handleGetPendingTasks(data);
    }
    
    if (action === 'updateTaskStatus') {
      return handleUpdateTaskStatus(data);
    }
    
    if (action === 'checkCooling') {
      return handleCheckCooling(data);
    }
    
    // ... 既存の処理 ...
  }
}
*/


// ============================================
// タスク登録ハンドラー
// ============================================

/**
 * タスク登録リクエストを処理
 * @param {Object} data - リクエストデータ
 * @return {TextOutput} JSON応答
 */
function handleRegisterTask(data) {
  try {
    let result;
    
    if (data.source === 'AI提案') {
      result = createTaskFromAISuggestion(
        data.pageUrl,
        data.pageTitle,
        data.taskType,
        data.taskDetail,
        data.priorityRank || 0
      );
    } else {
      result = createCustomTask(
        data.pageUrl,
        data.pageTitle,
        data.taskType,
        data.taskDetail,
        data.notes || ''
      );
    }
    
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}


/**
 * タスク完了リクエストを処理
 * @param {Object} data - リクエストデータ
 * @return {TextOutput} JSON応答
 */
function handleCompleteTask(data) {
  try {
    const result = completeTask(data.taskId, data.actualChange);
    
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}


/**
 * ページ別タスク取得リクエストを処理
 * @param {Object} data - リクエストデータ
 * @return {TextOutput} JSON応答
 */
function handleGetPageTasks(data) {
  try {
    const tasks = getTasksByPage(data.pageUrl);
    const coolingStatus = checkCoolingStatus(data.pageUrl);
    
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      tasks: tasks,
      coolingStatus: coolingStatus
    })).setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}


/**
 * 未完了タスク取得リクエストを処理
 * @param {Object} data - リクエストデータ
 * @return {TextOutput} JSON応答
 */
function handleGetPendingTasks(data) {
  try {
    const tasks = getPendingTasks(data.status);
    const summary = getTaskSummary();
    
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      tasks: tasks,
      summary: summary
    })).setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}


/**
 * ステータス更新リクエストを処理
 * @param {Object} data - リクエストデータ
 * @return {TextOutput} JSON応答
 */
function handleUpdateTaskStatus(data) {
  try {
    const result = updateTaskStatus(data.taskId, data.newStatus);
    
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}


/**
 * 冷却状態チェックリクエストを処理
 * @param {Object} data - リクエストデータ
 * @return {TextOutput} JSON応答
 */
function handleCheckCooling(data) {
  try {
    const coolingStatus = checkCoolingStatus(data.pageUrl, data.taskType);
    
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      coolingStatus: coolingStatus
    })).setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}


// ============================================
// チャットからのタスク登録連携
// ============================================

/**
 * AIチャットの提案をタスク登録可能な形式に変換
 * handleChatMessage()から呼び出し
 * 
 * @param {string} pageUrl - ページURL
 * @param {string} pageTitle - ページタイトル
 * @param {string} aiResponse - Claude APIからの応答
 * @return {Object} タスク登録用データ
 */
function parseAIResponseForTasks(pageUrl, pageTitle, aiResponse) {
  // AI応答から提案を抽出するパターン
  const suggestions = [];
  
  // タイトル提案を検出
  const titleMatch = aiResponse.match(/タイトル[変更修正案：:].{0,10}[「『](.+?)[」』]/);
  if (titleMatch) {
    suggestions.push({
      taskType: 'タイトル変更',
      taskDetail: titleMatch[1],
      priorityRank: 1
    });
  }
  
  // H2提案を検出
  const h2Matches = aiResponse.matchAll(/H2[追加：:].{0,10}[「『](.+?)[」』]/g);
  let h2Rank = 2;
  for (const match of h2Matches) {
    suggestions.push({
      taskType: 'H2追加',
      taskDetail: match[1],
      priorityRank: h2Rank++
    });
  }
  
  // Q&A提案を検出
  const qaMatch = aiResponse.match(/Q&A[追加：:]|よくある質問/);
  if (qaMatch) {
    suggestions.push({
      taskType: 'Q&A追加',
      taskDetail: 'FAQ構造化データを含むQ&Aセクション追加',
      priorityRank: 3
    });
  }
  
  // 本文追加を検出
  const contentMatch = aiResponse.match(/本文[追加修正：:]|コンテンツ[追加：:]/);
  if (contentMatch) {
    suggestions.push({
      taskType: '本文追加',
      taskDetail: 'コンテンツ拡充',
      priorityRank: 4
    });
  }
  
  // 冷却期間でフィルタリング
  const filtered = filterSuggestionsByCooling(pageUrl, suggestions);
  
  return {
    pageUrl: pageUrl,
    pageTitle: pageTitle,
    suggestions: filtered.available,
    excludedByCooling: filtered.excluded
  };
}


/**
 * チャット応答にタスク登録ボタン情報を追加
 * @param {string} response - 元の応答
 * @param {Object} taskData - parseAIResponseForTasks()の結果
 * @return {string} 拡張された応答
 */
function addTaskButtonsToResponse(response, taskData) {
  if (!taskData.suggestions || taskData.suggestions.length === 0) {
    return response;
  }
  
  // 応答の最後にタスク登録セクションを追加
  let taskSection = '\n\n---\n### 📋 タスク登録\n';
  taskSection += '以下の提案をタスクとして登録できます：\n\n';
  
  taskData.suggestions.forEach((suggestion, index) => {
    taskSection += `**${index + 1}. ${suggestion.taskType}**\n`;
    taskSection += `${suggestion.taskDetail}\n`;
    // フロントエンドでボタンを生成するためのマーカー
    taskSection += `<!--TASK_BUTTON:${JSON.stringify({
      pageUrl: taskData.pageUrl,
      pageTitle: taskData.pageTitle,
      taskType: suggestion.taskType,
      taskDetail: suggestion.taskDetail,
      priorityRank: suggestion.priorityRank
    })}-->\n\n`;
  });
  
  // 冷却中の項目があれば表示
  if (taskData.excludedByCooling && taskData.excludedByCooling.length > 0) {
    taskSection += generateCoolingMessage({ excluded: taskData.excludedByCooling });
  }
  
  // カスタムタスク追加のマーカー
  taskSection += `\n➕ **カスタムタスクを追加**\n`;
  taskSection += `<!--CUSTOM_TASK_BUTTON:${JSON.stringify({
    pageUrl: taskData.pageUrl,
    pageTitle: taskData.pageTitle
  })}-->\n`;
  
  return response + taskSection;
}


// ============================================
// 週次分析との連携
// ============================================

/**
 * 週次分析で冷却中ページを除外
 * Scoring.gsの getTopPriorityPages() から呼び出し
 * 
 * @param {Array} pages - ページ一覧
 * @return {Array} 冷却中でないページ一覧
 */
function excludeCoolingPagesFromAnalysis(pages) {
  return pages.filter(page => {
    const lastCompleted = getLastCompletedDate(page.url || page.page_url);
    if (!lastCompleted) return true;
    
    const today = new Date();
    const daysSince = Math.floor((today - lastCompleted) / (1000 * 60 * 60 * 24));
    
    // 最低30日は除外（タスク種別に関係なく）
    return daysSince >= 30;
  });
}


/**
 * 冷却情報を含めた優先ページ一覧を取得
 * @param {number} limit - 取得件数
 * @return {Array} ページ一覧（冷却情報付き）
 */
function getTopPriorityPagesWithCooling(limit) {
  // 既存の関数から優先ページ取得
  const pages = getTopPriorityPages ? getTopPriorityPages(limit * 2) : [];
  
  // 冷却情報を付加
  const pagesWithCooling = pages.map(page => {
    const url = page.url || page.page_url;
    const coolingStatus = checkCoolingStatus(url);
    
    return {
      ...page,
      coolingStatus: coolingStatus,
      isCooling: coolingStatus.isCooling
    };
  });
  
  // 冷却中でないものを優先、冷却中は後ろに
  const notCooling = pagesWithCooling.filter(p => !p.isCooling);
  const cooling = pagesWithCooling.filter(p => p.isCooling);
  
  return [...notCooling, ...cooling].slice(0, limit);
}

/**
 * WebApp.gs に追加するコード
 * doPost() 関数とタスク管理API
 * 追加日: 2025年12月3日
 */

// ============================================
// doPost() 関数（WebApp.gsに追加）
// ============================================

/**
 * POSTリクエストを処理
 * @param {Object} e - リクエストイベント
 * @return {TextOutput} JSON応答
 */
function doPost(e) {
  try {
    // リクエストデータを解析
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    
    Logger.log('doPost受信: action=' + action);
    
    // ============================================
    // チャットメッセージ処理
    // ============================================
    if (action === 'chat') {
      const response = handleChatMessage(data.message);
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        response: response
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // ============================================
    // タスク管理API
    // ============================================
    
    // タスク登録
    if (action === 'registerTask') {
      return handleRegisterTask(data);
    }
    
    // タスク完了
    if (action === 'completeTask') {
      return handleCompleteTask(data);
    }
    
    // ページ別タスク取得
    if (action === 'getPageTasks') {
      return handleGetPageTasks(data);
    }
    
    // 未完了タスク取得
    if (action === 'getPendingTasks') {
      return handleGetPendingTasks(data);
    }
    
    // ステータス更新
    if (action === 'updateTaskStatus') {
      return handleUpdateTaskStatus(data);
    }
    
    // 冷却状態チェック
    if (action === 'checkCooling') {
      return handleCheckCooling(data);
    }
    
    // ============================================
    // WordPress連携API
    // ============================================
    
    // WordPress投稿情報取得
    if (action === 'getWordPressPost') {
      return handleGetWordPressPost(data);
    }
    
    // WordPress更新適用
    if (action === 'applyToWordPress') {
      return handleApplyToWordPress(data);
    }
    
    // ============================================
    // データ取得API
    // ============================================
    
    // 優先ページ取得
    if (action === 'getTopPages') {
      const pages = getTopPriorityPagesWithCooling(data.limit || 10);
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        pages: pages
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // タスクサマリー取得
    if (action === 'getTaskSummary') {
      const summary = getTaskSummary();
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        summary: summary
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // ============================================
    // 不明なアクション
    // ============================================
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: '不明なアクション: ' + action
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    Logger.log('doPostエラー: ' + error.message);
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}


// ============================================
// WordPress連携ハンドラー
// ============================================

/**
 * WordPress投稿情報取得リクエストを処理
 * @param {Object} data - リクエストデータ
 * @return {TextOutput} JSON応答
 */
function handleGetWordPressPost(data) {
  try {
    const result = getPostInfoForRewrite(data.pageUrl);
    
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}


/**
 * WordPress更新適用リクエストを処理
 * @param {Object} data - リクエストデータ
 * @return {TextOutput} JSON応答
 */
function handleApplyToWordPress(data) {
  try {
    const result = applyRewriteToWordPress(data.taskId, data.changes);
    
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * WebApp.gs に追加するコード
 * doPost() 関数とタスク管理API
 * 追加日: 2025年12月3日
 */

// ============================================
// doPost() 関数（WebApp.gsに追加）
// ============================================

/**
 * POSTリクエストを処理
 * @param {Object} e - リクエストイベント
 * @return {TextOutput} JSON応答
 */
function doPost(e) {
  try {
    // リクエストデータを解析
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    
    Logger.log('doPost受信: action=' + action);
    
    // ============================================
    // チャットメッセージ処理
    // ============================================
    if (action === 'chat') {
      const response = handleChatMessage(data.message);
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        response: response
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // ============================================
    // タスク管理API
    // ============================================
    
    // タスク登録
    if (action === 'registerTask') {
      return handleRegisterTask(data);
    }
    
    // タスク完了
    if (action === 'completeTask') {
      return handleCompleteTask(data);
    }
    
    // ページ別タスク取得
    if (action === 'getPageTasks') {
      return handleGetPageTasks(data);
    }
    
    // 未完了タスク取得
    if (action === 'getPendingTasks') {
      return handleGetPendingTasks(data);
    }
    
    // ステータス更新
    if (action === 'updateTaskStatus') {
      return handleUpdateTaskStatus(data);
    }
    
    // 冷却状態チェック
    if (action === 'checkCooling') {
      return handleCheckCooling(data);
    }
    
    // ============================================
    // WordPress連携API
    // ============================================
    
    // WordPress投稿情報取得
    if (action === 'getWordPressPost') {
      return handleGetWordPressPost(data);
    }
    
    // WordPress更新適用
    if (action === 'applyToWordPress') {
      return handleApplyToWordPress(data);
    }
    
    // ============================================
    // データ取得API
    // ============================================
    
    // 優先ページ取得
    if (action === 'getTopPages') {
      const pages = getTopPriorityPagesWithCooling(data.limit || 10);
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        pages: pages
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // タスクサマリー取得
    if (action === 'getTaskSummary') {
      const summary = getTaskSummary();
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        summary: summary
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // ============================================
    // 不明なアクション
    // ============================================
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: '不明なアクション: ' + action
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    Logger.log('doPostエラー: ' + error.message);
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}


// ============================================
// WordPress連携ハンドラー
// ============================================

/**
 * WordPress投稿情報取得リクエストを処理
 * @param {Object} data - リクエストデータ
 * @return {TextOutput} JSON応答
 */
function handleGetWordPressPost(data) {
  try {
    const result = getPostInfoForRewrite(data.pageUrl);
    
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}


/**
 * WordPress更新適用リクエストを処理
 * @param {Object} data - リクエストデータ
 * @return {TextOutput} JSON応答
 */
function handleApplyToWordPress(data) {
  try {
    const result = applyRewriteToWordPress(data.taskId, data.changes);
    
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: error.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================
// GSCズレ分析チャット連携（Day 22追加）
// ============================================

/**
 * GSCズレ分析リクエストかどうかを判定
 * @param {string} message - ユーザーメッセージ
 * @return {boolean} GSCズレ分析リクエストかどうか
 */
function isGSCGapAnalysisRequest(message) {
  var lowerMessage = message.toLowerCase();
  
  var keywords = [
    'gscとターゲット',
    'gsc ターゲット',
    'gscズレ',
    'gsc ズレ',
    'ターゲットkwのズレ',
    'ターゲットキーワードのズレ',
    'クエリとターゲット',
    '実クエリとターゲット',
    'gsc分析',
    'キーワードのズレ',
    'gscと登録kw',
    '検索クエリとターゲット'
  ];
  
  for (var i = 0; i < keywords.length; i++) {
    if (lowerMessage.indexOf(keywords[i]) !== -1) {
      return true;
    }
  }
  
  return false;
}

/**
 * GSCズレ分析リクエストを処理（チャットから呼び出し）
 * @param {string} message - ユーザーメッセージ
 * @return {string} レスポンス
 */
function handleGSCGapAnalysisFromChat(message) {
  try {
    // URLを抽出
    var urlMatch = message.match(/\/[\w\-\/]+/);
    
    if (!urlMatch) {
      // URLが指定されていない場合、優先度上位ページを提案
      return 'GSCズレ分析を行うページURLを指定してください。\n\n' +
             '**使用例**:\n' +
             '- 「/insurance/recommend/ のGSCとターゲットKWのズレを教えて」\n' +
             '- 「/iphone-repair/ のGSC分析して」\n\n' +
             '💡 **ヒント**: リライト提案と一緒にGSCズレ分析も表示されます。';
    }
    
    var pageUrl = urlMatch[0];
    Logger.log('GSCズレ分析対象URL: ' + pageUrl);
    
    // SuggestionGenerator.gsのgetGSCTargetKWGapForChat関数を呼び出し
    if (typeof getGSCTargetKWGapForChat === 'function') {
      var gscGapText = getGSCTargetKWGapForChat(pageUrl);
      
      // トレンド分析も追加
      var trendText = '';
      if (typeof applyTrendModifier === 'function') {
        var targetKW = getTargetKeywordForPage(pageUrl);
        var trendResult = applyTrendModifier(pageUrl, targetKW, 50);
        
        if (trendResult.trend) {
          trendText = '\n\n📈 **順位トレンド（過去4週間）**\n';
          trendText += 'トレンド: **' + trendResult.trendLabel + '**\n';
          trendText += trendResult.message + '\n';
          
          if (trendResult.weeklyRanks) {
            var ranks = trendResult.weeklyRanks.map(function(w) {
              return w.rank || '圏外';
            }).join(' → ');
            trendText += '推移: ' + ranks + '\n';
          }
        }
      }
      
      return '## 🔍 GSCズレ分析結果\n\n' +
             '**対象ページ**: ' + pageUrl + '\n' +
             gscGapText + trendText;
    } else {
      return '⚠️ GSCズレ分析機能が利用できません。\n' +
             'SuggestionGenerator.gsにgetGSCTargetKWGapForChat関数が実装されているか確認してください。';
    }
    
  } catch (error) {
    Logger.log('GSCズレ分析エラー: ' + error.message);
    return 'GSCズレ分析中にエラーが発生しました: ' + error.message;
  }
}

/**
 * ページのターゲットキーワードを取得（補助関数）
 * @param {string} pageUrl - ページURL
 * @return {string} ターゲットキーワード
 */
function getTargetKeywordForPage(pageUrl) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('統合データ');
    
    if (!sheet) return '';
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var urlIdx = headers.indexOf('page_url');
    var kwIdx = headers.indexOf('target_keyword');
    
    if (urlIdx === -1 || kwIdx === -1) return '';
    
    // URLを正規化
    var normalizedInput = pageUrl.toLowerCase().replace(/\/$/, '');
    
    for (var i = 1; i < data.length; i++) {
      var rowUrl = (data[i][urlIdx] || '').toString().toLowerCase().replace(/\/$/, '');
      
      if (rowUrl === normalizedInput || rowUrl.indexOf(normalizedInput) !== -1 || normalizedInput.indexOf(rowUrl) !== -1) {
        return data[i][kwIdx] || '';
      }
    }
    
    return '';
  } catch (error) {
    Logger.log('ターゲットKW取得エラー: ' + error.message);
    return '';
  }
}

/**
 * GSCズレ分析のテスト
 */
function testGSCGapAnalysisChat() {
  Logger.log('=== GSCズレ分析チャットテスト ===');
  
  // テスト1: リクエスト判定
  var testMessages = [
    'GSCとターゲットKWのズレを教えて',
    '/iphone-insurance/ のGSC分析して',
    'リライト提案して',
    '競合分析して'
  ];
  
  Logger.log('\n--- リクエスト判定テスト ---');
  for (var i = 0; i < testMessages.length; i++) {
    var msg = testMessages[i];
    var isGSCGap = isGSCGapAnalysisRequest(msg);
    Logger.log('「' + msg + '」→ GSCズレ分析: ' + isGSCGap);
  }
  
  // テスト2: 実際の分析（統合データの最初のページ）
  Logger.log('\n--- 分析実行テスト ---');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('統合データ');
  
  if (sheet) {
    var data = sheet.getDataRange().getValues();
    var testUrl = data[1][0]; // 最初のページURL
    
    Logger.log('テスト対象URL: ' + testUrl);
    var result = handleGSCGapAnalysisFromChat(testUrl + ' のGSCズレを分析して');
    Logger.log('結果（先頭500文字）:');
    Logger.log(result.substring(0, 500));
  }
  
  Logger.log('\n=== テスト完了 ===');
}