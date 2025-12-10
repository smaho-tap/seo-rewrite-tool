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

    // ★アウトライン生成リクエスト（フェーズ1追加）
    if (userMessage.indexOf('__GENERATE_OUTLINE__') === 0) {
      Logger.log('アウトライン生成リクエスト検出');
      try {
        var jsonPart = userMessage.replace('__GENERATE_OUTLINE__', '');
        var params = JSON.parse(jsonPart);
        
        var result = generateOutline(params.pageUrl, params.suggestionTitle, params.suggestionType, params.suggestionContent);
        
        if (result.success) {
          var response = '## ✅ アウトラインを生成しました\n\n';
          response += '**ページ**: ' + params.pageUrl + '\n';
          response += '**種別**: ' + params.suggestionType + '\n\n---\n\n';
          response += result.outline;
          response += '\n\n---\n💡 このアウトラインを参考に、コンテンツを作成してください。';
          return response;
        } else {
          return '❌ アウトライン生成に失敗しました: ' + result.error;
        }
      } catch (e) {
        return '❌ アウトライン生成でエラー: ' + e.message;
      }
    }
    
    // ★タスク追加リクエスト（フェーズ1追加）
    if (userMessage.indexOf('__ADD_TASK__') === 0) {
      Logger.log('タスク追加リクエスト検出');
      try {
        var jsonPart = userMessage.replace('__ADD_TASK__', '');
        var params = JSON.parse(jsonPart);
        
        var result = registerTaskFromSuggestion(params.pageUrl, params.taskType, params.taskContent, params.priority || 3);
        
        if (result.success) {
          var priorityEmoji = ['', '🥇', '🥈', '🥉', '⭐', '☆'][params.priority] || '⭐';
          var response = '## ✅ タスクを登録しました\n\n';
          response += '| 項目 | 内容 |\n|------|------|\n';
          response += '| タスクID | ' + result.taskId + ' |\n';
          response += '| ページ | ' + params.pageUrl + ' |\n';
          response += '| 種別 | ' + params.taskType + ' |\n';
          response += '| 優先度 | ' + priorityEmoji + ' 優先度' + params.priority + ' |\n';
          response += '| ステータス | 未着手 |\n\n';
          response += '📋 「タスク管理」シートで確認・編集できます。';
          return response;
        } else {
          if (result.existingTaskId) {
            return '⚠️ 同じタスクが既に登録されています（ID: ' + result.existingTaskId + '）';
          }
          return '❌ タスク登録に失敗: ' + result.error;
        }
      } catch (e) {
        return '❌ タスク登録でエラー: ' + e.message;
      }
    }
    // ★詳細表示リクエスト（フェーズ1追加）
    if (userMessage.indexOf('__VIEW_DETAIL__') === 0) {
      Logger.log('詳細表示リクエスト検出');
      try {
        var pageUrl = userMessage.replace('__VIEW_DETAIL__', '').trim();
        var result = generateRewriteSuggestions(pageUrl);
        
        if (result.success) {
          return '## 📋 ' + pageUrl + ' の詳細リライト提案\n\n' + result.suggestion;
        } else {
          return '❌ 詳細取得に失敗しました: ' + result.suggestion;
        }
      } catch (e) {
        return '❌ 詳細取得でエラー: ' + e.message;
      }
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
          return result.suggestion;
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
          response += '   (機会損失: ' + page.opportunityScore + ' / パフォーマンス: ' + page.performanceScore + ' / ビジネス: ' + page.businessImpactScore + ')\n';
          response += '<button class="view-detail-btn" data-page-url="' + escapeHtmlAttr(page.url) + '">📋 詳細を見る</button>\n\n';
        }
        
        response += '---\n\n';
        response += '【1位ページの詳細リライト提案】\n\n';
        
        if (result.success) {
          response += result.suggestion;
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
  
  // ★「リライト」だけでなく「設置」「変更」「改善」も効果測定対象に
  var hasActionKeyword = message.includes('リライト') || 
                         message.includes('設置') || 
                         message.includes('変更') || 
                         message.includes('改善');
  
  var hasEffectKeyword = message.includes('効果') || message.includes('測定') || 
                         message.includes('比較') || message.includes('前後') ||
                         message.includes('滞在時間') || message.includes('離脱率') ||
                         message.includes('直帰率');
  
  if (hasActionKeyword && hasEffectKeyword) {
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
 * ページURLを抽出
 */
function extractPageUrl(userMessage) {
  // 「/xxx」形式のURLを抽出
  var match = userMessage.match(/\/[a-zA-Z0-9\-_\/]+/);
  if (match) {
    return match[0];
  }
  
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
  // ★効果測定キーワードが含まれている場合は競合分析ではない
  var effectMeasurementKeywords = [
    '設置', '変更', '改善', '前後', '効果', '測定',
    '滞在時間', '離脱率', '直帰率', 'PV', 'ページビュー'
  ];
  
  for (var i = 0; i < effectMeasurementKeywords.length; i++) {
    if (message.includes(effectMeasurementKeywords[i])) {
      Logger.log('効果測定キーワード検出: ' + effectMeasurementKeywords[i] + ' → 競合分析ではない');
      return false;
    }
  }
  
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