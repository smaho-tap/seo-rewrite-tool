/**
 * ClarityGA4Integration.gs
 * Clarity GA4連携 - GA4 Data API経由でClarityイベントを取得
 * 
 * 作成日: 2025年12月4日
 * バージョン: 1.0
 * 
 * 背景:
 * - Clarity Data Export APIは過去1-3日間のみ取得可能（不十分）
 * - Day 15でGA4連携を有効化済み
 * - GA4経由なら過去30日以上のデータ取得可能
 * 
 * 機能:
 * - GA4からClarityイベント取得（RageClick, DeadClick, QuickBack, ErrorClick, ExcessiveScroll）
 * - URL単位でUXスコア算出
 * - 統合データシートのUX列更新
 * - 週次自動実行対応
 */

// ========================================
// 設定
// ========================================

const CLARITY_GA4_CONFIG = {
  // Clarityイベント名（GA4に送信されるイベント）
  events: [
    'ClarityRageClick',
    'ClarityDeadClick', 
    'ClarityQuickBack',
    'ClarityErrorClick',
    'ClarityExcessiveScroll'
  ],
  
  // データ取得期間（日数）
  dateRange: 30,
  
  // UXスコア重み付け
  weights: {
    rageClick: 0.25,
    deadClick: 0.25,
    quickBack: 0.20,
    errorClick: 0.15,
    excessiveScroll: 0.15
  }
};

// ========================================
// メイン関数
// ========================================

/**
 * GA4からClarityイベントを取得してUXスコアを更新（メイン関数）
 */
function updateUXScoresFromClarityGA4() {
  const startTime = new Date();
  Logger.log('=== Clarity GA4連携開始 ===');
  
  try {
    // 1. GA4プロパティID取得
    const propertyId = PropertiesService.getScriptProperties().getProperty('GA4_PROPERTY_ID');
    if (!propertyId) {
      throw new Error('GA4_PROPERTY_IDが設定されていません');
    }
    Logger.log('GA4プロパティID: ' + propertyId);
    
    // 2. GA4からClarityイベント取得
    const clarityEvents = fetchClarityEventsFromGA4(propertyId);
    
    if (!clarityEvents || Object.keys(clarityEvents).length === 0) {
      Logger.log('⚠️ Clarityイベントが見つかりませんでした');
      Logger.log('確認事項:');
      Logger.log('  1. ClarityでGA4連携が有効になっているか');
      Logger.log('  2. データ蓄積に7-30日必要');
      return {
        success: false,
        message: 'Clarityイベントなし（データ蓄積待ちの可能性）',
        updatedCount: 0
      };
    }
    
    Logger.log('取得URL数: ' + Object.keys(clarityEvents).length);
    
    // 3. UXスコア算出
    const uxScores = calculateUXScoresFromClarityEvents(clarityEvents);
    Logger.log('UXスコア算出URL数: ' + Object.keys(uxScores).length);
    
    // 4. 統合データシート更新
    const updatedCount = updateIntegratedSheetWithClarityGA4(uxScores);
    
    // 5. ログ記録
    const elapsed = ((new Date() - startTime) / 1000).toFixed(1);
    Logger.log('=== Clarity GA4連携完了（' + elapsed + '秒、' + updatedCount + '件更新） ===');
    
    return {
      success: true,
      message: updatedCount + '件のUXスコアを更新しました',
      updatedCount: updatedCount,
      totalUrls: Object.keys(clarityEvents).length
    };
    
  } catch (error) {
    Logger.log('❌ エラー: ' + error.message);
    Logger.log(error.stack);
    return {
      success: false,
      message: error.message,
      updatedCount: 0
    };
  }
}

// ========================================
// GA4データ取得
// ========================================

/**
 * GA4 Data APIからClarityイベントを取得
 */
function fetchClarityEventsFromGA4(propertyId) {
  // 日付範囲設定
  const endDate = 'yesterday';
  const startDate = CLARITY_GA4_CONFIG.dateRange + 'daysAgo';
  
  Logger.log('取得期間: ' + startDate + ' ～ ' + endDate);
  
  // 結果格納用
  const urlEvents = {};
  
  // 各Clarityイベントを取得
  for (const eventName of CLARITY_GA4_CONFIG.events) {
    try {
      Logger.log('取得中: ' + eventName);
      
      const request = {
        dateRanges: [{
          startDate: startDate,
          endDate: endDate
        }],
        dimensions: [
          { name: 'pagePath' }
        ],
        metrics: [
          { name: 'eventCount' },
          { name: 'sessions' }
        ],
        dimensionFilter: {
          filter: {
            fieldName: 'eventName',
            stringFilter: {
              matchType: 'EXACT',
              value: eventName
            }
          }
        },
        limit: 10000
      };
      
      const response = AnalyticsData.Properties.runReport(request, 'properties/' + propertyId);
      
      if (!response.rows || response.rows.length === 0) {
        Logger.log('  → データなし');
        continue;
      }
      
      Logger.log('  → ' + response.rows.length + '件');
      
      // URL別に集計
      for (const row of response.rows) {
        const pagePath = row.dimensionValues[0].value;
        const url = normalizeClarityUrl(pagePath);
        if (!url) continue;
        
        const eventCount = parseInt(row.metricValues[0].value) || 0;
        const sessions = parseInt(row.metricValues[1].value) || 0;
        
        if (!urlEvents[url]) {
          urlEvents[url] = {
            rageClick: 0,
            deadClick: 0,
            quickBack: 0,
            errorClick: 0,
            excessiveScroll: 0,
            sessions: 0
          };
        }
        
        // イベント名に応じてカウント
        switch (eventName) {
          case 'ClarityRageClick':
            urlEvents[url].rageClick += eventCount;
            break;
          case 'ClarityDeadClick':
            urlEvents[url].deadClick += eventCount;
            break;
          case 'ClarityQuickBack':
            urlEvents[url].quickBack += eventCount;
            break;
          case 'ClarityErrorClick':
            urlEvents[url].errorClick += eventCount;
            break;
          case 'ClarityExcessiveScroll':
            urlEvents[url].excessiveScroll += eventCount;
            break;
        }
        
        // セッション数を更新（最大値を採用）
        if (sessions > urlEvents[url].sessions) {
          urlEvents[url].sessions = sessions;
        }
      }
      
    } catch (error) {
      Logger.log('  → エラー: ' + error.message);
    }
    
    // API制限対策
    Utilities.sleep(300);
  }
  
  return urlEvents;
}

// ========================================
// UXスコア算出
// ========================================

/**
 * イベントデータからUXスコアを算出
 */
function calculateUXScoresFromClarityEvents(urlEvents) {
  const uxScores = {};
  
  // 全URL統計（正規化用）
  const allRates = {
    rageClick: [],
    deadClick: [],
    quickBack: [],
    errorClick: [],
    excessiveScroll: []
  };
  
  // 各URLのレートを計算
  for (const [url, events] of Object.entries(urlEvents)) {
    if (events.sessions < 10) continue; // セッション少なすぎは除外
    
    allRates.rageClick.push(events.rageClick / events.sessions);
    allRates.deadClick.push(events.deadClick / events.sessions);
    allRates.quickBack.push(events.quickBack / events.sessions);
    allRates.errorClick.push(events.errorClick / events.sessions);
    allRates.excessiveScroll.push(events.excessiveScroll / events.sessions);
  }
  
  // 90パーセンタイルを算出（正規化の基準）
  const p90 = {
    rageClick: getPercentileValue(allRates.rageClick, 90),
    deadClick: getPercentileValue(allRates.deadClick, 90),
    quickBack: getPercentileValue(allRates.quickBack, 90),
    errorClick: getPercentileValue(allRates.errorClick, 90),
    excessiveScroll: getPercentileValue(allRates.excessiveScroll, 90)
  };
  
  // 各URLのUXスコアを算出
  for (const [url, events] of Object.entries(urlEvents)) {
    if (events.sessions < 10) continue;
    
    // 各イベントの発生率
    const rates = {
      rageClick: events.rageClick / events.sessions,
      deadClick: events.deadClick / events.sessions,
      quickBack: events.quickBack / events.sessions,
      errorClick: events.errorClick / events.sessions,
      excessiveScroll: events.excessiveScroll / events.sessions
    };
    
    // 正規化スコア（0-100、高いほど問題あり）
    const normalized = {
      rageClick: normalizeToScore(rates.rageClick, p90.rageClick),
      deadClick: normalizeToScore(rates.deadClick, p90.deadClick),
      quickBack: normalizeToScore(rates.quickBack, p90.quickBack),
      errorClick: normalizeToScore(rates.errorClick, p90.errorClick),
      excessiveScroll: normalizeToScore(rates.excessiveScroll, p90.excessiveScroll)
    };
    
    // 総合UXスコア（加重平均）
    const totalScore = 
      normalized.rageClick * CLARITY_GA4_CONFIG.weights.rageClick +
      normalized.deadClick * CLARITY_GA4_CONFIG.weights.deadClick +
      normalized.quickBack * CLARITY_GA4_CONFIG.weights.quickBack +
      normalized.errorClick * CLARITY_GA4_CONFIG.weights.errorClick +
      normalized.excessiveScroll * CLARITY_GA4_CONFIG.weights.excessiveScroll;
    
    uxScores[url] = {
      rageClickCount: events.rageClick,
      deadClickCount: events.deadClick,
      quickBackCount: events.quickBack,
      sessions: events.sessions,
      uxScore: Math.round(totalScore)
    };
  }
  
  return uxScores;
}

/**
 * パーセンタイル算出
 */
function getPercentileValue(arr, percentile) {
  if (arr.length === 0) return 0.1;
  const sorted = arr.slice().sort((a, b) => a - b);
  const index = Math.ceil((percentile / 100) * sorted.length) - 1;
  return sorted[Math.max(0, index)] || 0.1;
}

/**
 * レートを0-100スコアに正規化
 */
function normalizeToScore(rate, p90) {
  if (p90 === 0) return 0;
  const normalized = (rate / p90) * 100;
  return Math.min(100, Math.max(0, normalized));
}

// ========================================
// 統合データシート更新
// ========================================

/**
 * 統合データシートのUX列を更新
 */
function updateIntegratedSheetWithClarityGA4(uxScores) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('統合データ');
  
  if (!sheet) {
    throw new Error('統合データシートが見つかりません');
  }
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  // 列インデックス取得
  const urlCol = headers.indexOf('page_path');
  if (urlCol === -1) {
    // page_urlも試す
    const urlCol2 = headers.indexOf('page_url');
    if (urlCol2 === -1) {
      throw new Error('page_path/page_url列が見つかりません');
    }
  }
  const actualUrlCol = urlCol !== -1 ? urlCol : headers.indexOf('page_url');
  
  // Clarity関連列を特定
  const rageClickCol = headers.indexOf('clarity_rage_clicks');
  const deadClickCol = headers.indexOf('clarity_dead_clicks');
  const quickBackCol = headers.indexOf('clarity_quick_backs');
  const uxScoreCol = headers.indexOf('clarity_ux_score');
  
  if (uxScoreCol === -1) {
    Logger.log('⚠️ clarity_ux_score列が見つかりません');
    return 0;
  }
  
  let updatedCount = 0;
  
  for (let i = 1; i < data.length; i++) {
    const pageUrl = data[i][actualUrlCol];
    const url = normalizeClarityUrl(pageUrl);
    
    if (uxScores[url]) {
      const scores = uxScores[url];
      
      // 各列を更新
      if (rageClickCol !== -1) {
        sheet.getRange(i + 1, rageClickCol + 1).setValue(scores.rageClickCount);
      }
      if (deadClickCol !== -1) {
        sheet.getRange(i + 1, deadClickCol + 1).setValue(scores.deadClickCount);
      }
      if (quickBackCol !== -1) {
        sheet.getRange(i + 1, quickBackCol + 1).setValue(scores.quickBackCount);
      }
      if (uxScoreCol !== -1) {
        sheet.getRange(i + 1, uxScoreCol + 1).setValue(scores.uxScore);
      }
      
      updatedCount++;
    }
  }
  
  return updatedCount;
}

// ========================================
// ユーティリティ
// ========================================

/**
 * URLを正規化
 */
function normalizeClarityUrl(url) {
  if (!url) return '';
  
  let path = String(url);
  
  // 絶対URLから相対パスを抽出
  if (path.startsWith('http')) {
    try {
      const urlObj = new URL(path);
      path = urlObj.pathname;
    } catch (e) {
      // パースエラーの場合はそのまま
    }
  }
  
  // 末尾スラッシュを除去
  path = path.replace(/\/$/, '');
  
  // 空の場合はトップページ
  if (!path || path === '') {
    path = '/';
  }
  
  return path;
}

// ========================================
// テスト関数
// ========================================

/**
 * Clarity GA4連携テスト
 */
function testClarityGA4Integration() {
  Logger.log('=== Clarity GA4連携テスト ===\n');
  
  // 設定確認
  const propertyId = PropertiesService.getScriptProperties().getProperty('GA4_PROPERTY_ID');
  Logger.log('GA4プロパティID: ' + (propertyId ? propertyId : '未設定'));
  
  if (!propertyId) {
    Logger.log('❌ GA4_PROPERTY_IDが設定されていません');
    return;
  }
  
  // 各イベントの存在確認
  Logger.log('\n--- Clarityイベント存在確認 ---');
  
  let totalEvents = 0;
  
  for (const eventName of CLARITY_GA4_CONFIG.events) {
    try {
      const request = {
        dateRanges: [{
          startDate: '7daysAgo',
          endDate: 'yesterday'
        }],
        dimensions: [
          { name: 'eventName' }
        ],
        metrics: [
          { name: 'eventCount' }
        ],
        dimensionFilter: {
          filter: {
            fieldName: 'eventName',
            stringFilter: {
              matchType: 'EXACT',
              value: eventName
            }
          }
        },
        limit: 1
      };
      
      const response = AnalyticsData.Properties.runReport(request, 'properties/' + propertyId);
      
      if (response.rows && response.rows.length > 0) {
        const count = parseInt(response.metricValues[0].value) || 0;
        Logger.log('✅ ' + eventName + ': ' + count + '件');
        totalEvents += count;
      } else {
        Logger.log('❌ ' + eventName + ': データなし');
      }
    } catch (error) {
      Logger.log('❌ ' + eventName + ': エラー - ' + error.message);
    }
    
    Utilities.sleep(300);
  }
  
  Logger.log('\n--- 結果 ---');
  if (totalEvents > 0) {
    Logger.log('✅ Clarityイベントが検出されました');
    Logger.log('updateUXScoresFromClarityGA4() を実行してUXスコアを更新できます');
  } else {
    Logger.log('⚠️ Clarityイベントが見つかりませんでした');
    Logger.log('確認事項:');
    Logger.log('  1. ClarityでGA4連携が有効になっているか');
    Logger.log('  2. 連携後7-30日のデータ蓄積が必要');
  }
  
  Logger.log('\n=== テスト完了 ===');
}

/**
 * ドライラン（実際には更新しない）
 */
function dryRunClarityGA4() {
  Logger.log('=== ドライラン開始 ===\n');
  
  const propertyId = PropertiesService.getScriptProperties().getProperty('GA4_PROPERTY_ID');
  if (!propertyId) {
    Logger.log('❌ GA4_PROPERTY_IDが設定されていません');
    return;
  }
  
  // イベント取得
  const clarityEvents = fetchClarityEventsFromGA4(propertyId);
  Logger.log('\n取得URL数: ' + Object.keys(clarityEvents).length);
  
  if (Object.keys(clarityEvents).length === 0) {
    Logger.log('⚠️ データなし');
    return;
  }
  
  // UX問題が多いURL TOP5
  const sorted = Object.entries(clarityEvents)
    .sort((a, b) => {
      const totalA = a[1].rageClick + a[1].deadClick + a[1].quickBack;
      const totalB = b[1].rageClick + b[1].deadClick + b[1].quickBack;
      return totalB - totalA;
    })
    .slice(0, 5);
  
  Logger.log('\n--- UX問題が多いURL TOP5 ---');
  for (const [url, events] of sorted) {
    Logger.log(url);
    Logger.log('  Rage: ' + events.rageClick + ', Dead: ' + events.deadClick + ', QuickBack: ' + events.quickBack);
  }
  
  // UXスコア算出
  const uxScores = calculateUXScoresFromClarityEvents(clarityEvents);
  
  // スコアが高いURL TOP5
  const sortedScores = Object.entries(uxScores)
    .sort((a, b) => b[1].uxScore - a[1].uxScore)
    .slice(0, 5);
  
  Logger.log('\n--- UXスコアが高いURL TOP5（改善優先度高） ---');
  for (const [url, scores] of sortedScores) {
    Logger.log(url + ': ' + scores.uxScore + '点');
  }
  
  Logger.log('\n=== ドライラン完了 ===');
}

// ========================================
// 週次実行用
// ========================================

/**
 * 週次Clarity GA4更新（トリガーから呼び出し）
 */
function weeklyClarityGA4Update() {
  Logger.log('=== 週次Clarity GA4連携 ===');
  
  const result = updateUXScoresFromClarityGA4();
  
  if (result.success) {
    Logger.log('✅ 成功: ' + result.message);
  } else {
    Logger.log('⚠️ ' + result.message);
  }
  
  return result;
}