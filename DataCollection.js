/**
 * SEOリライト支援ツール - データ収集スクリプト
 * Day 2: GA4/GSCデータ取得 + Phase 2: 期間指定対応
 */

// ========================================
// 設定関数
// ========================================

/**
 * GA4プロパティIDを設定
 * 初回のみ実行してください
 */
function setGA4PropertyId() {
  const propertyId = 'properties/388689745';
  
  PropertiesService.getScriptProperties()
    .setProperty('GA4_PROPERTY_ID', propertyId);
  
  Logger.log(`✅ GA4_PROPERTY_IDを設定しました: ${propertyId}`);
  
  // 確認
  const saved = PropertiesService.getScriptProperties().getProperty('GA4_PROPERTY_ID');
  Logger.log(`確認: ${saved}`);
}

/**
 * Search Console サイトURLを設定
 * 初回のみ実行してください
 */
function setGSCSiteUrl() {
  const siteUrl = 'https://smaho-tap.com';
  
  PropertiesService.getScriptProperties()
    .setProperty('GSC_SITE_URL', siteUrl);
  
  Logger.log(`✅ GSC_SITE_URLを設定しました: ${siteUrl}`);
  
  // 確認
  const saved = PropertiesService.getScriptProperties().getProperty('GSC_SITE_URL');
  Logger.log(`確認: ${saved}`);
}

// ========================================
// GA4データ取得（定期実行用）
// ========================================

/**
 * GA4からページ別データを取得してGA4_RAWシートに保存
 */
function fetchGA4Data() {
  const startTime = new Date();
  Logger.log('=== GA4データ取得開始 ===');
  
  try {
    // 設定取得
    const propertyId = PropertiesService.getScriptProperties().getProperty('GA4_PROPERTY_ID');
    
    if (!propertyId) {
      throw new Error('GA4_PROPERTY_IDが設定されていません');
    }
    
    // リクエストボディ作成
    const request = {
      dateRanges: [
        {
          startDate: '30daysAgo',
          endDate: 'yesterday'
        }
      ],
      dimensions: [
        { name: 'pagePath' },
        { name: 'pageTitle' }
      ],
      metrics: [
        { name: 'screenPageViews' },
        { name: 'activeUsers' },
        { name: 'averageSessionDuration' },
        { name: 'bounceRate' },
        { name: 'conversions' }
      ],
      limit: 10000,
      orderBys: [
        {
          metric: { metricName: 'screenPageViews' },
          desc: true
        }
      ]
    };
    
    Logger.log(`プロパティID: ${propertyId}`);
    Logger.log('GA4 API呼び出し中...');
    
    // API呼び出し
    const response = AnalyticsData.Properties.runReport(request, propertyId);
    
    // データ処理
    const rows = response.rows || [];
    Logger.log(`取得行数: ${rows.length}行`);
    
    if (rows.length === 0) {
      Logger.log('警告: データが0件でした');
      return { success: true, rowCount: 0 };
    }
    
    // データを配列に変換
    const timestamp = new Date();
    const data = rows.map(row => [
      timestamp,                                    // date
      row.dimensionValues[0].value,                // page_path
      row.dimensionValues[1].value,                // page_title
      parseInt(row.metricValues[0].value),         // page_views
      parseInt(row.metricValues[1].value),         // unique_users
      parseFloat(row.metricValues[2].value),       // avg_session_duration
      parseFloat(row.metricValues[3].value) * 100, // bounce_rate (0-1 → %)
      parseInt(row.metricValues[4].value),         // conversions
      0,                                            // conversion_rate (後で計算)
      timestamp                                     // last_updated
    ]);
    
    // コンバージョン率を計算
    data.forEach(row => {
      const pageViews = row[3];
      const conversions = row[7];
      row[8] = pageViews > 0 ? (conversions / pageViews * 100) : 0;
    });
    
    // スプレッドシートに書き込み
    writeToGA4Sheet(data);
    
    const endTime = new Date();
    const duration = (endTime - startTime) / 1000;
    
    Logger.log(`=== GA4データ取得完了 ===`);
    Logger.log(`処理時間: ${duration}秒`);
    Logger.log(`取得件数: ${data.length}件`);
    
    return { success: true, rowCount: data.length, duration: duration };
    
  } catch (error) {
    Logger.log(`❌ GA4データ取得エラー: ${error.message}`);
    Logger.log(`スタックトレース: ${error.stack}`);
    throw error;
  }
}

/**
 * GA4データをGA4_RAWシートに書き込み
 */
function writeToGA4Sheet(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('GA4_RAW');
  
  if (!sheet) {
    throw new Error('GA4_RAWシートが見つかりません');
  }
  
  // ヘッダー行の確認・作成
  if (sheet.getLastRow() === 0) {
    const headers = [
      'date', 'page_path', 'page_title', 'page_views', 'unique_users',
      'avg_session_duration', 'bounce_rate', 'conversions', 
      'conversion_rate', 'last_updated'
    ];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  
  // データ書き込み（バッチ処理）
  if (data.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, data.length, data[0].length).setValues(data);
    
    Logger.log(`${data.length}行をGA4_RAWシートに書き込みました`);
  }
}

// ========================================
// Search Consoleデータ取得（定期実行用）
// ========================================

/**
 * Google Search Consoleからデータを取得してGSC_RAWシートに保存
 * UrlFetchAppを使用
 */
function fetchGSCData() {
  const startTime = new Date();
  Logger.log('=== Search Consoleデータ取得開始 ===');
  
  try {
    // 設定取得
    const siteUrl = PropertiesService.getScriptProperties().getProperty('GSC_SITE_URL');
    
    if (!siteUrl) {
      throw new Error('GSC_SITE_URLが設定されていません');
    }
    
    // 日付計算
    const endDate = getDateString(1);   // 昨日
    const startDate = getDateString(30); // 30日前
    
    Logger.log(`サイトURL: ${siteUrl}`);
    Logger.log(`期間: ${startDate} ～ ${endDate}`);
    
    // リクエストボディ作成
    const requestBody = {
      startDate: startDate,
      endDate: endDate,
      dimensions: ['page', 'query'],
      rowLimit: 25000,
      dataState: 'final'
    };
    
    // API URL（URLエンコード）
    const encodedSiteUrl = encodeURIComponent(siteUrl);
    const apiUrl = `https://www.googleapis.com/webmasters/v3/sites/${encodedSiteUrl}/searchAnalytics/query`;
    
    Logger.log('Search Console API呼び出し中...');
    
    // アクセストークン取得
    const token = ScriptApp.getOAuthToken();
    
    // APIリクエスト
    const options = {
      method: 'post',
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(requestBody),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(apiUrl, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode !== 200) {
      const errorText = response.getContentText();
      Logger.log(`API Error Response: ${errorText}`);
      throw new Error(`API Error: ${responseCode} - ${errorText}`);
    }
    
    const jsonData = JSON.parse(response.getContentText());
    const rows = jsonData.rows || [];
    
    Logger.log(`取得行数: ${rows.length}行`);
    
    if (rows.length === 0) {
      Logger.log('警告: データが0件でした');
      return { success: true, rowCount: 0 };
    }
    
    // データを配列に変換
    const timestamp = new Date();
    const data = rows.map(row => [
      timestamp,                          // date
      row.keys[0],                        // page_url
      row.keys[1],                        // query
      parseFloat(row.position.toFixed(1)), // position
      row.clicks,                         // clicks
      row.impressions,                    // impressions
      parseFloat((row.ctr * 100).toFixed(2)), // ctr (0-1 → %)
      timestamp                           // last_updated
    ]);
    
    // スプレッドシートに書き込み
    writeToGSCSheet(data);
    
    const endTime = new Date();
    const duration = (endTime - startTime) / 1000;
    
    Logger.log(`=== Search Consoleデータ取得完了 ===`);
    Logger.log(`処理時間: ${duration}秒`);
    Logger.log(`取得件数: ${data.length}件`);
    
    return { success: true, rowCount: data.length, duration: duration };
    
  } catch (error) {
    Logger.log(`❌ Search Consoleデータ取得エラー: ${error.message}`);
    Logger.log(`スタックトレース: ${error.stack}`);
    throw error;
  }
}

/**
 * Search ConsoleデータをGSC_RAWシートに書き込み
 */
function writeToGSCSheet(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('GSC_RAW');
  
  if (!sheet) {
    throw new Error('GSC_RAWシートが見つかりません');
  }
  
  // ヘッダー行の確認・作成
  if (sheet.getLastRow() === 0) {
    const headers = [
      'date', 'page_url', 'query', 'position', 
      'clicks', 'impressions', 'ctr', 'last_updated'
    ];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  
  // データ書き込み（バッチ処理）
  if (data.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, data.length, data[0].length).setValues(data);
    
    Logger.log(`${data.length}行をGSC_RAWシートに書き込みました`);
  }
}

// ========================================
// Phase 2: 期間指定データ取得
// ========================================

/**
 * 指定期間のGA4データを取得
 * 
 * @param {string} startDate - 開始日（YYYY-MM-DD形式）
 * @param {string} endDate - 終了日（YYYY-MM-DD形式）
 * @param {string} pageUrl - オプション: 特定ページのみ取得
 * @return {Array} ページ別データ配列
 */
function fetchGA4DataForDateRange(startDate, endDate, pageUrl) {
  try {
    Logger.log('=== GA4期間指定データ取得開始 ===');
    Logger.log('期間: ' + startDate + ' 〜 ' + endDate);
    if (pageUrl) Logger.log('対象URL: ' + pageUrl);
    
    var propertyId = PropertiesService.getScriptProperties().getProperty('GA4_PROPERTY_ID');
    
    if (!propertyId) {
      throw new Error('GA4_PROPERTY_IDが設定されていません');
    }
    
    // リクエスト構築
    var request = {
      dateRanges: [{
        startDate: startDate,
        endDate: endDate
      }],
      dimensions: [
        { name: 'pagePath' },
        { name: 'pageTitle' }
      ],
      metrics: [
        { name: 'screenPageViews' },
        { name: 'activeUsers' },
        { name: 'averageSessionDuration' },
        { name: 'bounceRate' },
        { name: 'conversions' }
      ],
      limit: 10000
    };
    
    // 特定ページのフィルタ追加
    if (pageUrl) {
      request.dimensionFilter = {
        filter: {
          fieldName: 'pagePath',
          stringFilter: {
            matchType: 'EXACT',
            value: pageUrl
          }
        }
      };
    }
    
    // API呼び出し
    var response = AnalyticsData.Properties.runReport(request, propertyId);
    
    if (!response.rows || response.rows.length === 0) {
      Logger.log('データが見つかりませんでした');
      return [];
    }
    
    // データ整形
    var data = response.rows.map(function(row) {
      return {
        page_url: row.dimensionValues[0].value,
        page_title: row.dimensionValues[1].value,
        page_views: parseInt(row.metricValues[0].value) || 0,
        active_users: parseInt(row.metricValues[1].value) || 0,
        avg_session_duration: parseFloat(row.metricValues[2].value) || 0,
        bounce_rate: parseFloat(row.metricValues[3].value) * 100 || 0,
        conversions: parseInt(row.metricValues[4].value) || 0
      };
    });
    
    Logger.log('GA4データ取得完了: ' + data.length + '件');
    return data;
    
  } catch (error) {
    Logger.log('GA4期間指定データ取得エラー: ' + error.message);
    throw error;
  }
}

/**
 * 指定期間のSearch Consoleデータを取得（UrlFetchApp方式）
 * 
 * @param {string} startDate - 開始日（YYYY-MM-DD形式）
 * @param {string} endDate - 終了日（YYYY-MM-DD形式）
 * @param {string} pageUrl - オプション: 特定ページのみ取得
 * @return {Array} ページ別データ配列
 */
function fetchGSCDataForDateRange(startDate, endDate, pageUrl) {
  try {
    Logger.log('=== GSC期間指定データ取得開始 ===');
    Logger.log('期間: ' + startDate + ' 〜 ' + endDate);
    if (pageUrl) Logger.log('対象URL: ' + pageUrl);
    
    var siteUrl = PropertiesService.getScriptProperties().getProperty('GSC_SITE_URL');
    
    if (!siteUrl) {
      throw new Error('GSC_SITE_URLが設定されていません');
    }
    
    // リクエスト構築
    var requestBody = {
      startDate: startDate,
      endDate: endDate,
      dimensions: ['page'],
      rowLimit: 25000,
      dataState: 'final'
    };
    
    // 特定ページのフィルタ追加
    if (pageUrl) {
      var fullUrl = siteUrl.replace(/\/$/, '') + pageUrl;
      requestBody.dimensionFilterGroups = [{
        filters: [{
          dimension: 'page',
          operator: 'equals',
          expression: fullUrl
        }]
      }];
    }
    
    // API URL
    var apiUrl = 'https://www.googleapis.com/webmasters/v3/sites/' + 
                 encodeURIComponent(siteUrl) + '/searchAnalytics/query';
    
    // API呼び出し（UrlFetchApp）
    var options = {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
      },
      payload: JSON.stringify(requestBody),
      muteHttpExceptions: true
    };
    
    var response = UrlFetchApp.fetch(apiUrl, options);
    var responseCode = response.getResponseCode();
    
    if (responseCode !== 200) {
      throw new Error('GSC API エラー: ' + responseCode + ' - ' + response.getContentText());
    }
    
    var result = JSON.parse(response.getContentText());
    
    if (!result.rows || result.rows.length === 0) {
      Logger.log('データが見つかりませんでした');
      return [];
    }
    
    // データ整形（URL単位で集計）
    var pageData = {};
    
    result.rows.forEach(function(row) {
      var url = row.keys[0];
      var shortUrl = url.replace(siteUrl.replace(/\/$/, ''), '');
      
      if (!pageData[shortUrl]) {
        pageData[shortUrl] = {
          page_url: shortUrl,
          clicks: 0,
          impressions: 0,
          position_sum: 0,
          count: 0
        };
      }
      
      pageData[shortUrl].clicks += row.clicks || 0;
      pageData[shortUrl].impressions += row.impressions || 0;
      pageData[shortUrl].position_sum += row.position || 0;
      pageData[shortUrl].count += 1;
    });
    
    // 配列に変換・平均計算
    var data = Object.keys(pageData).map(function(url) {
      var page = pageData[url];
      return {
        page_url: page.page_url,
        clicks: page.clicks,
        impressions: page.impressions,
        avg_position: page.position_sum / page.count,
        ctr: page.impressions > 0 ? (page.clicks / page.impressions) : 0
      };
    });
    
    Logger.log('GSCデータ取得完了: ' + data.length + '件');
    return data;
    
  } catch (error) {
    Logger.log('GSC期間指定データ取得エラー: ' + error.message);
    throw error;
  }
}

/**
 * 指定期間の統合データを取得
 * GA4とGSCデータを結合
 * 
 * @param {string} startDate - 開始日（YYYY-MM-DD形式）
 * @param {string} endDate - 終了日（YYYY-MM-DD形式）
 * @param {string} pageUrl - オプション: 特定ページのみ取得
 * @return {Array} 統合データ配列
 */
function fetchIntegratedDataForDateRange(startDate, endDate, pageUrl) {
  try {
    Logger.log('=== 期間指定統合データ取得開始 ===');
    
    // GA4とGSCデータを並行取得
    var ga4Data = fetchGA4DataForDateRange(startDate, endDate, pageUrl);
    var gscData = fetchGSCDataForDateRange(startDate, endDate, pageUrl);
    
    // URL単位でデータ統合
    var integratedData = {};
    
    // GA4データを追加
    ga4Data.forEach(function(row) {
      var url = row.page_url;
      if (!integratedData[url]) {
        integratedData[url] = {};
      }
      Object.keys(row).forEach(function(key) {
        integratedData[url][key] = row[key];
      });
    });
    
    // GSCデータを追加
    gscData.forEach(function(row) {
      var url = row.page_url;
      if (!integratedData[url]) {
        integratedData[url] = { page_url: url };
      }
      Object.keys(row).forEach(function(key) {
        if (key !== 'page_url') {
          integratedData[url][key] = row[key];
        }
      });
    });
    
    // 配列に変換
    var result = Object.keys(integratedData).map(function(url) {
      return integratedData[url];
    });
    
    Logger.log('統合データ作成完了: ' + result.length + '件');
    return result;
    
  } catch (error) {
    Logger.log('期間指定統合データ取得エラー: ' + error.message);
    throw error;
  }
}

/**
 * Before/Afterの統計値を計算
 * 
 * @param {Array} data - データ配列
 * @return {Object} 統計値
 */
function calculatePeriodStats(data) {
  if (!data || data.length === 0) {
    return {
      total_pages: 0,
      total_page_views: 0,
      avg_page_views: 0,
      total_clicks: 0,
      total_impressions: 0,
      avg_position: 0,
      avg_ctr: 0,
      avg_bounce_rate: 0,
      avg_session_duration: 0
    };
  }
  
  var totalPV = 0;
  var totalClicks = 0;
  var totalImpressions = 0;
  var totalPosition = 0;
  var totalBounceRate = 0;
  var totalSessionDuration = 0;
  var countPosition = 0;
  var countBounce = 0;
  var countSession = 0;
  
  data.forEach(function(row) {
    totalPV += row.page_views || 0;
    totalClicks += row.clicks || 0;
    totalImpressions += row.impressions || 0;
    
    if (row.avg_position && row.avg_position > 0) {
      totalPosition += row.avg_position;
      countPosition++;
    }
    
    if (row.bounce_rate && row.bounce_rate > 0) {
      totalBounceRate += row.bounce_rate;
      countBounce++;
    }
    
    if (row.avg_session_duration && row.avg_session_duration > 0) {
      totalSessionDuration += row.avg_session_duration;
      countSession++;
    }
  });
  
  return {
    total_pages: data.length,
    total_page_views: totalPV,
    avg_page_views: data.length > 0 ? totalPV / data.length : 0,
    total_clicks: totalClicks,
    total_impressions: totalImpressions,
    avg_position: countPosition > 0 ? totalPosition / countPosition : 0,
    avg_ctr: totalImpressions > 0 ? (totalClicks / totalImpressions) * 100 : 0,
    avg_bounce_rate: countBounce > 0 ? totalBounceRate / countBounce : 0,
    avg_session_duration: countSession > 0 ? totalSessionDuration / countSession : 0
  };
}

// ========================================
// データ統合（定期実行用）
// ========================================

/**
 * GA4とGSCのデータをURL単位で統合
 */
function integrateData() {
  const startTime = new Date();
  Logger.log('=== データ統合開始 ===');
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // RAWシート取得
    const ga4Sheet = ss.getSheetByName('GA4_RAW');
    const gscSheet = ss.getSheetByName('GSC_RAW');
    const integratedSheet = ss.getSheetByName('統合データ');
    
    if (!ga4Sheet || !gscSheet || !integratedSheet) {
      throw new Error('必要なシートが見つかりません');
    }
    
    // GA4データ読み込み
    Logger.log('GA4データ読み込み中...');
    const ga4Data = ga4Sheet.getDataRange().getValues();
    const ga4Headers = ga4Data[0];
    const ga4Rows = ga4Data.slice(1);
    
    // GSCデータ読み込み
    Logger.log('GSCデータ読み込み中...');
    const gscData = gscSheet.getDataRange().getValues();
    const gscHeaders = gscData[0];
    const gscRows = gscData.slice(1);
    
    Logger.log(`GA4: ${ga4Rows.length}行, GSC: ${gscRows.length}行`);
    
    // URLをキーにしたマップを作成
    Logger.log('データ集計中...');
    const urlMap = new Map();
    
    // GA4データをマップに追加
    ga4Rows.forEach(row => {
      const url = normalizeUrl(row[1]); // page_path
      if (!urlMap.has(url)) {
        urlMap.set(url, {
          page_path: row[1],
          page_title: row[2],
          page_views: row[3],
          unique_users: row[4],
          avg_session_duration: row[5],
          bounce_rate: row[6],
          conversions: row[7],
          // GSC指標（後で追加）
          total_clicks: 0,
          total_impressions: 0,
          avg_position: 0,
          avg_ctr: 0,
          top_queries: []
        });
      }
    });
    
    // GSCデータを集計してマップに追加
    gscRows.forEach(row => {
      const url = normalizeUrl(row[1]); // page_url
      const query = row[2];
      const position = row[3];
      const clicks = row[4];
      const impressions = row[5];
      const ctr = row[6];
      
      if (urlMap.has(url)) {
        const pageData = urlMap.get(url);
        pageData.total_clicks += clicks;
        pageData.total_impressions += impressions;
        
        // 上位クエリを保存（クリック数上位5件）
        if (pageData.top_queries.length < 5) {
          pageData.top_queries.push({ query, clicks, position });
        }
      } else {
        // GA4にないがGSCにあるページ
        urlMap.set(url, {
          page_path: url,
          page_title: '',
          page_views: 0,
          unique_users: 0,
          avg_session_duration: 0,
          bounce_rate: 0,
          conversions: 0,
          total_clicks: clicks,
          total_impressions: impressions,
          avg_position: position,
          avg_ctr: ctr,
          top_queries: [{ query, clicks, position }]
        });
      }
    });
    
    // GSC指標の平均を計算
    gscRows.forEach(row => {
      const url = normalizeUrl(row[1]);
      if (urlMap.has(url)) {
        const pageData = urlMap.get(url);
        const urlGscData = gscRows.filter(r => normalizeUrl(r[1]) === url);
        
        // 平均順位（表示回数で加重平均）
        let totalWeightedPosition = 0;
        let totalImpressions = 0;
        urlGscData.forEach(r => {
          totalWeightedPosition += r[3] * r[5]; // position * impressions
          totalImpressions += r[5];
        });
        pageData.avg_position = totalImpressions > 0 
          ? parseFloat((totalWeightedPosition / totalImpressions).toFixed(1))
          : 0;
        
        // 平均CTR
        pageData.avg_ctr = pageData.total_impressions > 0
          ? parseFloat(((pageData.total_clicks / pageData.total_impressions) * 100).toFixed(2))
          : 0;
      }
    });
    
    Logger.log(`統合URL数: ${urlMap.size}件`);
    
    // 統合データシートに書き込み
    Logger.log('統合データシート書き込み中...');
    
    // ヘッダー行作成（初回のみ）
    if (integratedSheet.getLastRow() === 0) {
      const headers = [
        'page_url', 'page_title', 'category', 'publish_date', 'last_modified',
        // GA4指標
        'avg_page_views_30d', 'avg_session_duration', 'bounce_rate', 'conversions_30d',
        // GSC指標
        'avg_position', 'total_clicks_30d', 'total_impressions_30d', 'avg_ctr', 'top_queries',
        // その他
        'last_analyzed'
      ];
      integratedSheet.appendRow(headers);
      integratedSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    }
    
    // データ行作成
    const timestamp = new Date();
    const outputData = [];
    
    urlMap.forEach((data, url) => {
      const topQueriesStr = data.top_queries
        .map(q => q.query)
        .join(', ');
      
      outputData.push([
        url,                              // page_url
        data.page_title,                  // page_title
        '',                               // category (後で手動入力)
        '',                               // publish_date (後で手動入力)
        '',                               // last_modified (後で手動入力)
        data.page_views,                  // avg_page_views_30d
        data.avg_session_duration,        // avg_session_duration
        data.bounce_rate,                 // bounce_rate
        data.conversions,                 // conversions_30d
        data.avg_position,                // avg_position
        data.total_clicks,                // total_clicks_30d
        data.total_impressions,           // total_impressions_30d
        data.avg_ctr,                     // avg_ctr
        topQueriesStr,                    // top_queries
        timestamp                         // last_analyzed
      ]);
    });
    
    // バッチ書き込み
    if (outputData.length > 0) {
      const startRow = integratedSheet.getLastRow() + 1;
      integratedSheet.getRange(startRow, 1, outputData.length, outputData[0].length)
        .setValues(outputData);
      
      Logger.log(`${outputData.length}行を統合データシートに書き込みました`);
    }
    
    const endTime = new Date();
    const duration = (endTime - startTime) / 1000;
    
    Logger.log(`=== データ統合完了 ===`);
    Logger.log(`処理時間: ${duration}秒`);
    Logger.log(`統合URL数: ${outputData.length}件`);
    
    return { success: true, urlCount: outputData.length, duration: duration };
    
  } catch (error) {
    Logger.log(`❌ データ統合エラー: ${error.message}`);
    Logger.log(`スタックトレース: ${error.stack}`);
    throw error;
  }
}

// ========================================
// ユーティリティ関数
// ========================================

/**
 * 日付文字列を取得（yyyy-MM-dd形式）
 */
function getDateString(daysAgo) {
  const date = new Date();
  date.setDate(date.getDate() - daysAgo);
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/**
 * URLを正規化（比較用）
 */
function normalizeUrl(url) {
  if (!url) return '';
  
  // プロトコルとドメインを除去してパスのみに
  let normalized = url.toString()
    .replace(/^https?:\/\/[^\/]+/, '')  // プロトコルとドメイン削除
    .replace(/\/$/, '')                  // 末尾のスラッシュ削除
    .replace(/\?.*$/, '')                // クエリパラメータ削除
    .replace(/#.*$/, '');                // ハッシュ削除
  
  // 空の場合はトップページ
  if (normalized === '' || normalized === '/') {
    normalized = '/';
  }
  
  return normalized;
}

/**
 * エラー通知を送信
 */
function sendErrorNotification(processName, error) {
  try {
    const email = Session.getActiveUser().getEmail();
    const subject = `[SEOツール] ${processName}でエラー発生`;
    const body = `
エラーが発生しました:

プロセス: ${processName}
エラーメッセージ: ${error.message}
発生時刻: ${new Date()}

スタックトレース:
${error.stack}
`;
    
    MailApp.sendEmail(email, subject, body);
  } catch (e) {
    Logger.log(`エラー通知の送信に失敗: ${e.message}`);
  }
}

// ========================================
// 定期実行
// ========================================

/**
 * 毎日のデータ更新処理
 * トリガーで自動実行
 */
function dailyUpdate() {
  const startTime = new Date();
  Logger.log('========================================');
  Logger.log('=== 毎日のデータ更新開始 ===');
  Logger.log(`実行日時: ${startTime}`);
  Logger.log('========================================');
  
  try {
    // 1. GA4データ取得
    Logger.log('');
    Logger.log('【1/4】GA4データ取得');
    const ga4Result = fetchGA4Data();
    Logger.log(`✅ GA4: ${ga4Result.rowCount}行取得完了`);
    
    // 少し待機（API制限対策）
    Utilities.sleep(2000);
    
    // 2. Search Consoleデータ取得
    Logger.log('');
    Logger.log('【2/4】Search Consoleデータ取得');
    const gscResult = fetchGSCData();
    Logger.log(`✅ GSC: ${gscResult.rowCount}行取得完了`);
    
    // 少し待機
    Utilities.sleep(2000);
    
    // 3. データ統合
    Logger.log('');
    Logger.log('【3/4】データ統合');
    const integrationResult = integrateData();
    Logger.log(`✅ 統合: ${integrationResult.urlCount}件完了`);
    
    // 4. GyronSEO順位同期（1位は自動除外）
    Logger.log('');
    Logger.log('【4/4】GyronSEO順位同期');
    const gyronResult = syncGyronPosition();
    Logger.log(`✅ GyronSEO: ${gyronResult.updated}件同期、${gyronResult.excluded}件除外`);
    
    const endTime = new Date();
    const totalDuration = (endTime - startTime) / 1000;
    
    Logger.log('');
    Logger.log('========================================');
    Logger.log('=== 毎日のデータ更新完了 ===');
    Logger.log(`総処理時間: ${totalDuration}秒（${(totalDuration/60).toFixed(1)}分）`);
    Logger.log(`完了日時: ${endTime}`);
    Logger.log('========================================');
    
    // 成功通知（オプション）
    sendSuccessNotification({
      ga4Rows: ga4Result.rowCount,
      gscRows: gscResult.rowCount,
      integratedUrls: integrationResult.urlCount,
      gyronUpdated: gyronResult.updated,
      gyronExcluded: gyronResult.excluded,
      duration: totalDuration
    });
    
    return { success: true };
    
  } catch (error) {
    Logger.log('');
    Logger.log('❌❌❌ エラー発生 ❌❌❌');
    Logger.log(`エラー: ${error.message}`);
    Logger.log(`スタックトレース: ${error.stack}`);
    
    // エラー通知
    sendErrorNotification('毎日のデータ更新', error);
    
    throw error;
  }
}

/**
 * 成功通知を送信
 */
function sendSuccessNotification(results) {
  try {
    const email = Session.getActiveUser().getEmail();
    const subject = '[SEOツール] データ更新完了';
    const body = `
データ更新が正常に完了しました。

【更新内容】
- GA4データ: ${results.ga4Rows}行
- Search Consoleデータ: ${results.gscRows}行
- 統合URL数: ${results.integratedUrls}件
- GyronSEO順位同期: ${results.gyronUpdated}件
- 1位で除外: ${results.gyronExcluded}件

処理時間: ${results.duration}秒（${(results.duration/60).toFixed(1)}分）
実行日時: ${new Date()}

スプレッドシートを確認してください。
`;
    
    MailApp.sendEmail(email, subject, body);
  } catch (e) {
    Logger.log(`成功通知の送信に失敗: ${e.message}`);
  }
}

/**
 * トリガーを作成（初回のみ実行）
 */
function createDailyTrigger() {
  // 既存のdailyUpdateトリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'dailyUpdate') {
      ScriptApp.deleteTrigger(trigger);
      Logger.log('既存のトリガーを削除しました');
    }
  });
  
  // 新規トリガー作成（毎日午前5時）
  ScriptApp.newTrigger('dailyUpdate')
    .timeBased()
    .atHour(5)
    .everyDays(1)
    .create();
  
  Logger.log('✅ 毎日午前5時に実行するトリガーを作成しました');
}

// ========================================
// テスト関数
// ========================================

/**
 * テスト: 期間指定データ取得
 */
function testFetchDataForDateRange() {
  var startDate = '2025-10-15';
  var endDate = '2025-10-21';
  
  Logger.log('=== 期間指定データ取得テスト ===');
  Logger.log('期間: ' + startDate + ' 〜 ' + endDate);
  
  try {
    var data = fetchIntegratedDataForDateRange(startDate, endDate);
    
    Logger.log('取得件数: ' + data.length);
    
    if (data.length > 0) {
      Logger.log('サンプルデータ:');
      Logger.log(JSON.stringify(data[0], null, 2));
    }
    
    var stats = calculatePeriodStats(data);
    Logger.log('\n統計値:');
    Logger.log('総PV: ' + stats.total_page_views);
    Logger.log('平均PV: ' + stats.avg_page_views.toFixed(1));
    Logger.log('平均順位: ' + stats.avg_position.toFixed(1) + '位');
    Logger.log('平均CTR: ' + stats.avg_ctr.toFixed(2) + '%');
    Logger.log('平均直帰率: ' + stats.avg_bounce_rate.toFixed(1) + '%');
    
    return data;
    
  } catch (error) {
    Logger.log('テスト失敗: ' + error.message);
    return null;
  }
}

/**
 * GA4から初回PV日を取得してpublish_dateを更新
 */
function updatePublishDatesFromGA4() {
  Logger.log('=== publish_date更新開始（GA4から取得） ===');
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ga4Sheet = ss.getSheetByName('GA4_RAW');
    const integratedSheet = ss.getSheetByName('統合データ');
    
    if (!ga4Sheet || !integratedSheet) {
      throw new Error('必要なシートが見つかりません');
    }
    
    // GA4_RAWからデータ取得
    const ga4Data = ga4Sheet.getDataRange().getValues();
    const ga4Headers = ga4Data[0];
    const dateIdx = ga4Headers.indexOf('date');
    const pathIdx = ga4Headers.indexOf('page_path');
    
    if (dateIdx === -1 || pathIdx === -1) {
      throw new Error('GA4_RAWに必要な列がありません');
    }
    
    // ページごとの最古の日付を計算
    const firstPVDates = {};
    
    for (let i = 1; i < ga4Data.length; i++) {
      const pagePath = String(ga4Data[i][pathIdx] || '').trim();
      const date = ga4Data[i][dateIdx];
      
      if (!pagePath || !date) continue;
      
      // 初めて見るページ、または既存より古い日付の場合
      if (!firstPVDates[pagePath] || date < firstPVDates[pagePath]) {
        firstPVDates[pagePath] = date;
      }
    }
    
    Logger.log(`初回PV日を計算: ${Object.keys(firstPVDates).length}ページ`);
    
    // 統合データシートを更新
    const integratedData = integratedSheet.getDataRange().getValues();
    const integratedHeaders = integratedData[0];
    const urlIdx = integratedHeaders.indexOf('page_url');
    const publishDateIdx = integratedHeaders.indexOf('publish_date');
    
    if (urlIdx === -1 || publishDateIdx === -1) {
      throw new Error('統合データシートに必要な列がありません');
    }
    
    let updatedCount = 0;
    
    for (let i = 1; i < integratedData.length; i++) {
      const pageUrl = String(integratedData[i][urlIdx] || '').trim();
      
      if (firstPVDates[pageUrl]) {
        // publish_date列を更新
        integratedSheet.getRange(i + 1, publishDateIdx + 1).setValue(firstPVDates[pageUrl]);
        updatedCount++;
      }
    }
    
    Logger.log(`publish_date更新完了: ${updatedCount}件`);
    Logger.log('=== 処理完了 ===');
    
    return {
      success: true,
      updatedCount: updatedCount,
      totalPages: Object.keys(firstPVDates).length
    };
    
  } catch (error) {
    Logger.log(`エラー: ${error.message}`);
    throw error;
  }
}

/**
 * テスト実行（最初の10ページのみ）
 */
function testUpdatePublishDates() {
  Logger.log('=== テスト実行: publish_date更新 ===');
  
  const result = updatePublishDatesFromGA4();
  
  if (result.success) {
    Logger.log('✅ 正常完了');
    Logger.log(`更新件数: ${result.updatedCount}`);
  }
}


/**
 * GyronSEO順位連携関数（v6）
 * 
 * ★v6の変更点★
 * 1ページに複数キーワードがある場合の除外ロジックを改善
 * 
 * 【除外ロジック】
 * 1. 1ページに1KWのみ && 1位 → リライト対象外
 * 2. 複数KW && 1位のKWが最大検索ボリューム → リライト対象外
 * 3. 複数KW && 1位のKW検索ボリューム >= 500 → リライト対象外
 * 4. それ以外 → リライト対象
 * 
 * 追加日: 2025/12/02
 * 修正日: 2025/12/02（複数KW対応・検索ボリューム考慮）
 */

/**
 * GyronSEO順位を統合データシートに同期（v6: 複数KW対応）
 * @return {Object} 同期結果
 */
function syncGyronPosition() {
  Logger.log('=== GyronSEO順位同期開始（v6: 複数KW対応） ===');
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var integratedSheet = ss.getSheetByName('統合データ');
  var gyronSheet = ss.getSheetByName('GyronSEO_RAW');
  var volumeSheet = ss.getSheetByName('検索ボリューム_RAW');
  
  if (!integratedSheet) {
    Logger.log('エラー: 統合データシートが見つかりません');
    return { success: false, error: '統合データシートが見つかりません', excluded: 0 };
  }
  
  if (!gyronSheet) {
    Logger.log('エラー: GyronSEO_RAWシートが見つかりません');
    return { success: false, error: 'GyronSEO_RAWシートが見つかりません', excluded: 0 };
  }
  
  if (!volumeSheet) {
    Logger.log('警告: 検索ボリューム_RAWシートが見つかりません（検索ボリューム考慮なしで続行）');
  }
  
  // ========================================
  // 1. 検索ボリュームデータを読み込み
  // ========================================
  var volumeMap = {};
  if (volumeSheet) {
    var volumeData = volumeSheet.getDataRange().getValues();
    var volumeHeaders = volumeData[0];
    var kwIndex = 0;  // ランクインキーワード
    var volIndex = 1; // 検索ボリューム
    
    for (var v = 1; v < volumeData.length; v++) {
      var kw = String(volumeData[v][kwIndex] || '').trim();
      var vol = parseFloat(volumeData[v][volIndex]) || 0;
      
      if (kw) {
        var normalizedKW = normalizeKeywordForGyron(kw);
        volumeMap[normalizedKW] = vol;
      }
    }
    Logger.log('検索ボリュームデータ: ' + Object.keys(volumeMap).length + '件');
  }
  
  // ========================================
  // 2. GyronSEO_RAWからURL別キーワードをグルーピング
  // ========================================
  var gyronData = gyronSheet.getDataRange().getValues();
  var gyronHeaders = gyronData[0];
  
  // キーワード列（A列）
  var gyronKeywordIndex = gyronHeaders.indexOf('キーワード');
  if (gyronKeywordIndex === -1) gyronKeywordIndex = 0;
  
  // URL列（B列）
  var gyronUrlIndex = gyronHeaders.indexOf('URL');
  if (gyronUrlIndex === -1) gyronUrlIndex = 1;
  
  // 最新の日付列（順位データ）
  var latestDateColIndex = -1;
  var latestDate = null;
  
  for (var col = 0; col < gyronHeaders.length; col++) {
    var header = gyronHeaders[col];
    if (header instanceof Date) {
      if (!latestDate || header > latestDate) {
        latestDate = header;
        latestDateColIndex = col;
      }
    }
  }
  
  if (latestDateColIndex === -1) {
    Logger.log('エラー: 日付列が見つかりません');
    return { success: false, error: '日付列が見つかりません', excluded: 0 };
  }
  
  var latestDateStr = Utilities.formatDate(latestDate, 'Asia/Tokyo', 'yyyy-MM-dd');
  Logger.log('最新日付列: ' + (latestDateColIndex + 1) + '列目（' + latestDateStr + '）');
  
  // URL別にキーワード・順位・検索ボリュームをグルーピング
  var urlKeywordsMap = {};  // { url: [{ keyword, position, volume }, ...] }
  
  for (var i = 1; i < gyronData.length; i++) {
    var keyword = String(gyronData[i][gyronKeywordIndex] || '').trim();
    var url = String(gyronData[i][gyronUrlIndex] || '').trim();
    var position = gyronData[i][latestDateColIndex];
    
    if (!keyword) continue;
    
    // 順位を数値に変換
    var positionNum = null;
    if (position !== '' && position !== null && position !== undefined) {
      if (String(position).includes('圏外')) {
        positionNum = 101;
      } else {
        positionNum = parseFloat(position);
        if (isNaN(positionNum)) positionNum = null;
      }
    }
    
    // URLからパスを抽出
    var path = extractPathFromUrl(url);
    if (!path) continue;
    
    // 検索ボリュームを取得
    var normalizedKW = normalizeKeywordForGyron(keyword);
    var volume = volumeMap[normalizedKW] || 0;
    
    if (!urlKeywordsMap[path]) {
      urlKeywordsMap[path] = [];
    }
    
    urlKeywordsMap[path].push({
      keyword: keyword,
      normalizedKeyword: normalizedKW,
      position: positionNum,
      volume: volume
    });
  }
  
  Logger.log('URL別キーワードグループ: ' + Object.keys(urlKeywordsMap).length + '件');
  
  // ========================================
  // 3. 統合データシートのヘッダー取得
  // ========================================
  var integratedHeaders = integratedSheet.getRange(1, 1, 1, integratedSheet.getLastColumn()).getValues()[0];
  var urlIndex = integratedHeaders.indexOf('page_url');
  var targetKWIndex = integratedHeaders.indexOf('target_keyword');
  var gyronPositionIndex = integratedHeaders.indexOf('gyron_position');
  var totalScoreIndex = integratedHeaders.indexOf('total_priority_score');
  var exclusionReasonIndex = integratedHeaders.indexOf('exclusion_reason');
  
  // gyron_position列がなければ追加
  if (gyronPositionIndex === -1) {
    var newColIndex = integratedSheet.getLastColumn() + 1;
    integratedSheet.getRange(1, newColIndex).setValue('gyron_position');
    gyronPositionIndex = newColIndex - 1;
    Logger.log('gyron_position列を追加しました（' + newColIndex + '列目）');
  }
  
  // ========================================
  // 4. 各ページの除外判定
  // ========================================
  var integratedData = integratedSheet.getDataRange().getValues();
  var updateCount = 0;
  var excludedCount = 0;
  var excludedDetails = [];
  
  for (var j = 1; j < integratedData.length; j++) {
    var pageUrl = String(integratedData[j][urlIndex] || '').trim();
    var targetKW = String(integratedData[j][targetKWIndex] || '').trim();
    
    if (!pageUrl) continue;
    
    // URLを正規化
    var pagePath = pageUrl.startsWith('/') ? pageUrl : '/' + pageUrl;
    
    // このURLに紐づくキーワード情報を取得
    var keywordsForPage = urlKeywordsMap[pagePath] || [];
    
    // target_keywordの順位を取得
    var normalizedTargetKW = normalizeKeywordForGyron(targetKW);
    var targetKWInfo = null;
    
    for (var k = 0; k < keywordsForPage.length; k++) {
      if (keywordsForPage[k].normalizedKeyword === normalizedTargetKW) {
        targetKWInfo = keywordsForPage[k];
        break;
      }
    }
    
    var gyronPosition = targetKWInfo ? targetKWInfo.position : null;
    
    // 圏外は空にする
    if (gyronPosition === 101) {
      gyronPosition = null;
    }
    
    // gyron_position列に書き込み
    integratedSheet.getRange(j + 1, gyronPositionIndex + 1).setValue(gyronPosition || '');
    
    if (gyronPosition) {
      updateCount++;
    }
    
    // ========================================
    // 除外判定ロジック
    // ========================================
    var shouldExclude = false;
    var exclusionReason = '';
    
    // 1位のキーワードを探す
    var rank1Keywords = keywordsForPage.filter(function(kw) {
      return kw.position === 1;
    });
    
    if (rank1Keywords.length > 0) {
      // 1位のキーワードが存在する
      var kwCount = keywordsForPage.length;
      
      // 最大検索ボリュームを持つKWを特定
      var maxVolumeKW = keywordsForPage.reduce(function(max, kw) {
        return (kw.volume > max.volume) ? kw : max;
      }, { volume: 0 });
      
      // 1位KWの中で最大検索ボリュームのものを特定
      var maxRank1KW = rank1Keywords.reduce(function(max, kw) {
        return (kw.volume > max.volume) ? kw : max;
      }, { volume: 0 });
      
      if (kwCount === 1) {
        // ケース1: 1ページに1KWのみ && 1位
        shouldExclude = true;
        exclusionReason = '1位獲得中のため除外（単一KW）';
      } else if (maxRank1KW.normalizedKeyword === maxVolumeKW.normalizedKeyword && maxVolumeKW.volume > 0) {
        // ケース2: 1位のKWが最大検索ボリューム
        shouldExclude = true;
        exclusionReason = '1位獲得中のため除外（最大ボリュームKW: ' + maxRank1KW.volume + '）';
      } else if (maxRank1KW.volume >= 500) {
        // ケース3: 1位のKW検索ボリューム >= 500
        shouldExclude = true;
        exclusionReason = '1位獲得中のため除外（ボリューム' + maxRank1KW.volume + ' >= 500）';
      } else {
        // ケース4: 1位だが小ボリュームで最大ではない → リライト対象
        shouldExclude = false;
      }
    }
    
    // 除外処理
    if (shouldExclude) {
      if (totalScoreIndex >= 0) {
        integratedSheet.getRange(j + 1, totalScoreIndex + 1).setValue(0);
      }
      if (exclusionReasonIndex >= 0) {
        integratedSheet.getRange(j + 1, exclusionReasonIndex + 1).setValue(exclusionReason);
      }
      excludedCount++;
      if (excludedDetails.length < 5) {
        excludedDetails.push(pageUrl + ' (' + exclusionReason + ')');
      }
    } else {
      // 除外されない場合、既存の除外理由が「1位獲得中」なら解除
      if (exclusionReasonIndex >= 0) {
        var currentReason = String(integratedData[j][exclusionReasonIndex] || '');
        if (currentReason.includes('1位獲得中')) {
          integratedSheet.getRange(j + 1, exclusionReasonIndex + 1).setValue('');
        }
      }
    }
  }
  
  Logger.log('更新完了: ' + updateCount + '件');
  Logger.log('★除外: ' + excludedCount + '件');
  if (excludedDetails.length > 0) {
    Logger.log('除外例: ');
    excludedDetails.forEach(function(detail) {
      Logger.log('  - ' + detail);
    });
  }
  Logger.log('=== GyronSEO順位同期完了 ===');
  
  return {
    success: true,
    updated: updateCount,
    excluded: excludedCount
  };
}

/**
 * URLからパス部分を抽出
 * @param {String} url - フルURL
 * @return {String} パス部分
 */
function extractPathFromUrl(url) {
  if (!url) return '';
  
  try {
    if (url.startsWith('/')) {
      return url;
    }
    
    var match = url.match(/https?:\/\/[^\/]+(\/[^\?#]*)?/);
    if (match && match[1]) {
      return match[1];
    }
    
    if (!url.includes('://')) {
      return '/' + url;
    }
    
    return '';
  } catch (e) {
    return '';
  }
}

/**
 * キーワードを正規化（GyronSEO用）
 * @param {String} keyword - キーワード
 * @return {String} 正規化されたキーワード
 */
function normalizeKeywordForGyron(keyword) {
  if (!keyword) return '';
  
  return String(keyword)
    .toLowerCase()
    .replace(/　/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * GyronSEO順位同期のテスト
 */
function testSyncGyronPosition() {
  var result = syncGyronPosition();
  Logger.log('同期結果: ' + JSON.stringify(result));
}

/**
 * 除外ロジックのデバッグ
 * 複数KWを持つページの除外判定を確認
 */
function debugExclusionLogic() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var gyronSheet = ss.getSheetByName('GyronSEO_RAW');
  var volumeSheet = ss.getSheetByName('検索ボリューム_RAW');
  
  // 検索ボリュームマップ作成
  var volumeMap = {};
  if (volumeSheet) {
    var volumeData = volumeSheet.getDataRange().getValues();
    for (var v = 1; v < volumeData.length; v++) {
      var kw = String(volumeData[v][0] || '').trim();
      var vol = parseFloat(volumeData[v][1]) || 0;
      if (kw) {
        volumeMap[normalizeKeywordForGyron(kw)] = vol;
      }
    }
  }
  
  // GyronSEOデータ読み込み
  var gyronData = gyronSheet.getDataRange().getValues();
  var gyronHeaders = gyronData[0];
  
  // 最新日付列
  var latestDateColIndex = -1;
  var latestDate = null;
  for (var col = 0; col < gyronHeaders.length; col++) {
    if (gyronHeaders[col] instanceof Date) {
      if (!latestDate || gyronHeaders[col] > latestDate) {
        latestDate = gyronHeaders[col];
        latestDateColIndex = col;
      }
    }
  }
  
  // URL別にグルーピング
  var urlKeywordsMap = {};
  for (var i = 1; i < gyronData.length; i++) {
    var keyword = String(gyronData[i][0] || '').trim();
    var url = String(gyronData[i][1] || '').trim();
    var position = gyronData[i][latestDateColIndex];
    
    if (!keyword || !url) continue;
    
    var positionNum = null;
    if (position !== '' && position !== null) {
      if (String(position).includes('圏外')) {
        positionNum = 101;
      } else {
        positionNum = parseFloat(position) || null;
      }
    }
    
    var path = extractPathFromUrl(url);
    if (!path) continue;
    
    var normalizedKW = normalizeKeywordForGyron(keyword);
    var volume = volumeMap[normalizedKW] || 0;
    
    if (!urlKeywordsMap[path]) {
      urlKeywordsMap[path] = [];
    }
    
    urlKeywordsMap[path].push({
      keyword: keyword,
      position: positionNum,
      volume: volume
    });
  }
  
  // 複数KWを持つURLを抽出
  Logger.log('=== 複数KWを持つURL（1位あり） ===');
  var count = 0;
  
  Object.keys(urlKeywordsMap).forEach(function(url) {
    var keywords = urlKeywordsMap[url];
    if (keywords.length > 1) {
      var rank1Keywords = keywords.filter(function(kw) { return kw.position === 1; });
      
      if (rank1Keywords.length > 0) {
        count++;
        Logger.log('');
        Logger.log(count + '. ' + url);
        Logger.log('   KW数: ' + keywords.length);
        
        keywords.forEach(function(kw) {
          var mark = kw.position === 1 ? '★1位' : kw.position + '位';
          Logger.log('   - ' + kw.keyword + ' (vol: ' + kw.volume + ', ' + mark + ')');
        });
        
        // 除外判定
        var maxVolumeKW = keywords.reduce(function(max, kw) {
          return (kw.volume > max.volume) ? kw : max;
        }, { volume: 0 });
        
        var maxRank1KW = rank1Keywords.reduce(function(max, kw) {
          return (kw.volume > max.volume) ? kw : max;
        }, { volume: 0 });
        
        var shouldExclude = false;
        var reason = '';
        
        if (normalizeKeywordForGyron(maxRank1KW.keyword) === normalizeKeywordForGyron(maxVolumeKW.keyword) && maxVolumeKW.volume > 0) {
          shouldExclude = true;
          reason = '最大ボリュームKWで1位';
        } else if (maxRank1KW.volume >= 500) {
          shouldExclude = true;
          reason = '1位KWのボリューム >= 500';
        } else {
          reason = '1位だが小ボリューム（' + maxRank1KW.volume + '）で最大ではない';
        }
        
        Logger.log('   → 判定: ' + (shouldExclude ? '除外' : 'リライト対象') + ' (' + reason + ')');
      }
    }
  });
  
  Logger.log('');
  Logger.log('複数KW + 1位あり: ' + count + '件');
}

/**
 * 統合データシートのgyron_position列を確認
 */
function checkGyronPositionColumn() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('統合データ');
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  var gyronPositionIndex = headers.indexOf('gyron_position');
  var urlIndex = headers.indexOf('page_url');
  var targetKWIndex = headers.indexOf('target_keyword');
  var totalScoreIndex = headers.indexOf('total_priority_score');
  var exclusionReasonIndex = headers.indexOf('exclusion_reason');
  
  if (gyronPositionIndex === -1) {
    Logger.log('gyron_position列は存在しません');
    return;
  }
  
  Logger.log('gyron_position列: ' + (gyronPositionIndex + 1) + '列目');
  
  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(2, 1, Math.min(lastRow - 1, 20), sheet.getLastColumn()).getValues();
  
  Logger.log('=== 上位20件のgyron_position ===');
  for (var i = 0; i < data.length; i++) {
    var url = data[i][urlIndex];
    var position = data[i][gyronPositionIndex];
    var targetKW = targetKWIndex >= 0 ? data[i][targetKWIndex] : '';
    var totalScore = totalScoreIndex >= 0 ? data[i][totalScoreIndex] : '';
    var exclusionReason = exclusionReasonIndex >= 0 ? data[i][exclusionReasonIndex] : '';
    
    var status = '';
    if (exclusionReason) {
      status = ' ★除外: ' + exclusionReason;
    }
    
    Logger.log((i + 1) + '. ' + url + ' → ' + (position || 'N/A') + '位 [' + targetKW + '] スコア:' + totalScore + status);
  }
}

/**
 * キーワードを正規化（GyronSEO用）
 * - 小文字に変換
 * - 全角スペースを半角に
 * - 連続スペースを1つに
 * - 前後の空白を削除
 * @param {String} keyword - キーワード
 * @return {String} 正規化されたキーワード
 */
function normalizeKeywordForGyron(keyword) {
  if (!keyword) return '';
  
  return String(keyword)
    .toLowerCase()
    .replace(/　/g, ' ')  // 全角スペースを半角に
    .replace(/\s+/g, ' ') // 連続スペースを1つに
    .trim();
}

/**
 * GyronSEO順位同期のテスト
 */
function testSyncGyronPosition() {
  var result = syncGyronPosition();
  Logger.log('同期結果: ' + JSON.stringify(result));
}

/**
 * 1位除外のテスト
 * 除外されたページを確認
 */
function testExcludedPages() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('統合データ');
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  var urlIndex = headers.indexOf('page_url');
  var targetKWIndex = headers.indexOf('target_keyword');
  var gyronPositionIndex = headers.indexOf('gyron_position');
  var totalScoreIndex = headers.indexOf('total_priority_score');
  var exclusionReasonIndex = headers.indexOf('exclusion_reason');
  
  var data = sheet.getDataRange().getValues();
  
  Logger.log('=== 1位で除外されたページ ===');
  var count = 0;
  
  for (var i = 1; i < data.length; i++) {
    var gyronPosition = data[i][gyronPositionIndex];
    
    if (gyronPosition === 1) {
      count++;
      var url = data[i][urlIndex];
      var targetKW = data[i][targetKWIndex];
      var totalScore = data[i][totalScoreIndex];
      var exclusionReason = exclusionReasonIndex >= 0 ? data[i][exclusionReasonIndex] : '';
      
      Logger.log(count + '. ' + url);
      Logger.log('   ターゲットKW: ' + targetKW);
      Logger.log('   順位: ' + gyronPosition + '位');
      Logger.log('   スコア: ' + totalScore + '点');
      Logger.log('   除外理由: ' + exclusionReason);
    }
  }
  
  Logger.log('合計: ' + count + '件が1位で除外');
}

/**
 * 統合データシートのgyron_position列を確認
 */
function checkGyronPositionColumn() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('統合データ');
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  var gyronPositionIndex = headers.indexOf('gyron_position');
  var urlIndex = headers.indexOf('page_url');
  var targetKWIndex = headers.indexOf('target_keyword');
  var totalScoreIndex = headers.indexOf('total_priority_score');
  
  if (gyronPositionIndex === -1) {
    Logger.log('gyron_position列は存在しません');
    return;
  }
  
  Logger.log('gyron_position列: ' + (gyronPositionIndex + 1) + '列目');
  
  // 順位データがある行を確認
  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(2, 1, Math.min(lastRow - 1, 20), sheet.getLastColumn()).getValues();
  
  Logger.log('=== 上位20件のgyron_position ===');
  var hasPositionCount = 0;
  for (var i = 0; i < data.length; i++) {
    var url = data[i][urlIndex];
    var position = data[i][gyronPositionIndex];
    var targetKW = targetKWIndex >= 0 ? data[i][targetKWIndex] : '';
    var totalScore = totalScoreIndex >= 0 ? data[i][totalScoreIndex] : '';
    
    if (position) {
      hasPositionCount++;
    }
    
    var status = '';
    if (position === 1) {
      status = ' ★除外（スコア: ' + totalScore + '）';
    }
    
    Logger.log((i + 1) + '. ' + url + ' → ' + (position || 'N/A') + '位 [' + targetKW + ']' + status);
  }
  Logger.log('順位データあり: ' + hasPositionCount + '件');
}