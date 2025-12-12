/**
 * DailyDataToSupabase.gs
 * GA4/GSCの日次データをSupabaseに蓄積するスクリプト
 * 
 * 【使い方】
 * 1. このファイルをGASに追加
 * 2. setupDailySupabaseTrigger() を1回実行してトリガー設定
 * 3. 毎朝5時に自動実行される
 * 
 * 【手動実行】
 * - runDailySupabaseUpdate() を実行
 */

const DAILY_CONFIG = {
  SUPABASE_URL: 'https://dgzfdugpineqnoihopsl.supabase.co',
  SITE_ID: '853ea711-7644-451e-872b-dea1b54fa8c7',
  GA4_PROPERTY_ID: 'properties/388689745',
  GSC_SITE_URL: 'https://smaho-tap.com'
};

/**
 * 日次更新のメイン関数（トリガーから呼ばれる）
 */
function runDailySupabaseUpdate() {
  Logger.log('=== 日次Supabase更新開始 ===');
  Logger.log(`実行日時: ${new Date().toLocaleString('ja-JP')}`);
  
  const serviceRoleKey = PropertiesService.getScriptProperties()
    .getProperty('SUPABASE_SERVICE_ROLE_KEY');
  
  if (!serviceRoleKey) {
    Logger.log('❌ Service Role Keyが設定されていません');
    return;
  }
  
  // ページマッピング取得
  const pageMapping = getPageMappingForDaily(serviceRoleKey);
  Logger.log(`ページマッピング: ${Object.keys(pageMapping).length}件`);
  
  // 前日の日付
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  const dateStr = formatDateForAPI(yesterday);
  
  Logger.log(`対象日: ${dateStr}`);
  
  // GA4データ取得・保存
  try {
    const ga4Count = fetchAndSaveGA4Daily(serviceRoleKey, pageMapping, dateStr);
    Logger.log(`✅ GA4: ${ga4Count}件保存`);
  } catch (e) {
    Logger.log(`❌ GA4エラー: ${e.message}`);
  }
  
  // GSCデータ取得・保存
  try {
    const gscCount = fetchAndSaveGSCDaily(serviceRoleKey, pageMapping, dateStr);
    Logger.log(`✅ GSC: ${gscCount}件保存`);
  } catch (e) {
    Logger.log(`❌ GSCエラー: ${e.message}`);
  }
  
  Logger.log('=== 日次更新完了 ===');
}

/**
 * ページマッピング取得（path → page_id）
 */
function getPageMappingForDaily(serviceRoleKey) {
  const url = `${DAILY_CONFIG.SUPABASE_URL}/rest/v1/pages?site_id=eq.${DAILY_CONFIG.SITE_ID}&select=id,path`;
  
  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: {
      'apikey': serviceRoleKey,
      'Authorization': `Bearer ${serviceRoleKey}`,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  });
  
  if (response.getResponseCode() !== 200) {
    throw new Error(`ページ取得エラー: ${response.getContentText()}`);
  }
  
  const pages = JSON.parse(response.getContentText());
  const mapping = {};
  
  pages.forEach(page => {
    // パスを正規化（先頭スラッシュなし）
    let path = page.path;
    if (path.startsWith('/')) {
      path = path.substring(1);
    }
    mapping[path] = page.id;
  });
  
  return mapping;
}

/**
 * GA4日次データ取得・保存
 */
function fetchAndSaveGA4Daily(serviceRoleKey, pageMapping, dateStr) {
  // GA4 Data API呼び出し
  const request = {
    dimensions: [{ name: 'pagePath' }],
    metrics: [
      { name: 'screenPageViews' },
      { name: 'sessions' },
      { name: 'averageSessionDuration' },
      { name: 'bounceRate' }
    ],
    dateRanges: [{ startDate: dateStr, endDate: dateStr }]
  };
  
  const report = AnalyticsData.Properties.runReport(request, DAILY_CONFIG.GA4_PROPERTY_ID);
  
  if (!report.rows || report.rows.length === 0) {
    return 0;
  }
  
  // Supabase形式に変換（テーブルのカラム名に合わせる）
  const records = [];
  
  report.rows.forEach(row => {
    let pagePath = row.dimensionValues[0].value;
    
    // パス正規化
    if (pagePath.startsWith('/')) {
      pagePath = pagePath.substring(1);
    }
    
    const pageId = pageMapping[pagePath];
    if (!pageId) return;
    
    records.push({
      page_id: pageId,
      date: dateStr,
      pageviews: parseInt(row.metricValues[0].value) || 0,
      unique_pageviews: parseInt(row.metricValues[1].value) || 0,
      avg_time_on_page: parseFloat(row.metricValues[2].value) || 0,
      bounce_rate: parseFloat(row.metricValues[3].value) || 0
    });
  });
  
  if (records.length === 0) return 0;
  
  // 既存データ削除（同じ日付）
  deleteExistingRecords(serviceRoleKey, 'ga4_metrics_daily', dateStr);
  
  // Supabaseに保存
  return saveToSupabase(serviceRoleKey, 'ga4_metrics_daily', records);
}

/**
 * GSC日次データ取得・保存
 */
function fetchAndSaveGSCDaily(serviceRoleKey, pageMapping, dateStr) {
  // GSC API呼び出し
  const payload = {
    startDate: dateStr,
    endDate: dateStr,
    dimensions: ['page'],
    rowLimit: 25000
  };
  
  const response = UrlFetchApp.fetch(
    `https://www.googleapis.com/webmasters/v3/sites/${encodeURIComponent(DAILY_CONFIG.GSC_SITE_URL)}/searchAnalytics/query`,
    {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    }
  );
  
  if (response.getResponseCode() !== 200) {
    throw new Error(`GSC APIエラー: ${response.getContentText()}`);
  }
  
  const data = JSON.parse(response.getContentText());
  
  if (!data.rows || data.rows.length === 0) {
    return 0;
  }
  
  // Supabase形式に変換
  const records = [];
  const siteUrlBase = DAILY_CONFIG.GSC_SITE_URL.replace(/\/$/, '');
  
  data.rows.forEach(row => {
    const fullUrl = row.keys[0];
    let path = fullUrl.replace(siteUrlBase, '');
    
    // パス正規化
    if (path.startsWith('/')) {
      path = path.substring(1);
    }
    
    const pageId = pageMapping[path];
    if (!pageId) return;
    
    records.push({
      page_id: pageId,
      date: dateStr,
      clicks: Math.round(row.clicks) || 0,
      impressions: Math.round(row.impressions) || 0,
      ctr: row.ctr || 0,
      position: row.position || 0
    });
  });
  
  if (records.length === 0) return 0;
  
  // 既存データ削除（同じ日付）
  deleteExistingRecords(serviceRoleKey, 'gsc_metrics_daily', dateStr);
  
  // Supabaseに保存
  return saveToSupabase(serviceRoleKey, 'gsc_metrics_daily', records);
}

/**
 * 既存レコード削除
 */
function deleteExistingRecords(serviceRoleKey, tableName, dateStr) {
  const deleteUrl = `${DAILY_CONFIG.SUPABASE_URL}/rest/v1/${tableName}?date=eq.${dateStr}`;
  
  UrlFetchApp.fetch(deleteUrl, {
    method: 'delete',
    headers: {
      'apikey': serviceRoleKey,
      'Authorization': `Bearer ${serviceRoleKey}`,
      'Content-Type': 'application/json',
      'Prefer': 'return=minimal'
    },
    muteHttpExceptions: true
  });
}

/**
 * Supabaseに保存
 */
function saveToSupabase(serviceRoleKey, tableName, records) {
  const url = `${DAILY_CONFIG.SUPABASE_URL}/rest/v1/${tableName}`;
  
  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: {
      'apikey': serviceRoleKey,
      'Authorization': `Bearer ${serviceRoleKey}`,
      'Content-Type': 'application/json',
      'Prefer': 'return=minimal'
    },
    payload: JSON.stringify(records),
    muteHttpExceptions: true
  });
  
  if (response.getResponseCode() === 201 || response.getResponseCode() === 200) {
    return records.length;
  } else {
    throw new Error(`保存エラー: ${response.getContentText()}`);
  }
}

/**
 * 日付フォーマット（YYYY-MM-DD）
 */
function formatDateForAPI(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

/**
 * 日次トリガー設定（1回実行）
 */
function setupDailySupabaseTrigger() {
  // 既存トリガー削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'runDailySupabaseUpdate') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // 毎日午前5時に実行
  ScriptApp.newTrigger('runDailySupabaseUpdate')
    .timeBased()
    .atHour(5)
    .everyDays(1)
    .create();
  
  Logger.log('✅ 日次トリガー設定完了（毎朝5時）');
}

/**
 * トリガー削除
 */
function removeDailySupabaseTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'runDailySupabaseUpdate') {
      ScriptApp.deleteTrigger(trigger);
      Logger.log('トリガー削除: runDailySupabaseUpdate');
    }
  });
}

/**
 * 手動テスト用（特定日付を指定）
 */
function testDailyUpdateForDate() {
  const testDate = '2025-12-11';  // テストしたい日付
  
  Logger.log(`=== テスト実行: ${testDate} ===`);
  
  const serviceRoleKey = PropertiesService.getScriptProperties()
    .getProperty('SUPABASE_SERVICE_ROLE_KEY');
  
  const pageMapping = getPageMappingForDaily(serviceRoleKey);
  
  const ga4Count = fetchAndSaveGA4Daily(serviceRoleKey, pageMapping, testDate);
  Logger.log(`GA4: ${ga4Count}件`);
  
  const gscCount = fetchAndSaveGSCDaily(serviceRoleKey, pageMapping, testDate);
  Logger.log(`GSC: ${gscCount}件`);
}