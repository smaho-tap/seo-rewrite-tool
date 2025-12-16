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
  
  // ページマッピング取得（投稿ページのみ）
  const pageMapping = getPageMappingForDaily(serviceRoleKey);
  Logger.log(`ページマッピング: ${Object.keys(pageMapping).length}件（投稿ページのみ）`);
  
 // 前日の日付（GA4用）
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  const dateStr = formatDateForAPI(yesterday);
  
  // 3日前の日付（GSC用 - GSCは2-3日遅れるため）
  const threeDaysAgo = new Date();
  threeDaysAgo.setDate(threeDaysAgo.getDate() - 3);
  const dateStrGSC = formatDateForAPI(threeDaysAgo);
  
  Logger.log(`対象日: GA4=${dateStr}, GSC=${dateStrGSC}`);
  
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
    const gscCount = fetchAndSaveGSCDaily(serviceRoleKey, pageMapping, dateStrGSC);
    Logger.log(`✅ GSC: ${gscCount}件保存`);
  } catch (e) {
    Logger.log(`❌ GSCエラー: ${e.message}`);
  }

  // GSCクエリデータ取得・保存
  try {
    const queryCount = fetchAndSaveGSCQueriesDaily(serviceRoleKey, pageMapping, dateStrGSC);
    Logger.log(`✅ GSCクエリ: ${queryCount}件保存`);
  } catch (e) {
    Logger.log(`❌ GSCクエリエラー: ${e.message}`);
  }

  // WordPress投稿日同期
  try {
    const wpCount = syncWordPressPublishDates(serviceRoleKey);
    Logger.log(`✅ WordPress投稿日同期: ${wpCount}件更新`);
  } catch (e) {
    Logger.log(`❌ WordPress同期エラー: ${e.message}`);
  }
  Logger.log('=== 日次更新完了 ===');
}

/**
 * ページマッピング取得（path → page_id）
 * ★ status=active（投稿ページ）のみ取得
 */
function getPageMappingForDaily(serviceRoleKey) {
  // status=eq.active で投稿ページのみフィルタリング
  const url = `${DAILY_CONFIG.SUPABASE_URL}/rest/v1/pages?site_id=eq.${DAILY_CONFIG.SITE_ID}&status=eq.active&select=id,path`;
  
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
    if (!pageId) return;  // 投稿ページ以外はスキップ
    
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
    if (!pageId) return;  // 投稿ページ以外はスキップ
    
    records.push({
      page_id: pageId,
      date: dateStr,
      clicks: Math.round(row.clicks) || 0,
      impressions: Math.round(row.impressions) || 0,
      ctr: row.ctr || 0,
      avg_position: row.position || 0
    });
  });
  
  if (records.length === 0) return 0;
  
  // 既存データ削除（同じ日付）
  deleteExistingRecords(serviceRoleKey, 'gsc_metrics_daily', dateStr);
  
  // Supabaseに保存
  return saveToSupabase(serviceRoleKey, 'gsc_metrics_daily', records);
}

/**
 * GSCクエリ単位データ取得・保存
 * 主要KWと実クエリの一致度分析用
 */
function fetchAndSaveGSCQueriesDaily(serviceRoleKey, pageMapping, dateStr) {
  Logger.log('--- GSCクエリデータ取得開始 ---');
  
  // GSC API呼び出し（ページ×クエリ）
  const payload = {
    startDate: dateStr,
    endDate: dateStr,
    dimensions: ['page', 'query'],
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
    Logger.log('クエリデータなし');
    return 0;
  }
  
  Logger.log(`GSC APIから${data.rows.length}件取得`);
  
  // Supabase形式に変換
  const records = [];
  const siteUrlBase = DAILY_CONFIG.GSC_SITE_URL.replace(/\/$/, '');
  
  data.rows.forEach(row => {
    const fullUrl = row.keys[0];
    const query = row.keys[1];
    
    let path = fullUrl.replace(siteUrlBase, '');
    if (path.startsWith('/')) {
      path = path.substring(1);
    }
    
    const pageId = pageMapping[path];
    if (!pageId) return;
    
    if (row.impressions < 5) return;
    
    records.push({
      page_id: pageId,
      query: query,
      date: dateStr,
      impressions: Math.round(row.impressions) || 0,
      clicks: Math.round(row.clicks) || 0,
      ctr: row.ctr || 0,
      position: row.position || 0
    });
  });
  
  Logger.log(`フィルタ後: ${records.length}件`);
  
  if (records.length === 0) return 0;
  
  deleteExistingQueryRecords(serviceRoleKey, dateStr);
  return saveQueriesToSupabase(serviceRoleKey, records);
}

/**
 * 既存のクエリレコード削除
 */
function deleteExistingQueryRecords(serviceRoleKey, dateStr) {
  const deleteUrl = `${DAILY_CONFIG.SUPABASE_URL}/rest/v1/gsc_queries?date=eq.${dateStr}`;
  
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
 * クエリデータをSupabaseに保存（バッチ処理）
 */
function saveQueriesToSupabase(serviceRoleKey, records) {
  const BATCH_SIZE = 500;
  let totalSaved = 0;
  
  for (let i = 0; i < records.length; i += BATCH_SIZE) {
    const batch = records.slice(i, i + BATCH_SIZE);
    
    const url = `${DAILY_CONFIG.SUPABASE_URL}/rest/v1/gsc_queries`;
    
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      headers: {
        'apikey': serviceRoleKey,
        'Authorization': `Bearer ${serviceRoleKey}`,
        'Content-Type': 'application/json',
        'Prefer': 'return=minimal'
      },
      payload: JSON.stringify(batch),
      muteHttpExceptions: true
    });
    
    const code = response.getResponseCode();
    
    if (code === 201 || code === 200) {
      totalSaved += batch.length;
    } else {
      Logger.log(`バッチ保存エラー（${code}）: ${response.getContentText().substring(0, 200)}`);
    }
    
    if (i + BATCH_SIZE < records.length) {
      Utilities.sleep(300);
    }
  }
  
  return totalSaved;
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

/**
 * WordPress REST APIから投稿日を取得してSupabaseに同期
 * 新規記事（first_published_atがnull）のみ更新
 */
function syncWordPressPublishDates(serviceRoleKey) {
  Logger.log('--- WordPress投稿日同期開始 ---');
  
  // 1. Supabaseからfirst_published_atがnullのページを取得
  const pagesUrl = `${DAILY_CONFIG.SUPABASE_URL}/rest/v1/pages?site_id=eq.${DAILY_CONFIG.SITE_ID}&status=eq.active&first_published_at=is.null&select=id,path`;
  
  const pagesResponse = UrlFetchApp.fetch(pagesUrl, {
    method: 'get',
    headers: {
      'apikey': serviceRoleKey,
      'Authorization': `Bearer ${serviceRoleKey}`,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  });
  
  if (pagesResponse.getResponseCode() !== 200) {
    throw new Error(`ページ取得エラー: ${pagesResponse.getContentText()}`);
  }
  
  const pagesWithoutDate = JSON.parse(pagesResponse.getContentText());
  
  if (pagesWithoutDate.length === 0) {
    Logger.log('投稿日未設定のページはありません');
    return 0;
  }
  
  Logger.log(`投稿日未設定ページ: ${pagesWithoutDate.length}件`);
  
  // 2. WordPress REST APIから全記事を取得
  const wpPosts = fetchAllWordPressPosts();
  Logger.log(`WordPress記事数: ${wpPosts.length}件`);
  
  // 3. slugでマッチングして更新
  let updatedCount = 0;
  
  pagesWithoutDate.forEach(page => {
    // pathからslugを抽出（先頭の/を除去）
    let slug = page.path;
    if (slug.startsWith('/')) {
      slug = slug.substring(1);
    }
    
    // WordPressの記事を検索
    const wpPost = wpPosts.find(post => post.slug === slug);
    
    if (wpPost && wpPost.published_date) {
      // Supabaseを更新
      const updateUrl = `${DAILY_CONFIG.SUPABASE_URL}/rest/v1/pages?id=eq.${page.id}`;
      
      const updateResponse = UrlFetchApp.fetch(updateUrl, {
        method: 'patch',
        headers: {
          'apikey': serviceRoleKey,
          'Authorization': `Bearer ${serviceRoleKey}`,
          'Content-Type': 'application/json',
          'Prefer': 'return=minimal'
        },
        payload: JSON.stringify({
          first_published_at: wpPost.published_date,
          updated_at: new Date().toISOString()
        }),
        muteHttpExceptions: true
      });
      
      if (updateResponse.getResponseCode() === 204 || updateResponse.getResponseCode() === 200) {
        Logger.log(`  更新: ${slug} → ${wpPost.published_date}`);
        updatedCount++;
      }
    }
  });
  
  Logger.log(`--- WordPress同期完了: ${updatedCount}件更新 ---`);
  return updatedCount;
}

/**
 * WordPress REST APIから全記事を取得
 */
function fetchAllWordPressPosts() {
  const allPosts = [];
  let page = 1;
  let hasMore = true;
  
  while (hasMore) {
    const url = `${DAILY_CONFIG.GSC_SITE_URL}/wp-json/wp/v2/posts?per_page=100&page=${page}&_fields=id,date,slug`;
    
    try {
      const response = UrlFetchApp.fetch(url, {
        method: 'get',
        muteHttpExceptions: true
      });
      
      if (response.getResponseCode() !== 200) {
        hasMore = false;
        break;
      }
      
      const posts = JSON.parse(response.getContentText());
      
      if (posts.length === 0) {
        hasMore = false;
      } else {
        posts.forEach(post => {
          // slugをデコード（日本語URLの場合）
          let decodedSlug = post.slug;
          try {
            decodedSlug = decodeURIComponent(post.slug);
          } catch (e) {
            // デコード失敗時はそのまま使用
          }
          
          allPosts.push({
            id: post.id,
            slug: decodedSlug,
            published_date: post.date
          });
        });
        page++;
      }
    } catch (e) {
      Logger.log(`WordPress APIエラー（page ${page}）: ${e.message}`);
      hasMore = false;
    }
  }
  
  return allPosts;
}

/**
 * WordPress同期の手動テスト
 */
function testWordPressSync() {
  Logger.log('=== WordPress同期テスト ===');
  
  const serviceRoleKey = PropertiesService.getScriptProperties()
    .getProperty('SUPABASE_SERVICE_ROLE_KEY');
  
  if (!serviceRoleKey) {
    Logger.log('❌ Service Role Keyが設定されていません');
    return;
  }
  
  const count = syncWordPressPublishDates(serviceRoleKey);
  Logger.log(`結果: ${count}件更新`);
}

/**
 * 過去30日分のGSCクエリデータを一括取得（初回移行用）
 * ★ 1回だけ実行してください
 */
function migrateGSCQueries30Days() {
  Logger.log('=== GSCクエリ 過去30日分移行開始 ===');
  
  const serviceRoleKey = PropertiesService.getScriptProperties()
    .getProperty('SUPABASE_SERVICE_ROLE_KEY');
  
  if (!serviceRoleKey) {
    Logger.log('❌ Service Role Keyが設定されていません');
    return;
  }
  
  const pageMapping = getPageMappingForDaily(serviceRoleKey);
  Logger.log(`ページマッピング: ${Object.keys(pageMapping).length}件`);
  
  let totalCount = 0;
  
  // 過去30日分を取得（3日前から33日前まで）
  for (let i = 3; i <= 33; i++) {
    const targetDate = new Date();
    targetDate.setDate(targetDate.getDate() - i);
    const dateStr = formatDateForAPI(targetDate);
    
    Logger.log(`\n--- ${dateStr} ---`);
    
    try {
      const count = fetchAndSaveGSCQueriesDaily(serviceRoleKey, pageMapping, dateStr);
      totalCount += count;
      Logger.log(`✅ ${count}件保存（累計: ${totalCount}件）`);
      
      // API制限対策
      Utilities.sleep(1000);
      
    } catch (e) {
      Logger.log(`❌ エラー: ${e.message}`);
    }
  }
  
  Logger.log(`\n=== 移行完了: 合計${totalCount}件 ===`);
}