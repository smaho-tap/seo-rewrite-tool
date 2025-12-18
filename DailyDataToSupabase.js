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
  // 修正後
const request = {
  dimensions: [{ name: 'pagePath' }],
  metrics: [
    { name: 'screenPageViews' },
    { name: 'sessions' },
    { name: 'userEngagementDuration' },
    { name: 'activeUsers' },
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
    
   // 修正後
const engagementDuration = parseFloat(row.metricValues[2].value) || 0;
const activeUsers = parseInt(row.metricValues[3].value) || 1;
const avgTimeOnPage = activeUsers > 0 ? engagementDuration / activeUsers : 0;

records.push({
  page_id: pageId,
  date: dateStr,
  pageviews: parseInt(row.metricValues[0].value) || 0,
  unique_pageviews: parseInt(row.metricValues[1].value) || 0,
  avg_time_on_page: avgTimeOnPage,
  bounce_rate: parseFloat(row.metricValues[4].value) || 0
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

/**
 * GA4欠損データ一括復旧（12/1〜12/10）
 * ★ 1回だけ実行してください
 */
function recoverGA4MissingData() {
  Logger.log('=== GA4 欠損データ復旧開始 ===');
  Logger.log(`実行日時: ${new Date().toLocaleString('ja-JP')}`);
  
  const serviceRoleKey = PropertiesService.getScriptProperties()
    .getProperty('SUPABASE_SERVICE_ROLE_KEY');
  
  if (!serviceRoleKey) {
    Logger.log('❌ Service Role Keyが設定されていません');
    return;
  }
  
  const pageMapping = getPageMappingForDaily(serviceRoleKey);
  Logger.log(`ページマッピング: ${Object.keys(pageMapping).length}件`);
  
  // 復旧対象の日付リスト（12/1〜12/10）
  const targetDates = [
    '2025-12-01',
    '2025-12-02',
    '2025-12-03',
    '2025-12-04',
    '2025-12-05',
    '2025-12-06',
    '2025-12-07',
    '2025-12-08',
    '2025-12-09',
    '2025-12-10'
  ];
  
  let totalCount = 0;
  let successDays = 0;
  let failedDays = [];
  
  targetDates.forEach((dateStr, index) => {
    Logger.log(`\n--- [${index + 1}/10] ${dateStr} ---`);
    
    try {
      const count = fetchAndSaveGA4Daily(serviceRoleKey, pageMapping, dateStr);
      totalCount += count;
      successDays++;
      Logger.log(`✅ ${count}件保存（累計: ${totalCount}件）`);
    } catch (e) {
      Logger.log(`❌ エラー: ${e.message}`);
      failedDays.push(dateStr);
    }
    
    // API制限対策（最後以外は1秒待機）
    if (index < targetDates.length - 1) {
      Utilities.sleep(1000);
    }
  });
  
  Logger.log('\n=============================');
  Logger.log('=== 復旧完了 ===');
  Logger.log(`成功: ${successDays}/10日`);
  Logger.log(`合計: ${totalCount}件`);
  
  if (failedDays.length > 0) {
    Logger.log(`失敗した日付: ${failedDays.join(', ')}`);
  }
  Logger.log('=============================');
}
/**
 * GA4データ全期間再取得（メトリクス修正後に1回実行）
 * 11/18〜12/17の30日分を再取得
 */
function refreshAllGA4Data() {
  Logger.log('=== GA4 全データ再取得開始 ===');
  Logger.log(`実行日時: ${new Date().toLocaleString('ja-JP')}`);
  
  const serviceRoleKey = PropertiesService.getScriptProperties()
    .getProperty('SUPABASE_SERVICE_ROLE_KEY');
  
  if (!serviceRoleKey) {
    Logger.log('❌ Service Role Keyが設定されていません');
    return;
  }
  
  const pageMapping = getPageMappingForDaily(serviceRoleKey);
  Logger.log(`ページマッピング: ${Object.keys(pageMapping).length}件`);
  
  let totalCount = 0;
  let successDays = 0;
  
  // 過去30日分を再取得
  for (let i = 1; i <= 30; i++) {
    const targetDate = new Date();
    targetDate.setDate(targetDate.getDate() - i);
    const dateStr = formatDateForAPI(targetDate);
    
    Logger.log(`[${i}/30] ${dateStr}`);
    
    try {
      const count = fetchAndSaveGA4Daily(serviceRoleKey, pageMapping, dateStr);
      totalCount += count;
      successDays++;
    } catch (e) {
      Logger.log(`  ❌ エラー: ${e.message}`);
    }
    
    // API制限対策
    if (i < 30) {
      Utilities.sleep(500);
    }
  }
  
  Logger.log('\n=============================');
  Logger.log(`完了: ${successDays}/30日, 合計${totalCount}件`);
  Logger.log('=============================');
}

/**
 * ========================================
 * リライト効果通知機能
 * ========================================
 */

// 通知先メールアドレス（ご自身のアドレスに変更してください）
const NOTIFICATION_EMAIL = 'foster_inc@icloud.com';

/**
 * リマインダーチェック＆メール送信（毎朝実行）
 */
function checkAndSendRewriteReminders() {
  Logger.log('=== リライト効果通知チェック開始 ===');
  
  const serviceRoleKey = PropertiesService.getScriptProperties()
    .getProperty('SUPABASE_SERVICE_ROLE_KEY');
  
  if (!serviceRoleKey) {
    Logger.log('❌ Service Role Keyが設定されていません');
    return;
  }
  
  // 今日通知すべきリマインダーを取得
  const reminders = getPendingReminders(serviceRoleKey);
  
  if (reminders.length === 0) {
    Logger.log('通知すべきリマインダーはありません');
    return;
  }
  
  Logger.log(`${reminders.length}件のリマインダーを処理します`);
  
  reminders.forEach(reminder => {
    try {
      // Before/After効果データを取得
      const effectData = getRewriteEffect(serviceRoleKey, reminder.page_id, reminder.implemented_at);
      
      // メール送信
      sendEffectEmail(reminder, effectData);
      
      // 送信済みにマーク
      markReminderSent(serviceRoleKey, reminder.reminder_id);
      
      Logger.log(`✅ 送信完了: ${reminder.page_path}`);
    } catch (e) {
      Logger.log(`❌ エラー: ${reminder.page_path} - ${e.message}`);
    }
  });
  
  Logger.log('=== リライト効果通知チェック完了 ===');
}

/**
 * 保留中のリマインダーを取得
 */
function getPendingReminders(serviceRoleKey) {
  const url = `${DAILY_CONFIG.SUPABASE_URL}/rest/v1/rpc/get_pending_reminders`;
  
  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: {
      'apikey': serviceRoleKey,
      'Authorization': `Bearer ${serviceRoleKey}`,
      'Content-Type': 'application/json'
    },
    payload: '{}',
    muteHttpExceptions: true
  });
  
  if (response.getResponseCode() !== 200) {
    Logger.log(`リマインダー取得エラー: ${response.getContentText()}`);
    return [];
  }
  
  return JSON.parse(response.getContentText());
}

/**
 * リライト効果データを取得（Before/After比較）
 */
function getRewriteEffect(serviceRoleKey, pageId, implementedAt) {
  const implementedDate = new Date(implementedAt);
  const beforeStart = new Date(implementedDate);
  beforeStart.setDate(beforeStart.getDate() - 7);
  const afterEnd = new Date();
  
  // GSCデータで比較（Before: 実施前7日間, After: 実施後〜現在）
  const beforeData = getGSCMetrics(serviceRoleKey, pageId, formatDateForAPI(beforeStart), formatDateForAPI(implementedDate));
  const afterData = getGSCMetrics(serviceRoleKey, pageId, formatDateForAPI(implementedDate), formatDateForAPI(afterEnd));
  
  return {
    before: beforeData,
    after: afterData,
    change: {
      clicks: afterData.clicks - beforeData.clicks,
      impressions: afterData.impressions - beforeData.impressions,
      avg_position: beforeData.avg_position - afterData.avg_position, // 順位は低いほど良い
      ctr: afterData.ctr - beforeData.ctr
    }
  };
}

/**
 * GSCメトリクス取得（期間集計）
 */
function getGSCMetrics(serviceRoleKey, pageId, startDate, endDate) {
  const url = `${DAILY_CONFIG.SUPABASE_URL}/rest/v1/gsc_metrics_daily?page_id=eq.${pageId}&date=gte.${startDate}&date=lt.${endDate}&select=clicks,impressions,ctr,avg_position`;
  
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
    return { clicks: 0, impressions: 0, ctr: 0, avg_position: 0, days: 0 };
  }
  
  const rows = JSON.parse(response.getContentText());
  
  if (rows.length === 0) {
    return { clicks: 0, impressions: 0, ctr: 0, avg_position: 0, days: 0 };
  }
  
  const totals = rows.reduce((acc, row) => {
    acc.clicks += row.clicks || 0;
    acc.impressions += row.impressions || 0;
    acc.positions.push(row.avg_position || 0);
    return acc;
  }, { clicks: 0, impressions: 0, positions: [] });
  
  const avgPosition = totals.positions.length > 0 
    ? totals.positions.reduce((a, b) => a + b, 0) / totals.positions.length 
    : 0;
  
  return {
    clicks: totals.clicks,
    impressions: totals.impressions,
    ctr: totals.impressions > 0 ? (totals.clicks / totals.impressions * 100) : 0,
    avg_position: avgPosition,
    days: rows.length
  };
}

/**
 * 効果レポートメール送信
 */
function sendEffectEmail(reminder, effectData) {
  const subject = `【リライト効果レポート】${reminder.page_path}`;
  
  const positionChange = effectData.change.avg_position;
  const positionEmoji = positionChange > 0 ? '📈' : (positionChange < 0 ? '📉' : '➡️');
  
  const body = `
リライト効果レポート
====================

■ ページ情報
パス: ${reminder.page_path}
タイトル: ${reminder.page_title}
リライト種別: ${reminder.rewrite_type}
実施日: ${new Date(reminder.implemented_at).toLocaleDateString('ja-JP')}

■ 変更内容
【Before】
${reminder.before_content || '(記録なし)'}

【After】
${reminder.after_content || '(記録なし)'}

■ 効果測定（GSCデータ）

【Before（実施前7日間）】
・クリック数: ${effectData.before.clicks}
・表示回数: ${effectData.before.impressions}
・平均順位: ${effectData.before.avg_position.toFixed(1)}位
・CTR: ${effectData.before.ctr.toFixed(2)}%

【After（実施後〜現在）】
・クリック数: ${effectData.after.clicks}
・表示回数: ${effectData.after.impressions}
・平均順位: ${effectData.after.avg_position.toFixed(1)}位
・CTR: ${effectData.after.ctr.toFixed(2)}%

■ 変化 ${positionEmoji}
・クリック数: ${effectData.change.clicks >= 0 ? '+' : ''}${effectData.change.clicks}
・表示回数: ${effectData.change.impressions >= 0 ? '+' : ''}${effectData.change.impressions}
・順位変動: ${positionChange >= 0 ? '+' : ''}${positionChange.toFixed(1)}位
・CTR変動: ${effectData.change.ctr >= 0 ? '+' : ''}${effectData.change.ctr.toFixed(2)}%

====================
SEOリライト支援ツール
`;
  
  MailApp.sendEmail({
    to: NOTIFICATION_EMAIL,
    subject: subject,
    body: body
  });
}

/**
 * リマインダーを送信済みにマーク
 */
function markReminderSent(serviceRoleKey, reminderId) {
  const url = `${DAILY_CONFIG.SUPABASE_URL}/rest/v1/rpc/mark_reminder_sent`;
  
  UrlFetchApp.fetch(url, {
    method: 'post',
    headers: {
      'apikey': serviceRoleKey,
      'Authorization': `Bearer ${serviceRoleKey}`,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify({ p_reminder_id: reminderId }),
    muteHttpExceptions: true
  });
}

/**
 * リマインダー通知トリガー設定（1回実行）
 */
function setupReminderTrigger() {
  // 既存トリガー削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'checkAndSendRewriteReminders') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // 毎日午前7時に実行（日次更新の後）
  ScriptApp.newTrigger('checkAndSendRewriteReminders')
    .timeBased()
    .atHour(7)
    .everyDays(1)
    .create();
  
  Logger.log('✅ リマインダー通知トリガー設定完了（毎朝7時）');
}

/**
 * リマインダー手動登録（チャットから呼び出し用）
 */
function registerReminderManual(rewriteHistoryId, daysAfter) {
  const serviceRoleKey = PropertiesService.getScriptProperties()
    .getProperty('SUPABASE_SERVICE_ROLE_KEY');
  
  const url = `${DAILY_CONFIG.SUPABASE_URL}/rest/v1/rpc/register_rewrite_reminder`;
  
  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: {
      'apikey': serviceRoleKey,
      'Authorization': `Bearer ${serviceRoleKey}`,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify({
      p_rewrite_history_id: rewriteHistoryId,
      p_days_after: daysAfter || 7
    }),
    muteHttpExceptions: true
  });
  
  Logger.log(response.getContentText());
  return JSON.parse(response.getContentText());
}

/**
 * ========================================
 * 古い情報検出＆通知機能
 * ========================================
 */

/**
 * 古い情報をチェックしてメール通知（月次実行）
 */
function checkOutdatedContentAndNotify() {
  Logger.log('=== 古い情報チェック開始 ===');
  
  const serviceRoleKey = PropertiesService.getScriptProperties()
    .getProperty('SUPABASE_SERVICE_ROLE_KEY');
  
  if (!serviceRoleKey) {
    Logger.log('❌ Service Role Keyが設定されていません');
    return;
  }
  
  // 現在の年を取得
  const currentYear = new Date().getFullYear();
  Logger.log(`現在の年: ${currentYear}`);
  
  // 検出結果を保存
  const savedCount = saveOutdatedAlerts(serviceRoleKey, currentYear);
  Logger.log(`検出・保存件数: ${savedCount}`);
  
  // 通知すべきアラートを取得
  const alerts = getOutdatedAlertsForNotification(serviceRoleKey);
  
  if (alerts.length === 0) {
    Logger.log('通知すべき古い情報はありません');
    return;
  }
  
  Logger.log(`通知対象: ${alerts.length}件`);
  
  // メール送信
  sendOutdatedContentEmail(alerts, currentYear);
  
  Logger.log('=== 古い情報チェック完了 ===');
}

/**
 * 検出結果を保存
 */
function saveOutdatedAlerts(serviceRoleKey, currentYear) {
  const url = `${DAILY_CONFIG.SUPABASE_URL}/rest/v1/rpc/save_outdated_alerts`;
  
  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: {
      'apikey': serviceRoleKey,
      'Authorization': `Bearer ${serviceRoleKey}`,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify({ p_current_year: currentYear }),
    muteHttpExceptions: true
  });
  
  if (response.getResponseCode() !== 200) {
    Logger.log(`保存エラー: ${response.getContentText()}`);
    return 0;
  }
  
  return JSON.parse(response.getContentText());
}

/**
 * 通知すべきアラートを取得
 */
function getOutdatedAlertsForNotification(serviceRoleKey) {
  const url = `${DAILY_CONFIG.SUPABASE_URL}/rest/v1/rpc/get_outdated_alerts_for_notification`;
  
  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: {
      'apikey': serviceRoleKey,
      'Authorization': `Bearer ${serviceRoleKey}`,
      'Content-Type': 'application/json'
    },
    payload: '{}',
    muteHttpExceptions: true
  });
  
  if (response.getResponseCode() !== 200) {
    Logger.log(`取得エラー: ${response.getContentText()}`);
    return [];
  }
  
  return JSON.parse(response.getContentText());
}

/**
 * 古い情報検出メールを送信
 */
function sendOutdatedContentEmail(alerts, currentYear) {
  const subject = `【確認依頼】古い情報が検出されました（${alerts.length}件）`;
  
  // 緊急度別に分類
  const highUrgency = alerts.filter(a => a.urgency_level === 'high');
  const mediumUrgency = alerts.filter(a => a.urgency_level === 'medium');
  const lowUrgency = alerts.filter(a => a.urgency_level === 'low');
  
  let body = `
古い情報検出レポート
====================
検出日: ${new Date().toLocaleDateString('ja-JP')}
現在の年: ${currentYear}年

以下のページに古い年号が検出されました。
更新が必要かどうかご確認ください。

`;

  if (highUrgency.length > 0) {
    body += `\n■ 要確認度：高（${highUrgency.length}件）\n`;
    body += `  「最新」「おすすめ」等を含むため更新推奨\n`;
    body += `-----------------------------------------\n`;
    highUrgency.forEach(alert => {
      body += `・${alert.path}\n`;
      body += `  タイトル: ${alert.title}\n`;
      body += `  検出年号: ${alert.detected_year}年\n\n`;
    });
  }
  
  if (mediumUrgency.length > 0) {
    body += `\n■ 要確認度：中（${mediumUrgency.length}件）\n`;
    body += `  1年前の情報\n`;
    body += `-----------------------------------------\n`;
    mediumUrgency.forEach(alert => {
      body += `・${alert.path}\n`;
      body += `  タイトル: ${alert.title}\n`;
      body += `  検出年号: ${alert.detected_year}年\n\n`;
    });
  }
  
  if (lowUrgency.length > 0) {
    body += `\n■ 要確認度：低（${lowUrgency.length}件）\n`;
    body += `  歴史的事実の可能性あり\n`;
    body += `-----------------------------------------\n`;
    lowUrgency.forEach(alert => {
      body += `・${alert.path}\n`;
      body += `  タイトル: ${alert.title}\n`;
      body += `  検出年号: ${alert.detected_year}年\n\n`;
    });
  }
  
  body += `
====================
【対応方法】
・更新が必要な場合 → リライトを実施
・更新不要の場合 → チャットで「〇〇は対応不要にして」とお伝えください

※対応不要にしたページは、次の年になるまで再通知されません。

SEOリライト支援ツール
`;

  MailApp.sendEmail({
    to: NOTIFICATION_EMAIL,
    subject: subject,
    body: body
  });
  
  Logger.log('✅ メール送信完了');
}

/**
 * 古い情報チェックのトリガー設定（1回実行）
 * 毎月1日の午前8時に実行
 */
function setupOutdatedContentTrigger() {
  // 既存トリガー削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'checkOutdatedContentAndNotify') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // 毎月1日に実行
  ScriptApp.newTrigger('checkOutdatedContentAndNotify')
    .timeBased()
    .onMonthDay(1)
    .atHour(8)
    .create();
  
  Logger.log('✅ 古い情報チェックトリガー設定完了（毎月1日 午前8時）');
}

/**
 * 手動テスト用
 */
function testOutdatedContentCheck() {
  checkOutdatedContentAndNotify();
}