/**
 * GA4ToSupabase.gs
 * GA4からSupabaseへの過去データ移行スクリプト
 * 
 * 【使い方】
 * 1. setSupabaseServiceRoleKey() を実行してService Role Keyを設定
 * 2. migrateGA4DataToSupabase() を実行してデータ移行
 * 
 * 【Service Role Keyの取得方法】
 * Supabase Dashboard > Project Settings > API > service_role (secret) をコピー
 */

// ========================================
// 設定
// ========================================

const CONFIG = {
  // Supabase設定
  SUPABASE_URL: 'https://dgzfdugpineqnoihopsl.supabase.co',
  SITE_ID: '853ea711-7644-451e-872b-dea1b54fa8c7',
  
  // GA4設定
  GA4_PROPERTY_ID: 'properties/388689745',
  
  // 移行期間（14ヶ月分）
  START_YEAR: 2024,
  START_MONTH: 10,  // 2024年10月
  END_YEAR: 2025,
  END_MONTH: 11,    // 2025年11月
  
  // バッチサイズ
  BATCH_SIZE: 500   // 1回のインサートで送信するレコード数
};

// ========================================
// 初期設定関数
// ========================================

/**
 * Supabase Service Role Keyを設定
 * ★★★ 初回のみ実行 ★★★
 * 
 * Supabase Dashboard > Project Settings > API > service_role (secret) からコピー
 */
function setSupabaseServiceRoleKey() {
  const serviceRoleKey = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImRnemZkdWdwaW5lcW5vaWhvcHNsIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NjU0NDUxNDUsImV4cCI6MjA4MTAyMTE0NX0.qw9EzW53FU1LacgI_VLUQ-5sFgkSiJdqGglSKGxRexU';  // ★編集してから実行
  
  if (serviceRoleKey === 'ここにService Role Keyを貼り付け') {
    Logger.log('❌ エラー: Service Role Keyを設定してください');
    Logger.log('1. Supabase Dashboard > Project Settings > API を開く');
    Logger.log('2. service_role (secret) の値をコピー');
    Logger.log('3. この関数内の serviceRoleKey 変数に貼り付け');
    Logger.log('4. 再度実行');
    return;
  }
  
  PropertiesService.getScriptProperties()
    .setProperty('SUPABASE_SERVICE_ROLE_KEY', serviceRoleKey);
  
  Logger.log('✅ SUPABASE_SERVICE_ROLE_KEY を設定しました');
  
  // 接続テスト
  testSupabaseConnection();
}

/**
 * Supabase接続テスト
 */
function testSupabaseConnection() {
  Logger.log('=== Supabase接続テスト ===');
  
  const serviceRoleKey = PropertiesService.getScriptProperties()
    .getProperty('SUPABASE_SERVICE_ROLE_KEY');
  
  if (!serviceRoleKey) {
    Logger.log('❌ Service Role Keyが設定されていません');
    Logger.log('setSupabaseServiceRoleKey() を実行してください');
    return false;
  }
  
  try {
    // pagesテーブルから1件取得してテスト
    const url = `${CONFIG.SUPABASE_URL}/rest/v1/pages?limit=1`;
    
    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: {
        'apikey': serviceRoleKey,
        'Authorization': `Bearer ${serviceRoleKey}`,
        'Content-Type': 'application/json'
      },
      muteHttpExceptions: true
    });
    
    const responseCode = response.getResponseCode();
    
    if (responseCode === 200) {
      const data = JSON.parse(response.getContentText());
      Logger.log('✅ Supabase接続成功！');
      Logger.log(`pagesテーブル確認: ${data.length}件取得`);
      return true;
    } else {
      Logger.log(`❌ 接続エラー: ${responseCode}`);
      Logger.log(response.getContentText());
      return false;
    }
  } catch (error) {
    Logger.log(`❌ 接続エラー: ${error.message}`);
    return false;
  }
}

// ========================================
// メイン移行関数
// ========================================

/**
 * GA4データをSupabaseに移行（メイン関数）
 * ★★★ これを実行 ★★★
 */
function migrateGA4DataToSupabase() {
  const startTime = new Date();
  Logger.log('========================================');
  Logger.log('=== GA4 → Supabase データ移行開始 ===');
  Logger.log(`開始時刻: ${startTime}`);
  Logger.log('========================================');
  
  try {
    // 1. 設定確認
    const serviceRoleKey = PropertiesService.getScriptProperties()
      .getProperty('SUPABASE_SERVICE_ROLE_KEY');
    
    if (!serviceRoleKey) {
      throw new Error('Service Role Keyが設定されていません。setSupabaseServiceRoleKey()を実行してください');
    }
    
    // 2. ページマッピング取得（path → page_id）
    Logger.log('\n【Step 1】ページマッピング取得中...');
    const pageMapping = getPageMapping(serviceRoleKey);
    Logger.log(`✅ ${Object.keys(pageMapping).length}ページのマッピング取得完了`);
    
    // 3. 既存データ削除（クリーンスタート）
    Logger.log('\n【Step 2】既存のGA4データを削除中...');
    deleteExistingGA4Data(serviceRoleKey);
    Logger.log('✅ 既存データ削除完了');
    
    // 4. 月ごとにデータ取得・保存
    Logger.log('\n【Step 3】月別データ取得・保存開始...');
    
    let totalRecords = 0;
    let processedMonths = 0;
    
    // 対象月のリストを生成
    const months = generateMonthList(
      CONFIG.START_YEAR, CONFIG.START_MONTH,
      CONFIG.END_YEAR, CONFIG.END_MONTH
    );
    
    Logger.log(`対象期間: ${months.length}ヶ月`);
    
    for (const month of months) {
      Logger.log(`\n--- ${month.year}年${month.month}月 ---`);
      
      try {
        const records = processMonth(month.year, month.month, pageMapping, serviceRoleKey);
        totalRecords += records;
        processedMonths++;
        
        Logger.log(`✅ ${records}件保存完了（累計: ${totalRecords}件）`);
        
        // API制限対策：月ごとに2秒待機
        Utilities.sleep(2000);
        
      } catch (monthError) {
        Logger.log(`⚠️ ${month.year}年${month.month}月でエラー: ${monthError.message}`);
        Logger.log('次の月に進みます...');
        continue;
      }
    }
    
    // 5. 完了
    const endTime = new Date();
    const duration = (endTime - startTime) / 1000;
    
    Logger.log('\n========================================');
    Logger.log('=== 移行完了 ===');
    Logger.log(`処理時間: ${duration}秒（${(duration/60).toFixed(1)}分）`);
    Logger.log(`処理月数: ${processedMonths}/${months.length}ヶ月`);
    Logger.log(`総レコード数: ${totalRecords}件`);
    Logger.log('========================================');
    
    return {
      success: true,
      totalRecords: totalRecords,
      processedMonths: processedMonths,
      duration: duration
    };
    
  } catch (error) {
    Logger.log(`\n❌❌❌ 移行エラー ❌❌❌`);
    Logger.log(`エラー: ${error.message}`);
    Logger.log(`スタックトレース: ${error.stack}`);
    throw error;
  }
}

// ========================================
// 個別処理関数
// ========================================

/**
 * ページマッピング取得（path → page_id）
 */
function getPageMapping(serviceRoleKey) {
  const url = `${CONFIG.SUPABASE_URL}/rest/v1/pages?site_id=eq.${CONFIG.SITE_ID}&status=eq.active&select=id,path`;
  
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
    throw new Error(`ページマッピング取得エラー: ${response.getContentText()}`);
  }
  
  const pages = JSON.parse(response.getContentText());
  const mapping = {};
  
  pages.forEach(page => {
    mapping[page.path] = page.id;
  });
  
  return mapping;
}

/**
 * 既存のGA4データを削除
 */
function deleteExistingGA4Data(serviceRoleKey) {
  // 全レコード削除（日付で絞り込み - 2024年以降）
  const deleteUrl = `${CONFIG.SUPABASE_URL}/rest/v1/ga4_metrics_daily?date=gte.2024-01-01`;
  
  const deleteResponse = UrlFetchApp.fetch(deleteUrl, {
    method: 'delete',
    headers: {
      'apikey': serviceRoleKey,
      'Authorization': `Bearer ${serviceRoleKey}`,
      'Content-Type': 'application/json',
      'Prefer': 'return=minimal'
    },
    muteHttpExceptions: true
  });
  
  const responseCode = deleteResponse.getResponseCode();
  if (responseCode !== 200 && responseCode !== 204) {
    Logger.log(`警告: 削除でエラー（${responseCode}）: ${deleteResponse.getContentText()}`);
  }
}

/**
 * 1ヶ月分のデータを処理
 */
function processMonth(year, month, pageMapping, serviceRoleKey) {
  // 日付範囲を計算
  const startDate = `${year}-${String(month).padStart(2, '0')}-01`;
  const lastDay = new Date(year, month, 0).getDate();
  const endDate = `${year}-${String(month).padStart(2, '0')}-${String(lastDay).padStart(2, '0')}`;
  
  Logger.log(`期間: ${startDate} 〜 ${endDate}`);
  
  // GA4からデータ取得
  const ga4Data = fetchGA4DailyData(startDate, endDate);
  
  if (ga4Data.length === 0) {
    Logger.log('データなし');
    return 0;
  }
  
  Logger.log(`GA4から${ga4Data.length}件取得`);
  
  // Supabase形式に変換
  const records = [];
  
  ga4Data.forEach(row => {
    const pagePath = row.pagePath;
    const pageId = pageMapping[pagePath];
    
    if (!pageId) {
      // マッピングにないページはスキップ
      return;
    }
    
    records.push({
      page_id: pageId,
      date: row.date,
      pageviews: row.pageviews,
      unique_pageviews: row.uniquePageviews,
      avg_time_on_page: row.avgTimeOnPage,
      bounce_rate: row.bounceRate,
      exits: row.exits || 0,
      entrances: row.entrances || 0
    });
  });
  
  Logger.log(`マッピング後: ${records.length}件`);
  
  if (records.length === 0) {
    return 0;
  }
  
  // バッチでSupabaseに保存
  const savedCount = batchInsertToSupabase(records, serviceRoleKey);
  
  return savedCount;
}

/**
 * GA4から日別データを取得
 */
function fetchGA4DailyData(startDate, endDate) {
  const request = {
    dateRanges: [{
      startDate: startDate,
      endDate: endDate
    }],
    dimensions: [
      { name: 'date' },
      { name: 'pagePath' }
    ],
    metrics: [
      { name: 'screenPageViews' },
      { name: 'activeUsers' },
      { name: 'averageSessionDuration' },
      { name: 'bounceRate' },
      { name: 'sessions' }
    ],
    limit: 50000,
    orderBys: [
      { dimension: { dimensionName: 'date' }, desc: false }
    ]
  };
  
  const response = AnalyticsData.Properties.runReport(request, CONFIG.GA4_PROPERTY_ID);
  
  if (!response.rows || response.rows.length === 0) {
    return [];
  }
  
  return response.rows.map(row => {
    // 日付フォーマット変換: 20241001 → 2024-10-01
    const rawDate = row.dimensionValues[0].value;
    const formattedDate = `${rawDate.substring(0,4)}-${rawDate.substring(4,6)}-${rawDate.substring(6,8)}`;
    
    return {
      date: formattedDate,
      pagePath: row.dimensionValues[1].value,
      pageviews: parseInt(row.metricValues[0].value) || 0,
      uniquePageviews: parseInt(row.metricValues[1].value) || 0,
      avgTimeOnPage: parseFloat(row.metricValues[2].value) || 0,
      bounceRate: parseFloat(row.metricValues[3].value) * 100 || 0,
      exits: 0,        // GA4 Data APIでは取得不可
      entrances: parseInt(row.metricValues[4].value) || 0  // sessionsで代用
    };
  });
}

/**
 * Supabaseにバッチインサート
 */
function batchInsertToSupabase(records, serviceRoleKey) {
  let totalInserted = 0;
  
  // バッチに分割
  for (let i = 0; i < records.length; i += CONFIG.BATCH_SIZE) {
    const batch = records.slice(i, i + CONFIG.BATCH_SIZE);
    
    const url = `${CONFIG.SUPABASE_URL}/rest/v1/ga4_metrics_daily`;
    
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
    
    const responseCode = response.getResponseCode();
    
    if (responseCode === 201 || responseCode === 200) {
      totalInserted += batch.length;
    } else {
      Logger.log(`バッチインサートエラー（${responseCode}）: ${response.getContentText()}`);
      // エラーでも続行
    }
    
    // API制限対策
    if (i + CONFIG.BATCH_SIZE < records.length) {
      Utilities.sleep(500);
    }
  }
  
  return totalInserted;
}

/**
 * 対象月のリストを生成
 */
function generateMonthList(startYear, startMonth, endYear, endMonth) {
  const months = [];
  
  let year = startYear;
  let month = startMonth;
  
  while (year < endYear || (year === endYear && month <= endMonth)) {
    months.push({ year, month });
    
    month++;
    if (month > 12) {
      month = 1;
      year++;
    }
  }
  
  return months;
}

// ========================================
// 個別月の再実行用関数
// ========================================

/**
 * 特定の月だけ再取得（エラー時の再実行用）
 * @param {number} year - 年（例: 2024）
 * @param {number} month - 月（例: 10）
 */
function reprocessSingleMonth(year, month) {
  Logger.log(`=== ${year}年${month}月 再処理 ===`);
  
  const serviceRoleKey = PropertiesService.getScriptProperties()
    .getProperty('SUPABASE_SERVICE_ROLE_KEY');
  
  if (!serviceRoleKey) {
    Logger.log('❌ Service Role Keyが設定されていません');
    return;
  }
  
  // ページマッピング取得
  const pageMapping = getPageMapping(serviceRoleKey);
  
  // 該当月のデータを削除
  const startDate = `${year}-${String(month).padStart(2, '0')}-01`;
  const lastDay = new Date(year, month, 0).getDate();
  const endDate = `${year}-${String(month).padStart(2, '0')}-${String(lastDay).padStart(2, '0')}`;
  
  Logger.log(`期間: ${startDate} 〜 ${endDate}`);
  Logger.log('既存データ削除中...');
  
  // 日付範囲で削除
  const pageIds = Object.values(pageMapping);
  const deleteUrl = `${CONFIG.SUPABASE_URL}/rest/v1/ga4_metrics_daily?page_id=in.(${pageIds.join(',')})&date=gte.${startDate}&date=lte.${endDate}`;
  
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
  
  // 再取得・保存
  const records = processMonth(year, month, pageMapping, serviceRoleKey);
  
  Logger.log(`✅ 完了: ${records}件保存`);
  
  return records;
}

/**
 * 2024年10月を再処理（例）
 */
function reprocess202410() {
  return reprocessSingleMonth(2024, 10);
}

/**
 * 2024年11月を再処理（例）
 */
function reprocess202411() {
  return reprocessSingleMonth(2024, 11);
}

// ========================================
// 確認用関数
// ========================================

/**
 * 移行結果を確認
 */
function checkMigrationResult() {
  Logger.log('=== 移行結果確認 ===');
  
  const serviceRoleKey = PropertiesService.getScriptProperties()
    .getProperty('SUPABASE_SERVICE_ROLE_KEY');
  
  if (!serviceRoleKey) {
    Logger.log('❌ Service Role Keyが設定されていません');
    return;
  }
  
  // 月別レコード数を集計
  const url = `${CONFIG.SUPABASE_URL}/rest/v1/rpc/count_ga4_by_month`;
  
  // RPCがない場合は直接クエリ
  const directUrl = `${CONFIG.SUPABASE_URL}/rest/v1/ga4_metrics_daily?select=date`;
  
  const response = UrlFetchApp.fetch(directUrl, {
    method: 'get',
    headers: {
      'apikey': serviceRoleKey,
      'Authorization': `Bearer ${serviceRoleKey}`,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  });
  
  if (response.getResponseCode() !== 200) {
    Logger.log(`エラー: ${response.getContentText()}`);
    return;
  }
  
  const data = JSON.parse(response.getContentText());
  
  // 月別に集計
  const monthCounts = {};
  data.forEach(row => {
    const month = row.date.substring(0, 7);  // YYYY-MM
    monthCounts[month] = (monthCounts[month] || 0) + 1;
  });
  
  Logger.log(`総レコード数: ${data.length}件`);
  Logger.log('\n月別レコード数:');
  
  Object.keys(monthCounts).sort().forEach(month => {
    Logger.log(`  ${month}: ${monthCounts[month]}件`);
  });
}

/**
 * Service Role Key設定状況を確認
 */
function checkServiceRoleKey() {
  const key = PropertiesService.getScriptProperties()
    .getProperty('SUPABASE_SERVICE_ROLE_KEY');
  
  if (key) {
    Logger.log('✅ Service Role Key: 設定済み');
    Logger.log(`先頭10文字: ${key.substring(0, 10)}...`);
  } else {
    Logger.log('❌ Service Role Key: 未設定');
    Logger.log('setSupabaseServiceRoleKey() を実行してください');
  }
}