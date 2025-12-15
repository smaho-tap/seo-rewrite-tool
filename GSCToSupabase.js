/**
 * GSCToSupabase.gs
 * Google Search ConsoleからSupabaseへの過去データ移行スクリプト
 * 
 * 【使い方】
 * 1. GA4ToSupabase.gsでService Role Keyを設定済みであれば、そのまま実行可能
 * 2. migrateGSCDataToSupabase() を実行してデータ移行
 * 
 * 【注意】
 * - GSC APIは過去16ヶ月分のデータのみ取得可能
 * - 日別×ページ別のデータを取得
 */

// ========================================
// 設定
// ========================================

const GSC_CONFIG = {
  // Supabase設定（GA4ToSupabaseと共通）
  SUPABASE_URL: 'https://dgzfdugpineqnoihopsl.supabase.co',
  SITE_ID: '853ea711-7644-451e-872b-dea1b54fa8c7',
  
  // GSC設定
  GSC_SITE_URL: 'https://smaho-tap.com',
  
  // 移行期間（16ヶ月分）
  START_YEAR: 2024,
  START_MONTH: 8,   // 2024年8月（16ヶ月前）
  END_YEAR: 2025,
  END_MONTH: 11,    // 2025年11月
  
  // バッチサイズ
  BATCH_SIZE: 500
};

// ========================================
// メイン移行関数
// ========================================

/**
 * GSCデータをSupabaseに移行（メイン関数）
 * ★★★ これを実行 ★★★
 */
function migrateGSCDataToSupabase() {
  const startTime = new Date();
  Logger.log('========================================');
  Logger.log('=== GSC → Supabase データ移行開始 ===');
  Logger.log(`開始時刻: ${startTime}`);
  Logger.log('========================================');
  
  try {
    // 1. 設定確認
    const serviceRoleKey = PropertiesService.getScriptProperties()
      .getProperty('SUPABASE_SERVICE_ROLE_KEY');
    
    if (!serviceRoleKey) {
      throw new Error('Service Role Keyが設定されていません。GA4ToSupabase.gsのsetSupabaseServiceRoleKey()を実行してください');
    }
    
    // 2. ページマッピング取得（path → page_id）
    Logger.log('\n【Step 1】ページマッピング取得中...');
    const pageMapping = getGSCPageMapping(serviceRoleKey);
    Logger.log(`✅ ${Object.keys(pageMapping).length}ページのマッピング取得完了`);
    
    // 3. 既存データ削除（クリーンスタート）
    Logger.log('\n【Step 2】既存のGSCデータを削除中...');
    deleteExistingGSCData(serviceRoleKey);
    Logger.log('✅ 既存データ削除完了');
    
    // 4. 月ごとにデータ取得・保存
    Logger.log('\n【Step 3】月別データ取得・保存開始...');
    
    let totalRecords = 0;
    let processedMonths = 0;
    
    // 対象月のリストを生成
    const months = generateGSCMonthList(
      GSC_CONFIG.START_YEAR, GSC_CONFIG.START_MONTH,
      GSC_CONFIG.END_YEAR, GSC_CONFIG.END_MONTH
    );
    
    Logger.log(`対象期間: ${months.length}ヶ月`);
    
    for (const month of months) {
      Logger.log(`\n--- ${month.year}年${month.month}月 ---`);
      
      try {
        const records = processGSCMonth(month.year, month.month, pageMapping, serviceRoleKey);
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
function getGSCPageMapping(serviceRoleKey) {
  const url = `${GSC_CONFIG.SUPABASE_URL}/rest/v1/pages?site_id=eq.${GSC_CONFIG.SITE_ID}&status=eq.active&select=id,path`;
  
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
 * 既存のGSCデータを削除
 */
function deleteExistingGSCData(serviceRoleKey) {
  // 全レコード削除（日付で絞り込み - 2024年以降）
  const deleteUrl = `${GSC_CONFIG.SUPABASE_URL}/rest/v1/gsc_metrics_daily?date=gte.2024-01-01`;
  
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
function processGSCMonth(year, month, pageMapping, serviceRoleKey) {
  // 日付範囲を計算
  const startDate = `${year}-${String(month).padStart(2, '0')}-01`;
  const lastDay = new Date(year, month, 0).getDate();
  const endDate = `${year}-${String(month).padStart(2, '0')}-${String(lastDay).padStart(2, '0')}`;
  
  Logger.log(`期間: ${startDate} 〜 ${endDate}`);
  
  // GSCからデータ取得
  const gscData = fetchGSCDailyData(startDate, endDate);
  
  if (gscData.length === 0) {
    Logger.log('データなし');
    return 0;
  }
  
  Logger.log(`GSCから${gscData.length}件取得`);
  
  // Supabase形式に変換
  const records = [];
  
  gscData.forEach(row => {
    const pagePath = row.pagePath;
    const pageId = pageMapping[pagePath];
    
    if (!pageId) {
      // マッピングにないページはスキップ
      return;
    }
    
    records.push({
      page_id: pageId,
      date: row.date,
      impressions: row.impressions,
      clicks: row.clicks,
      ctr: row.ctr,
      avg_position: row.position
    });
  });
  
  Logger.log(`マッピング後: ${records.length}件`);
  
  if (records.length === 0) {
    return 0;
  }
  
  // バッチでSupabaseに保存
  const savedCount = batchInsertGSCToSupabase(records, serviceRoleKey);
  
  return savedCount;
}

/**
 * GSCから日別データを取得
 */
function fetchGSCDailyData(startDate, endDate) {
  const siteUrl = GSC_CONFIG.GSC_SITE_URL;
  
  // リクエストボディ（日別×ページ別）
  const requestBody = {
    startDate: startDate,
    endDate: endDate,
    dimensions: ['date', 'page'],
    rowLimit: 25000,
    dataState: 'final'
  };
  
  // API URL
  const apiUrl = 'https://www.googleapis.com/webmasters/v3/sites/' + 
                 encodeURIComponent(siteUrl) + '/searchAnalytics/query';
  
  // API呼び出し
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
    },
    payload: JSON.stringify(requestBody),
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch(apiUrl, options);
  const responseCode = response.getResponseCode();
  
  if (responseCode !== 200) {
    throw new Error(`GSC API エラー: ${responseCode} - ${response.getContentText()}`);
  }
  
  const result = JSON.parse(response.getContentText());
  
  if (!result.rows || result.rows.length === 0) {
    return [];
  }
  
  // データ整形
  return result.rows.map(row => {
    const fullUrl = row.keys[1];  // page URL
    const pagePath = fullUrl.replace(siteUrl.replace(/\/$/, ''), '');
    
    return {
      date: row.keys[0],  // YYYY-MM-DD形式
      pagePath: pagePath || '/',
      impressions: row.impressions || 0,
      clicks: row.clicks || 0,
      ctr: (row.ctr || 0) * 100,  // 0-1 → %
      position: row.position || 0
    };
  });
}

/**
 * Supabaseにバッチインサート
 */
function batchInsertGSCToSupabase(records, serviceRoleKey) {
  let totalInserted = 0;
  
  // バッチに分割
  for (let i = 0; i < records.length; i += GSC_CONFIG.BATCH_SIZE) {
    const batch = records.slice(i, i + GSC_CONFIG.BATCH_SIZE);
    
    const url = `${GSC_CONFIG.SUPABASE_URL}/rest/v1/gsc_metrics_daily`;
    
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
    if (i + GSC_CONFIG.BATCH_SIZE < records.length) {
      Utilities.sleep(500);
    }
  }
  
  return totalInserted;
}

/**
 * 対象月のリストを生成
 */
function generateGSCMonthList(startYear, startMonth, endYear, endMonth) {
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
function reprocessGSCSingleMonth(year, month) {
  Logger.log(`=== GSC ${year}年${month}月 再処理 ===`);
  
  const serviceRoleKey = PropertiesService.getScriptProperties()
    .getProperty('SUPABASE_SERVICE_ROLE_KEY');
  
  if (!serviceRoleKey) {
    Logger.log('❌ Service Role Keyが設定されていません');
    return;
  }
  
  // ページマッピング取得
  const pageMapping = getGSCPageMapping(serviceRoleKey);
  
  // 該当月のデータを削除
  const startDate = `${year}-${String(month).padStart(2, '0')}-01`;
  const lastDay = new Date(year, month, 0).getDate();
  const endDate = `${year}-${String(month).padStart(2, '0')}-${String(lastDay).padStart(2, '0')}`;
  
  Logger.log(`期間: ${startDate} 〜 ${endDate}`);
  Logger.log('既存データ削除中...');
  
  // 日付範囲で削除
  const deleteUrl = `${GSC_CONFIG.SUPABASE_URL}/rest/v1/gsc_metrics_daily?date=gte.${startDate}&date=lte.${endDate}`;
  
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
  const records = processGSCMonth(year, month, pageMapping, serviceRoleKey);
  
  Logger.log(`✅ 完了: ${records}件保存`);
  
  return records;
}

// ========================================
// 確認用関数
// ========================================

/**
 * 移行結果を確認
 */
function checkGSCMigrationResult() {
  Logger.log('=== GSC移行結果確認 ===');
  
  const serviceRoleKey = PropertiesService.getScriptProperties()
    .getProperty('SUPABASE_SERVICE_ROLE_KEY');
  
  if (!serviceRoleKey) {
    Logger.log('❌ Service Role Keyが設定されていません');
    return;
  }
  
  // 総レコード数と月別件数を取得
  const url = `${GSC_CONFIG.SUPABASE_URL}/rest/v1/gsc_metrics_daily?select=date`;
  
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
 * GSC API接続テスト
 */
function testGSCConnection() {
  Logger.log('=== GSC API接続テスト ===');
  
  try {
    const endDate = new Date();
    endDate.setDate(endDate.getDate() - 3);  // 3日前
    const startDate = new Date(endDate);
    startDate.setDate(startDate.getDate() - 7);  // その7日前
    
    const startDateStr = Utilities.formatDate(startDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const endDateStr = Utilities.formatDate(endDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    
    Logger.log(`テスト期間: ${startDateStr} 〜 ${endDateStr}`);
    
    const data = fetchGSCDailyData(startDateStr, endDateStr);
    
    Logger.log(`✅ GSC API接続成功！`);
    Logger.log(`取得件数: ${data.length}件`);
    
    if (data.length > 0) {
      Logger.log('\nサンプルデータ:');
      Logger.log(JSON.stringify(data[0], null, 2));
    }
    
    return true;
    
  } catch (error) {
    Logger.log(`❌ GSC API接続エラー: ${error.message}`);
    return false;
  }
}

/**
 * 12月のGSCデータを再取得
 */
function fixDecemberGSC() {
  reprocessGSCSingleMonth(2025, 12);
}