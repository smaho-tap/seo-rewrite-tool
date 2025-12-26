/**
 * DataForSEORankTracking.gs
 * DataForSEO APIで確定KWの順位を毎日取得してSupabaseに保存
 * 
 * 【機能】
 * - 確定KW（is_confirmed=true）の順位を自動取得
 * - Google デスクトップ / モバイル 両方取得
 * - AIO（AI Overview）の有無・掲載状況も取得
 * - 日次でSupabaseに蓄積
 * 
 * 【実行タイミング】
 * - 毎日午前6時（日次トリガー）
 * 
 * 【API費用目安】
 * - 15KW × 2デバイス = 30回/日
 * - 月額約$0.54（約80円）
 * 
 * 作成日: 2025/12/26
 */

const RANK_CONFIG = {
  SUPABASE_URL: 'https://dgzfdugpineqnoihopsl.supabase.co',
  SITE_URL: 'smaho-tap.com',
  DATAFORSEO_API_URL: 'https://api.dataforseo.com/v3/serp/google/organic/live/advanced',
  LOCATION_CODE: 2392,  // Japan
  LANGUAGE_CODE: 'ja'
};

/**
 * メイン関数：確定KWの順位を取得してSupabaseに保存
 */
function fetchAndSaveKeywordRankings() {
  Logger.log('=== DataForSEO順位取得開始 ===');
  Logger.log(`実行日時: ${new Date().toLocaleString('ja-JP')}`);
  
  const serviceRoleKey = PropertiesService.getScriptProperties()
    .getProperty('SUPABASE_SERVICE_ROLE_KEY');
  const dataforseoLogin = PropertiesService.getScriptProperties()
    .getProperty('DATAFORSEO_LOGIN');
  const dataforseoPassword = PropertiesService.getScriptProperties()
    .getProperty('DATAFORSEO_PASSWORD');
  
  if (!serviceRoleKey) {
    Logger.log('❌ SUPABASE_SERVICE_ROLE_KEY が設定されていません');
    return;
  }
  
  if (!dataforseoLogin || !dataforseoPassword) {
    Logger.log('❌ DataForSEO認証情報が設定されていません');
    Logger.log('スクリプトプロパティに DATAFORSEO_LOGIN と DATAFORSEO_PASSWORD を設定してください');
    return;
  }
  
  // 1. 確定KWを取得
  const keywords = getConfirmedKeywords(serviceRoleKey);
  Logger.log(`確定KW: ${keywords.length}件`);
  
  if (keywords.length === 0) {
    Logger.log('取得対象のKWがありません');
    return;
  }
  
  const today = formatDate(new Date());
  let successCount = 0;
  let errorCount = 0;
  
  // 2. 各KWの順位を取得
  for (const kw of keywords) {
    Logger.log(`\n--- ${kw.keyword} ---`);
    
    try {
      // デスクトップ順位取得
      const desktopResult = fetchSerpRank(
        kw.keyword, 
        'desktop', 
        dataforseoLogin, 
        dataforseoPassword
      );
      
      // モバイル順位取得
      const mobileResult = fetchSerpRank(
        kw.keyword, 
        'mobile', 
        dataforseoLogin, 
        dataforseoPassword
      );
      
      // Supabaseに保存
      const record = {
        keyword_id: kw.keyword_id,
        date: today,
        google_desktop_rank: desktopResult.rank,
        google_desktop_url: desktopResult.url,
        google_mobile_rank: mobileResult.rank,
        google_mobile_url: mobileResult.url,
        has_aio: desktopResult.hasAio || mobileResult.hasAio,
        aio_position: desktopResult.aioPosition || mobileResult.aioPosition,
        in_aio: desktopResult.inAio || mobileResult.inAio,
        source: 'dataforseo'
      };
      
      saveRankingToSupabase(serviceRoleKey, record);
      
      Logger.log(`  デスクトップ: ${desktopResult.rank || '圏外'}位`);
      Logger.log(`  モバイル: ${mobileResult.rank || '圏外'}位`);
      if (desktopResult.hasAio) {
        Logger.log(`  AIO: あり（自社${desktopResult.inAio ? '掲載' : '未掲載'}）`);
      }
      
      successCount++;
      
      // API制限対策（1秒待機）
      Utilities.sleep(1000);
      
    } catch (e) {
      Logger.log(`  ❌ エラー: ${e.message}`);
      errorCount++;
    }
  }
  
  Logger.log('\n=== 順位取得完了 ===');
  Logger.log(`成功: ${successCount}件, エラー: ${errorCount}件`);
  
  // 完了通知メール
  if (successCount > 0) {
    sendCompletionNotification(successCount, errorCount, keywords.length);
  }
}

/**
 * 確定KWをSupabaseから取得
 */
function getConfirmedKeywords(serviceRoleKey) {
  const url = `${RANK_CONFIG.SUPABASE_URL}/rest/v1/confirmed_keywords?select=keyword_id,keyword,path`;
  
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
    throw new Error(`確定KW取得エラー: ${response.getContentText()}`);
  }
  
  return JSON.parse(response.getContentText());
}

/**
 * DataForSEO APIでSERP順位を取得
 */
function fetchSerpRank(keyword, device, login, password) {
  const payload = [{
    keyword: keyword,
    location_code: RANK_CONFIG.LOCATION_CODE,
    language_code: RANK_CONFIG.LANGUAGE_CODE,
    device: device,
    depth: 100  // 100位まで検索
  }];
  
  const auth = Utilities.base64Encode(`${login}:${password}`);
  
  const response = UrlFetchApp.fetch(RANK_CONFIG.DATAFORSEO_API_URL, {
    method: 'post',
    headers: {
      'Authorization': `Basic ${auth}`,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  
  if (response.getResponseCode() !== 200) {
    throw new Error(`DataForSEO APIエラー: ${response.getResponseCode()}`);
  }
  
  const data = JSON.parse(response.getContentText());
  
  if (!data.tasks || data.tasks.length === 0 || !data.tasks[0].result) {
    return { rank: null, url: null, hasAio: false, inAio: false, aioPosition: null };
  }
  
  const result = data.tasks[0].result[0];
  const items = result.items || [];
  
  // 自サイトの順位を検索
  let rank = null;
  let url = null;
  let hasAio = false;
  let inAio = false;
  let aioPosition = null;
  
  for (const item of items) {
    // AI Overview（AIO）チェック
    if (item.type === 'ai_overview') {
      hasAio = true;
      aioPosition = item.rank_absolute;
      
      // AIO内の参照サイトをチェック
      if (item.items) {
        for (const ref of item.items) {
          if (ref.domain && ref.domain.includes(RANK_CONFIG.SITE_URL)) {
            inAio = true;
            break;
          }
        }
      }
    }
    
    // オーガニック検索結果
    if (item.type === 'organic' && item.domain && item.domain.includes(RANK_CONFIG.SITE_URL)) {
      if (rank === null) {  // 最初にヒットした順位を採用
        rank = item.rank_absolute;
        url = item.url;
      }
    }
  }
  
  return { rank, url, hasAio, inAio, aioPosition };
}

/**
 * 順位データをSupabaseに保存
 */
function saveRankingToSupabase(serviceRoleKey, record) {
  const url = `${RANK_CONFIG.SUPABASE_URL}/rest/v1/keyword_rankings_daily`;
  
  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: {
      'apikey': serviceRoleKey,
      'Authorization': `Bearer ${serviceRoleKey}`,
      'Content-Type': 'application/json',
      'Prefer': 'resolution=merge-duplicates'  // 重複時は更新
    },
    payload: JSON.stringify(record),
    muteHttpExceptions: true
  });
  
  const code = response.getResponseCode();
  if (code !== 201 && code !== 200) {
    throw new Error(`保存エラー: ${response.getContentText()}`);
  }
}

/**
 * 日付フォーマット（YYYY-MM-DD）
 */
function formatDate(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

/**
 * 完了通知メール送信
 */
function sendCompletionNotification(successCount, errorCount, totalCount) {
  const email = 'foster_inc@icloud.com';
  const subject = `【順位取得完了】${successCount}/${totalCount}件`;
  
  const body = `
DataForSEO順位取得が完了しました。

実行日時: ${new Date().toLocaleString('ja-JP')}
成功: ${successCount}件
エラー: ${errorCount}件

※詳細はSupabaseの keyword_rankings_daily テーブルをご確認ください。
`;
  
  // 毎日通知は不要なのでコメントアウト
  // MailApp.sendEmail(email, subject, body);
}

/**
 * 日次トリガー設定（1回実行）
 */
function setupDailyRankingTrigger() {
  // 既存トリガー削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'fetchAndSaveKeywordRankings') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // 毎日午前6時に実行
  ScriptApp.newTrigger('fetchAndSaveKeywordRankings')
    .timeBased()
    .atHour(6)
    .everyDays(1)
    .create();
  
  Logger.log('✅ 日次トリガー設定完了（毎朝6時）');
}

/**
 * トリガー削除
 */
function removeDailyRankingTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'fetchAndSaveKeywordRankings') {
      ScriptApp.deleteTrigger(trigger);
      Logger.log('トリガー削除: fetchAndSaveKeywordRankings');
    }
  });
}

/**
 * 手動テスト用（1KWだけ取得）
 */
function testFetchSingleKeyword() {
  Logger.log('=== テスト実行 ===');
  
  const dataforseoLogin = PropertiesService.getScriptProperties()
    .getProperty('DATAFORSEO_LOGIN');
  const dataforseoPassword = PropertiesService.getScriptProperties()
    .getProperty('DATAFORSEO_PASSWORD');
  
  if (!dataforseoLogin || !dataforseoPassword) {
    Logger.log('❌ DataForSEO認証情報が設定されていません');
    return;
  }
  
  const testKeyword = 'iPhone 買い替え時期';
  
  Logger.log(`テストKW: ${testKeyword}`);
  
  const desktopResult = fetchSerpRank(testKeyword, 'desktop', dataforseoLogin, dataforseoPassword);
  Logger.log(`デスクトップ: ${desktopResult.rank || '圏外'}位`);
  Logger.log(`URL: ${desktopResult.url || 'なし'}`);
  Logger.log(`AIO: ${desktopResult.hasAio ? 'あり' : 'なし'}`);
  if (desktopResult.hasAio) {
    Logger.log(`AIO内自社: ${desktopResult.inAio ? '掲載' : '未掲載'}`);
  }
  
  const mobileResult = fetchSerpRank(testKeyword, 'mobile', dataforseoLogin, dataforseoPassword);
  Logger.log(`モバイル: ${mobileResult.rank || '圏外'}位`);
}

/**
 * 全KW手動実行（即時実行用）
 */
function runManualRankingFetch() {
  fetchAndSaveKeywordRankings();
}

/**
 * 順位レポート生成（週次サマリー用）
 */
function generateWeeklyRankingReport() {
  Logger.log('=== 週次順位レポート生成 ===');
  
  const serviceRoleKey = PropertiesService.getScriptProperties()
    .getProperty('SUPABASE_SERVICE_ROLE_KEY');
  
  if (!serviceRoleKey) {
    Logger.log('❌ Service Role Keyが設定されていません');
    return;
  }
  
  // 過去7日間の順位変動を取得
  const url = `${RANK_CONFIG.SUPABASE_URL}/rest/v1/rpc/get_weekly_ranking_summary`;
  
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
    Logger.log(`レポート取得エラー: ${response.getContentText()}`);
    return;
  }
  
  const report = JSON.parse(response.getContentText());
  
  // メール送信
  let body = '【週次順位レポート】\n\n';
  
  report.forEach(item => {
    const change = item.rank_change > 0 ? `↑${item.rank_change}` : 
                   item.rank_change < 0 ? `↓${Math.abs(item.rank_change)}` : '→';
    body += `${item.keyword}: ${item.current_rank}位 (${change})\n`;
  });
  
  MailApp.sendEmail({
    to: 'foster_inc@icloud.com',
    subject: '【SEOツール】週次順位レポート',
    body: body
  });
  
  Logger.log('✅ レポートメール送信完了');
}