/**
 * ClarityToSupabase.gs
 * Clarity APIから日次データをSupabaseに蓄積するスクリプト
 * 
 * 【注意】
 * - Clarity APIは過去3日分のデータのみ取得可能
 * - 毎日実行して蓄積する必要がある
 * - 投稿ページ（status=active）のみ保存
 * 
 * 【使い方】
 * 1. このファイルをGASに追加
 * 2. setupDailyClarityTrigger() を1回実行してトリガー設定
 * 3. 毎朝6時に自動実行される（GA4/GSCの後）
 * 
 * 【手動実行】
 * - runDailyClarityUpdate() を実行
 */

const CLARITY_CONFIG = {
  SUPABASE_URL: 'https://dgzfdugpineqnoihopsl.supabase.co',
  SITE_ID: '853ea711-7644-451e-872b-dea1b54fa8c7',
  SITE_BASE_URL: 'https://smaho-tap.com',
  CLARITY_API_URL: 'https://www.clarity.ms/export-data/api/v1/project-live-insights'
};

/**
 * 日次Clarity更新のメイン関数
 */
function runDailyClarityUpdate() {
  Logger.log('=== 日次Clarity更新開始 ===');
  Logger.log(`実行日時: ${new Date().toLocaleString('ja-JP')}`);
  
  const serviceRoleKey = PropertiesService.getScriptProperties()
    .getProperty('SUPABASE_SERVICE_ROLE_KEY');
  const clarityToken = PropertiesService.getScriptProperties()
    .getProperty('CLARITY_API_TOKEN');
  
  if (!serviceRoleKey) {
    Logger.log('❌ SUPABASE_SERVICE_ROLE_KEYが設定されていません');
    return;
  }
  
  if (!clarityToken) {
    Logger.log('❌ CLARITY_API_TOKENが設定されていません');
    return;
  }
  
  // ページマッピング取得（投稿ページのみ）
  const pageMapping = getPageMappingForClarity(serviceRoleKey);
  Logger.log(`ページマッピング: ${Object.keys(pageMapping).length}件（投稿ページのみ）`);
  
  // Clarityデータ取得
  try {
    const clarityData = fetchClarityDataFromAPI(clarityToken);
    
    if (clarityData.length === 0) {
      Logger.log('⚠️ Clarityデータなし');
      return;
    }
    
    Logger.log(`Clarityデータ取得: ${clarityData.length}件`);
    
    // Supabase形式に変換・保存
    const savedCount = saveClarityToSupabase(serviceRoleKey, pageMapping, clarityData);
    Logger.log(`✅ Clarity: ${savedCount}件保存`);
    
  } catch (e) {
    Logger.log(`❌ Clarityエラー: ${e.message}`);
  }
  
  Logger.log('=== 日次Clarity更新完了 ===');
}

/**
 * ページマッピング取得（path → page_id）
 * ★ status=active（投稿ページ）のみ取得
 */
function getPageMappingForClarity(serviceRoleKey) {
  const url = `${CLARITY_CONFIG.SUPABASE_URL}/rest/v1/pages?site_id=eq.${CLARITY_CONFIG.SITE_ID}&status=eq.active&select=id,path`;
  
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
 * Clarity Data Export APIからデータ取得
 */
function fetchClarityDataFromAPI(clarityToken) {
  // 前日分のみ取得（初回は3に変更して実行）
  const params = {
    numOfDays: 1,
    dimension1: 'URL'
  };
  
  const url = CLARITY_CONFIG.CLARITY_API_URL + '?' + 
    Object.keys(params).map(key => key + '=' + encodeURIComponent(params[key])).join('&');
  
  const response = UrlFetchApp.fetch(url, {
    method: 'GET',
    headers: {
      'Authorization': 'Bearer ' + clarityToken,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  });
  
  const statusCode = response.getResponseCode();
  
  if (statusCode === 200) {
    const rawData = JSON.parse(response.getContentText());
    return parseClarityAPIResponse(rawData);
  } else if (statusCode === 429) {
    Logger.log('⚠️ Clarity API制限超過（1日10リクエスト）');
    return [];
  } else if (statusCode === 401) {
    Logger.log('❌ Clarity認証エラー: APIトークンが無効');
    return [];
  } else {
    Logger.log(`❌ Clarity APIエラー: ${statusCode}`);
    Logger.log(response.getContentText());
    return [];
  }
}

/**
 * Clarity APIレスポンスをパース（修正版）
 * 
 * APIレスポンス構造:
 * - DeadClickCount: {sessionsCount, subTotal, Url}
 * - RageClickCount: {sessionsCount, subTotal, Url}
 * - QuickbackClick: {sessionsCount, subTotal, Url}
 * - ScrollDepth: {averageScrollDepth, Url}
 * - EngagementTime: {totalTime, activeTime, Url}
 * - Traffic: {totalSessionCount, Url}
 */
function parseClarityAPIResponse(rawData) {
  const result = {};
  const today = new Date();
  const dateStr = formatDateForClarity(today);
  
  if (!Array.isArray(rawData) || rawData.length === 0) {
    return [];
  }
  
  rawData.forEach(metric => {
    const metricName = metric.metricName;
    
    if (!metric.information) return;
    
    metric.information.forEach(item => {
      const rawUrl = item.Url;
      if (!rawUrl) return;
      
      // URLを正規化（パス部分のみ抽出、先頭スラッシュなし）
      let path = normalizeUrlToPath(rawUrl);
      
      // 空パス（トップページ）はスキップ
      if (!path || path === '') return;
      
      if (!result[path]) {
        result[path] = {
          date: dateStr,
          path: path,
          sessions: 0,
          scroll_depth: 0,
          dead_clicks: 0,
          rage_clicks: 0,
          quick_backs: 0,
          engagement_time: 0
        };
      }
      
      // メトリックごとに適切なフィールドから値を取得
      switch (metricName) {
        case 'DeadClickCount':
          result[path].dead_clicks = parseInt(item.subTotal) || 0;
          // sessionsCountも取得
          if (item.sessionsCount && result[path].sessions === 0) {
            result[path].sessions = parseInt(item.sessionsCount) || 0;
          }
          break;
          
        case 'RageClickCount':
          result[path].rage_clicks = parseInt(item.subTotal) || 0;
          if (item.sessionsCount && result[path].sessions === 0) {
            result[path].sessions = parseInt(item.sessionsCount) || 0;
          }
          break;
          
        case 'QuickbackClick':
          result[path].quick_backs = parseInt(item.subTotal) || 0;
          if (item.sessionsCount && result[path].sessions === 0) {
            result[path].sessions = parseInt(item.sessionsCount) || 0;
          }
          break;
          
        case 'ScrollDepth':
          // ScrollDepthはaverageScrollDepthフィールド
          result[path].scroll_depth = parseFloat(item.averageScrollDepth) || 0;
          break;
          
        case 'EngagementTime':
          // EngagementTimeはtotalTimeまたはactiveTime
          const engTime = parseFloat(item.totalTime) || parseFloat(item.activeTime) || 0;
          result[path].engagement_time = engTime;
          break;
          
        case 'Traffic':
          // Trafficはセッション数（totalSessionCount）
          if (item.totalSessionCount) {
            result[path].sessions = parseInt(item.totalSessionCount) || 0;
          }
          break;
      }
    });
  });
  
  return Object.values(result);
}

/**
 * URLをパスに正規化（GAS対応版）
 */
function normalizeUrlToPath(url) {
  if (!url) return '';
  
  let path = String(url);
  
  // フルURLからパス部分を抽出
  // https://smaho-tap.com/xxx → xxx
  // https://smaho-tap.com → ''
  if (path.indexOf('://') !== -1) {
    // プロトコルとドメインを除去
    // 例: https://smaho-tap.com/page → /page
    const parts = path.split('://');
    if (parts.length > 1) {
      const afterProtocol = parts[1];
      const slashIndex = afterProtocol.indexOf('/');
      if (slashIndex !== -1) {
        path = afterProtocol.substring(slashIndex);
      } else {
        // ドメインのみ（トップページ）
        path = '';
      }
    }
  }
  
  // 先頭スラッシュを除去
  if (path.startsWith('/')) {
    path = path.substring(1);
  }
  
  // 末尾スラッシュを除去
  if (path.endsWith('/') && path.length > 0) {
    path = path.slice(0, -1);
  }
  
  // アンカーを除去
  const hashIndex = path.indexOf('#');
  if (hashIndex !== -1) {
    path = path.substring(0, hashIndex);
  }
  
  // クエリパラメータを除去
  const queryIndex = path.indexOf('?');
  if (queryIndex !== -1) {
    path = path.substring(0, queryIndex);
  }
  
  return path;
}

/**
 * ClarityデータをSupabaseに保存
 */
function saveClarityToSupabase(serviceRoleKey, pageMapping, clarityData) {
  const records = [];
  
  clarityData.forEach(item => {
    const pageId = pageMapping[item.path];
    if (!pageId) return;  // 投稿ページ以外はスキップ
    
    // エンゲージメントスコア計算（0-100）
    // engagement_time（秒）を基にスコア化
    let engagementScore = 0;
    if (item.engagement_time > 0) {
      // 120秒以上で100点、0秒で0点の線形スコア
      engagementScore = Math.min(100, Math.round(item.engagement_time / 120 * 100));
    }
    
    records.push({
      page_id: pageId,
      date: item.date,
      total_sessions: item.sessions,
      pages_per_session: 0,  // Clarityからは取得不可
      scroll_depth: item.scroll_depth,
      rage_clicks: item.rage_clicks,
      dead_clicks: item.dead_clicks,
      quick_backs: item.quick_backs,
      engagement_score: engagementScore
    });
  });
  
  if (records.length === 0) return 0;
  
  // 既存データ削除（同じ日付）
  const dateStr = records[0].date;
  deleteClarityExistingRecords(serviceRoleKey, dateStr);
  
  // Supabaseに保存
  const url = `${CLARITY_CONFIG.SUPABASE_URL}/rest/v1/clarity_metrics_daily`;
  
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
 * 既存Clarityレコード削除
 */
function deleteClarityExistingRecords(serviceRoleKey, dateStr) {
  const deleteUrl = `${CLARITY_CONFIG.SUPABASE_URL}/rest/v1/clarity_metrics_daily?date=eq.${dateStr}`;
  
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
 * 日付フォーマット（YYYY-MM-DD）
 */
function formatDateForClarity(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

/**
 * 日次Clarityトリガー設定（1回実行）
 */
function setupDailyClarityTrigger() {
  // 既存トリガー削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'runDailyClarityUpdate') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // 毎日午前6時に実行（GA4/GSCの1時間後）
  ScriptApp.newTrigger('runDailyClarityUpdate')
    .timeBased()
    .atHour(6)
    .everyDays(1)
    .create();
  
  Logger.log('✅ 日次Clarityトリガー設定完了（毎朝6時）');
}

/**
 * Clarityトリガー削除
 */
function removeDailyClarityTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'runDailyClarityUpdate') {
      ScriptApp.deleteTrigger(trigger);
      Logger.log('トリガー削除: runDailyClarityUpdate');
    }
  });
}

/**
 * 手動テスト用
 */
function testClarityUpdate() {
  Logger.log('=== Clarityテスト実行 ===');
  
  const clarityToken = PropertiesService.getScriptProperties()
    .getProperty('CLARITY_API_TOKEN');
  
  if (!clarityToken) {
    Logger.log('❌ CLARITY_API_TOKENが設定されていません');
    return;
  }
  
  Logger.log('Clarity APIにリクエスト中...');
  const clarityData = fetchClarityDataFromAPI(clarityToken);
  
  Logger.log(`取得件数: ${clarityData.length}件`);
  
  if (clarityData.length > 0) {
    Logger.log('サンプルデータ（最初の5件）:');
    clarityData.slice(0, 5).forEach((item, i) => {
      Logger.log(`  ${i + 1}. ${item.path}`);
      Logger.log(`     sessions: ${item.sessions}, scroll: ${item.scroll_depth}%`);
      Logger.log(`     dead: ${item.dead_clicks}, rage: ${item.rage_clicks}, quick: ${item.quick_backs}`);
      Logger.log(`     engagement: ${item.engagement_time}秒`);
    });
  }
}

/**
 * 本番実行テスト（Supabase保存まで）
 */
function testClarityFullUpdate() {
  Logger.log('=== Clarity本番テスト実行 ===');
  runDailyClarityUpdate();
}