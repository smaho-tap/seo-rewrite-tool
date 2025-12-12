/**
 * GyronToSupabase.gs
 * GyronSEOデータ（順位・検索ボリューム）とpublish_dateをSupabaseに同期
 * 
 * 【実行タイミング】
 * - GyronCSVインポート後に手動実行
 * - または週次トリガー（月曜6:00）
 * 
 * 【同期対象】
 * 1. target_keywords.gyron_position - ターゲットKW順位
 * 2. target_keywords.search_volume - 検索ボリューム（更新）
 * 3. pages.first_published_at - 初回公開日
 * 
 * 作成日: 2025/12/13
 */

const GYRON_CONFIG = {
  SUPABASE_URL: 'https://dgzfdugpineqnoihopsl.supabase.co',
  SITE_ID: '853ea711-7644-451e-872b-dea1b54fa8c7'
};

/**
 * メイン関数: GyronデータとPublish DateをSupabaseに同期
 */
function syncGyronAndPublishDateToSupabase() {
  Logger.log('=== Gyron・PublishDate同期開始 ===');
  Logger.log(`実行日時: ${new Date().toLocaleString('ja-JP')}`);
  
  const serviceRoleKey = PropertiesService.getScriptProperties()
    .getProperty('SUPABASE_SERVICE_ROLE_KEY');
  
  if (!serviceRoleKey) {
    Logger.log('❌ Service Role Keyが設定されていません');
    return { success: false, error: 'Service Role Key未設定' };
  }
  
  const results = {
    gyronPosition: 0,
    publishDate: 0,
    errors: []
  };
  
  try {
    // 1. Gyron順位・検索ボリュームを同期
    results.gyronPosition = syncGyronPositionToSupabase(serviceRoleKey);
    Logger.log(`✅ Gyron順位同期: ${results.gyronPosition}件`);
  } catch (e) {
    Logger.log(`❌ Gyron同期エラー: ${e.message}`);
    results.errors.push(`Gyron: ${e.message}`);
  }
  
  try {
    // 2. Publish Dateを同期
    results.publishDate = syncPublishDateToSupabase(serviceRoleKey);
    Logger.log(`✅ PublishDate同期: ${results.publishDate}件`);
  } catch (e) {
    Logger.log(`❌ PublishDate同期エラー: ${e.message}`);
    results.errors.push(`PublishDate: ${e.message}`);
  }
  
  Logger.log('=== 同期完了 ===');
  Logger.log(`結果: Gyron=${results.gyronPosition}件, PublishDate=${results.publishDate}件`);
  
  return results;
}


/**
 * Gyron順位・検索ボリュームをSupabaseに同期
 */
function syncGyronPositionToSupabase(serviceRoleKey) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // ターゲットKW分析シートからデータ取得
  const kwSheet = ss.getSheetByName('ターゲットKW分析');
  if (!kwSheet) {
    throw new Error('ターゲットKW分析シートが見つかりません');
  }
  
  const kwData = kwSheet.getDataRange().getValues();
  const kwHeaders = kwData[0];
  
  // 列インデックス取得
  const urlIdx = kwHeaders.indexOf('page_url');
  const keywordIdx = kwHeaders.indexOf('target_keyword');
  const gyronPosIdx = kwHeaders.indexOf('gyron_position');
  const volumeIdx = kwHeaders.indexOf('search_volume');
  
  if (urlIdx === -1 || keywordIdx === -1) {
    throw new Error('必要な列（page_url, target_keyword）が見つかりません');
  }
  
  // Supabaseのページマッピング取得
  const pageMapping = getPageMappingForGyron(serviceRoleKey);
  
  // Supabaseのtarget_keywordsを取得（既存データ確認用）
  const existingKW = getExistingTargetKeywords(serviceRoleKey);
  
  let updateCount = 0;
  const now = new Date().toISOString();
  
  for (let i = 1; i < kwData.length; i++) {
    const row = kwData[i];
    let pageUrl = (row[urlIdx] || '').toString().trim();
    const keyword = (row[keywordIdx] || '').toString().trim();
    const gyronPosition = parseInt(row[gyronPosIdx]) || null;
    const searchVolume = parseInt(row[volumeIdx]) || 0;
    
    if (!pageUrl || !keyword) continue;
    
    // URLを正規化（パス形式に）
    pageUrl = normalizeToPath(pageUrl);
    
    const pageId = pageMapping[pageUrl];
    if (!pageId) {
      Logger.log(`ページ未登録: ${pageUrl}`);
      continue;
    }
    
    // 既存のtarget_keywordを検索
    const existingKey = `${pageId}_${keyword}`;
    const existing = existingKW[existingKey];
    
    if (existing) {
      // 更新（gyron_positionとsearch_volumeを更新）
      updateTargetKeyword(serviceRoleKey, existing.id, {
        gyron_position: gyronPosition,
        search_volume: searchVolume,
        gyron_updated_at: now
      });
      updateCount++;
    } else {
      // 新規作成はスキップ（既存のtarget_keywordsのみ更新）
      Logger.log(`target_keyword未登録: ${pageUrl} - ${keyword}`);
    }
  }
  
  return updateCount;
}


/**
 * Publish DateをSupabaseに同期
 */
function syncPublishDateToSupabase(serviceRoleKey) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 統合データシートからデータ取得
  const intSheet = ss.getSheetByName('統合データ');
  if (!intSheet) {
    throw new Error('統合データシートが見つかりません');
  }
  
  const intData = intSheet.getDataRange().getValues();
  const intHeaders = intData[0];
  
  // 列インデックス取得
  const urlIdx = intHeaders.indexOf('page_url');
  const publishDateIdx = intHeaders.indexOf('publish_date');
  const toukoubi = intHeaders.indexOf('投稿日');  // 別名の可能性
  
  // publish_dateがなければ投稿日列を使用
  const dateIdx = publishDateIdx !== -1 ? publishDateIdx : toukoubi;
  
  if (urlIdx === -1) {
    throw new Error('page_url列が見つかりません');
  }
  
  // Supabaseのページマッピング取得
  const pageMapping = getPageMappingForGyron(serviceRoleKey);
  
  let updateCount = 0;
  
  for (let i = 1; i < intData.length; i++) {
    const row = intData[i];
    let pageUrl = (row[urlIdx] || '').toString().trim();
    let publishDate = dateIdx !== -1 ? row[dateIdx] : null;
    
    if (!pageUrl) continue;
    
    // URLを正規化
    pageUrl = normalizeToPath(pageUrl);
    
    const pageId = pageMapping[pageUrl];
    if (!pageId) continue;
    
    // 日付を変換
    let isoDate = null;
    if (publishDate) {
      if (publishDate instanceof Date) {
        isoDate = publishDate.toISOString();
      } else if (typeof publishDate === 'string' && publishDate.trim()) {
        // 文字列の場合はパース
        const parsed = new Date(publishDate);
        if (!isNaN(parsed.getTime())) {
          isoDate = parsed.toISOString();
        }
      }
    }
    
    if (isoDate) {
      updatePagePublishDate(serviceRoleKey, pageId, isoDate);
      updateCount++;
    }
  }
  
  return updateCount;
}


/**
 * ページマッピング取得（path → page_id）
 */
function getPageMappingForGyron(serviceRoleKey) {
  const url = `${GYRON_CONFIG.SUPABASE_URL}/rest/v1/pages?site_id=eq.${GYRON_CONFIG.SITE_ID}&select=id,path`;
  
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
    // パスを正規化
    mapping[page.path] = page.id;
  });
  
  return mapping;
}


/**
 * 既存のtarget_keywordsを取得
 */
function getExistingTargetKeywords(serviceRoleKey) {
  const url = `${GYRON_CONFIG.SUPABASE_URL}/rest/v1/target_keywords?select=id,page_id,keyword`;
  
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
    throw new Error(`target_keywords取得エラー: ${response.getContentText()}`);
  }
  
  const keywords = JSON.parse(response.getContentText());
  const mapping = {};
  
  keywords.forEach(kw => {
    const key = `${kw.page_id}_${kw.keyword}`;
    mapping[key] = kw;
  });
  
  return mapping;
}


/**
 * target_keywordを更新
 */
function updateTargetKeyword(serviceRoleKey, keywordId, data) {
  const url = `${GYRON_CONFIG.SUPABASE_URL}/rest/v1/target_keywords?id=eq.${keywordId}`;
  
  const response = UrlFetchApp.fetch(url, {
    method: 'patch',
    headers: {
      'apikey': serviceRoleKey,
      'Authorization': `Bearer ${serviceRoleKey}`,
      'Content-Type': 'application/json',
      'Prefer': 'return=minimal'
    },
    payload: JSON.stringify(data),
    muteHttpExceptions: true
  });
  
  if (response.getResponseCode() !== 204 && response.getResponseCode() !== 200) {
    Logger.log(`target_keyword更新エラー: ${response.getContentText()}`);
  }
}


/**
 * pageのfirst_published_atを更新
 */
function updatePagePublishDate(serviceRoleKey, pageId, isoDate) {
  const url = `${GYRON_CONFIG.SUPABASE_URL}/rest/v1/pages?id=eq.${pageId}`;
  
  const response = UrlFetchApp.fetch(url, {
    method: 'patch',
    headers: {
      'apikey': serviceRoleKey,
      'Authorization': `Bearer ${serviceRoleKey}`,
      'Content-Type': 'application/json',
      'Prefer': 'return=minimal'
    },
    payload: JSON.stringify({ first_published_at: isoDate }),
    muteHttpExceptions: true
  });
  
  if (response.getResponseCode() !== 204 && response.getResponseCode() !== 200) {
    Logger.log(`publish_date更新エラー: ${response.getContentText()}`);
  }
}


/**
 * URLをパス形式に正規化
 */
function normalizeToPath(url) {
  if (!url) return '';
  
  let path = url.toString();
  
  // フルURLからパスを抽出
  if (path.includes('://')) {
    try {
      const urlObj = new URL(path);
      path = urlObj.pathname;
    } catch (e) {
      const match = path.match(/https?:\/\/[^\/]+(\/.*)?/);
      if (match && match[1]) {
        path = match[1];
      }
    }
  }
  
  // 先頭スラッシュを確保
  if (!path.startsWith('/')) {
    path = '/' + path;
  }
  
  // 末尾スラッシュを除去
  path = path.replace(/\/$/, '');
  
  return path;
}


// ===========================================
// 手動実行・トリガー設定
// ===========================================

/**
 * 手動実行用（メニューから呼び出し）
 */
function runGyronSync() {
  const ui = SpreadsheetApp.getUi();
  
  const result = ui.alert(
    'Gyron・PublishDate同期',
    'ターゲットKW分析シートと統合データシートからSupabaseに同期します。\n\n' +
    '同期対象:\n' +
    '1. target_keywords.gyron_position（順位）\n' +
    '2. target_keywords.search_volume（検索ボリューム）\n' +
    '3. pages.first_published_at（初回公開日）\n\n' +
    '実行しますか？',
    ui.ButtonSet.YES_NO
  );
  
  if (result !== ui.Button.YES) {
    ui.alert('キャンセルしました');
    return;
  }
  
  try {
    const execResult = syncGyronAndPublishDateToSupabase();
    
    if (execResult.errors.length === 0) {
      ui.alert(
        '完了',
        `Supabase同期が完了しました。\n\n` +
        `Gyron順位: ${execResult.gyronPosition}件\n` +
        `PublishDate: ${execResult.publishDate}件`,
        ui.ButtonSet.OK
      );
    } else {
      ui.alert(
        '一部エラー',
        `同期完了（一部エラーあり）\n\n` +
        `Gyron順位: ${execResult.gyronPosition}件\n` +
        `PublishDate: ${execResult.publishDate}件\n\n` +
        `エラー:\n${execResult.errors.join('\n')}`,
        ui.ButtonSet.OK
      );
    }
  } catch (error) {
    Logger.log(`エラー: ${error.message}`);
    ui.alert('エラー', error.toString(), ui.ButtonSet.OK);
  }
}


/**
 * 週次トリガー設定（GyronCSVインポート後に実行）
 * 毎週月曜 6:00
 */
function setupGyronSyncTrigger() {
  // 既存トリガー削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'syncGyronAndPublishDateToSupabase') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // 毎週月曜6時に実行（GyronCSVインポート後）
  ScriptApp.newTrigger('syncGyronAndPublishDateToSupabase')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(6)
    .nearMinute(0)
    .create();
  
  Logger.log('✅ 週次トリガー設定完了（毎週月曜6:00）');
}


/**
 * テスト: 同期実行（ドライラン）
 */
function testGyronSyncDryRun() {
  Logger.log('=== ドライランテスト ===');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // ターゲットKW分析シートの確認
  const kwSheet = ss.getSheetByName('ターゲットKW分析');
  if (kwSheet) {
    const kwData = kwSheet.getDataRange().getValues();
    Logger.log(`ターゲットKW分析: ${kwData.length - 1}件`);
    
    // サンプル出力
    if (kwData.length > 1) {
      const headers = kwData[0];
      Logger.log(`列: ${headers.join(', ')}`);
      Logger.log(`サンプル: ${kwData[1].slice(0, 5).join(', ')}`);
    }
  } else {
    Logger.log('ターゲットKW分析シートなし');
  }
  
  // 統合データシートの確認
  const intSheet = ss.getSheetByName('統合データ');
  if (intSheet) {
    const intData = intSheet.getDataRange().getValues();
    const headers = intData[0];
    
    const publishIdx = headers.indexOf('publish_date');
    const toukoubiIdx = headers.indexOf('投稿日');
    
    Logger.log(`統合データ: ${intData.length - 1}件`);
    Logger.log(`publish_date列: ${publishIdx !== -1 ? 'あり' : 'なし'}`);
    Logger.log(`投稿日列: ${toukoubiIdx !== -1 ? 'あり' : 'なし'}`);
    
    // publish_dateの入力状況確認
    let filledCount = 0;
    const dateIdx = publishIdx !== -1 ? publishIdx : toukoubiIdx;
    if (dateIdx !== -1) {
      for (let i = 1; i < intData.length; i++) {
        if (intData[i][dateIdx]) filledCount++;
      }
    }
    Logger.log(`日付入力済み: ${filledCount}/${intData.length - 1}件`);
  }
  
  Logger.log('=== テスト完了 ===');
}
/**
 * TargetKeywordsMigration.gs
 * ターゲットKW分析シートの全データをSupabaseに移行
 * 
 * GyronToSupabase.gs に追記してください
 */

/**
 * ターゲットKW分析シートの全データをSupabaseに移行
 */
function migrateAllTargetKeywordsToSupabase() {
  Logger.log('=== ターゲットKW全件移行開始 ===');
  
  const serviceRoleKey = PropertiesService.getScriptProperties()
    .getProperty('SUPABASE_SERVICE_ROLE_KEY');
  
  if (!serviceRoleKey) {
    Logger.log('❌ Service Role Keyが設定されていません');
    return { success: false, error: 'Service Role Key未設定' };
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const kwSheet = ss.getSheetByName('ターゲットKW分析');
  
  if (!kwSheet) {
    Logger.log('❌ ターゲットKW分析シートが見つかりません');
    return { success: false, error: 'シート未発見' };
  }
  
  const kwData = kwSheet.getDataRange().getValues();
  const kwHeaders = kwData[0];
  
  // 列インデックス
  const urlIdx = kwHeaders.indexOf('page_url');
  const keywordIdx = kwHeaders.indexOf('target_keyword');
  const gyronPosIdx = kwHeaders.indexOf('gyron_position');
  const volumeIdx = kwHeaders.indexOf('search_volume');
  const statusIdx = kwHeaders.indexOf('status');
  
  Logger.log(`列インデックス: url=${urlIdx}, kw=${keywordIdx}, pos=${gyronPosIdx}, vol=${volumeIdx}`);
  
  // ページマッピング取得
  const pageMapping = getPageMappingForMigration(serviceRoleKey);
  Logger.log(`ページマッピング: ${Object.keys(pageMapping).length}件`);
  
  // 既存のtarget_keywordsを取得
  const existingKW = getExistingTargetKeywordsMap(serviceRoleKey);
  Logger.log(`既存KW: ${Object.keys(existingKW).length}件`);
  
  const now = new Date().toISOString();
  let insertCount = 0;
  let updateCount = 0;
  let skipCount = 0;
  const errors = [];
  
  // バッチ処理用の配列
  const toInsert = [];
  const toUpdate = [];
  
  for (let i = 1; i < kwData.length; i++) {
    const row = kwData[i];
    let pageUrl = (row[urlIdx] || '').toString().trim();
    const keyword = (row[keywordIdx] || '').toString().trim();
    const gyronPosition = parseInt(row[gyronPosIdx]) || null;
    const searchVolume = parseInt(row[volumeIdx]) || 0;
    const status = (row[statusIdx] || '').toString().trim();
    
    if (!pageUrl || !keyword) {
      skipCount++;
      continue;
    }
    
    // URLを正規化
    pageUrl = normalizeToPathForMigration(pageUrl);
    
    const pageId = pageMapping[pageUrl];
    if (!pageId) {
      Logger.log(`ページ未登録: ${pageUrl}`);
      skipCount++;
      continue;
    }
    
    const existingKey = `${pageId}_${keyword}`;
    
    if (existingKW[existingKey]) {
      // 更新
      toUpdate.push({
        id: existingKW[existingKey].id,
        gyron_position: gyronPosition,
        search_volume: searchVolume,
        gyron_updated_at: now
      });
    } else {
      // 新規
      toInsert.push({
        page_id: pageId,
        keyword: keyword,
        gyron_position: gyronPosition,
        search_volume: searchVolume,
        status: status || 'active',
        gyron_updated_at: now,
        created_at: now
      });
    }
  }
  
  Logger.log(`処理予定: 新規=${toInsert.length}件, 更新=${toUpdate.length}件, スキップ=${skipCount}件`);
  
  // 新規挿入（バッチ）
  if (toInsert.length > 0) {
    try {
      insertCount = batchInsertTargetKeywords(serviceRoleKey, toInsert);
      Logger.log(`✅ 新規挿入: ${insertCount}件`);
    } catch (e) {
      Logger.log(`❌ 挿入エラー: ${e.message}`);
      errors.push(`挿入: ${e.message}`);
    }
  }
  
  // 更新（個別）
  for (const item of toUpdate) {
    try {
      updateTargetKeywordById(serviceRoleKey, item.id, {
        gyron_position: item.gyron_position,
        search_volume: item.search_volume,
        gyron_updated_at: item.gyron_updated_at
      });
      updateCount++;
    } catch (e) {
      errors.push(`更新(${item.id}): ${e.message}`);
    }
  }
  
  Logger.log('=== 移行完了 ===');
  Logger.log(`結果: 新規=${insertCount}件, 更新=${updateCount}件, スキップ=${skipCount}件`);
  
  return {
    success: errors.length === 0,
    inserted: insertCount,
    updated: updateCount,
    skipped: skipCount,
    errors: errors
  };
}


/**
 * ページマッピング取得
 */
function getPageMappingForMigration(serviceRoleKey) {
  const SUPABASE_URL = 'https://dgzfdugpineqnoihopsl.supabase.co';
  const SITE_ID = '853ea711-7644-451e-872b-dea1b54fa8c7';
  
  const url = `${SUPABASE_URL}/rest/v1/pages?site_id=eq.${SITE_ID}&select=id,path`;
  
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
    mapping[page.path] = page.id;
  });
  
  return mapping;
}


/**
 * 既存target_keywordsをマップで取得
 */
function getExistingTargetKeywordsMap(serviceRoleKey) {
  const SUPABASE_URL = 'https://dgzfdugpineqnoihopsl.supabase.co';
  
  const url = `${SUPABASE_URL}/rest/v1/target_keywords?select=id,page_id,keyword`;
  
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
    throw new Error(`target_keywords取得エラー: ${response.getContentText()}`);
  }
  
  const keywords = JSON.parse(response.getContentText());
  const mapping = {};
  
  keywords.forEach(kw => {
    const key = `${kw.page_id}_${kw.keyword}`;
    mapping[key] = kw;
  });
  
  return mapping;
}


/**
 * target_keywordsをバッチ挿入
 */
function batchInsertTargetKeywords(serviceRoleKey, records) {
  const SUPABASE_URL = 'https://dgzfdugpineqnoihopsl.supabase.co';
  const url = `${SUPABASE_URL}/rest/v1/target_keywords`;
  
  // 50件ずつ分割して挿入
  const BATCH_SIZE = 50;
  let totalInserted = 0;
  
  for (let i = 0; i < records.length; i += BATCH_SIZE) {
    const batch = records.slice(i, i + BATCH_SIZE);
    
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
    
    if (response.getResponseCode() === 201 || response.getResponseCode() === 200) {
      totalInserted += batch.length;
    } else {
      throw new Error(`バッチ挿入エラー: ${response.getContentText()}`);
    }
    
    // API制限対策
    Utilities.sleep(100);
  }
  
  return totalInserted;
}


/**
 * target_keywordを個別更新
 */
function updateTargetKeywordById(serviceRoleKey, keywordId, data) {
  const SUPABASE_URL = 'https://dgzfdugpineqnoihopsl.supabase.co';
  const url = `${SUPABASE_URL}/rest/v1/target_keywords?id=eq.${keywordId}`;
  
  const response = UrlFetchApp.fetch(url, {
    method: 'patch',
    headers: {
      'apikey': serviceRoleKey,
      'Authorization': `Bearer ${serviceRoleKey}`,
      'Content-Type': 'application/json',
      'Prefer': 'return=minimal'
    },
    payload: JSON.stringify(data),
    muteHttpExceptions: true
  });
  
  if (response.getResponseCode() !== 204 && response.getResponseCode() !== 200) {
    throw new Error(response.getContentText());
  }
}


/**
 * URLをパス形式に正規化
 */
function normalizeToPathForMigration(url) {
  if (!url) return '';
  
  let path = url.toString();
  
  if (path.includes('://')) {
    try {
      const urlObj = new URL(path);
      path = urlObj.pathname;
    } catch (e) {
      const match = path.match(/https?:\/\/[^\/]+(\/.*)?/);
      if (match && match[1]) {
        path = match[1];
      }
    }
  }
  
  if (!path.startsWith('/')) {
    path = '/' + path;
  }
  
  path = path.replace(/\/$/, '');
  
  return path;
}


/**
 * 手動実行用（UIダイアログ付き）
 */
function runTargetKeywordsMigration() {
  const ui = SpreadsheetApp.getUi();
  
  const result = ui.alert(
    'ターゲットKW全件移行',
    'ターゲットKW分析シートの全データをSupabaseに移行します。\n\n' +
    '・既存データは更新されます\n' +
    '・新規データは追加されます\n\n' +
    '実行しますか？',
    ui.ButtonSet.YES_NO
  );
  
  if (result !== ui.Button.YES) {
    ui.alert('キャンセルしました');
    return;
  }
  
  try {
    const execResult = migrateAllTargetKeywordsToSupabase();
    
    if (execResult.success) {
      ui.alert(
        '完了',
        `ターゲットKW移行が完了しました。\n\n` +
        `新規追加: ${execResult.inserted}件\n` +
        `更新: ${execResult.updated}件\n` +
        `スキップ: ${execResult.skipped}件`,
        ui.ButtonSet.OK
      );
    } else {
      ui.alert(
        '一部エラー',
        `移行完了（一部エラーあり）\n\n` +
        `新規追加: ${execResult.inserted}件\n` +
        `更新: ${execResult.updated}件\n\n` +
        `エラー:\n${execResult.errors.slice(0, 5).join('\n')}`,
        ui.ButtonSet.OK
      );
    }
  } catch (error) {
    Logger.log(`エラー: ${error.message}`);
    ui.alert('エラー', error.toString(), ui.ButtonSet.OK);
  }
}