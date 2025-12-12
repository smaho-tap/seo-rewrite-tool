/**
 * DAHistoryToSupabase.gs
 * DA履歴データをSupabaseに移行するスクリプト
 * 
 * 【使い方】
 * migrateDAHistoryToSupabase() を実行
 */

const DA_CONFIG = {
  SUPABASE_URL: 'https://dgzfdugpineqnoihopsl.supabase.co',
  BATCH_SIZE: 100
};

/**
 * DA履歴をSupabaseに移行
 */
function migrateDAHistoryToSupabase() {
  Logger.log('=== DA履歴 → Supabase 移行開始 ===');
  
  const serviceRoleKey = PropertiesService.getScriptProperties()
    .getProperty('SUPABASE_SERVICE_ROLE_KEY');
  
  if (!serviceRoleKey) {
    Logger.log('❌ Service Role Keyが設定されていません');
    return;
  }
  
  try {
    // 1. DA履歴シートからデータ取得
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('DA履歴');
    
    if (!sheet) {
      throw new Error('DA履歴シートが見つかりません');
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const rows = data.slice(1);
    
    Logger.log(`DA履歴シート: ${rows.length}件`);
    
    // 2. 既存データ削除
    Logger.log('既存データ削除中...');
    deleteExistingDAHistory(serviceRoleKey);
    
    // 3. Supabase形式に変換
    const records = rows.map(row => {
      // domain, da, pa, last_updated, cache_until
      const domain = String(row[0]).trim();
      const da = parseInt(row[1]) || 0;
      const lastUpdated = row[3];
      
      // 日付をISO形式に変換
      let fetchedAt;
      if (lastUpdated instanceof Date) {
        fetchedAt = lastUpdated.toISOString();
      } else if (typeof lastUpdated === 'string' && lastUpdated) {
        // "2025-11-27 14:09:16" 形式をパース
        fetchedAt = new Date(lastUpdated.replace(' ', 'T') + '+09:00').toISOString();
      } else {
        fetchedAt = new Date().toISOString();
      }
      
      return {
        domain: domain,
        da: da,
        source: 'moz',
        fetched_at: fetchedAt
      };
    }).filter(r => r.domain && r.da > 0);  // 空のドメインやDA=0は除外
    
    Logger.log(`有効レコード: ${records.length}件`);
    
    // 4. バッチでSupabaseに保存
    let totalInserted = 0;
    
    for (let i = 0; i < records.length; i += DA_CONFIG.BATCH_SIZE) {
      const batch = records.slice(i, i + DA_CONFIG.BATCH_SIZE);
      
      const url = `${DA_CONFIG.SUPABASE_URL}/rest/v1/da_history`;
      
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
        Logger.log(`バッチ ${Math.floor(i / DA_CONFIG.BATCH_SIZE) + 1}: ${batch.length}件保存`);
      } else {
        Logger.log(`バッチエラー（${responseCode}）: ${response.getContentText()}`);
      }
      
      // API制限対策
      Utilities.sleep(300);
    }
    
    Logger.log('');
    Logger.log('=== 移行完了 ===');
    Logger.log(`総レコード数: ${totalInserted}件`);
    
    return { success: true, totalRecords: totalInserted };
    
  } catch (error) {
    Logger.log(`❌ エラー: ${error.message}`);
    throw error;
  }
}

/**
 * 既存のDA履歴データを削除
 */
function deleteExistingDAHistory(serviceRoleKey) {
  const deleteUrl = `${DA_CONFIG.SUPABASE_URL}/rest/v1/da_history?fetched_at=gte.2024-01-01`;
  
  const response = UrlFetchApp.fetch(deleteUrl, {
    method: 'delete',
    headers: {
      'apikey': serviceRoleKey,
      'Authorization': `Bearer ${serviceRoleKey}`,
      'Content-Type': 'application/json',
      'Prefer': 'return=minimal'
    },
    muteHttpExceptions: true
  });
  
  const responseCode = response.getResponseCode();
  if (responseCode !== 200 && responseCode !== 204) {
    Logger.log(`警告: 削除エラー（${responseCode}）`);
  }
}

/**
 * 移行結果を確認
 */
function checkDAHistoryMigration() {
  Logger.log('=== DA履歴移行結果確認 ===');
  
  const serviceRoleKey = PropertiesService.getScriptProperties()
    .getProperty('SUPABASE_SERVICE_ROLE_KEY');
  
  if (!serviceRoleKey) {
    Logger.log('❌ Service Role Keyが設定されていません');
    return;
  }
  
  const url = `${DA_CONFIG.SUPABASE_URL}/rest/v1/da_history?select=domain,da&limit=10`;
  
  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: {
      'apikey': serviceRoleKey,
      'Authorization': `Bearer ${serviceRoleKey}`,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  });
  
  if (response.getResponseCode() === 200) {
    const data = JSON.parse(response.getContentText());
    Logger.log(`サンプルデータ（10件）:`);
    data.forEach(row => {
      Logger.log(`  ${row.domain}: DA ${row.da}`);
    });
  }
  
  // 総件数を別途取得
  const countUrl = `${DA_CONFIG.SUPABASE_URL}/rest/v1/da_history?select=id`;
  const countResponse = UrlFetchApp.fetch(countUrl, {
    method: 'get',
    headers: {
      'apikey': serviceRoleKey,
      'Authorization': `Bearer ${serviceRoleKey}`,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  });
  
  if (countResponse.getResponseCode() === 200) {
    const countData = JSON.parse(countResponse.getContentText());
    Logger.log(`\n総件数: ${countData.length}件`);
  }
}