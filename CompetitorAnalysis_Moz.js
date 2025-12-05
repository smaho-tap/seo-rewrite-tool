/**
 * ============================================================================
 * Day 11-12: Moz API DA取得（完全版・末尾スラッシュ対応）
 * ============================================================================
 * 日付比較と末尾スラッシュの問題を修正した完全版
 */

/**
 * URLからドメイン名を抽出
 */
function extractDomain(url) {
  if (!url) return '';
  
  try {
    let domain = url.replace(/^https?:\/\//, '');
    domain = domain.split('/')[0];
    domain = domain.split('?')[0];
    domain = domain.replace(/^www\./, '');
    
    return domain;
  } catch (error) {
    Logger.log(`✗ ドメイン抽出エラー: ${url} - ${error.message}`);
    return '';
  }
}

/**
 * DA履歴シートからキャッシュを取得（修正版・末尾スラッシュ対応）
 */
function getDAFromCache(domain) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('DA履歴');
  
  if (!sheet) {
    return null;
  }
  
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    let cachedDomain = row[0];
    const da = row[1];
    const pa = row[2];
    const lastUpdated = row[3];
    const cacheUntil = row[4];
    
    // 末尾スラッシュを削除（重要！）
    if (cachedDomain && cachedDomain.endsWith('/')) {
      cachedDomain = cachedDomain.slice(0, -1);
    }
    
    if (cachedDomain === domain) {
      // キャッシュが有効か確認（修正版）
      const now = new Date();
      
      // cacheUntilがDate型かどうか確認
      let expiryDate;
      if (cacheUntil instanceof Date) {
        expiryDate = cacheUntil;
      } else if (typeof cacheUntil === 'string') {
        expiryDate = new Date(cacheUntil);
      } else {
        // 日付型でもない場合は、とりあえず有効とする
        return {
          domain: domain,
          da: da,
          pa: pa,
          last_updated: lastUpdated,
          cache_until: cacheUntil,
          cached: true
        };
      }
      
      if (now < expiryDate) {
        // キャッシュ有効
        return {
          domain: domain,
          da: da,
          pa: pa,
          last_updated: lastUpdated,
          cache_until: cacheUntil,
          cached: true
        };
      } else {
        // キャッシュ期限切れ
        return null;
      }
    }
  }
  
  return null;
}

/**
 * DA履歴シートにキャッシュを保存
 */
function saveDAToCache(daResults) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('DA履歴');
  
  if (!sheet) {
    Logger.log('⚠ DA履歴シートが見つかりません');
    return;
  }
  
  const now = new Date();
  const cacheUntil = new Date(now.getTime() + 90 * 24 * 60 * 60 * 1000);
  
  const data = sheet.getDataRange().getValues();
  const existingDomains = {};
  
  for (let i = 1; i < data.length; i++) {
    let domain = data[i][0];
    // 末尾スラッシュを削除して比較
    if (domain && domain.endsWith('/')) {
      domain = domain.slice(0, -1);
    }
    existingDomains[domain] = i + 1;
  }
  
  const newRows = [];
  
  daResults.forEach(result => {
    const rowData = [
      result.domain,
      result.da || 0,
      result.pa || 0,
      now,
      cacheUntil
    ];
    
    if (existingDomains[result.domain]) {
      const rowIndex = existingDomains[result.domain];
      sheet.getRange(rowIndex, 1, 1, 5).setValues([rowData]);
    } else {
      newRows.push(rowData);
    }
  });
  
  if (newRows.length > 0) {
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, newRows.length, 5).setValues(newRows);
  }
  
  Logger.log(`✓ DA履歴シートに ${daResults.length} 件保存（90日間有効）`);
}

/**
 * Moz APIでドメインのDA/PAを取得
 */
function fetchDomainAuthority(domains) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const apiKey = scriptProperties.getProperty('MOZ_API_KEY');
  
  if (!apiKey) {
    throw new Error('Moz API認証情報が設定されていません');
  }
  
  const validDomains = domains.filter(d => d && d.trim() !== '');
  
  if (validDomains.length === 0) {
    return [];
  }
  
  const domainsToFetch = validDomains.slice(0, 50);
  
  const url = 'https://lsapi.seomoz.com/v2/url_metrics';
  
  const requestBody = {
    "targets": domainsToFetch
  };
  
  const options = {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${apiKey}`,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(requestBody),
    muteHttpExceptions: true
  };
  
  const maxRetries = 3;
  let lastError = null;
  
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      Logger.log(`Moz API呼び出し中... (試行 ${attempt}/${maxRetries}、${domainsToFetch.length}件)`);
      
      const response = UrlFetchApp.fetch(url, options);
      const responseCode = response.getResponseCode();
      
      if (responseCode === 200) {
        const data = JSON.parse(response.getContentText());
        const results = data.results || [];
        
        Logger.log(`✓ Moz API取得成功: ${results.length} 件`);
        
        return results.map(result => ({
          domain: result.page || result.subdomain || '',
          da: result.domain_authority || 0,
          pa: result.page_authority || 0
        }));
      } else if (responseCode === 429) {
        const waitTime = Math.pow(2, attempt) * 1000;
        Logger.log(`⚠ レート制限（429）、${waitTime/1000}秒待機...`);
        Utilities.sleep(waitTime);
        continue;
      } else {
        throw new Error(`HTTP ${responseCode}: ${response.getContentText()}`);
      }
    } catch (error) {
      lastError = error;
      Logger.log(`✗ 試行 ${attempt}/${maxRetries} 失敗: ${error.message}`);
      
      if (attempt < maxRetries) {
        const waitTime = Math.pow(2, attempt) * 1000;
        Logger.log(`${waitTime/1000}秒待機後に再試行...`);
        Utilities.sleep(waitTime);
      }
    }
  }
  
  throw new Error(`Moz API取得失敗（${maxRetries}回試行）: ${lastError.message}`);
}

/**
 * スマートキャッシング方式でDAを取得
 */
function fetchDAWithSmartCaching(urls) {
  const domains = urls.map(url => extractDomain(url)).filter(d => d);
  
  if (domains.length === 0) {
    return {};
  }
  
  Logger.log(`=== スマートキャッシング開始（${domains.length}件） ===`);
  
  const cachedDAs = {};
  const domainsToFetch = [];
  
  domains.forEach(domain => {
    const cached = getDAFromCache(domain);
    
    if (cached) {
      cachedDAs[domain] = cached;
      Logger.log(`✓ [${domain}] キャッシュヒット（DA: ${cached.da}）`);
    } else {
      domainsToFetch.push(domain);
      Logger.log(`⏳ [${domain}] キャッシュなし、API取得予定`);
    }
  });
  
  Logger.log('');
  Logger.log(`キャッシュヒット: ${Object.keys(cachedDAs).length} / ${domains.length}`);
  Logger.log(`API取得必要: ${domainsToFetch.length} / ${domains.length}`);
  
  if (domainsToFetch.length > 0) {
    Logger.log('');
    Logger.log('Moz APIで新規ドメインのDAを取得中...');
    
    try {
      const newDAs = fetchDomainAuthority(domainsToFetch);
      
      saveDAToCache(newDAs);
      
      newDAs.forEach(result => {
        cachedDAs[result.domain] = result;
      });
      
      Logger.log(`✓ ${newDAs.length} 件の新規DA取得完了`);
    } catch (error) {
      Logger.log(`✗ Moz API取得エラー: ${error.message}`);
    }
  }
  
  Logger.log('');
  Logger.log('=== スマートキャッシング完了 ===');
  Logger.log(`合計: ${Object.keys(cachedDAs).length} 件のDA取得`);
  
  return cachedDAs;
}

/**
 * 競合分析シートのDA列を更新（修正版）
 */
function updateCompetitorDA(startRow = 2, endRow = null) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('競合分析');
  
  if (!sheet) {
    throw new Error('競合分析シートが見つかりません');
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('⚠ データがありません');
    return;
  }
  
  const actualEndRow = endRow || lastRow;
  const rowCount = actualEndRow - startRow + 1;
  
  Logger.log(`=== 競合分析シートのDA更新（${rowCount}行） ===`);
  Logger.log('');
  
  const data = sheet.getRange(startRow, 1, rowCount, 27).getValues();
  
  const allUrls = [];
  
  data.forEach(row => {
    const pageUrl = row[2]; // C列
    if (pageUrl) {
      allUrls.push(pageUrl);
    }
    
    // G, I, K, M, O, Q, S, U, W, Y列（配列index: 6, 8, 10, 12, 14, 16, 18, 20, 22, 24）
    for (let i = 6; i <= 24; i += 2) {
      const url = row[i];
      if (url) {
        allUrls.push(url);
      }
    }
  });
  
  Logger.log(`全URL数: ${allUrls.length}`);
  Logger.log('');
  
  const daMap = fetchDAWithSmartCaching(allUrls);
  
  Logger.log('');
  Logger.log('競合分析シートを更新中...');
  
  let updateCount = 0;
  
  data.forEach((row, index) => {
    const rowIndex = startRow + index;
    const keyword = row[1]; // B列
    
    // 自社サイトDA（E列 = 5列目）
    const pageUrl = row[2]; // C列
    if (pageUrl) {
      const pageDomain = extractDomain(pageUrl);
      const pageDA = daMap[pageDomain];
      
      if (pageDA) {
        sheet.getRange(rowIndex, 5).setValue(pageDA.da);
        updateCount++;
      }
    }
    
    // 競合サイトDA
    for (let i = 0; i < 10; i++) {
      const urlArrayIndex = 6 + (i * 2); // 6, 8, 10, 12...
      const daColNumber = 8 + (i * 2);    // 8, 10, 12, 14...（H, J, L, N列...）
      
      const url = row[urlArrayIndex];
      
      if (url) {
        const domain = extractDomain(url);
        const daData = daMap[domain];
        
        if (daData) {
          sheet.getRange(rowIndex, daColNumber).setValue(daData.da);
          updateCount++;
        }
      }
    }
    
    Logger.log(`✓ [${keyword}] DA更新完了`);
  });
  
  Logger.log('');
  Logger.log('=== DA更新完了 ===');
  Logger.log(`更新数: ${updateCount} 件`);
}

/**
 * テスト実行: 競合分析シートの最初の5行のDAを更新
 */
function testMozDAUpdate() {
  Logger.log('=== Moz API DA更新テスト ===');
  Logger.log('');
  
  updateCompetitorDA(2, 6);
  
  Logger.log('');
  Logger.log('=== テスト完了 ===');
  Logger.log('競合分析シートのDA列（E列、H列、J列...）を確認してください');
}

/**
 * 全行のDA更新
 */
function updateAllCompetitorDA() {
  Logger.log('=== 全行DA更新開始 ===');
  Logger.log('');
  
  updateCompetitorDA();
  
  Logger.log('');
  Logger.log('=== 全行DA更新完了 ===');
}

/**
 * DA履歴シートの統計情報を表示
 */
function showDAHistoryStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('DA履歴');
  
  if (!sheet) {
    Logger.log('✗ DA履歴シートが見つかりません');
    return;
  }
  
  const data = sheet.getDataRange().getValues();
  const now = new Date();
  
  let validCount = 0;
  let expiredCount = 0;
  
  for (let i = 1; i < data.length; i++) {
    const cacheUntil = data[i][4];
    const expiryDate = cacheUntil instanceof Date ? cacheUntil : new Date(cacheUntil);
    
    if (now < expiryDate) {
      validCount++;
    } else {
      expiredCount++;
    }
  }
  
  Logger.log('=== DA履歴シート統計 ===');
  Logger.log(`総ドメイン数: ${data.length - 1}`);
  Logger.log(`有効キャッシュ: ${validCount}`);
  Logger.log(`期限切れ: ${expiredCount}`);
}

/**
 * キャッシュをクリア（期限切れのみ）
 */
function clearExpiredDACache() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('DA履歴');
  
  if (!sheet) {
    Logger.log('✗ DA履歴シートが見つかりません');
    return;
  }
  
  const data = sheet.getDataRange().getValues();
  const now = new Date();
  
  const rowsToDelete = [];
  
  for (let i = data.length - 1; i >= 1; i--) {
    const cacheUntil = data[i][4];
    const expiryDate = cacheUntil instanceof Date ? cacheUntil : new Date(cacheUntil);
    
    if (now >= expiryDate) {
      rowsToDelete.push(i + 1);
    }
  }
  
  rowsToDelete.forEach(rowIndex => {
    sheet.deleteRow(rowIndex);
  });
  
  Logger.log(`✓ 期限切れキャッシュ ${rowsToDelete.length} 件を削除`);
}