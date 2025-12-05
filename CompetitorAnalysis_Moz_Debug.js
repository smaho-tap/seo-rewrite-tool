/**
 * ============================================================================
 * Day 11-12: Moz API DA取得（デバッグ版）
 * ============================================================================
 * 詳細なログを出力して問題を特定
 */

/**
 * 競合分析シートのDA列を更新（デバッグ版）
 */
function updateCompetitorDADebug() {
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
  
  Logger.log(`=== デバッグ開始 ===`);
  Logger.log(`最終行: ${lastRow}`);
  Logger.log('');
  
  // 2行目のデータを取得してデバッグ
  const testRow = sheet.getRange(2, 1, 1, 27).getValues()[0]; // 27列取得
  
  Logger.log('=== 2行目のデータ ===');
  Logger.log(`A列（analysis_id）: ${testRow[0]}`);
  Logger.log(`B列（target_keyword）: ${testRow[1]}`);
  Logger.log(`C列（page_url）: ${testRow[2]}`);
  Logger.log(`D列（analysis_date）: ${testRow[3]}`);
  Logger.log(`E列（own_site_da）: ${testRow[4]}`);
  Logger.log(`F列（own_site_current_rank）: ${testRow[5]}`);
  Logger.log(`G列（rank_1_url）: ${testRow[6]}`);
  Logger.log(`H列（rank_1_da）: ${testRow[7]}`);
  Logger.log('');
  
  // URLを収集
  const allUrls = [];
  
  // 自社サイトURL（C列 = index 2）
  const pageUrl = testRow[2];
  Logger.log(`自社サイトURL（C列）: "${pageUrl}"`);
  
  if (pageUrl) {
    allUrls.push(pageUrl);
    Logger.log(`✓ 自社サイトURLを追加`);
  } else {
    Logger.log(`⚠ 自社サイトURLが空欄`);
  }
  
  // 競合サイトURL（G, I, K, M, O, Q, S, U, W, Y列）
  Logger.log('');
  Logger.log('=== 競合サイトURL ===');
  
  for (let i = 0; i < 10; i++) {
    const urlIndex = 6 + (i * 2); // 6=G列, 8=I列, 10=K列...
    const url = testRow[urlIndex];
    
    Logger.log(`rank_${i+1}_url（配列index ${urlIndex}）: "${url}"`);
    
    if (url) {
      allUrls.push(url);
    }
  }
  
  Logger.log('');
  Logger.log(`全URL数: ${allUrls.length}`);
  Logger.log('');
  
  // ドメイン抽出テスト
  Logger.log('=== ドメイン抽出テスト ===');
  
  allUrls.forEach((url, index) => {
    const domain = extractDomain(url);
    Logger.log(`[${index}] ${url} → ${domain}`);
  });
  
  Logger.log('');
  
  // DA履歴シートからデータ取得
  Logger.log('=== DA履歴シートからデータ取得 ===');
  
  const daSheet = ss.getSheetByName('DA履歴');
  if (!daSheet) {
    Logger.log('✗ DA履歴シートが見つかりません');
    return;
  }
  
  const daData = daSheet.getDataRange().getValues();
  Logger.log(`DA履歴シート行数: ${daData.length - 1}`);
  Logger.log('');
  
  // 最初の5件を表示
  Logger.log('=== DA履歴シート（最初の5件） ===');
  for (let i = 1; i < Math.min(6, daData.length); i++) {
    const row = daData[i];
    Logger.log(`${row[0]} → DA: ${row[1]}, PA: ${row[2]}`);
  }
  
  Logger.log('');
  
  // マッチングテスト
  Logger.log('=== マッチングテスト ===');
  
  allUrls.forEach(url => {
    const domain = extractDomain(url);
    const cached = getDAFromCache(domain);
    
    if (cached) {
      Logger.log(`✓ [${domain}] マッチ → DA: ${cached.da}`);
    } else {
      Logger.log(`✗ [${domain}] マッチなし`);
    }
  });
  
  Logger.log('');
  Logger.log('=== デバッグ完了 ===');
}

/**
 * シート構造を確認
 */
function checkSheetStructure() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('競合分析');
  
  if (!sheet) {
    Logger.log('✗ 競合分析シートが見つかりません');
    return;
  }
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  Logger.log('=== 競合分析シート構造 ===');
  Logger.log(`列数: ${headers.length}`);
  Logger.log('');
  Logger.log('列名:');
  
  headers.forEach((header, index) => {
    const colLetter = String.fromCharCode(65 + index); // A, B, C...
    Logger.log(`${colLetter}列（index ${index}）: ${header}`);
  });
}

/**
 * DA書き込みテスト（1行のみ）
 */
function testDAWriteToSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('競合分析');
  
  if (!sheet) {
    throw new Error('競合分析シートが見つかりません');
  }
  
  Logger.log('=== DA書き込みテスト ===');
  Logger.log('');
  
  // 2行目のデータを取得
  const testRow = sheet.getRange(2, 1, 1, 27).getValues()[0];
  const keyword = testRow[1];
  
  Logger.log(`キーワード: ${keyword}`);
  Logger.log('');
  
  // URLを収集
  const urls = [];
  
  // 自社サイトURL
  const pageUrl = testRow[2];
  if (pageUrl) {
    urls.push({ type: 'own_site', url: pageUrl, col: 5 }); // E列
  }
  
  // 競合サイトURL
  for (let i = 0; i < 10; i++) {
    const urlIndex = 6 + (i * 2);
    const daColNumber = 8 + (i * 2); // H, J, L, N, P, R, T, V, X, Z列
    const url = testRow[urlIndex];
    
    if (url) {
      urls.push({ 
        type: `rank_${i+1}`, 
        url: url, 
        col: daColNumber 
      });
    }
  }
  
  Logger.log(`URL数: ${urls.length}`);
  Logger.log('');
  
  // DAを取得して書き込み
  let writeCount = 0;
  
  urls.forEach(item => {
    const domain = extractDomain(item.url);
    const cached = getDAFromCache(domain);
    
    if (cached) {
      Logger.log(`✓ [${item.type}] ${domain} → DA: ${cached.da} を ${String.fromCharCode(64 + item.col)}列に書き込み`);
      
      // 実際に書き込み
      sheet.getRange(2, item.col).setValue(cached.da);
      writeCount++;
    } else {
      Logger.log(`✗ [${item.type}] ${domain} → DAが見つかりません`);
    }
  });
  
  Logger.log('');
  Logger.log(`=== 書き込み完了: ${writeCount} 件 ===`);
}