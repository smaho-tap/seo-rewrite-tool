/**
 * ============================================================================
 * Day 11-12: 最終デバッグ - 手動DA書き込みテスト
 * ============================================================================
 * DA履歴シートから直接読み取って、競合分析シートに書き込む
 */

/**
 * DA履歴シートから全データを取得してマップ作成（末尾スラッシュ対応）
 */
function getDAMapFromHistory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('DA履歴');
  
  if (!sheet) {
    return {};
  }
  
  const data = sheet.getDataRange().getValues();
  const daMap = {};
  
  for (let i = 1; i < data.length; i++) {
    let domain = data[i][0];
    const da = data[i][1];
    const pa = data[i][2];
    
    // 末尾スラッシュを削除（重要！）
    if (domain && domain.endsWith('/')) {
      domain = domain.slice(0, -1);
    }
    
    daMap[domain] = {
      domain: domain,
      da: da,
      pa: pa
    };
  }
  
  Logger.log(`DA履歴から ${Object.keys(daMap).length} 件のドメインを取得`);
  
  return daMap;
}

/**
 * 競合分析シートに直接DA書き込み（簡略版）
 */
function directWriteDA() {
  Logger.log('=== 直接DA書き込みテスト ===');
  Logger.log('');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('競合分析');
  
  if (!sheet) {
    Logger.log('✗ 競合分析シートが見つかりません');
    return;
  }
  
  // DA履歴シートから全データ取得
  const daMap = getDAMapFromHistory();
  
  Logger.log(`DAマップ件数: ${Object.keys(daMap).length}`);
  Logger.log('');
  
  // 最初の3件を表示
  Logger.log('DAマップ（最初の3件）:');
  let count = 0;
  for (const domain in daMap) {
    if (count < 3) {
      Logger.log(`  ${domain}: DA ${daMap[domain].da}`);
      count++;
    }
  }
  Logger.log('');
  
  // 2行目のデータを取得
  const row = 2;
  const data = sheet.getRange(row, 1, 1, 27).getValues()[0];
  
  const keyword = data[1]; // B列
  Logger.log(`キーワード: ${keyword}`);
  Logger.log('');
  
  let writeCount = 0;
  
  // rank_1_url（G列 = 配列index 6）のドメインを抽出してDA取得
  const rank1Url = data[6];
  Logger.log(`rank_1_url: ${rank1Url}`);
  
  if (rank1Url) {
    const domain1 = extractDomain(rank1Url);
    Logger.log(`抽出ドメイン: ${domain1}`);
    Logger.log(`daMap[${domain1}]: ${daMap[domain1] ? 'あり' : 'なし'}`);
    
    if (daMap[domain1]) {
      const da = daMap[domain1].da;
      Logger.log(`✓ DA取得: ${da}`);
      Logger.log(`H列（8列目）に書き込み中...`);
      
      sheet.getRange(row, 8).setValue(da);
      writeCount++;
      
      Logger.log(`✓ 書き込み完了`);
    } else {
      Logger.log(`✗ DAが見つかりません`);
    }
  }
  
  Logger.log('');
  Logger.log(`=== 書き込み完了: ${writeCount} 件 ===`);
  Logger.log('競合分析シートのH列（2行目）を確認してください');
}

/**
 * 全行に対してDA書き込み（確実版）
 */
function writeAllDA() {
  Logger.log('=== 全行DA書き込み開始 ===');
  Logger.log('');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('競合分析');
  
  if (!sheet) {
    Logger.log('✗ 競合分析シートが見つかりません');
    return;
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('⚠ データがありません');
    return;
  }
  
  // DA履歴から全データ取得
  const daMap = getDAMapFromHistory();
  Logger.log(`DAマップ: ${Object.keys(daMap).length} 件`);
  Logger.log('');
  
  // 全行処理
  const data = sheet.getRange(2, 1, lastRow - 1, 27).getValues();
  
  let totalWriteCount = 0;
  
  data.forEach((row, index) => {
    const rowNum = index + 2;
    const keyword = row[1]; // B列
    
    Logger.log(`[行${rowNum}] ${keyword}`);
    
    let rowWriteCount = 0;
    
    // rank_1_url ～ rank_10_url を処理
    for (let i = 0; i < 10; i++) {
      const urlIndex = 6 + (i * 2); // 6, 8, 10, 12, 14, 16, 18, 20, 22, 24
      const daCol = 8 + (i * 2);    // 8, 10, 12, 14, 16, 18, 20, 22, 24, 26
      
      const url = row[urlIndex];
      
      if (url) {
        const domain = extractDomain(url);
        
        if (daMap[domain]) {
          const da = daMap[domain].da;
          sheet.getRange(rowNum, daCol).setValue(da);
          rowWriteCount++;
        }
      }
    }
    
    Logger.log(`  書き込み: ${rowWriteCount} 件`);
    totalWriteCount += rowWriteCount;
  });
  
  Logger.log('');
  Logger.log(`=== 全行DA書き込み完了 ===`);
  Logger.log(`合計書き込み数: ${totalWriteCount} 件`);
}