/**
 * ============================================================================
 * AIOTracking.gs v1.0
 * ============================================================================
 * AI Overview（AIO）の順位追跡・履歴管理
 * 
 * 機能:
 * - 自社サイトのAIO引用順位を取得
 * - AIO順位履歴シートへの保存
 * - リライト前後のAIO順位比較
 * - AIO順位レポート生成
 * 
 * @version 1.0
 * @date 2025-12-02
 * @author SEOリライト支援ツール
 */

// =================================
// 定数定義
// =================================

/**
 * 自社サイトのドメイン（設定・マスタシートから取得、またはデフォルト値）
 */
function getOwnDomain() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const ownDomain = scriptProperties.getProperty('OWN_DOMAIN');
  
  if (ownDomain) {
    return ownDomain;
  }
  
  // 設定・マスタシートから取得を試みる
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = ss.getSheetByName('設定・マスタ');
    if (settingsSheet) {
      const data = settingsSheet.getDataRange().getValues();
      for (let i = 0; i < data.length; i++) {
        if (data[i][0] === 'OWN_DOMAIN' || data[i][0] === '自社ドメイン') {
          return data[i][1];
        }
      }
    }
  } catch (error) {
    Logger.log(`自社ドメイン取得エラー: ${error.message}`);
  }
  
  // デフォルト値（必要に応じて変更）
  return '';
}

/**
 * 自社ドメインを設定
 * @param {string} domain - 自社ドメイン（例: 'example.com'）
 */
function setOwnDomain(domain) {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('OWN_DOMAIN', domain);
  Logger.log(`✓ 自社ドメインを設定しました: ${domain}`);
}

// =================================
// AIO順位履歴シート管理
// =================================

/**
 * AIO順位履歴シートを作成（存在しない場合）
 * @return {Sheet} シートオブジェクト
 */
function createAIOHistorySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('AIO順位履歴');
  
  if (sheet) {
    Logger.log('AIO順位履歴シートは既に存在します');
    return sheet;
  }
  
  // シート作成
  sheet = ss.insertSheet('AIO順位履歴');
  
  // ヘッダー設定
  const headers = [
    'record_id',           // A: レコードID
    'keyword',             // B: ターゲットキーワード
    'check_date',          // C: 取得日時
    'has_aio',             // D: AIO表示有無
    'own_site_in_aio',     // E: 自社サイト引用有無
    'aio_reference_position', // F: 引用順位（1〜N）
    'referenced_url',      // G: 引用されたページURL
    'referenced_text',     // H: 引用されたテキスト
    'total_references',    // I: 総引用数
    'organic_rank',        // J: 通常検索での順位
    'aio_position_change', // K: 前回からの順位変動
    'notes'                // L: メモ
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // ヘッダースタイル
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  
  // 列幅調整
  sheet.setColumnWidth(1, 150);  // record_id
  sheet.setColumnWidth(2, 200);  // keyword
  sheet.setColumnWidth(3, 150);  // check_date
  sheet.setColumnWidth(4, 80);   // has_aio
  sheet.setColumnWidth(5, 100);  // own_site_in_aio
  sheet.setColumnWidth(6, 100);  // aio_reference_position
  sheet.setColumnWidth(7, 300);  // referenced_url
  sheet.setColumnWidth(8, 200);  // referenced_text
  sheet.setColumnWidth(9, 80);   // total_references
  sheet.setColumnWidth(10, 80);  // organic_rank
  sheet.setColumnWidth(11, 100); // aio_position_change
  sheet.setColumnWidth(12, 200); // notes
  
  // 行を固定
  sheet.setFrozenRows(1);
  
  Logger.log('✓ AIO順位履歴シートを作成しました');
  return sheet;
}

// =================================
// AIO順位検出・保存
// =================================

/**
 * 自社サイトのAIO引用位置を検出
 * 
 * @param {Object} aioData - fetchSearchResultsから取得したAIOデータ
 * @param {string} ownDomain - 自社ドメイン（省略時は設定から取得）
 * @return {Object} AIO引用位置情報
 */
function findOwnSiteInAIO(aioData, ownDomain = null) {
  if (!ownDomain) {
    ownDomain = getOwnDomain();
  }
  
  if (!ownDomain) {
    Logger.log('⚠ 自社ドメインが設定されていません');
    return {
      found: false,
      position: null,
      url: null,
      text: null,
      error: '自社ドメイン未設定'
    };
  }
  
  // AIOがない場合
  if (!aioData || !aioData.hasAIO) {
    return {
      found: false,
      position: null,
      url: null,
      text: null,
      hasAIO: false
    };
  }
  
  // 自社サイトを検索
  const ownDomainLower = ownDomain.toLowerCase();
  
  for (let i = 0; i < aioData.references.length; i++) {
    const ref = aioData.references[i];
    const refDomain = (ref.domain || '').toLowerCase();
    
    // ドメイン一致チェック（サブドメイン含む）
    if (refDomain === ownDomainLower || 
        refDomain.endsWith('.' + ownDomainLower)) {
      return {
        found: true,
        position: i + 1,  // 1-indexed
        url: ref.url,
        title: ref.title,
        text: ref.text || '',
        location: ref.location,  // 'inline' or 'footer'
        totalReferences: aioData.totalReferences,
        hasAIO: true
      };
    }
  }
  
  // 自社サイトが見つからなかった
  return {
    found: false,
    position: null,
    url: null,
    text: null,
    totalReferences: aioData.totalReferences,
    hasAIO: true
  };
}

/**
 * AIO順位データを履歴シートに保存
 * 
 * @param {string} keyword - キーワード
 * @param {Object} aioData - AIOデータ
 * @param {Object} ownSiteInfo - 自社サイト引用情報
 * @param {number} organicRank - 通常検索順位（省略可）
 * @return {boolean} 保存成功フラグ
 */
function saveAIOHistory(keyword, aioData, ownSiteInfo, organicRank = null) {
  try {
    // シートを取得または作成
    const sheet = createAIOHistorySheet();
    
    // レコードID生成
    const recordId = `aio_${keyword.replace(/\s+/g, '_')}_${Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmmss')}`;
    
    // 前回の順位を取得して変動を計算
    const previousPosition = getPreviousAIOPosition(keyword);
    let positionChange = '';
    
    if (ownSiteInfo.found && previousPosition !== null) {
      const change = previousPosition - ownSiteInfo.position;
      if (change > 0) {
        positionChange = `↑${change}`;
      } else if (change < 0) {
        positionChange = `↓${Math.abs(change)}`;
      } else {
        positionChange = '→';
      }
    } else if (ownSiteInfo.found && previousPosition === null) {
      positionChange = 'NEW';
    } else if (!ownSiteInfo.found && previousPosition !== null) {
      positionChange = 'OUT';
    }
    
    // データ行作成
    const rowData = [
      recordId,
      keyword,
      new Date(),
      aioData.hasAIO,
      ownSiteInfo.found,
      ownSiteInfo.found ? ownSiteInfo.position : '',
      ownSiteInfo.url || '',
      ownSiteInfo.text || '',
      aioData.totalReferences || 0,
      organicRank || '',
      positionChange,
      ''  // notes
    ];
    
    // 最終行に追加
    sheet.appendRow(rowData);
    
    Logger.log(`✓ [${keyword}] AIO順位履歴を保存: 引用${ownSiteInfo.found ? ownSiteInfo.position + '位' : 'なし'} (${positionChange})`);
    
    return true;
    
  } catch (error) {
    Logger.log(`AIO履歴保存エラー: ${error.message}`);
    return false;
  }
}

/**
 * 特定キーワードの前回AIO順位を取得
 * 
 * @param {string} keyword - キーワード
 * @return {number|null} 前回の順位（なければnull）
 */
function getPreviousAIOPosition(keyword) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('AIO順位履歴');
    
    if (!sheet || sheet.getLastRow() <= 1) {
      return null;
    }
    
    const data = sheet.getDataRange().getValues();
    
    // 最新のレコードから逆順で検索
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][1] === keyword && data[i][4] === true) {
        // own_site_in_aio が true で、aio_reference_position を返す
        return data[i][5];
      }
    }
    
    return null;
    
  } catch (error) {
    Logger.log(`前回AIO順位取得エラー: ${error.message}`);
    return null;
  }
}

// =================================
// バッチ処理・統合機能
// =================================

/**
 * 検索結果からAIO順位を抽出して保存（1キーワード）
 * 
 * @param {Object} searchResult - fetchSearchResultsの戻り値
 * @return {Object} 処理結果
 */
function processAIOForKeyword(searchResult) {
  if (!searchResult || searchResult.error) {
    return {
      keyword: searchResult ? searchResult.keyword : 'unknown',
      success: false,
      error: searchResult ? searchResult.error : 'データなし'
    };
  }
  
  const keyword = searchResult.keyword;
  const aioData = searchResult.aio;
  
  // 自社サイトの引用位置を検索
  const ownSiteInfo = findOwnSiteInAIO(aioData);
  
  // 通常検索での自社順位を取得
  const ownDomain = getOwnDomain();
  let organicRank = null;
  
  if (ownDomain) {
    const ownDomainLower = ownDomain.toLowerCase();
    for (let i = 0; i < searchResult.results.length; i++) {
      const resultDomain = (searchResult.results[i].domain || '').toLowerCase();
      if (resultDomain === ownDomainLower || resultDomain.endsWith('.' + ownDomainLower)) {
        organicRank = searchResult.results[i].rank;
        break;
      }
    }
  }
  
  // 履歴に保存
  const saved = saveAIOHistory(keyword, aioData, ownSiteInfo, organicRank);
  
  return {
    keyword: keyword,
    success: saved,
    hasAIO: aioData.hasAIO,
    ownSiteInAIO: ownSiteInfo.found,
    aioPosition: ownSiteInfo.position,
    organicRank: organicRank,
    totalReferences: aioData.totalReferences
  };
}

/**
 * 複数キーワードのAIO順位を一括処理
 * 
 * @param {Array} searchResults - fetchMultipleSearchResultsの戻り値
 * @return {Object} 処理結果サマリー
 */
function processAIOForMultipleKeywords(searchResults) {
  Logger.log(`=== ${searchResults.length} キーワードのAIO順位処理開始 ===`);
  
  const results = [];
  let aioCount = 0;
  let ownSiteInAIOCount = 0;
  
  searchResults.forEach((searchResult, index) => {
    const result = processAIOForKeyword(searchResult);
    results.push(result);
    
    if (result.hasAIO) aioCount++;
    if (result.ownSiteInAIO) ownSiteInAIOCount++;
    
    Logger.log(`[${index + 1}/${searchResults.length}] ${result.keyword}: AIO=${result.hasAIO ? 'あり' : 'なし'}, 自社引用=${result.ownSiteInAIO ? result.aioPosition + '位' : 'なし'}`);
  });
  
  Logger.log('');
  Logger.log('=== AIO順位処理サマリー ===');
  Logger.log(`処理完了: ${results.length} キーワード`);
  Logger.log(`AIOあり: ${aioCount} / ${results.length}`);
  Logger.log(`自社サイト引用あり: ${ownSiteInAIOCount} / ${aioCount}`);
  
  return {
    total: results.length,
    aioCount: aioCount,
    ownSiteInAIOCount: ownSiteInAIOCount,
    results: results
  };
}

// =================================
// 履歴取得・分析
// =================================

/**
 * 特定キーワードのAIO順位履歴を取得
 * 
 * @param {string} keyword - キーワード
 * @param {number} limit - 取得件数（デフォルト: 10）
 * @return {Array} 履歴データ配列
 */
function getAIOHistoryForKeyword(keyword, limit = 10) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('AIO順位履歴');
    
    if (!sheet || sheet.getLastRow() <= 1) {
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const history = [];
    
    // 最新から逆順で取得
    for (let i = data.length - 1; i >= 1 && history.length < limit; i--) {
      if (data[i][1] === keyword) {
        history.push({
          recordId: data[i][0],
          keyword: data[i][1],
          checkDate: data[i][2],
          hasAIO: data[i][3],
          ownSiteInAIO: data[i][4],
          aioPosition: data[i][5],
          referencedUrl: data[i][6],
          referencedText: data[i][7],
          totalReferences: data[i][8],
          organicRank: data[i][9],
          positionChange: data[i][10],
          notes: data[i][11]
        });
      }
    }
    
    return history;
    
  } catch (error) {
    Logger.log(`AIO履歴取得エラー: ${error.message}`);
    return [];
  }
}

/**
 * リライト前後のAIO順位を比較
 * 
 * @param {string} keyword - キーワード
 * @param {Date} rewriteDate - リライト日時
 * @return {Object} 比較結果
 */
function compareAIOBeforeAfter(keyword, rewriteDate) {
  const history = getAIOHistoryForKeyword(keyword, 50);
  
  if (history.length === 0) {
    return {
      keyword: keyword,
      error: 'AIO履歴データがありません'
    };
  }
  
  const rewriteTime = rewriteDate.getTime();
  
  // リライト前後で分類
  const before = history.filter(h => new Date(h.checkDate).getTime() < rewriteTime);
  const after = history.filter(h => new Date(h.checkDate).getTime() >= rewriteTime);
  
  // 直近のデータを取得
  const latestBefore = before.length > 0 ? before[0] : null;
  const latestAfter = after.length > 0 ? after[after.length - 1] : null;
  
  // 比較結果を生成
  const result = {
    keyword: keyword,
    rewriteDate: rewriteDate,
    before: {
      date: latestBefore ? latestBefore.checkDate : null,
      hasAIO: latestBefore ? latestBefore.hasAIO : null,
      ownSiteInAIO: latestBefore ? latestBefore.ownSiteInAIO : null,
      aioPosition: latestBefore ? latestBefore.aioPosition : null,
      organicRank: latestBefore ? latestBefore.organicRank : null
    },
    after: {
      date: latestAfter ? latestAfter.checkDate : null,
      hasAIO: latestAfter ? latestAfter.hasAIO : null,
      ownSiteInAIO: latestAfter ? latestAfter.ownSiteInAIO : null,
      aioPosition: latestAfter ? latestAfter.aioPosition : null,
      organicRank: latestAfter ? latestAfter.organicRank : null
    },
    improvement: {}
  };
  
  // 改善判定
  if (latestBefore && latestAfter) {
    // AIO引用の改善
    if (!latestBefore.ownSiteInAIO && latestAfter.ownSiteInAIO) {
      result.improvement.aioAppeared = true;
      result.improvement.message = `AIO引用獲得！（${latestAfter.aioPosition}位）`;
    } else if (latestBefore.ownSiteInAIO && latestAfter.ownSiteInAIO) {
      const posChange = latestBefore.aioPosition - latestAfter.aioPosition;
      if (posChange > 0) {
        result.improvement.aioImproved = true;
        result.improvement.aioChange = posChange;
        result.improvement.message = `AIO順位 ${posChange}位アップ！（${latestBefore.aioPosition}位→${latestAfter.aioPosition}位）`;
      } else if (posChange < 0) {
        result.improvement.aioDeclined = true;
        result.improvement.aioChange = posChange;
        result.improvement.message = `AIO順位 ${Math.abs(posChange)}位ダウン（${latestBefore.aioPosition}位→${latestAfter.aioPosition}位）`;
      } else {
        result.improvement.message = `AIO順位変化なし（${latestAfter.aioPosition}位）`;
      }
    } else if (latestBefore.ownSiteInAIO && !latestAfter.ownSiteInAIO) {
      result.improvement.aioLost = true;
      result.improvement.message = 'AIO引用から除外されました';
    }
  }
  
  return result;
}

// =================================
// レポート生成
// =================================

/**
 * AIO順位サマリーレポートを生成
 * 
 * @return {Object} サマリーレポート
 */
function generateAIOSummaryReport() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('AIO順位履歴');
    
    if (!sheet || sheet.getLastRow() <= 1) {
      return {
        error: 'AIO順位履歴データがありません'
      };
    }
    
    const data = sheet.getDataRange().getValues();
    
    // 最新日付を特定
    let latestDate = null;
    const latestRecords = {};
    
    for (let i = data.length - 1; i >= 1; i--) {
      const keyword = data[i][1];
      const checkDate = new Date(data[i][2]);
      
      if (!latestRecords[keyword]) {
        latestRecords[keyword] = {
          keyword: keyword,
          checkDate: checkDate,
          hasAIO: data[i][3],
          ownSiteInAIO: data[i][4],
          aioPosition: data[i][5],
          organicRank: data[i][9],
          positionChange: data[i][10]
        };
        
        if (!latestDate || checkDate > latestDate) {
          latestDate = checkDate;
        }
      }
    }
    
    const keywords = Object.values(latestRecords);
    
    // 集計
    const summary = {
      reportDate: new Date(),
      dataDate: latestDate,
      totalKeywords: keywords.length,
      aioDisplayed: keywords.filter(k => k.hasAIO).length,
      ownSiteInAIO: keywords.filter(k => k.ownSiteInAIO).length,
      aioTop3: keywords.filter(k => k.ownSiteInAIO && k.aioPosition <= 3).length,
      improved: keywords.filter(k => k.positionChange && k.positionChange.startsWith('↑')).length,
      declined: keywords.filter(k => k.positionChange && k.positionChange.startsWith('↓')).length,
      newAppearance: keywords.filter(k => k.positionChange === 'NEW').length,
      lostAppearance: keywords.filter(k => k.positionChange === 'OUT').length
    };
    
    // 詳細リスト
    summary.keywordsInAIO = keywords
      .filter(k => k.ownSiteInAIO)
      .sort((a, b) => a.aioPosition - b.aioPosition)
      .map(k => ({
        keyword: k.keyword,
        aioPosition: k.aioPosition,
        organicRank: k.organicRank,
        change: k.positionChange
      }));
    
    summary.keywordsWithAIONoAppearance = keywords
      .filter(k => k.hasAIO && !k.ownSiteInAIO)
      .map(k => ({
        keyword: k.keyword,
        organicRank: k.organicRank
      }));
    
    return summary;
    
  } catch (error) {
    Logger.log(`AIOサマリーレポート生成エラー: ${error.message}`);
    return { error: error.message };
  }
}

// =================================
// テスト関数
// =================================

/**
 * AIO順位追跡のテスト
 */
function testAIOTracking() {
  Logger.log('=== AIO順位追跡テスト ===');
  Logger.log('');
  
  // 自社ドメイン設定確認
  const ownDomain = getOwnDomain();
  Logger.log(`自社ドメイン: ${ownDomain || '未設定'}`);
  
  if (!ownDomain) {
    Logger.log('⚠ 自社ドメインを設定してください: setOwnDomain("example.com")');
    return;
  }
  
  // テストキーワードで検索
  const testKeywords = ['iphone 保険'];
  
  Logger.log(`テストキーワード: ${testKeywords.join(', ')}`);
  Logger.log('');
  
  // 検索結果取得（AIO対応版）
  const searchResults = fetchMultipleSearchResults(testKeywords);
  
  // AIO順位処理
  const aioResults = processAIOForMultipleKeywords(searchResults);
  
  Logger.log('');
  Logger.log('=== テスト結果 ===');
  aioResults.results.forEach(r => {
    Logger.log(`${r.keyword}:`);
    Logger.log(`  AIO表示: ${r.hasAIO ? 'あり' : 'なし'}`);
    Logger.log(`  自社引用: ${r.ownSiteInAIO ? r.aioPosition + '位' : 'なし'}`);
    Logger.log(`  通常順位: ${r.organicRank || '圏外'}`);
  });
  
  Logger.log('');
  Logger.log('AIO順位履歴シートを確認してください');
}

/**
 * AIOサマリーレポートのテスト
 */
function testAIOSummaryReport() {
  Logger.log('=== AIOサマリーレポートテスト ===');
  
  const report = generateAIOSummaryReport();
  
  if (report.error) {
    Logger.log(`エラー: ${report.error}`);
    return;
  }
  
  Logger.log('');
  Logger.log('【AIOサマリー】');
  Logger.log(`データ日時: ${report.dataDate}`);
  Logger.log(`対象キーワード数: ${report.totalKeywords}`);
  Logger.log(`AIO表示あり: ${report.aioDisplayed}`);
  Logger.log(`自社引用あり: ${report.ownSiteInAIO}`);
  Logger.log(`AIO Top3: ${report.aioTop3}`);
  Logger.log(`順位上昇: ${report.improved}`);
  Logger.log(`順位下落: ${report.declined}`);
  Logger.log(`新規掲載: ${report.newAppearance}`);
  Logger.log(`掲載終了: ${report.lostAppearance}`);
  
  if (report.keywordsInAIO.length > 0) {
    Logger.log('');
    Logger.log('【AIO掲載キーワード】');
    report.keywordsInAIO.forEach(k => {
      Logger.log(`  ${k.keyword}: AIO ${k.aioPosition}位 (通常${k.organicRank}位) ${k.change}`);
    });
  }
}

/**
 * 自社ドメイン設定テスト
 */
function testSetOwnDomain() {
  // 例: setOwnDomain('your-domain.com');
  Logger.log('自社ドメインを設定するには、以下のコードを実行してください:');
  Logger.log('setOwnDomain("your-domain.com")');
  Logger.log('');
  Logger.log(`現在の設定: ${getOwnDomain() || '未設定'}`);
}

function initializeOwnDomain() {
  setOwnDomain("smaho-tap.com");
  Logger.log("設定完了: " + getOwnDomain());
}

/**
 * iPad mini 安く買う でAIOテスト
 */
function testAIOiPadMini() {
  Logger.log('=== iPad mini 安く買う AIOテスト ===');
  Logger.log('自社ドメイン: ' + getOwnDomain());
  
  var testKeywords = ['iPad mini 安く買う'];
  
  // 検索結果取得（AIO対応）
  var searchResults = fetchMultipleSearchResults(testKeywords);
  
  // AIO順位処理
  var aioResults = processAIOForMultipleKeywords(searchResults);
  
  // 結果表示
  Logger.log('\n=== テスト結果 ===');
  for (var i = 0; i < aioResults.length; i++) {
    var result = aioResults[i];
    Logger.log(result.keyword + ':');
    Logger.log('  AIO表示: ' + (result.hasAIO ? 'あり' : 'なし'));
    Logger.log('  自社引用: ' + (result.ownSiteFound ? 'あり（' + result.ownSitePosition + '位）' : 'なし'));
    Logger.log('  引用URL: ' + (result.ownSiteUrl || '-'));
    Logger.log('  通常順位: ' + (result.organicRank || '圏外'));
  }
  
  Logger.log('\nAIO順位履歴シートを確認してください');
}

/**
 * セピア色キーワードでAIOテスト
 */
function testAIOSepia() {
  Logger.log('=== iphone 画面がセピア色になる AIOテスト ===');
  Logger.log('自社ドメイン: ' + getOwnDomain());
  
  var testKeywords = ['iphone 画面がセピア色になる'];
  
  var searchResults = fetchMultipleSearchResults(testKeywords);
  var aioResults = processAIOForMultipleKeywords(searchResults);
  
  Logger.log('\n=== テスト結果 ===');
  for (var i = 0; i < aioResults.length; i++) {
    var result = aioResults[i];
    Logger.log(result.keyword + ':');
    Logger.log('  AIO表示: ' + (result.hasAIO ? 'あり' : 'なし'));
    Logger.log('  自社引用: ' + (result.ownSiteFound ? 'あり（' + result.ownSitePosition + '位）' : 'なし'));
    Logger.log('  引用URL: ' + (result.ownSiteUrl || '-'));
    Logger.log('  通常順位: ' + (result.organicRank || '圏外'));
  }
}