/**
 * GyronCSVImport.gs
 * GyronSEOのCSVデータを取り込み、ターゲットKW分析シートを更新
 * 
 * 対応CSV:
 * 1. rank-Google_all_*.csv - 検索順位（日別）
 * 2. 検索ボリューム_*.csv - 月間検索ボリューム
 * 
 * 作成日: 2025/11/30
 */

/**
 * 両方のCSVを取り込んでターゲットKW分析シートを更新（メイン関数）
 */
function importGyronCSVsAndUpdateAll() {
  console.log('=== GyronSEO CSVインポート開始 ===');
  const startTime = new Date();
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. 検索順位CSVを取り込み
  console.log('\n--- 検索順位CSV取り込み ---');
  const rankData = importRankCSV(ss);
  
  // 2. 検索ボリュームCSVを取り込み
  console.log('\n--- 検索ボリュームCSV取り込み ---');
  const volumeData = importVolumeCSV(ss);
  
  // 3. データをマージ
  console.log('\n--- データマージ ---');
  const mergedData = mergeRankAndVolumeData(rankData, volumeData);
  
  // 4. ターゲットKW分析シートを上書き
  console.log('\n--- ターゲットKW分析シート更新 ---');
  updateTargetKWSheet(ss, mergedData);
  
  // 5. GyronSEO_RAWシートも更新
  console.log('\n--- GyronSEO_RAWシート更新 ---');
  updateGyronSEORawSheet(ss, mergedData);
  
  // 6. 統合データシートのtarget_keywordを再連携
  console.log('\n--- 統合データシート連携 ---');
  updateIntegratedSheetTargetKeyword(ss, mergedData);
  
  const endTime = new Date();
  const duration = (endTime - startTime) / 1000;
  
  console.log('\n=== GyronSEO CSVインポート完了 ===');
  console.log(`総KW数: ${mergedData.length}件`);
  console.log(`所要時間: ${duration.toFixed(1)}秒`);
  
  return {
    success: true,
    totalKeywords: mergedData.length,
    duration: duration
  };
}


/**
 * 検索順位CSVを取り込み
 */
function importRankCSV(ss) {
  const rawSheet = ss.getSheetByName('GyronSEO_RAW');
  
  if (!rawSheet) {
    console.log('GyronSEO_RAWシートが見つかりません。新規作成します。');
    ss.insertSheet('GyronSEO_RAW');
    return [];
  }
  
  // GyronSEO_RAWシートから検索順位データを読み込む
  // CSVをGyronSEO_RAWシートに直接貼り付けてある前提
  const data = rawSheet.getDataRange().getValues();
  
  if (data.length < 2) {
    console.log('GyronSEO_RAWシートにデータがありません');
    return [];
  }
  
  const headers = data[0];
  console.log('検索順位CSV列数:', headers.length);
  
  // 列インデックスを特定
  const kwIdx = 0; // キーワード
  const urlIdx = 1; // URL
  
  // 日付列を特定（最新の日付を取得）
  const dateColumns = [];
  for (let i = 7; i < headers.length; i++) {
    const header = headers[i];
    if (header && (header instanceof Date || /^\d{4}-\d{2}-\d{2}/.test(header.toString()))) {
      dateColumns.push(i);
    }
  }
  
  console.log(`日付列数: ${dateColumns.length}`);
  
  // 最新の順位データを取得（末尾から探す）
  const rankData = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const keyword = (row[kwIdx] || '').toString().trim();
    let url = (row[urlIdx] || '').toString().trim();
    
    if (!keyword) continue;
    
    // URLを正規化（パス形式に）
    url = normalizeUrlToPath(url);
    
    // 最新の順位を取得（末尾から非空の値を探す）
    let latestPosition = null;
    let positionDate = null;
    
    for (let j = dateColumns.length - 1; j >= 0; j--) {
      const colIdx = dateColumns[j];
      const val = row[colIdx];
      
      if (val !== '' && val !== null && val !== undefined) {
        if (val === '圏外' || val === '-' || val === '--') {
          latestPosition = 101;
        } else if (typeof val === 'number') {
          latestPosition = val;
        } else {
          const parsed = parseInt(val);
          if (!isNaN(parsed)) {
            latestPosition = parsed;
          }
        }
        
        if (latestPosition !== null) {
          positionDate = headers[colIdx];
          break;
        }
      }
    }
    
    // 7日前、30日前、90日前の順位も取得
    const position7dAgo = getPositionAtDaysAgo(row, headers, dateColumns, 7);
    const position30dAgo = getPositionAtDaysAgo(row, headers, dateColumns, 30);
    const position90dAgo = getPositionAtDaysAgo(row, headers, dateColumns, 90);
    
    // トレンド計算
    const trend = calculateTrend(latestPosition, position7dAgo);
    
    rankData.push({
      keyword: keyword,
      url: url,
      gyronPosition: latestPosition || 101,
      positionDate: positionDate,
      position7dAgo: position7dAgo,
      position30dAgo: position30dAgo,
      position90dAgo: position90dAgo,
      trend: trend
    });
  }
  
  console.log(`検索順位データ: ${rankData.length}件`);
  console.log(`URL付き: ${rankData.filter(d => d.url).length}件`);
  
  return rankData;
}


/**
 * 検索ボリュームCSVを取り込み
 */
function importVolumeCSV(ss) {
  // 検索ボリュームシートを探す（または作成）
  let volumeSheet = ss.getSheetByName('検索ボリューム_RAW');
  
  if (!volumeSheet) {
    // シートがない場合は、ユーザーにアップロードを促す
    console.log('検索ボリューム_RAWシートが見つかりません');
    console.log('検索ボリュームCSVをGyronSEO_RAWの隣にシートを作成して貼り付けてください');
    return new Map();
  }
  
  const data = volumeSheet.getDataRange().getValues();
  
  if (data.length < 2) {
    console.log('検索ボリューム_RAWシートにデータがありません');
    return new Map();
  }
  
  const headers = data[0];
  console.log('検索ボリュームCSV列数:', headers.length);
  
  // 列インデックス
  const kwIdx = 0; // ランクインキーワード
  const volumeIdx = 1; // 検索ボリューム
  const trendIdx = 2; // 年間トレンド
  const competitionIdx = 3; // 競合性
  
  const volumeMap = new Map();
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const keyword = (row[kwIdx] || '').toString().trim();
    
    if (!keyword) continue;
    
    const volume = parseNumber(row[volumeIdx]);
    const yearlyTrend = row[trendIdx] || '';
    const competition = parseNumber(row[competitionIdx]);
    
    volumeMap.set(keyword.toLowerCase(), {
      searchVolume: volume,
      yearlyTrend: yearlyTrend,
      competition: competition
    });
  }
  
  console.log(`検索ボリュームデータ: ${volumeMap.size}件`);
  
  return volumeMap;
}


/**
 * 検索順位と検索ボリュームデータをマージ
 */
function mergeRankAndVolumeData(rankData, volumeMap) {
  const mergedData = [];
  
  for (const rank of rankData) {
    const kwLower = rank.keyword.toLowerCase();
    const volume = volumeMap.get(kwLower) || {};
    
    mergedData.push({
      keyword: rank.keyword,
      url: rank.url,
      gyronPosition: rank.gyronPosition,
      position7dAgo: rank.position7dAgo,
      position30dAgo: rank.position30dAgo,
      position90dAgo: rank.position90dAgo,
      trend: rank.trend,
      searchVolume: volume.searchVolume || 0,
      yearlyTrend: volume.yearlyTrend || '',
      competition: volume.competition || 0
    });
  }
  
  console.log(`マージ後データ: ${mergedData.length}件`);
  
  // 検索ボリューム取得率
  const withVolume = mergedData.filter(d => d.searchVolume > 0).length;
  console.log(`検索ボリューム取得: ${withVolume}/${mergedData.length}件`);
  
  return mergedData;
}


/**
 * ターゲットKW分析シートを上書き更新
 */
function updateTargetKWSheet(ss, mergedData) {
  let sheet = ss.getSheetByName('ターゲットKW分析');
  
  if (!sheet) {
    sheet = ss.insertSheet('ターゲットKW分析');
  }
  
  // シートをクリア
  sheet.clear();
  
  // ヘッダー
  const headers = [
    'keyword_id',
    'page_url',
    'target_keyword',
    'gyron_position',
    'gsc_position',
    'gsc_clicks',
    'gsc_impressions',
    'gsc_ctr',
    'search_volume',
    'competition_level',
    'kw_score',
    'performance_score',
    'search_volume_score',
    'strategic_value_score',
    'removal_score',
    'status',
    'notes',
    'last_updated'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // データ行を作成
  const rows = [];
  const now = new Date();
  
  for (let i = 0; i < mergedData.length; i++) {
    const d = mergedData[i];
    
    // KWスコアを計算
    const kwScore = calculateKWScore(d);
    const performanceScore = calculatePerformanceScore(d);
    const volumeScore = calculateVolumeScore(d.searchVolume);
    const strategicScore = calculateStrategicScore(d);
    
    // ステータス判定
    const status = determineStatus(d.gyronPosition, kwScore);
    
    rows.push([
      `KW-${String(i + 1).padStart(4, '0')}`, // keyword_id
      d.url || '', // page_url
      d.keyword, // target_keyword
      d.gyronPosition, // gyron_position
      '', // gsc_position（GSCからの更新用）
      '', // gsc_clicks
      '', // gsc_impressions
      '', // gsc_ctr
      d.searchVolume, // search_volume
      getCompetitionLevel(d.competition), // competition_level
      kwScore, // kw_score
      performanceScore, // performance_score
      volumeScore, // search_volume_score
      strategicScore, // strategic_value_score
      0, // removal_score
      status, // status
      '', // notes
      now // last_updated
    ]);
  }
  
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
  
  console.log(`ターゲットKW分析シート更新: ${rows.length}件`);
}


/**
 * GyronSEO_RAWシートを整形して更新
 */
function updateGyronSEORawSheet(ss, mergedData) {
  // GyronSEO_RAWシートは元のCSVデータを保持するため、
  // ここでは追加の整形は行わない
  console.log('GyronSEO_RAWシートは元データを保持');
}


/**
 * 統合データシートのtarget_keywordを更新
 */
function updateIntegratedSheetTargetKeyword(ss, mergedData) {
  const integratedSheet = ss.getSheetByName('統合データ');
  
  if (!integratedSheet) {
    console.log('統合データシートが見つかりません');
    return;
  }
  
  // URLからKWへのマップを作成（検索ボリューム最大のKWを選択）
  const urlToKW = new Map();
  
  for (const d of mergedData) {
    if (!d.url) continue;
    
    const normalizedUrl = d.url.toLowerCase().replace(/\/$/, '');
    
    if (!urlToKW.has(normalizedUrl) || 
        (urlToKW.get(normalizedUrl).searchVolume < d.searchVolume)) {
      urlToKW.set(normalizedUrl, {
        keyword: d.keyword,
        searchVolume: d.searchVolume,
        gyronPosition: d.gyronPosition
      });
    }
  }
  
  console.log(`URLマップ構築: ${urlToKW.size}件`);
  
  // 統合データシートのヘッダーを取得
  const intHeaders = integratedSheet.getRange(1, 1, 1, integratedSheet.getLastColumn()).getValues()[0];
  
  // target_keyword列を探す
  let targetKwIdx = intHeaders.indexOf('target_keyword');
  
  if (targetKwIdx === -1) {
    // 列がなければ追加（C列に挿入）
    integratedSheet.insertColumnBefore(3);
    integratedSheet.getRange(1, 3).setValue('target_keyword');
    targetKwIdx = 2; // 0-indexed
    console.log('target_keyword列を追加しました（列C）');
  }
  
  // page_url列を探す
  const pageUrlIdx = intHeaders.indexOf('page_url');
  
  if (pageUrlIdx === -1) {
    console.log('page_url列が見つかりません');
    return;
  }
  
  // 統合データシートのデータを取得
  const intData = integratedSheet.getDataRange().getValues();
  
  let matchCount = 0;
  let unmatchCount = 0;
  let errorCount = 0;
  
  // ヘッダー更新後の列位置を再計算
  const newHeaders = integratedSheet.getRange(1, 1, 1, integratedSheet.getLastColumn()).getValues()[0];
  const newTargetKwIdx = newHeaders.indexOf('target_keyword');
  
  for (let i = 1; i < intData.length; i++) {
    let pageUrl = intData[i][pageUrlIdx];
    
    if (!pageUrl) continue;
    
    // URLを正規化
    pageUrl = normalizeUrlToPath(pageUrl).toLowerCase().replace(/\/$/, '');
    
    const kwData = urlToKW.get(pageUrl);
    
    if (kwData) {
      try {
        integratedSheet.getRange(i + 1, newTargetKwIdx + 1).setValue(kwData.keyword);
        matchCount++;
      } catch (e) {
        // 入力規則エラーなどをスキップ
        errorCount++;
        console.log(`行${i + 1}でエラー: ${e.message}`);
      }
    } else {
      unmatchCount++;
    }
  }
  
  console.log(`統合データシート連携: マッチ${matchCount}件, 未マッチ${unmatchCount}件, エラー${errorCount}件`);
}


// ===========================================
// ユーティリティ関数
// ===========================================

/**
 * URLをパス形式に正規化
 */
function normalizeUrlToPath(url) {
  if (!url) return '';
  
  let path = url.toString();
  
  // フルURLからパスを抽出
  if (path.includes('://')) {
    try {
      const urlObj = new URL(path);
      path = urlObj.pathname;
    } catch (e) {
      // パース失敗時は文字列操作
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


/**
 * 数値をパース
 */
function parseNumber(val) {
  if (typeof val === 'number') return val;
  if (!val) return 0;
  
  const str = val.toString().replace(/[,￥$%]/g, '');
  const num = parseFloat(str);
  return isNaN(num) ? 0 : num;
}


/**
 * N日前の順位を取得
 */
function getPositionAtDaysAgo(row, headers, dateColumns, daysAgo) {
  const targetDate = new Date();
  targetDate.setDate(targetDate.getDate() - daysAgo);
  
  // 最も近い日付の列を探す
  let closestIdx = -1;
  let closestDiff = Infinity;
  
  for (const colIdx of dateColumns) {
    const header = headers[colIdx];
    let headerDate;
    
    if (header instanceof Date) {
      headerDate = header;
    } else {
      headerDate = new Date(header);
    }
    
    if (isNaN(headerDate.getTime())) continue;
    
    const diff = Math.abs(headerDate.getTime() - targetDate.getTime());
    if (diff < closestDiff) {
      closestDiff = diff;
      closestIdx = colIdx;
    }
  }
  
  if (closestIdx === -1) return null;
  
  const val = row[closestIdx];
  
  if (val === '圏外' || val === '-' || val === '--') return 101;
  if (typeof val === 'number') return val;
  
  const parsed = parseInt(val);
  return isNaN(parsed) ? null : parsed;
}


/**
 * トレンドを計算
 */
function calculateTrend(latest, previous) {
  if (latest === null || previous === null) return '--';
  if (latest === 101 && previous === 101) return '--';
  if (latest < previous) return '↑';
  if (latest > previous) return '↓';
  return '→';
}


/**
 * KWスコアを計算
 */
function calculateKWScore(data) {
  let score = 0;
  
  // 順位スコア（40%）
  if (data.gyronPosition <= 3) score += 40;
  else if (data.gyronPosition <= 10) score += 30;
  else if (data.gyronPosition <= 20) score += 20;
  else if (data.gyronPosition <= 50) score += 10;
  else score += 0;
  
  // 検索ボリュームスコア（30%）
  if (data.searchVolume >= 1000) score += 30;
  else if (data.searchVolume >= 500) score += 20;
  else if (data.searchVolume >= 100) score += 15;
  else if (data.searchVolume >= 10) score += 10;
  else score += 5;
  
  // URL有無スコア（30%）
  if (data.url) score += 30;
  
  return Math.min(100, score);
}


/**
 * パフォーマンススコアを計算
 */
function calculatePerformanceScore(data) {
  let score = 50; // 基準点
  
  // 順位による調整
  if (data.gyronPosition <= 3) score += 30;
  else if (data.gyronPosition <= 10) score += 20;
  else if (data.gyronPosition <= 20) score += 10;
  else if (data.gyronPosition > 50) score -= 20;
  
  // トレンドによる調整
  if (data.trend === '↑') score += 10;
  else if (data.trend === '↓') score -= 10;
  
  return Math.max(0, Math.min(100, score));
}


/**
 * 検索ボリュームスコアを計算
 */
function calculateVolumeScore(volume) {
  if (volume >= 10000) return 100;
  if (volume >= 5000) return 90;
  if (volume >= 1000) return 80;
  if (volume >= 500) return 60;
  if (volume >= 100) return 40;
  if (volume >= 10) return 20;
  return 10;
}


/**
 * 戦略的価値スコアを計算
 */
function calculateStrategicScore(data) {
  let score = 50;
  
  // 検索ボリュームと順位の組み合わせ
  if (data.searchVolume >= 500 && data.gyronPosition <= 20) {
    score = 90; // 高ボリューム＋上位
  } else if (data.searchVolume >= 100 && data.gyronPosition <= 10) {
    score = 80; // 中ボリューム＋TOP10
  } else if (data.gyronPosition <= 5) {
    score = 70; // TOP5
  } else if (data.gyronPosition > 50 && data.searchVolume < 100) {
    score = 20; // 低ボリューム＋圏外近く
  }
  
  return score;
}


/**
 * ステータスを判定
 */
function determineStatus(position, kwScore) {
  if (position <= 10 && kwScore >= 60) return '維持';
  if (position <= 20) return '要改善';
  if (position <= 50) return '要強化';
  return '除外候補';
}


/**
 * 競合レベルを取得
 */
function getCompetitionLevel(competition) {
  if (competition >= 80) return '激戦';
  if (competition >= 50) return '難';
  if (competition >= 20) return '中';
  return '易';
}


// ===========================================
// 手動実行用関数
// ===========================================

/**
 * GyronSEO CSVインポート実行（メニューから呼び出し用）
 */
function runGyronCSVImport() {
  const ui = SpreadsheetApp.getUi();
  
  const result = ui.alert(
    'GyronSEO CSVインポート',
    '検索順位CSVと検索ボリュームCSVを取り込みます。\n\n' +
    '事前準備:\n' +
    '1. GyronSEO_RAWシートに検索順位CSVを貼り付け\n' +
    '2. 検索ボリューム_RAWシートを作成して検索ボリュームCSVを貼り付け\n\n' +
    '既存のターゲットKW分析シートは上書きされます。\n実行しますか？',
    ui.ButtonSet.YES_NO
  );
  
  if (result !== ui.Button.YES) {
    ui.alert('キャンセルしました');
    return;
  }
  
  try {
    const execResult = importGyronCSVsAndUpdateAll();
    
    if (execResult.success) {
      ui.alert(
        '完了',
        `GyronSEO CSVインポートが完了しました。\n\n` +
        `総KW数: ${execResult.totalKeywords}件\n` +
        `所要時間: ${execResult.duration.toFixed(1)}秒`,
        ui.ButtonSet.OK
      );
    }
  } catch (error) {
    console.error('エラー:', error);
    ui.alert('エラー', error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * ターゲットKW分析シートのみ更新（統合データシート連携なし）
 */
function importGyronCSVsOnly() {
  console.log('=== GyronSEO CSVインポート開始（ターゲットKWのみ）===');
  const startTime = new Date();
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. 検索順位CSVを取り込み
  console.log('\n--- 検索順位CSV取り込み ---');
  const rankData = importRankCSV(ss);
  
  // 2. 検索ボリュームCSVを取り込み
  console.log('\n--- 検索ボリュームCSV取り込み ---');
  const volumeData = importVolumeCSV(ss);
  
  // 3. データをマージ
  console.log('\n--- データマージ ---');
  const mergedData = mergeRankAndVolumeData(rankData, volumeData);
  
  // 4. ターゲットKW分析シートを上書き
  console.log('\n--- ターゲットKW分析シート更新 ---');
  updateTargetKWSheet(ss, mergedData);
  
  const endTime = new Date();
  const duration = (endTime - startTime) / 1000;
  
  console.log('\n=== 完了 ===');
  console.log(`総KW数: ${mergedData.length}件`);
  console.log(`所要時間: ${duration.toFixed(1)}秒`);
}