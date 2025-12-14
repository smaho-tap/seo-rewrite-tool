/**
 * TrendAnalysis.gs - 順位トレンド判定ロジック
 * 
 * 4週間の順位変動を分析し、スコアに加減算を適用
 * 
 * 【トレンド判定】
 * - 安定: 4週間の変動幅±5位以内 → ±0点
 * - 上昇傾向: 6位以上改善 → -20点（好調なので触らない）
 * - 下降傾向: 6位以上悪化 → +15点（要リライト）
 * - 不安定: 週ごとに±6位以上の乱高下 → -25点（様子見）
 * 
 * 作成日: 2025/12/14
 */

// ===========================================
// 定数定義
// ===========================================

const TREND_CONFIG = {
  // 分析期間（週）
  WEEKS_TO_ANALYZE: 4,
  
  // 安定判定の閾値
  STABLE_THRESHOLD: 5,  // ±5位以内
  
  // 上昇/下降判定の閾値
  SIGNIFICANT_CHANGE: 6,  // 6位以上の変動
  
  // 不安定判定の閾値（週間変動）
  UNSTABLE_THRESHOLD: 6,  // 週ごとに±6位以上
  
  // スコア加減算
  ADJUSTMENTS: {
    STABLE: 0,
    RISING: -20,      // 上昇傾向: 優先度下げ
    FALLING: 15,      // 下降傾向: 優先度上げ
    UNSTABLE: -25     // 不安定: 様子見
  },
  
  // 圏外の扱い
  OUT_OF_RANK_VALUE: 101
};


// ===========================================
// メイン関数: トレンド判定
// ===========================================

/**
 * 指定キーワードのトレンドを判定
 * @param {String} keyword - ターゲットキーワード
 * @return {Object} { trend: String, adjustment: Number, details: Object }
 */
function analyzeTrend(keyword) {
  if (!keyword) {
    return { trend: '不明', adjustment: 0, details: null };
  }
  
  // 過去4週間の順位データを取得
  const rankingHistory = getRankingHistory(keyword);
  
  if (!rankingHistory || rankingHistory.length < 7) {
    // データ不足
    return { trend: 'データ不足', adjustment: 0, details: { dataPoints: rankingHistory ? rankingHistory.length : 0 } };
  }
  
  // 週ごとの中央値を計算
  const weeklyMedians = calculateWeeklyMedians(rankingHistory);
  
  if (weeklyMedians.length < 2) {
    return { trend: 'データ不足', adjustment: 0, details: { weeks: weeklyMedians.length } };
  }
  
  // トレンド判定
  const trendResult = determineTrendFromMedians(weeklyMedians);
  
  return {
    trend: trendResult.trend,
    adjustment: TREND_CONFIG.ADJUSTMENTS[trendResult.trendType],
    details: {
      weeklyMedians: weeklyMedians,
      week1Median: weeklyMedians[0],
      week4Median: weeklyMedians[weeklyMedians.length - 1],
      change: trendResult.change,
      maxWeeklyChange: trendResult.maxWeeklyChange
    }
  };
}


/**
 * 全キーワードのトレンドをまとめて分析
 * @return {Map} キーワード → トレンド結果のマップ
 */
function analyzeAllTrends() {
  console.log('=== 全キーワードトレンド分析開始 ===');
  const startTime = new Date();
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gyronSheet = ss.getSheetByName('GyronSEO_RAW');
  
  if (!gyronSheet) {
    console.error('GyronSEO_RAWシートが見つかりません');
    return new Map();
  }
  
  const data = gyronSheet.getDataRange().getValues();
  const trendMap = new Map();
  
  // 各キーワードを分析（ヘッダー行をスキップ）
  for (let i = 1; i < data.length; i++) {
    const keyword = data[i][0];
    if (!keyword) continue;
    
    const trendResult = analyzeTrend(keyword);
    trendMap.set(keyword.toLowerCase().trim(), trendResult);
  }
  
  const endTime = new Date();
  console.log(`トレンド分析完了: ${trendMap.size}件, ${(endTime - startTime) / 1000}秒`);
  
  return trendMap;
}


// ===========================================
// データ取得関数
// ===========================================

/**
 * GyronSEO_RAWから指定キーワードの過去4週間の順位を取得
 * @param {String} keyword - ターゲットキーワード
 * @return {Array} [{ date: Date, rank: Number }, ...]
 */
function getRankingHistory(keyword) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gyronSheet = ss.getSheetByName('GyronSEO_RAW');
  
  if (!gyronSheet) {
    console.warn('GyronSEO_RAWシートが見つかりません');
    return [];
  }
  
  const data = gyronSheet.getDataRange().getValues();
  const headers = data[0];
  
  // キーワード列（0列目）から該当行を検索
  let targetRow = null;
  const keywordLower = keyword.toLowerCase().trim();
  
  for (let i = 1; i < data.length; i++) {
    const rowKeyword = (data[i][0] || '').toString().toLowerCase().trim();
    if (rowKeyword === keywordLower) {
      targetRow = data[i];
      break;
    }
  }
  
  if (!targetRow) {
    console.warn(`キーワード「${keyword}」が見つかりません`);
    return [];
  }
  
  // 日付列（7列目以降）から過去28日分を取得
  const today = new Date();
  const fourWeeksAgo = new Date(today.getTime() - 28 * 24 * 60 * 60 * 1000);
  
  const rankings = [];
  
  for (let col = 7; col < headers.length; col++) {
    const dateStr = headers[col];
    if (!dateStr) continue;
    
    // 日付をパース
    const date = parseDateFromHeader(dateStr);
    if (!date || date < fourWeeksAgo) continue;
    
    // 順位を取得
    const rankValue = targetRow[col];
    const rank = parseRank(rankValue);
    
    if (rank !== null) {
      rankings.push({ date: date, rank: rank });
    }
  }
  
  // 日付順にソート（古い順）
  rankings.sort((a, b) => a.date - b.date);
  
  return rankings;
}


/**
 * ヘッダーの日付文字列をDateオブジェクトに変換
 * @param {String} dateStr - "2025-11-30" 形式
 * @return {Date|null}
 */
function parseDateFromHeader(dateStr) {
  if (!dateStr) return null;
  
  const str = dateStr.toString().trim();
  
  // "2025-11-30" 形式
  const match = str.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (match) {
    return new Date(parseInt(match[1]), parseInt(match[2]) - 1, parseInt(match[3]));
  }
  
  // Date型の場合
  if (dateStr instanceof Date) {
    return dateStr;
  }
  
  return null;
}


/**
 * 順位値をパース
 * @param {*} value - 順位値（数字、"圏外"、空白など）
 * @return {Number|null}
 */
function parseRank(value) {
  if (value === null || value === undefined || value === '') {
    return null;  // データなし
  }
  
  const str = value.toString().trim();
  
  if (str === '圏外') {
    return TREND_CONFIG.OUT_OF_RANK_VALUE;
  }
  
  const num = parseInt(str, 10);
  if (!isNaN(num) && num > 0) {
    return num;
  }
  
  return null;
}


// ===========================================
// 計算関数
// ===========================================

/**
 * 週ごとの中央値を計算
 * @param {Array} rankings - [{ date: Date, rank: Number }, ...]
 * @return {Array} [week1Median, week2Median, week3Median, week4Median]
 */
function calculateWeeklyMedians(rankings) {
  if (!rankings || rankings.length === 0) return [];
  
  // 日付でソート（古い順）
  const sorted = [...rankings].sort((a, b) => a.date - b.date);
  
  // 最新日を基準に週を区切る
  const latestDate = sorted[sorted.length - 1].date;
  
  const weeks = [[], [], [], []];  // Week1(最古) 〜 Week4(最新)
  
  for (const item of sorted) {
    const daysDiff = Math.floor((latestDate - item.date) / (24 * 60 * 60 * 1000));
    
    if (daysDiff < 7) {
      weeks[3].push(item.rank);  // Week4（最新週）
    } else if (daysDiff < 14) {
      weeks[2].push(item.rank);  // Week3
    } else if (daysDiff < 21) {
      weeks[1].push(item.rank);  // Week2
    } else if (daysDiff < 28) {
      weeks[0].push(item.rank);  // Week1（最古週）
    }
  }
  
  // 各週の中央値を計算
  const medians = [];
  for (const weekRanks of weeks) {
    if (weekRanks.length > 0) {
      medians.push(calculateMedian(weekRanks));
    }
  }
  
  return medians;
}


/**
 * 中央値を計算
 * @param {Array} values - 数値配列
 * @return {Number}
 */
function calculateMedian(values) {
  if (!values || values.length === 0) return null;
  
  const sorted = [...values].sort((a, b) => a - b);
  const mid = Math.floor(sorted.length / 2);
  
  if (sorted.length % 2 === 0) {
    return (sorted[mid - 1] + sorted[mid]) / 2;
  } else {
    return sorted[mid];
  }
}


/**
 * 週ごとの中央値からトレンドを判定
 * @param {Array} weeklyMedians - [week1, week2, week3, week4]
 * @return {Object} { trend: String, trendType: String, change: Number, maxWeeklyChange: Number }
 */
function determineTrendFromMedians(weeklyMedians) {
  if (weeklyMedians.length < 2) {
    return { trend: 'データ不足', trendType: 'STABLE', change: 0, maxWeeklyChange: 0 };
  }
  
  const week1 = weeklyMedians[0];  // 最古
  const weekLatest = weeklyMedians[weeklyMedians.length - 1];  // 最新
  
  // 全体の変化量（最新 - 最古）
  // 順位なので、マイナス = 上昇（良い）、プラス = 下降（悪い）
  const totalChange = weekLatest - week1;
  
  // 週間の最大変動幅を計算（不安定判定用）
  let maxWeeklyChange = 0;
  for (let i = 1; i < weeklyMedians.length; i++) {
    const weeklyChange = Math.abs(weeklyMedians[i] - weeklyMedians[i - 1]);
    if (weeklyChange > maxWeeklyChange) {
      maxWeeklyChange = weeklyChange;
    }
  }
  
  // 不安定判定（週ごとに±6位以上の乱高下）
  if (maxWeeklyChange >= TREND_CONFIG.UNSTABLE_THRESHOLD) {
    // 最終的に変化が小さくても、途中で大きく乱高下していれば不安定
    const hasSignificantFluctuation = weeklyMedians.length >= 3 && 
      weeklyMedians.some((_, i) => {
        if (i === 0 || i === weeklyMedians.length - 1) return false;
        const prevChange = weeklyMedians[i] - weeklyMedians[i - 1];
        const nextChange = weeklyMedians[i + 1] - weeklyMedians[i];
        // 上昇→下降 または 下降→上昇 の反転があるか
        return (prevChange * nextChange < 0) && 
               (Math.abs(prevChange) >= TREND_CONFIG.UNSTABLE_THRESHOLD || 
                Math.abs(nextChange) >= TREND_CONFIG.UNSTABLE_THRESHOLD);
      });
    
    if (hasSignificantFluctuation) {
      return { 
        trend: '不安定', 
        trendType: 'UNSTABLE', 
        change: totalChange, 
        maxWeeklyChange: maxWeeklyChange 
      };
    }
  }
  
  // 上昇傾向（6位以上改善 = totalChange <= -6）
  if (totalChange <= -TREND_CONFIG.SIGNIFICANT_CHANGE) {
    return { 
      trend: '上昇傾向', 
      trendType: 'RISING', 
      change: totalChange, 
      maxWeeklyChange: maxWeeklyChange 
    };
  }
  
  // 下降傾向（6位以上悪化 = totalChange >= 6）
  if (totalChange >= TREND_CONFIG.SIGNIFICANT_CHANGE) {
    return { 
      trend: '下降傾向', 
      trendType: 'FALLING', 
      change: totalChange, 
      maxWeeklyChange: maxWeeklyChange 
    };
  }
  
  // 安定（±5位以内）
  return { 
    trend: '安定', 
    trendType: 'STABLE', 
    change: totalChange, 
    maxWeeklyChange: maxWeeklyChange 
  };
}


// ===========================================
// スコア調整関数
// ===========================================

/**
 * トレンドに基づいてスコアを調整
 * @param {Number} originalScore - 元のスコア
 * @param {String} keyword - ターゲットキーワード
 * @return {Object} { adjustedScore: Number, trend: String, adjustment: Number }
 */
function adjustScoreByTrend(originalScore, keyword) {
  const trendResult = analyzeTrend(keyword);
  
  let adjustedScore = originalScore + trendResult.adjustment;
  
  // 0-100の範囲に収める
  adjustedScore = Math.max(0, Math.min(100, adjustedScore));
  
  return {
    adjustedScore: adjustedScore,
    trend: trendResult.trend,
    adjustment: trendResult.adjustment,
    details: trendResult.details
  };
}


/**
 * トレンドマップを使って高速にスコア調整（バッチ処理用）
 * @param {Number} originalScore - 元のスコア
 * @param {String} keyword - ターゲットキーワード
 * @param {Map} trendMap - 事前計算されたトレンドマップ
 * @return {Object} { adjustedScore: Number, trend: String, adjustment: Number }
 */
function adjustScoreByTrendFast(originalScore, keyword, trendMap) {
  if (!keyword || !trendMap) {
    return { adjustedScore: originalScore, trend: '不明', adjustment: 0 };
  }
  
  const keywordLower = keyword.toString().toLowerCase().trim();
  const trendResult = trendMap.get(keywordLower);
  
  if (!trendResult) {
    return { adjustedScore: originalScore, trend: '不明', adjustment: 0 };
  }
  
  let adjustedScore = originalScore + trendResult.adjustment;
  adjustedScore = Math.max(0, Math.min(100, adjustedScore));
  
  return {
    adjustedScore: adjustedScore,
    trend: trendResult.trend,
    adjustment: trendResult.adjustment
  };
}


// ===========================================
// 統合データシートへの反映
// ===========================================

/**
 * 統合データシートにトレンド情報を追加し、スコアを調整
 */
function applyTrendToIntegratedData() {
  console.log('=== トレンド反映開始 ===');
  const startTime = new Date();
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('統合データ');
  
  if (!sheet) {
    console.error('統合データシートが見つかりません');
    return { success: false, error: '統合データシートが見つかりません' };
  }
  
  // 列を確保
  ensureTrendColumns(sheet);
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  // 列インデックスを取得
  const targetKwIdx = headers.indexOf('target_keyword');
  const totalScoreIdx = headers.indexOf('total_priority_score');
  const trendIdx = headers.indexOf('trend');
  const trendAdjustmentIdx = headers.indexOf('trend_adjustment');
  const adjustedScoreIdx = headers.indexOf('adjusted_score');
  
  if (targetKwIdx === -1 || totalScoreIdx === -1) {
    console.error('必要な列が見つかりません');
    return { success: false, error: '必要な列が見つかりません' };
  }
  
  // 全キーワードのトレンドを事前計算
  const trendMap = analyzeAllTrends();
  
  let processedCount = 0;
  let adjustedCount = 0;
  
  // 各行を処理
  for (let i = 1; i < data.length; i++) {
    const keyword = data[i][targetKwIdx];
    const originalScore = parseFloat(data[i][totalScoreIdx]) || 0;
    
    if (!keyword) continue;
    
    const result = adjustScoreByTrendFast(originalScore, keyword, trendMap);
    
    // トレンド情報を書き込み
    if (trendIdx >= 0) {
      sheet.getRange(i + 1, trendIdx + 1).setValue(result.trend);
    }
    if (trendAdjustmentIdx >= 0) {
      sheet.getRange(i + 1, trendAdjustmentIdx + 1).setValue(result.adjustment);
    }
    if (adjustedScoreIdx >= 0) {
      sheet.getRange(i + 1, adjustedScoreIdx + 1).setValue(result.adjustedScore);
    }
    
    processedCount++;
    if (result.adjustment !== 0) {
      adjustedCount++;
    }
  }
  
  const endTime = new Date();
  const duration = (endTime - startTime) / 1000;
  
  console.log('=== トレンド反映完了 ===');
  console.log(`処理件数: ${processedCount}件`);
  console.log(`調整件数: ${adjustedCount}件`);
  console.log(`所要時間: ${duration}秒`);
  
  return {
    success: true,
    processedCount: processedCount,
    adjustedCount: adjustedCount,
    duration: duration
  };
}


/**
 * トレンド関連の列を確保
 */
function ensureTrendColumns(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  const requiredColumns = ['trend', 'trend_adjustment', 'adjusted_score'];
  let lastCol = sheet.getLastColumn();
  
  for (const colName of requiredColumns) {
    if (!headers.includes(colName)) {
      lastCol++;
      sheet.getRange(1, lastCol).setValue(colName);
      console.log(`列追加: ${colName} (列${lastCol})`);
    }
  }
}


// ===========================================
// テスト関数
// ===========================================

/**
 * 単一キーワードのトレンド分析テスト
 */
function testSingleTrend() {
  const testKeyword = 'amazon で iphone を 買う メリット';
  
  console.log('=== トレンド分析テスト ===');
  console.log(`キーワード: ${testKeyword}`);
  
  const result = analyzeTrend(testKeyword);
  
  console.log(`トレンド: ${result.trend}`);
  console.log(`スコア調整: ${result.adjustment}点`);
  
  if (result.details) {
    console.log(`週別中央値: ${JSON.stringify(result.details.weeklyMedians)}`);
    console.log(`Week1→Week4の変化: ${result.details.change}位`);
  }
  
  console.log('=== テスト完了 ===');
}


/**
 * スコア調整テスト
 */
function testScoreAdjustment() {
  console.log('=== スコア調整テスト ===');
  
  const testCases = [
    { keyword: 'amazon で iphone を 買う メリット', originalScore: 60 },
    { keyword: 'apple mfi認証とは', originalScore: 40 }
  ];
  
  for (const tc of testCases) {
    const result = adjustScoreByTrend(tc.originalScore, tc.keyword);
    console.log(`\nキーワード: ${tc.keyword}`);
    console.log(`元スコア: ${tc.originalScore} → 調整後: ${result.adjustedScore}`);
    console.log(`トレンド: ${result.trend}, 調整: ${result.adjustment}点`);
  }
  
  console.log('\n=== テスト完了 ===');
}


/**
 * 全キーワードのトレンド分析サマリー
 */
function testAllTrendsSummary() {
  console.log('=== 全キーワードトレンドサマリー ===');
  
  const trendMap = analyzeAllTrends();
  
  const summary = {
    '安定': 0,
    '上昇傾向': 0,
    '下降傾向': 0,
    '不安定': 0,
    'データ不足': 0,
    '不明': 0
  };
  
  for (const [keyword, result] of trendMap) {
    const trend = result.trend || '不明';
    summary[trend] = (summary[trend] || 0) + 1;
  }
  
  console.log('\n【トレンド分布】');
  for (const [trend, count] of Object.entries(summary)) {
    console.log(`${trend}: ${count}件`);
  }
  
  console.log('\n=== サマリー完了 ===');
}

// ===========================================
// CSV鮮度チェック機能
// ===========================================

const FRESHNESS_CONFIG = {
  // 警告閾値（日数）
  WARNING_THRESHOLD: 7,      // 7日以上で警告
  SKIP_THRESHOLD: 14,        // 14日以上でトレンド判定スキップ
  CRITICAL_THRESHOLD: 28     // 28日以上で強い警告
};


/**
 * GyronSEO_RAWの最新データ日付を取得
 * @return {Object} { latestDate: Date, daysOld: Number, status: String }
 */
function checkCsvFreshness() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gyronSheet = ss.getSheetByName('GyronSEO_RAW');
  
  if (!gyronSheet) {
    return { 
      latestDate: null, 
      daysOld: 999, 
      status: 'ERROR',
      message: 'GyronSEO_RAWシートが見つかりません'
    };
  }
  
  const headers = gyronSheet.getRange(1, 1, 1, gyronSheet.getLastColumn()).getValues()[0];
  
  // 最後の日付列を探す（右端から）
  let latestDate = null;
  for (let i = headers.length - 1; i >= 7; i--) {
    const date = parseDateFromHeader(headers[i]);
    if (date) {
      latestDate = date;
      break;
    }
  }
  
  if (!latestDate) {
    return { 
      latestDate: null, 
      daysOld: 999, 
      status: 'ERROR',
      message: '日付データが見つかりません'
    };
  }
  
  // 今日との差分を計算
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  latestDate.setHours(0, 0, 0, 0);
  
  const daysOld = Math.floor((today - latestDate) / (24 * 60 * 60 * 1000));
  
  // ステータス判定
  let status, message;
  
  if (daysOld <= FRESHNESS_CONFIG.WARNING_THRESHOLD) {
    status = 'OK';
    message = `順位データは最新です（最終更新: ${formatDate(latestDate)}）`;
  } else if (daysOld <= FRESHNESS_CONFIG.SKIP_THRESHOLD) {
    status = 'WARNING';
    message = `⚠️ 順位データが${daysOld}日前です（最終更新: ${formatDate(latestDate)}）。GyronSEOから最新CSVをインポートすることを推奨します。`;
  } else if (daysOld <= FRESHNESS_CONFIG.CRITICAL_THRESHOLD) {
    status = 'SKIP';
    message = `⚠️ 順位データが${daysOld}日以上古いため、トレンド判定をスキップします（最終更新: ${formatDate(latestDate)}）。GyronSEOから最新CSVをインポートしてください。`;
  } else {
    status = 'CRITICAL';
    message = `🚨 順位データが${daysOld}日以上更新されていません（最終更新: ${formatDate(latestDate)}）。トレンド判定は無効です。GyronSEOから最新CSVをインポートしてください。`;
  }
  
  return {
    latestDate: latestDate,
    daysOld: daysOld,
    status: status,
    message: message
  };
}


/**
 * 日付を "YYYY/MM/DD" 形式でフォーマット
 */
function formatDate(date) {
  if (!date) return '不明';
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `${y}/${m}/${d}`;
}


/**
 * CSV鮮度を考慮したトレンド判定
 * @param {String} keyword - ターゲットキーワード
 * @return {Object} { trend: String, adjustment: Number, details: Object, freshness: Object }
 */
function analyzeTrendWithFreshnessCheck(keyword) {
  // まず鮮度チェック
  const freshness = checkCsvFreshness();
  
  // データが古すぎる場合はスキップ
  if (freshness.status === 'SKIP' || freshness.status === 'CRITICAL' || freshness.status === 'ERROR') {
    return {
      trend: 'データ古い',
      adjustment: 0,  // スコア調整しない
      details: null,
      freshness: freshness
    };
  }
  
  // 通常のトレンド分析
  const trendResult = analyzeTrend(keyword);
  
  return {
    trend: trendResult.trend,
    adjustment: trendResult.adjustment,
    details: trendResult.details,
    freshness: freshness
  };
}


/**
 * 全キーワードのトレンドを鮮度チェック付きで分析
 * @return {Object} { trendMap: Map, freshness: Object }
 */
function analyzeAllTrendsWithFreshnessCheck() {
  console.log('=== 全キーワードトレンド分析（鮮度チェック付き）開始 ===');
  
  // 鮮度チェック
  const freshness = checkCsvFreshness();
  console.log(`CSV鮮度: ${freshness.status} - ${freshness.message}`);
  
  // データが古すぎる場合は空のマップを返す
  if (freshness.status === 'SKIP' || freshness.status === 'CRITICAL' || freshness.status === 'ERROR') {
    console.warn('データが古いため、トレンド分析をスキップします');
    return {
      trendMap: new Map(),
      freshness: freshness,
      skipped: true
    };
  }
  
  // 通常の分析
  const trendMap = analyzeAllTrends();
  
  return {
    trendMap: trendMap,
    freshness: freshness,
    skipped: false
  };
}


/**
 * 統合データシートへのトレンド反映（鮮度チェック付き）
 */
function applyTrendToIntegratedDataWithFreshnessCheck() {
  console.log('=== トレンド反映（鮮度チェック付き）開始 ===');
  const startTime = new Date();
  
  // 鮮度チェック
  const freshnessResult = analyzeAllTrendsWithFreshnessCheck();
  
  if (freshnessResult.skipped) {
    console.warn(freshnessResult.freshness.message);
    return {
      success: false,
      skipped: true,
      freshness: freshnessResult.freshness,
      message: freshnessResult.freshness.message
    };
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('統合データ');
  
  if (!sheet) {
    console.error('統合データシートが見つかりません');
    return { success: false, error: '統合データシートが見つかりません' };
  }
  
  // 列を確保
  ensureTrendColumns(sheet);
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  // 列インデックスを取得
  const targetKwIdx = headers.indexOf('target_keyword');
  const totalScoreIdx = headers.indexOf('total_priority_score');
  const trendIdx = headers.indexOf('trend');
  const trendAdjustmentIdx = headers.indexOf('trend_adjustment');
  const adjustedScoreIdx = headers.indexOf('adjusted_score');
  
  if (targetKwIdx === -1 || totalScoreIdx === -1) {
    console.error('必要な列が見つかりません');
    return { success: false, error: '必要な列が見つかりません' };
  }
  
  const trendMap = freshnessResult.trendMap;
  
  let processedCount = 0;
  let adjustedCount = 0;
  
  // 各行を処理
  for (let i = 1; i < data.length; i++) {
    const keyword = data[i][targetKwIdx];
    const originalScore = parseFloat(data[i][totalScoreIdx]) || 0;
    
    if (!keyword) continue;
    
    const result = adjustScoreByTrendFast(originalScore, keyword, trendMap);
    
    // トレンド情報を書き込み
    if (trendIdx >= 0) {
      sheet.getRange(i + 1, trendIdx + 1).setValue(result.trend);
    }
    if (trendAdjustmentIdx >= 0) {
      sheet.getRange(i + 1, trendAdjustmentIdx + 1).setValue(result.adjustment);
    }
    if (adjustedScoreIdx >= 0) {
      sheet.getRange(i + 1, adjustedScoreIdx + 1).setValue(result.adjustedScore);
    }
    
    processedCount++;
    if (result.adjustment !== 0) {
      adjustedCount++;
    }
  }
  
  const endTime = new Date();
  const duration = (endTime - startTime) / 1000;
  
  console.log('=== トレンド反映完了 ===');
  console.log(`CSV鮮度: ${freshnessResult.freshness.status}`);
  console.log(`処理件数: ${processedCount}件`);
  console.log(`調整件数: ${adjustedCount}件`);
  console.log(`所要時間: ${duration}秒`);
  
  return {
    success: true,
    processedCount: processedCount,
    adjustedCount: adjustedCount,
    duration: duration,
    freshness: freshnessResult.freshness
  };
}


// ===========================================
// 鮮度チェックテスト関数
// ===========================================

/**
 * CSV鮮度チェックのテスト
 */
function testCsvFreshness() {
  console.log('=== CSV鮮度チェックテスト ===');
  
  const result = checkCsvFreshness();
  
  console.log(`最終更新日: ${result.latestDate ? formatDate(result.latestDate) : '不明'}`);
  console.log(`経過日数: ${result.daysOld}日`);
  console.log(`ステータス: ${result.status}`);
  console.log(`メッセージ: ${result.message}`);
  
  console.log('=== テスト完了 ===');
  
  return result;
}


/**
 * 鮮度チェック付きトレンド分析のテスト
 */
function testTrendWithFreshness() {
  console.log('=== 鮮度チェック付きトレンド分析テスト ===');
  
  const testKeyword = 'amazon で iphone を 買う メリット';
  const result = analyzeTrendWithFreshnessCheck(testKeyword);
  
  console.log('\n【CSV鮮度】');
  console.log(`ステータス: ${result.freshness.status}`);
  console.log(`メッセージ: ${result.freshness.message}`);
  
  console.log('\n【トレンド分析】');
  console.log(`キーワード: ${testKeyword}`);
  console.log(`トレンド: ${result.trend}`);
  console.log(`スコア調整: ${result.adjustment}点`);
  
  if (result.details) {
    console.log(`週別中央値: ${JSON.stringify(result.details.weeklyMedians)}`);
    console.log(`Week1→Week4の変化: ${result.details.change}位`);
  }
  
  console.log('\n=== テスト完了 ===');
  
  return result;
}

// ===========================================
// ターゲットKWを統合データに反映
// ===========================================

/**
 * ターゲットKW分析シートから統合データシートにtarget_keywordを反映
 * @return {Object} { success: Boolean, updatedCount: Number }
 */
function syncTargetKeywordsToIntegratedData() {
  console.log('=== ターゲットKW同期開始 ===');
  const startTime = new Date();
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // ターゲットKW分析シートを取得
  const kwSheet = ss.getSheetByName('ターゲットKW分析');
  if (!kwSheet) {
    console.error('ターゲットKW分析シートが見つかりません');
    return { success: false, error: 'ターゲットKW分析シートが見つかりません' };
  }
  
  // 統合データシートを取得
  const integratedSheet = ss.getSheetByName('統合データ');
  if (!integratedSheet) {
    console.error('統合データシートが見つかりません');
    return { success: false, error: '統合データシートが見つかりません' };
  }
  
  // ターゲットKW分析からURL→キーワードのマップを作成
  const kwData = kwSheet.getDataRange().getValues();
  const kwHeaders = kwData[0];
  const kwPageUrlIdx = kwHeaders.indexOf('page_url');
  const kwTargetKwIdx = kwHeaders.indexOf('target_keyword');
  const kwSearchVolIdx = kwHeaders.indexOf('search_volume');
  
  if (kwPageUrlIdx === -1 || kwTargetKwIdx === -1) {
    console.error('ターゲットKW分析シートの列が見つかりません');
    return { success: false, error: '必要な列が見つかりません' };
  }
  
  // URL → {keyword, searchVolume} のマップを作成
  // 同じURLに複数KWがある場合は検索ボリューム最大のものを選択
  const urlToKeyword = new Map();
  
  for (let i = 1; i < kwData.length; i++) {
    const pageUrl = normalizeUrl(kwData[i][kwPageUrlIdx]);
    const targetKw = kwData[i][kwTargetKwIdx];
    const searchVol = parseInt(kwData[i][kwSearchVolIdx]) || 0;
    
    if (!pageUrl || !targetKw) continue;
    
    const existing = urlToKeyword.get(pageUrl);
    if (!existing || searchVol > existing.searchVolume) {
      urlToKeyword.set(pageUrl, {
        keyword: targetKw,
        searchVolume: searchVol
      });
    }
  }
  
  console.log(`URL→KWマップ作成完了: ${urlToKeyword.size}件`);
  
  // 統合データシートを更新
  const intData = integratedSheet.getDataRange().getValues();
  const intHeaders = intData[0];
  const intPageUrlIdx = intHeaders.indexOf('page_url');
  const intTargetKwIdx = intHeaders.indexOf('target_keyword');
  
  if (intPageUrlIdx === -1 || intTargetKwIdx === -1) {
    console.error('統合データシートの列が見つかりません');
    return { success: false, error: '必要な列が見つかりません' };
  }
  
  let updatedCount = 0;
  const updates = [];
  
  for (let i = 1; i < intData.length; i++) {
    const pageUrl = normalizeUrl(intData[i][intPageUrlIdx]);
    const currentKw = intData[i][intTargetKwIdx];
    
    if (!pageUrl) continue;
    
    const kwInfo = urlToKeyword.get(pageUrl);
    
    if (kwInfo && kwInfo.keyword !== currentKw) {
      updates.push({
        row: i + 1,
        col: intTargetKwIdx + 1,
        value: kwInfo.keyword
      });
      updatedCount++;
    }
  }
  
  // バッチ更新
  for (const update of updates) {
    integratedSheet.getRange(update.row, update.col).setValue(update.value);
  }
  
  const endTime = new Date();
  const duration = (endTime - startTime) / 1000;
  
  console.log('=== ターゲットKW同期完了 ===');
  console.log(`更新件数: ${updatedCount}件`);
  console.log(`所要時間: ${duration}秒`);
  
  return {
    success: true,
    updatedCount: updatedCount,
    totalMapped: urlToKeyword.size,
    duration: duration
  };
}


/**
 * URLを正規化（末尾スラッシュの統一、ドメイン除去）
 * @param {String} url - URL
 * @return {String} 正規化されたパス
 */
function normalizeUrl(url) {
  if (!url) return '';
  
  let normalized = url.toString().trim();
  
  // ドメインを除去してパスのみに
  normalized = normalized.replace(/^https?:\/\/[^\/]+/, '');
  
  // 先頭にスラッシュがなければ追加
  if (!normalized.startsWith('/')) {
    normalized = '/' + normalized;
  }
  
  // 末尾のスラッシュを除去
  normalized = normalized.replace(/\/$/, '');
  
  return normalized;
}


/**
 * ターゲットKW同期のテスト
 */
function testSyncTargetKeywords() {
  console.log('=== ターゲットKW同期テスト ===');
  
  const result = syncTargetKeywordsToIntegratedData();
  
  console.log(`成功: ${result.success}`);
  console.log(`更新件数: ${result.updatedCount}件`);
  console.log(`マッピング総数: ${result.totalMapped}件`);
  
  console.log('=== テスト完了 ===');
  
  return result;
}