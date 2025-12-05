/**
 * ScoringV2.gs - 5軸スコアリング完全版
 * Day 13: ページ総合スコアリング5軸版
 * 
 * 5軸評価:
 * 1. 機会損失スコア（20%）- 既存Scoring.gsから
 * 2. パフォーマンススコア（20%、UX含む）- 既存Scoring.gsから
 * 3. ビジネスインパクトスコア（20%）- 既存Scoring.gsから
 * 4. キーワード戦略スコア（20%）- 新規実装
 * 5. 競合難易度スコア（20%）- winnable_scoreをそのまま使用
 * 
 * 【重要】競合難易度スコアについて（Day 11-12決定事項）:
 * - 平均DA比較は使用しない（意味がないため）
 * - winnable_score（v6）をそのまま使用
 * - v6ロジック: 1位〜10位それぞれのDAと自社DAを個別比較
 *   - 1位が弱い（自社DA以下）→ +20点（超狙い目）
 *   - 2位が弱い → +15点
 *   - 3位が弱い → +12点
 *   - TOP3に弱サイト3個以上 → +30点（混戦型）
 *   - 1位の強さによる上限制限あり
 * - 7段階判定: 超狙い目, 易, 中, やや難, 難, 厳しい, 激戦
 * 
 * 作成日: 2025/11/30
 */

// ===========================================
// 定数定義
// ===========================================

const SCORING_CONFIG = {
  // 5軸の重み（各20%）
  WEIGHTS: {
    OPPORTUNITY: 0.20,
    PERFORMANCE: 0.20,
    BUSINESS_IMPACT: 0.20,
    KEYWORD_STRATEGY: 0.20,
    COMPETITOR_DIFFICULTY: 0.20
  },
  
  // 優先度判定閾値
  PRIORITY_THRESHOLDS: {
    HIGHEST: 80,  // 最優先（今すぐリライト）
    HIGH: 60,     // 高優先（今週中）
    MEDIUM: 40,   // 中優先（今月中）
    LOW: 0        // 低優先（様子見）
  },
  
  // CV近接キーワード
  CV_PROXIMITY_KEYWORDS: {
    HIGH: ['おすすめ', '比較', 'ランキング', '評判', '口コミ', 'レビュー', '最安', '格安', '購入', '申込', '契約', '見積'],
    MEDIUM: ['選び方', 'メリット', 'デメリット', '違い', '特徴', '料金', '価格', '費用', 'プラン'],
    LOW: ['とは', '意味', '仕組み', '歴史', '基礎', '入門', '初心者']
  },
  
  // 検索ボリューム閾値
  SEARCH_VOLUME_THRESHOLDS: {
    HIGH: 1000,
    MEDIUM: 500,
    LOW: 100
  }
};


// ===========================================
// メイン関数: 全ページ5軸再スコアリング
// ===========================================

/**
 * 全144ページを5軸スコアリングで再計算
 */
function recalculateAllScoresV2() {
  console.log('=== 5軸スコアリング開始 ===');
  const startTime = new Date();
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const integratedSheet = ss.getSheetByName('統合データ');
  const competitorSheet = ss.getSheetByName('競合分析');
  const targetKWSheet = ss.getSheetByName('ターゲットKW分析');
  const querySheet = ss.getSheetByName('クエリ分析');
  
  if (!integratedSheet) {
    console.error('統合データシートが見つかりません');
    return { success: false, error: '統合データシートが見つかりません' };
  }
  
  // 統合データシートの列構成を確認・更新
  ensureIntegratedSheetColumns(integratedSheet);
  
  // データ取得
  const integratedData = integratedSheet.getDataRange().getValues();
  const headers = integratedData[0];
  
  // 列インデックスを取得
  const colIndex = getColumnIndices(headers);
  
  // 競合分析データをマップ化（高速検索用）
  const competitorMap = buildCompetitorMap(competitorSheet);
  
  // ターゲットKWデータをマップ化
  const targetKWMap = buildTargetKWMap(targetKWSheet);
  
  // クエリ分析データをマップ化
  const queryMap = buildQueryMap(querySheet);
  
  let processedCount = 0;
  let excludedCount = 0;
  const results = [];
  
  // 各ページをスコアリング
  for (let i = 1; i < integratedData.length; i++) {
    const row = integratedData[i];
    const pageUrl = row[colIndex.page_url];
    
    if (!pageUrl) continue;
    
    // 既存の3軸スコアを取得
    const opportunityScore = row[colIndex.opportunity_score] || 0;
    const performanceScore = row[colIndex.performance_score] || 0;
    const businessImpactScore = row[colIndex.business_impact_score] || 0;
    
    // ターゲットKWを取得
    const targetKeyword = row[colIndex.target_keyword] || '';
    
    // 現在の順位を取得
    const currentPosition = row[colIndex.avg_position] || 999;
    
    // 4. キーワード戦略スコアを計算
    const keywordStrategyScore = calculateKeywordStrategyScore(
      pageUrl, 
      targetKeyword, 
      targetKWMap, 
      queryMap
    );
    
    // 5. 競合難易度スコアを取得（winnable_scoreをそのまま使用）
    // URLをパス形式に正規化してマッチング
    const normalizedPageUrl = normalizeUrlToPath(pageUrl);
    const competitorData = getCompetitorData(normalizedPageUrl, targetKeyword, competitorMap);
    const competitorDifficultyScore = competitorData.winnableScore;
    const competitorLevel = competitorData.competitorLevel;
    
    // 絶対的足切り条件チェック
    const exclusionResult = checkAbsoluteExclusion(currentPosition, competitorLevel);
    
    let totalScore;
    let exclusionReason = '';
    
    if (exclusionResult.excluded) {
      // 足切り対象
      totalScore = 0;
      exclusionReason = exclusionResult.reason;
      excludedCount++;
    } else {
      // 5軸総合スコア計算
      totalScore = calculateTotalPriorityScore5Axis(
        opportunityScore,
        performanceScore,
        businessImpactScore,
        keywordStrategyScore,
        competitorDifficultyScore
      );
    }
    
    // 優先度判定
    const priorityLevel = getPriorityLevel(totalScore);
    
    // 結果を保存
    results.push({
      rowIndex: i + 1, // 1-indexed for sheet
      keywordStrategyScore: keywordStrategyScore,
      competitorDifficultyScore: competitorDifficultyScore,
      totalScore: totalScore,
      priorityLevel: priorityLevel,
      exclusionReason: exclusionReason
    });
    
    processedCount++;
    
    // 進捗ログ（20件ごと）
    if (processedCount % 20 === 0) {
      console.log(`処理中: ${processedCount}件完了`);
    }
  }
  
  // 結果をシートに書き込み
  writeScoresToSheet(integratedSheet, results, colIndex);
  
  // 優先度スコアでソート
  sortByPriorityScore(integratedSheet, colIndex);
  
  const endTime = new Date();
  const duration = (endTime - startTime) / 1000;
  
  console.log('=== 5軸スコアリング完了 ===');
  console.log(`処理件数: ${processedCount}件`);
  console.log(`除外件数: ${excludedCount}件`);
  console.log(`所要時間: ${duration}秒`);
  
  return {
    success: true,
    processedCount: processedCount,
    excludedCount: excludedCount,
    duration: duration
  };
}


// ===========================================
// 統合データシート列管理
// ===========================================

/**
 * 統合データシートに必要な列が存在することを確認し、なければ追加
 */
function ensureIntegratedSheetColumns(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  const requiredColumns = [
    'keyword_strategy_score',
    'competitor_difficulty_score',
    'exclusion_reason'
  ];
  
  let lastCol = sheet.getLastColumn();
  
  for (const colName of requiredColumns) {
    if (!headers.includes(colName)) {
      lastCol++;
      sheet.getRange(1, lastCol).setValue(colName);
      console.log(`列追加: ${colName} (列${lastCol})`);
    }
  }
}

/**
 * 列インデックスを取得
 */
function getColumnIndices(headers) {
  const indices = {};
  
  const columnNames = [
    'page_url',
    'target_keyword',
    'avg_position',
    'opportunity_score',
    'performance_score',
    'business_impact_score',
    'keyword_strategy_score',
    'competitor_difficulty_score',
    'total_priority_score',
    'exclusion_reason',
    'status'
  ];
  
  for (const name of columnNames) {
    const index = headers.indexOf(name);
    indices[name] = index >= 0 ? index : -1;
  }
  
  return indices;
}


// ===========================================
// データマップ構築（高速検索用）
// ===========================================

/**
 * 競合分析データをマップ化
 * キー: "pageUrl|targetKeyword"
 */
function buildCompetitorMap(sheet) {
  const map = new Map();
  
  if (!sheet) {
    console.warn('競合分析シートが見つかりません');
    return map;
  }
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  // 列インデックス
  const targetKwIdx = headers.indexOf('target_keyword');
  const pageUrlIdx = headers.indexOf('page_url');
  const winnableScoreIdx = headers.indexOf('winnable_score');
  const competitorLevelIdx = headers.indexOf('competitor_level');
  
  if (winnableScoreIdx === -1) {
    console.warn('winnable_score列が見つかりません');
    return map;
  }
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const targetKw = row[targetKwIdx] || '';
    // URLをパス形式に正規化
    const pageUrl = normalizeUrlToPath(row[pageUrlIdx] || '');
    const winnableScore = row[winnableScoreIdx] || 0;
    const competitorLevel = row[competitorLevelIdx] || '';
    
    // キーを作成（URLとKWの両方でマッチング）
    const key = `${pageUrl}|${targetKw.toLowerCase()}`;
    
    map.set(key, {
      winnableScore: winnableScore,
      competitorLevel: competitorLevel
    });
    
    // KWのみでもマッチングできるようにする（バックアップ）
    const kwOnlyKey = `kw|${targetKw.toLowerCase()}`;
    if (!map.has(kwOnlyKey)) {
      map.set(kwOnlyKey, {
        winnableScore: winnableScore,
        competitorLevel: competitorLevel
      });
    }
  }
  
  console.log(`競合分析マップ構築: ${map.size}件`);
  return map;
}

/**
 * ターゲットKW分析データをマップ化
 * URLではなくtarget_keywordでマッチング（GyronSEOで圏外のKWはURLが空のため）
 */
function buildTargetKWMap(sheet) {
  const map = new Map();
  
  if (!sheet) {
    console.warn('ターゲットKW分析シートが見つかりません');
    return map;
  }
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const pageUrlIdx = headers.indexOf('page_url');
  const targetKwIdx = headers.indexOf('target_keyword');
  const searchVolumeIdx = headers.indexOf('search_volume');
  const gscPositionIdx = headers.indexOf('gsc_position');
  const gyronPositionIdx = headers.indexOf('gyron_position');
  const kwScoreIdx = headers.indexOf('kw_score');
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const targetKw = (row[targetKwIdx] || '').toString().trim().toLowerCase();
    const pageUrl = normalizeUrlToPath(row[pageUrlIdx] || '');
    
    if (!targetKw) continue;
    
    const kwData = {
      targetKeyword: row[targetKwIdx] || '',
      pageUrl: pageUrl,
      searchVolume: row[searchVolumeIdx] || 0,
      gscPosition: row[gscPositionIdx] || 999,
      gyronPosition: row[gyronPositionIdx] || 999,
      kwScore: row[kwScoreIdx] || 0
    };
    
    // target_keywordでマッチング（メイン）
    if (!map.has(targetKw)) {
      map.set(targetKw, kwData);
    }
    
    // URLがある場合はURLでもマッチングできるようにする（バックアップ）
    if (pageUrl && !map.has(pageUrl)) {
      map.set(pageUrl, kwData);
    }
  }
  
  console.log(`ターゲットKWマップ構築: ${map.size}件`);
  return map;
}

/**
 * URLをパス形式に正規化（/で始まる形式）
 */
function normalizeUrlToPath(url) {
  if (!url) return '';
  
  let path = url.toString().toLowerCase();
  
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
 * クエリ分析データをマップ化
 */
function buildQueryMap(sheet) {
  const map = new Map();
  
  if (!sheet) {
    console.warn('クエリ分析シートが見つかりません');
    return map;
  }
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const pageUrlIdx = headers.indexOf('page_url');
  const queryIdx = headers.indexOf('query');
  const positionIdx = headers.indexOf('position');
  const clicksIdx = headers.indexOf('clicks');
  const impressionsIdx = headers.indexOf('impressions');
  const ctrIdx = headers.indexOf('ctr');
  const queryScoreIdx = headers.indexOf('query_score');
  const cvProximityIdx = headers.indexOf('cv_proximity');
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    // URLをパス形式に正規化
    const pageUrl = normalizeUrlToPath(row[pageUrlIdx] || '');
    
    if (!pageUrl) continue;
    
    if (!map.has(pageUrl)) {
      map.set(pageUrl, []);
    }
    
    map.get(pageUrl).push({
      query: row[queryIdx] || '',
      position: row[positionIdx] || 999,
      clicks: row[clicksIdx] || 0,
      impressions: row[impressionsIdx] || 0,
      ctr: row[ctrIdx] || 0,
      queryScore: row[queryScoreIdx] || 0,
      cvProximity: row[cvProximityIdx] || ''
    });
  }
  
  console.log(`クエリ分析マップ構築: ${map.size}件`);
  return map;
}


// ===========================================
// キーワード戦略スコア計算
// ===========================================

/**
 * キーワード戦略スコアを計算
 * = (キーワード品質スコア × 0.50) + (ターゲットKW達成度スコア × 0.50)
 */
function calculateKeywordStrategyScore(pageUrl, targetKeyword, targetKWMap, queryMap) {
  // URLをパス形式に正規化
  const normalizedUrl = normalizeUrlToPath(pageUrl);
  
  // キーワード品質スコア
  const qualityScore = calculateKeywordQualityScore(normalizedUrl, targetKeyword, targetKWMap);
  
  // ターゲットKW達成度スコア
  const achievementScore = calculateTargetKWAchievementScore(normalizedUrl, targetKeyword, queryMap, targetKWMap);
  
  // 合計（各50%）
  const totalScore = (qualityScore * 0.50) + (achievementScore * 0.50);
  
  return Math.round(totalScore);
}

/**
 * キーワード品質スコアを計算
 * = (ターゲットKWと実クエリの一致度 × 0.30) + (CVへの近さ × 0.40) + (検索ボリューム × 0.30)
 */
function calculateKeywordQualityScore(pageUrl, targetKeyword, targetKWMap) {
  // target_keywordでマッチング
  const targetKwLower = (targetKeyword || '').toString().trim().toLowerCase();
  let kwData = targetKWMap.get(targetKwLower);
  
  // target_keywordでマッチしない場合はURLで試す
  if (!kwData) {
    kwData = targetKWMap.get(pageUrl);
  }
  
  if (!kwData) {
    return 30; // データがない場合はデフォルト値
  }
  
  // 1. ターゲットKWと実クエリの一致度（30%）
  // ※この部分はクエリ分析との突合が必要だが、簡易版として既存のkwScoreを使用
  const matchScore = kwData.kwScore || 50;
  
  // 2. CVへの近さ（40%）
  const cvProximityScore = calculateCVProximityScore(targetKeyword);
  
  // 3. 検索ボリューム（30%）
  const volumeScore = calculateSearchVolumeScore(kwData.searchVolume);
  
  const totalScore = (matchScore * 0.30) + (cvProximityScore * 0.40) + (volumeScore * 0.30);
  
  return Math.min(100, Math.max(0, totalScore));
}

/**
 * CVへの近さスコアを計算
 */
function calculateCVProximityScore(keyword) {
  if (!keyword) return 20;
  
  const lowerKW = keyword.toLowerCase();
  
  // 高CV（100点）
  for (const word of SCORING_CONFIG.CV_PROXIMITY_KEYWORDS.HIGH) {
    if (lowerKW.includes(word)) return 100;
  }
  
  // 中CV（60点）
  for (const word of SCORING_CONFIG.CV_PROXIMITY_KEYWORDS.MEDIUM) {
    if (lowerKW.includes(word)) return 60;
  }
  
  // 低CV（20点）
  for (const word of SCORING_CONFIG.CV_PROXIMITY_KEYWORDS.LOW) {
    if (lowerKW.includes(word)) return 20;
  }
  
  // デフォルト（40点）
  return 40;
}

/**
 * 検索ボリュームスコアを計算
 */
function calculateSearchVolumeScore(volume) {
  if (!volume || volume === 0) return 20;
  
  if (volume >= SCORING_CONFIG.SEARCH_VOLUME_THRESHOLDS.HIGH) return 100;
  if (volume >= SCORING_CONFIG.SEARCH_VOLUME_THRESHOLDS.MEDIUM) return 70;
  if (volume >= SCORING_CONFIG.SEARCH_VOLUME_THRESHOLDS.LOW) return 40;
  
  return 20;
}

/**
 * ターゲットKW達成度スコアを計算
 */
function calculateTargetKWAchievementScore(pageUrl, targetKeyword, queryMap, targetKWMap) {
  if (!targetKeyword) return 0;
  
  const queries = queryMap.get(pageUrl) || [];
  
  // target_keywordでマッチング
  const targetKwLower = (targetKeyword || '').toString().trim().toLowerCase();
  let kwData = targetKWMap.get(targetKwLower);
  
  // target_keywordでマッチしない場合はURLで試す
  if (!kwData) {
    kwData = targetKWMap.get(pageUrl);
  }
  
  // ターゲットKWの現在順位を取得
  const currentPosition = kwData ? Math.min(kwData.gscPosition || 999, kwData.gyronPosition || 999) : 999;
  
  // クエリの中でターゲットKWがどの順位にいるか
  let queryRank = -1;
  
  // クエリを表示回数順にソート
  const sortedQueries = [...queries].sort((a, b) => b.impressions - a.impressions);
  
  for (let i = 0; i < sortedQueries.length; i++) {
    const query = sortedQueries[i].query.toLowerCase();
    if (query.includes(targetKwLower) || targetKwLower.includes(query)) {
      queryRank = i + 1;
      break;
    }
  }
  
  // 基本スコア
  let baseScore = 0;
  
  if (queryRank >= 1 && queryRank <= 3) {
    baseScore = 100; // トップ3に含まれる
  } else if (queryRank >= 4 && queryRank <= 10) {
    baseScore = 70; // トップ10に含まれる
  } else if (queryRank >= 11 && queryRank <= 20) {
    baseScore = 40; // トップ20に含まれる
  } else {
    baseScore = 0; // 含まれない
  }
  
  // 順位ボーナス
  let bonus = 0;
  if (currentPosition <= 5) {
    bonus = 40; // 5位以内
  } else if (currentPosition <= 10) {
    bonus = 20; // 10位以内
  }
  
  return Math.min(100, baseScore + bonus);
}


// ===========================================
// 競合難易度スコア取得
// ===========================================

/**
 * 競合分析データを取得
 * 
 * 【重要】Day 11-12で決定した方針:
 * - 競合難易度スコア = winnable_score（そのまま使用）
 * - 平均DA比較は使用しない（意味がないため）
 * - winnable_scoreはDay 11-12で実装済みのv6ロジックで算出済み
 *   - 1位が弱い → +20点
 *   - TOP3に弱サイト3個以上 → +30点
 *   - 1位の強さによる上限制限あり
 *   - 7段階判定（超狙い目〜激戦）
 */
function getCompetitorData(pageUrl, targetKeyword, competitorMap) {
  // URLをパス形式に正規化
  const normalizedUrl = normalizeUrlToPath(pageUrl);
  const normalizedKW = (targetKeyword || '').toLowerCase();
  
  // URL + KWでマッチング
  const key = `${normalizedUrl}|${normalizedKW}`;
  let data = competitorMap.get(key);
  
  // 見つからない場合はKWのみでマッチング
  if (!data && normalizedKW) {
    const kwOnlyKey = `kw|${normalizedKW}`;
    data = competitorMap.get(kwOnlyKey);
  }
  
  if (data) {
    return {
      // winnable_scoreをそのまま競合難易度スコアとして使用
      winnableScore: data.winnableScore || 0,
      competitorLevel: data.competitorLevel || ''
    };
  }
  
  // データがない場合はデフォルト値（競合分析未実施）
  return {
    winnableScore: 50, // 中間値（未分析扱い）
    competitorLevel: '未分析'
  };
}


// ===========================================
// 絶対的足切り条件
// ===========================================

/**
 * 絶対的足切り条件をチェック
 * 条件: 圏外（順位101以上） AND 競合レベル「激戦」
 */
function checkAbsoluteExclusion(position, competitorLevel) {
  const isOutOfRank = position >= 101 || position === 999 || !position;
  const isIntenseCompetition = competitorLevel === '激戦';
  
  if (isOutOfRank && isIntenseCompetition) {
    return {
      excluded: true,
      reason: '圏外＋激戦区のため除外'
    };
  }
  
  return {
    excluded: false,
    reason: ''
  };
}


// ===========================================
// 5軸総合スコア計算
// ===========================================

/**
 * 5軸総合優先度スコアを計算
 */
function calculateTotalPriorityScore5Axis(
  opportunityScore,
  performanceScore,
  businessImpactScore,
  keywordStrategyScore,
  competitorDifficultyScore
) {
  const weights = SCORING_CONFIG.WEIGHTS;
  
  const totalScore = 
    (opportunityScore * weights.OPPORTUNITY) +
    (performanceScore * weights.PERFORMANCE) +
    (businessImpactScore * weights.BUSINESS_IMPACT) +
    (keywordStrategyScore * weights.KEYWORD_STRATEGY) +
    (competitorDifficultyScore * weights.COMPETITOR_DIFFICULTY);
  
  return Math.round(totalScore);
}

/**
 * 優先度レベルを判定
 */
function getPriorityLevel(score) {
  const thresholds = SCORING_CONFIG.PRIORITY_THRESHOLDS;
  
  if (score >= thresholds.HIGHEST) return '最優先';
  if (score >= thresholds.HIGH) return '高優先';
  if (score >= thresholds.MEDIUM) return '中優先';
  return '低優先';
}


// ===========================================
// 結果書き込み
// ===========================================

/**
 * スコア結果をシートに書き込み
 */
function writeScoresToSheet(sheet, results, colIndex) {
  console.log('結果をシートに書き込み中...');
  
  // 列インデックスを再取得（列追加後）
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const kwStrategyIdx = headers.indexOf('keyword_strategy_score');
  const compDiffIdx = headers.indexOf('competitor_difficulty_score');
  const totalScoreIdx = headers.indexOf('total_priority_score');
  const exclusionIdx = headers.indexOf('exclusion_reason');
  
  for (const result of results) {
    const row = result.rowIndex;
    
    // キーワード戦略スコア
    if (kwStrategyIdx >= 0) {
      sheet.getRange(row, kwStrategyIdx + 1).setValue(result.keywordStrategyScore);
    }
    
    // 競合難易度スコア
    if (compDiffIdx >= 0) {
      sheet.getRange(row, compDiffIdx + 1).setValue(result.competitorDifficultyScore);
    }
    
    // 総合優先度スコア
    if (totalScoreIdx >= 0) {
      sheet.getRange(row, totalScoreIdx + 1).setValue(result.totalScore);
    }
    
    // 除外理由
    if (exclusionIdx >= 0) {
      sheet.getRange(row, exclusionIdx + 1).setValue(result.exclusionReason);
    }
  }
  
  console.log(`${results.length}件のスコアを書き込みました`);
}

/**
 * 優先度スコアで降順ソート
 */
function sortByPriorityScore(sheet, colIndex) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const totalScoreIdx = headers.indexOf('total_priority_score');
  
  if (totalScoreIdx < 0) {
    console.warn('total_priority_score列が見つかりません。ソートをスキップします。');
    return;
  }
  
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  if (lastRow <= 1) return;
  
  // データ範囲をソート（ヘッダー除く）
  const dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
  dataRange.sort({ column: totalScoreIdx + 1, ascending: false });
  
  console.log('優先度スコアで降順ソートしました');
}


// ===========================================
// ユーティリティ関数
// ===========================================

// normalizeUrlToPath関数は上部で定義済み（buildTargetKWMap関数の後）


// ===========================================
// テスト関数
// ===========================================

/**
 * 5軸スコアリングのテスト実行
 */
function testScoring5Axis() {
  console.log('=== 5軸スコアリング テスト開始 ===');
  
  // テスト1: キーワード品質スコア
  console.log('\n--- テスト1: CV近接度スコア ---');
  console.log('iphone おすすめ:', calculateCVProximityScore('iphone おすすめ')); // 100
  console.log('iphone 選び方:', calculateCVProximityScore('iphone 選び方')); // 60
  console.log('iphone とは:', calculateCVProximityScore('iphone とは')); // 20
  console.log('iphone 保険:', calculateCVProximityScore('iphone 保険')); // 40
  
  // テスト2: 検索ボリュームスコア
  console.log('\n--- テスト2: 検索ボリュームスコア ---');
  console.log('1500:', calculateSearchVolumeScore(1500)); // 100
  console.log('700:', calculateSearchVolumeScore(700)); // 70
  console.log('200:', calculateSearchVolumeScore(200)); // 40
  console.log('50:', calculateSearchVolumeScore(50)); // 20
  
  // テスト3: 絶対的足切り条件
  console.log('\n--- テスト3: 絶対的足切り条件 ---');
  console.log('圏外+激戦:', checkAbsoluteExclusion(150, '激戦')); // excluded: true
  console.log('圏外+易:', checkAbsoluteExclusion(150, '易')); // excluded: false
  console.log('10位+激戦:', checkAbsoluteExclusion(10, '激戦')); // excluded: false
  
  // テスト4: 5軸総合スコア
  console.log('\n--- テスト4: 5軸総合スコア ---');
  const score = calculateTotalPriorityScore5Axis(80, 70, 60, 75, 85);
  console.log('5軸スコア (80,70,60,75,85):', score); // 74
  console.log('優先度:', getPriorityLevel(score)); // 高優先
  
  console.log('\n=== テスト完了 ===');
}

/**
 * 単一ページのスコアリングテスト
 */
function testSinglePageScoring() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const integratedSheet = ss.getSheetByName('統合データ');
  const competitorSheet = ss.getSheetByName('競合分析');
  const targetKWSheet = ss.getSheetByName('ターゲットKW分析');
  const querySheet = ss.getSheetByName('クエリ分析');
  
  // マップ構築
  const competitorMap = buildCompetitorMap(competitorSheet);
  const targetKWMap = buildTargetKWMap(targetKWSheet);
  const queryMap = buildQueryMap(querySheet);
  
  // テストデータ（最初のページ）
  const testUrl = integratedSheet.getRange(2, 1).getValue();
  const testKW = integratedSheet.getRange(2, integratedSheet.getRange(1, 1, 1, integratedSheet.getLastColumn()).getValues()[0].indexOf('target_keyword') + 1).getValue();
  
  console.log('テストURL:', testUrl);
  console.log('テストKW:', testKW);
  
  // スコア計算
  const kwScore = calculateKeywordStrategyScore(testUrl, testKW, targetKWMap, queryMap);
  console.log('キーワード戦略スコア:', kwScore);
  
  const compData = getCompetitorData(testUrl, testKW, competitorMap);
  console.log('競合難易度スコア:', compData.winnableScore);
  console.log('競合レベル:', compData.competitorLevel);
}


/**
 * シート構造確認（デバッグ用）
 */
function debugSheetStructure() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  console.log('=== シート構造確認 ===\n');
  
  // 統合データシート
  const integratedSheet = ss.getSheetByName('統合データ');
  if (integratedSheet) {
    const headers = integratedSheet.getRange(1, 1, 1, integratedSheet.getLastColumn()).getValues()[0];
    console.log('【統合データシート】');
    console.log('列数:', headers.length);
    console.log('行数:', integratedSheet.getLastRow());
    console.log('列名:', headers.join(', '));
    console.log('');
  }
  
  // 競合分析シート
  const competitorSheet = ss.getSheetByName('競合分析');
  if (competitorSheet) {
    const headers = competitorSheet.getRange(1, 1, 1, competitorSheet.getLastColumn()).getValues()[0];
    console.log('【競合分析シート】');
    console.log('列数:', headers.length);
    console.log('行数:', competitorSheet.getLastRow());
    console.log('列名:', headers.join(', '));
    
    // winnable_score列の確認
    const winnableIdx = headers.indexOf('winnable_score');
    const levelIdx = headers.indexOf('competitor_level');
    console.log('winnable_score列:', winnableIdx >= 0 ? `あり（列${winnableIdx + 1}）` : 'なし');
    console.log('competitor_level列:', levelIdx >= 0 ? `あり（列${levelIdx + 1}）` : 'なし');
    console.log('');
  }
  
  // ターゲットKW分析シート
  const targetKWSheet = ss.getSheetByName('ターゲットKW分析');
  if (targetKWSheet) {
    const headers = targetKWSheet.getRange(1, 1, 1, targetKWSheet.getLastColumn()).getValues()[0];
    console.log('【ターゲットKW分析シート】');
    console.log('列数:', headers.length);
    console.log('行数:', targetKWSheet.getLastRow());
    console.log('列名:', headers.join(', '));
    console.log('');
  }
  
  // クエリ分析シート
  const querySheet = ss.getSheetByName('クエリ分析');
  if (querySheet) {
    const headers = querySheet.getRange(1, 1, 1, querySheet.getLastColumn()).getValues()[0];
    console.log('【クエリ分析シート】');
    console.log('列数:', headers.length);
    console.log('行数:', querySheet.getLastRow());
    console.log('列名:', headers.join(', '));
    console.log('');
  }
  
  console.log('=== 確認完了 ===');
}


/**
 * 競合分析データのサンプル確認（デバッグ用）
 */
function debugCompetitorData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const competitorSheet = ss.getSheetByName('競合分析');
  
  if (!competitorSheet) {
    console.log('競合分析シートが見つかりません');
    return;
  }
  
  const data = competitorSheet.getDataRange().getValues();
  const headers = data[0];
  
  console.log('=== 競合分析データサンプル ===\n');
  console.log('総行数:', data.length - 1);
  
  // 列インデックス
  const targetKwIdx = headers.indexOf('target_keyword');
  const pageUrlIdx = headers.indexOf('page_url');
  const winnableScoreIdx = headers.indexOf('winnable_score');
  const competitorLevelIdx = headers.indexOf('competitor_level');
  
  // 最初の5件を表示
  for (let i = 1; i <= Math.min(5, data.length - 1); i++) {
    const row = data[i];
    console.log(`\n--- データ${i} ---`);
    console.log('ターゲットKW:', row[targetKwIdx]);
    console.log('ページURL:', row[pageUrlIdx]);
    console.log('勝算度スコア:', row[winnableScoreIdx]);
    console.log('競合レベル:', row[competitorLevelIdx]);
  }
  
  // 7段階判定の分布
  console.log('\n=== 7段階判定の分布 ===');
  const levelCounts = {};
  for (let i = 1; i < data.length; i++) {
    const level = data[i][competitorLevelIdx] || '未設定';
    levelCounts[level] = (levelCounts[level] || 0) + 1;
  }
  for (const [level, count] of Object.entries(levelCounts)) {
    console.log(`${level}: ${count}件`);
  }
}


// ===========================================
// 手動実行用関数
// ===========================================

/**
 * 5軸スコアリングを実行（メニューから呼び出し用）
 */
function runScoring5Axis() {
  const ui = SpreadsheetApp.getUi();
  
  const result = ui.alert(
    '5軸スコアリング実行',
    '全144ページを5軸でスコアリングします。\n既存のスコアは上書きされます。\n実行しますか？',
    ui.ButtonSet.YES_NO
  );
  
  if (result !== ui.Button.YES) {
    ui.alert('キャンセルしました');
    return;
  }
  
  try {
    const execResult = recalculateAllScoresV2();
    
    if (execResult.success) {
      ui.alert(
        '完了',
        `5軸スコアリングが完了しました。\n\n` +
        `処理件数: ${execResult.processedCount}件\n` +
        `除外件数: ${execResult.excludedCount}件\n` +
        `所要時間: ${execResult.duration.toFixed(1)}秒`,
        ui.ButtonSet.OK
      );
    } else {
      ui.alert('エラー', execResult.error, ui.ButtonSet.OK);
    }
  } catch (error) {
    console.error('エラー:', error);
    ui.alert('エラー', error.toString(), ui.ButtonSet.OK);
  }
}
function debugTargetKWMatch() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 統合データシートのtarget_keyword取得
  const intSheet = ss.getSheetByName('統合データ');
  const intData = intSheet.getDataRange().getValues();
  const intHeaders = intData[0];
  const intTargetKwIdx = intHeaders.indexOf('target_keyword');
  
  // ターゲットKW分析シートのtarget_keyword取得
  const kwSheet = ss.getSheetByName('ターゲットKW分析');
  const kwData = kwSheet.getDataRange().getValues();
  const kwHeaders = kwData[0];
  const kwTargetKwIdx = kwHeaders.indexOf('target_keyword');
  
  // ターゲットKW分析のKWをセットに
  const kwSet = new Set();
  for (let i = 1; i < kwData.length; i++) {
    const kw = (kwData[i][kwTargetKwIdx] || '').toString().trim().toLowerCase();
    if (kw) kwSet.add(kw);
  }
  
  console.log('ターゲットKW分析のKW数:', kwSet.size);
  
  // 統合データのtarget_keywordをチェック
  let matchCount = 0;
  let unmatchCount = 0;
  const unmatchedKWs = [];
  
  for (let i = 1; i < intData.length; i++) {
    const kw = (intData[i][intTargetKwIdx] || '').toString().trim();
    if (!kw) continue;
    
    const kwLower = kw.toLowerCase();
    
    if (kwSet.has(kwLower)) {
      matchCount++;
    } else {
      unmatchCount++;
      if (unmatchedKWs.length < 10) {
        unmatchedKWs.push(kw);
      }
    }
  }
  
  console.log('マッチ:', matchCount);
  console.log('アンマッチ:', unmatchCount);
  console.log('アンマッチ例:', unmatchedKWs);
}