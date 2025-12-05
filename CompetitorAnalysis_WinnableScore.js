/**
 * ============================================================================
 * Day 11-12: 勝算度スコア算出（完全版v6・7段階判定+1位制限）
 * ============================================================================
 * 7段階の詳細判定 + 1位の強さによる上限制限
 */

/**
 * 自社サイトのURLかどうかを判定
 */
function isOwnSite(url, ownDomain) {
  if (!url || !ownDomain) return false;
  
  try {
    const domain = extractDomain(url);
    const normalizedDomain = domain.replace(/\/$/, '');
    const normalizedOwnDomain = ownDomain.replace(/\/$/, '');
    
    return normalizedDomain === normalizedOwnDomain;
  } catch (error) {
    return false;
  }
}

/**
 * DA差分に基づく競合強度を判定
 * @param {number} daDiff - 競合DA - 自社DA
 * @return {string} 競合強度
 */
function classifyDAStrength(daDiff) {
  if (daDiff >= 20) {
    return 'かなり強い';
  } else if (daDiff >= 11) {
    return '強い';
  } else if (daDiff >= 6) {
    return 'やや強い';
  } else if (Math.abs(daDiff) <= 5) {
    return '同格';
  } else {
    return '弱い';
  }
}

/**
 * 競合DA分析（完全版v6）
 */
function analyzeCompetitorDA(ownDA, competitorDAs, competitorUrls, ownDomain) {
  // 自社サイトの現在順位を検出
  let ownSiteCurrentRank = null;
  
  for (let i = 0; i < competitorUrls.length; i++) {
    if (isOwnSite(competitorUrls[i], ownDomain)) {
      ownSiteCurrentRank = i + 1;
      break;
    }
  }
  
  // 各順位のDA差分を計算（自社サイトを除外）
  const daDiffs = competitorDAs.map((da, index) => {
    const url = competitorUrls[index];
    const isOwn = isOwnSite(url, ownDomain);
    
    // 競合DA - 自社DA（正の値 = 競合が強い）
    const diff = (da && da > 0 && !isOwn) ? (da - ownDA) : null;
    
    return {
      rank: index + 1,
      da: da || 0,
      url: url || '',
      is_own: isOwn,
      diff: diff,
      strength: diff !== null ? classifyDAStrength(diff) : null,
      exists: (da && da > 0 && !isOwn)
    };
  });
  
  // 競合サイトのみを抽出
  const competitorOnly = daDiffs.filter(d => !d.is_own && d.exists);
  
  // 1位データ
  const rank1Data = competitorOnly.length > 0 ? competitorOnly[0] : null;
  const rank1Missing = !rank1Data;
  
  // 弱いサイト（diff < 0）
  const weakSites = competitorOnly.filter(d => d.diff < 0);
  // 同格サイト（±5）
  const equalSites = competitorOnly.filter(d => Math.abs(d.diff) <= 5);
  // やや強いサイト（+6〜+10）
  const slightlyStrongSites = competitorOnly.filter(d => d.diff >= 6 && d.diff <= 10);
  
  const top3WeakCount = weakSites.filter(d => d.rank <= 3).length + 
                        equalSites.filter(d => d.rank <= 3).length;
  const top10WeakCount = weakSites.filter(d => d.rank <= 10).length + 
                        equalSites.filter(d => d.rank <= 10).length;
  
  const weakestRank = weakSites.length > 0 ? 
    Math.min(...weakSites.map(s => s.rank)) : null;
  
  return {
    da_diffs: daDiffs,
    competitor_only: competitorOnly,
    own_site_current_rank: ownSiteCurrentRank,
    rank_1_missing: rank1Missing,
    rank_1_data: rank1Data,
    weak_sites: weakSites,
    equal_sites: equalSites,
    slightly_strong_sites: slightlyStrongSites,
    top3_weak_count: top3WeakCount,
    top10_weak_count: top10WeakCount,
    total_weak_count: weakSites.length + equalSites.length,
    weakest_rank: weakestRank
  };
}

/**
 * 勝算度スコア算出（完全版v6・調整版）
 */
function calculateWinnableScore(analysis) {
  // 既に1位の場合は100点
  if (analysis.own_site_current_rank === 1) {
    return 100;
  }
  
  let score = 0;
  
  // 既にランクイン済みのボーナス（調整版）
  if (analysis.own_site_current_rank === 2) {
    score += 25;  // 30 → 25
  } else if (analysis.own_site_current_rank === 3) {
    score += 20;  // 30 → 20
  } else if (analysis.own_site_current_rank && analysis.own_site_current_rank <= 5) {
    score += 15;  // 20 → 15
  } else if (analysis.own_site_current_rank && analysis.own_site_current_rank <= 10) {
    score += 10;
  }
  
  // 1位攻略スコア（50点満点・調整版）
  if (analysis.rank_1_data) {
    const diff = analysis.rank_1_data.diff;  // 1位DA - 自社DA
    
    if (diff <= -6) {
      score += 50;  // 弱い（最優先）
    } else if (Math.abs(diff) <= 5) {
      score += 45;  // 同格（積極的）
    } else if (diff <= 10) {
      score += 25;  // やや強い（40 → 25に調整）
    } else if (diff <= 19) {
      score += 15;  // 強い（25 → 15に調整）
    } else if (diff <= 29) {
      score += 5;   // かなり強い（10 → 5に調整）
    } else {
      score += 0;   // 超強い
    }
  }
  
  // TOP3攻略スコア（30点満点）
  let top3Score = 0;
  for (let i = 0; i < 3 && i < analysis.competitor_only.length; i++) {
    const competitor = analysis.competitor_only[i];
    const diff = competitor.diff;
    
    if (diff <= -6) {
      top3Score += 10;
    } else if (Math.abs(diff) <= 5) {
      top3Score += 8;
    } else if (diff <= 10) {
      top3Score += 5;  // 6 → 5に調整
    }
  }
  score += Math.min(30, top3Score);
  
  // 弱サイト分布ボーナス（20点満点）
  score += Math.min(20, analysis.total_weak_count * 3);
  
  return Math.min(100, Math.round(score));
}

/**
 * 競合レベル判定（7段階+1位制限版）
 */
function classifyCompetitorLevel(winnableScore, rank1Strength, ownSiteCurrentRank) {
  // 既に1位の場合
  if (ownSiteCurrentRank === 1) {
    return '維持';
  }
  
  // 基本的な7段階判定
  let baseLevel;
  if (winnableScore >= 95) {
    baseLevel = '超狙い目';
  } else if (winnableScore >= 80) {
    baseLevel = '易';
  } else if (winnableScore >= 65) {
    baseLevel = '中';
  } else if (winnableScore >= 50) {
    baseLevel = 'やや難';
  } else if (winnableScore >= 35) {
    baseLevel = '難';
  } else if (winnableScore >= 20) {
    baseLevel = '厳しい';
  } else {
    baseLevel = '激戦';
  }
  
  // 1位の強さによる上限制限
  if (!rank1Strength) {
    return baseLevel;  // 1位データなしの場合はそのまま
  }
  
  const levelOrder = ['激戦', '厳しい', '難', 'やや難', '中', '易', '超狙い目'];
  
  let maxLevel;
  if (rank1Strength === '弱い' || rank1Strength === '同格') {
    maxLevel = '超狙い目';  // 制限なし
  } else if (rank1Strength === 'やや強い') {
    maxLevel = '易';  // 「易」まで
  } else if (rank1Strength === '強い') {
    maxLevel = '中';  // 「中」まで
  } else if (rank1Strength === 'かなり強い') {
    maxLevel = 'やや難';  // 「やや難」まで
  } else {
    maxLevel = '難';  // 「難」まで
  }
  
  // baseLevelとmaxLevelを比較して、厳しい方を採用
  const baseLevelIndex = levelOrder.indexOf(baseLevel);
  const maxLevelIndex = levelOrder.indexOf(maxLevel);
  
  if (baseLevelIndex > maxLevelIndex) {
    // baseLevelの方が楽観的すぎる → maxLevelにダウングレード
    return maxLevel;
  } else {
    return baseLevel;
  }
}

/**
 * コメント生成
 */
function generateAnalysisNote(analysis, winnableScore, competitorLevel) {
  let note = '';
  
  // 既にランクインしている場合
  if (analysis.own_site_current_rank) {
    if (analysis.own_site_current_rank === 1) {
      return '🏆現在1位です。このまま維持。リライト不要。';
    }
    note = `📊現在${analysis.own_site_current_rank}位。`;
  }
  
  // 1位データなし
  if (analysis.rank_1_missing) {
    note += '⚠️1位データなし。';
    return note;
  }
  
  const rank1 = analysis.rank_1_data;
  const strength = rank1.strength;
  
  // 競合レベルに応じた絵文字
  const emoji = {
    '超狙い目': '🎯',
    '易': '⭐',
    '中': '💡',
    'やや難': '🔧',
    '難': '⚡',
    '厳しい': '⚠️',
    '激戦': '🔥'
  }[competitorLevel] || '📊';
  
  // 1位の強さに応じたコメント
  if (strength === '弱い') {
    note += `${emoji}1位（DA ${rank1.da}）弱い！積極的に狙える。`;
  } else if (strength === '同格') {
    note += `${emoji}1位（DA ${rank1.da}）同格！積極的に狙える。`;
  } else if (strength === 'やや強い') {
    note += `${emoji}1位（DA ${rank1.da}）やや強い。勝算あり。`;
  } else if (strength === '強い') {
    note += `${emoji}1位（DA ${rank1.da}）強い。厳しいが可能。`;
  } else {
    note += `${emoji}1位（DA ${rank1.da}）かなり強い。長期戦。`;
  }
  
  return note;
}

/**
 * 競合分析シートの勝算度スコアを更新
 */
function updateWinnableScores(startRow = 2, endRow = null) {
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
  
  Logger.log('=== 勝算度スコア算出開始 ===');
  Logger.log(`処理行数: ${rowCount}行`);
  Logger.log('');
  
  const ownDomain = 'smaho-tap.com';
  const data = sheet.getRange(startRow, 1, rowCount, 36).getValues();
  const results = [];
  
  data.forEach((row, index) => {
    const rowIndex = startRow + index;
    const keyword = row[1];
    const ownDA = row[4];
    
    const competitorDAs = [
      row[7], row[9], row[11], row[13], row[15],
      row[17], row[19], row[21], row[23], row[25]
    ];
    
    const competitorUrls = [
      row[6], row[8], row[10], row[12], row[14],
      row[16], row[18], row[20], row[22], row[24]
    ];
    
    if (!ownDA || ownDA === 0) {
      Logger.log(`⚠ [行${rowIndex}] ${keyword}: 自社DAなし`);
      results.push([0, 0, null, false, true, null, null, 0, '激戦', 'データ不足']);
      return;
    }
    
    const analysis = analyzeCompetitorDA(ownDA, competitorDAs, competitorUrls, ownDomain);
    const winnableScore = calculateWinnableScore(analysis);
    const rank1Strength = analysis.rank_1_data ? analysis.rank_1_data.strength : null;
    const competitorLevel = classifyCompetitorLevel(winnableScore, rank1Strength, analysis.own_site_current_rank);
    const analysisNote = generateAnalysisNote(analysis, winnableScore, competitorLevel);
    
    const rank1Diff = analysis.rank_1_data ? analysis.rank_1_data.diff : null;
    
    results.push([
      analysis.top3_weak_count,
      analysis.top10_weak_count,
      analysis.weakest_rank,
      analysis.rank_1_data ? (analysis.rank_1_data.diff < 0) : false,
      analysis.rank_1_missing,
      rank1Diff,
      analysis.own_site_current_rank,
      winnableScore,
      competitorLevel,
      analysisNote
    ]);
    
    const rank1Info = analysis.rank_1_data ? 
      `DA ${analysis.rank_1_data.da}（${analysis.rank_1_data.strength}）` : '不明';
    const currentRankInfo = analysis.own_site_current_rank ? 
      `現在${analysis.own_site_current_rank}位` : '圏外';
    
    Logger.log(`✓ [行${rowIndex}] ${keyword}: 勝算度${winnableScore}点（${competitorLevel}）`);
    Logger.log(`  ${currentRankInfo}, 自社DA: ${ownDA}, 1位: ${rank1Info}`);
    Logger.log(`  ${analysisNote}`);
    Logger.log('');
  });
  
  sheet.getRange(startRow, 27, rowCount, 10).setValues(results);
  
  Logger.log('=== 勝算度スコア算出完了 ===');
  Logger.log(`処理完了: ${rowCount}行`);
}

function testWinnableScoreCalculation() {
  Logger.log('=== 勝算度スコア算出テスト ===\n');
  updateWinnableScores(2, 6);
  Logger.log('\n=== テスト完了 ===');
}

function updateAllWinnableScores() {
  Logger.log('=== 全行勝算度スコア算出開始 ===\n');
  updateWinnableScores();
  Logger.log('\n=== 全行勝算度スコア算出完了 ===');
}