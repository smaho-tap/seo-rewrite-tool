/**
 * SEOリライト支援ツール - SuggestionGenerator.gs
 * バージョン: 1.2（5軸完全対応版）
 * 作成日: 2025年12月1日
 * 更新日: 2025年12月1日
 * 
 * 機能（5軸対応）:
 * - クエリベース提案生成（機会損失スコア対応）
 * - UXベース提案生成（パフォーマンススコア対応）
 * - イベントベース提案生成（ビジネスインパクトスコア対応）
 * - キーワード戦略ベース提案生成（キーワード戦略スコア対応）★NEW
 * - 競合分析ベース提案生成（競合難易度スコア対応）★NEW
 * - 統合リライトレポート生成（5軸統合）
 */

// ============================================
// 定数定義
// ============================================

const SUGGESTION_GENERATOR_CONFIG = {
  // クエリ分析の閾値
  MIN_IMPRESSIONS: 10,        // 最小表示回数
  MIN_CTR_GAP: 0.02,          // 最小CTRギャップ（2%）
  TOP_QUERY_LIMIT: 10,        // 上位クエリ数
  
  // UX分析の閾値
  LOW_SCROLL_DEPTH: 30,       // 低スクロール深度（%）
  HIGH_DEAD_CLICKS: 5,        // 高デッドクリック数
  HIGH_RAGE_CLICKS: 3,        // 高レイジクリック数
  HIGH_QUICK_BACKS: 5,        // 高クイックバック数
  
  // キーワード戦略の閾値
  HIGH_SEARCH_VOLUME: 1000,   // 高検索ボリューム
  MEDIUM_SEARCH_VOLUME: 100,  // 中検索ボリューム
  GOOD_POSITION: 10,          // 良好な順位
  
  // 競合分析の閾値
  HIGH_WINNABLE_SCORE: 70,    // 高勝算度
  MEDIUM_WINNABLE_SCORE: 40,  // 中勝算度
  
  // 期待CTR（順位別）
  EXPECTED_CTR: {
    1: 0.316,   // 31.6%
    2: 0.158,   // 15.8%
    3: 0.103,   // 10.3%
    4: 0.076,   // 7.6%
    5: 0.057,   // 5.7%
    6: 0.044,   // 4.4%
    7: 0.035,   // 3.5%
    8: 0.029,   // 2.9%
    9: 0.024,   // 2.4%
    10: 0.020   // 2.0%
  }
};


// ============================================
// メイン関数: クエリベース提案生成（機会損失スコア対応）
// ============================================

/**
 * クエリベース提案を生成
 * @param {string} pageUrl - 対象ページURL（相対パス）
 * @returns {Object} 提案結果オブジェクト
 */
function generateQueryBasedSuggestions(pageUrl) {
  Logger.log(`=== クエリベース提案生成開始: ${pageUrl} ===`);
  
  try {
    // 1. クエリデータを取得
    const queryData = getQueryDataForPage(pageUrl);
    
    if (!queryData || queryData.length === 0) {
      Logger.log('クエリデータが見つかりません');
      return {
        success: false,
        error: 'クエリデータが見つかりません',
        pageUrl: pageUrl
      };
    }
    
    Logger.log(`取得クエリ数: ${queryData.length}`);
    
    // 2. CTRギャップ分析
    const analyzedQueries = analyzeQueryCTRGap(queryData);
    
    // 3. 上位クエリを抽出（改善余地が大きい順）
    const topQueries = analyzedQueries
      .filter(q => q.impressions >= SUGGESTION_GENERATOR_CONFIG.MIN_IMPRESSIONS)
      .sort((a, b) => b.improvementPotential - a.improvementPotential)
      .slice(0, SUGGESTION_GENERATOR_CONFIG.TOP_QUERY_LIMIT);
    
    Logger.log(`分析対象クエリ数: ${topQueries.length}`);
    
    // 4. ページ情報を取得
    const pageInfo = getPageInfo(pageUrl);
    
    // 5. Claude APIで提案生成
    const suggestion = callClaudeForQuerySuggestion(pageUrl, pageInfo, topQueries);
    
    return {
      success: true,
      pageUrl: pageUrl,
      pageTitle: pageInfo.title || '',
      queryCount: topQueries.length,
      topQueries: topQueries,
      suggestion: suggestion,
      generatedAt: new Date().toISOString()
    };
    
  } catch (error) {
    Logger.log(`エラー: ${error.message}`);
    return {
      success: false,
      error: error.message,
      pageUrl: pageUrl
    };
  }
}


// ============================================
// メイン関数: UXベース提案生成（パフォーマンススコア対応）
// ============================================

/**
 * UXベース提案を生成
 * @param {string} pageUrl - 対象ページURL
 * @returns {Object} 提案結果オブジェクト
 */
function generateUXBasedSuggestions(pageUrl) {
  Logger.log(`=== UXベース提案生成開始: ${pageUrl} ===`);
  
  try {
    // 1. UXデータを取得
    const uxData = getUXDataForPage(pageUrl);
    
    if (!uxData) {
      Logger.log('UXデータが見つかりません');
      return {
        success: false,
        error: 'UXデータが見つかりません',
        pageUrl: pageUrl
      };
    }
    
    // 2. UX問題を分析
    const uxProblems = analyzeUXProblems(uxData);
    
    // 3. ページ情報を取得
    const pageInfo = getPageInfo(pageUrl);
    
    // 4. Claude APIで提案生成
    const suggestion = callClaudeForUXSuggestion(pageUrl, pageInfo, uxData, uxProblems);
    
    return {
      success: true,
      pageUrl: pageUrl,
      pageTitle: pageInfo.title || '',
      uxData: uxData,
      problems: uxProblems,
      suggestion: suggestion,
      generatedAt: new Date().toISOString()
    };
    
  } catch (error) {
    Logger.log(`エラー: ${error.message}`);
    return {
      success: false,
      error: error.message,
      pageUrl: pageUrl
    };
  }
}


// ============================================
// メイン関数: イベントベース提案生成（ビジネスインパクトスコア対応）
// ============================================

/**
 * イベントベース提案を生成
 * @param {string} pageUrl - 対象ページURL
 * @returns {Object} 提案結果オブジェクト
 */
function generateEventBasedSuggestions(pageUrl) {
  Logger.log(`=== イベントベース提案生成開始: ${pageUrl} ===`);
  
  try {
    // 1. イベントデータを取得
    const eventData = getEventData();
    
    if (!eventData || eventData.length === 0) {
      Logger.log('イベントデータが見つかりません');
      return {
        success: false,
        error: 'イベントデータが見つかりません',
        pageUrl: pageUrl
      };
    }
    
    // 2. ページ情報を取得
    const pageInfo = getPageInfo(pageUrl);
    
    // 3. Claude APIで提案生成
    const suggestion = callClaudeForEventSuggestion(pageUrl, pageInfo, eventData);
    
    return {
      success: true,
      pageUrl: pageUrl,
      pageTitle: pageInfo.title || '',
      eventCount: eventData.length,
      events: eventData,
      suggestion: suggestion,
      generatedAt: new Date().toISOString()
    };
    
  } catch (error) {
    Logger.log(`エラー: ${error.message}`);
    return {
      success: false,
      error: error.message,
      pageUrl: pageUrl
    };
  }
}


// ============================================
// メイン関数: キーワード戦略ベース提案生成（キーワード戦略スコア対応）★NEW
// ============================================

/**
 * キーワード戦略ベース提案を生成
 * @param {string} pageUrl - 対象ページURL
 * @returns {Object} 提案結果オブジェクト
 */
function generateKeywordStrategySuggestions(pageUrl) {
  Logger.log(`=== キーワード戦略ベース提案生成開始: ${pageUrl} ===`);
  
  try {
    // 1. ターゲットKWデータを取得
    const keywordData = getTargetKeywordDataForPage(pageUrl);
    
    if (!keywordData) {
      Logger.log('ターゲットKWデータが見つかりません');
      return {
        success: false,
        error: 'ターゲットKWデータが見つかりません',
        pageUrl: pageUrl
      };
    }
    
    // 2. キーワード戦略を分析
    const keywordAnalysis = analyzeKeywordStrategy(keywordData);
    
    // 3. ページ情報を取得
    const pageInfo = getPageInfo(pageUrl);
    
    // 4. 実クエリデータも取得（ターゲットKWとの比較用）
    const queryData = getQueryDataForPage(pageUrl);
    const topQueries = queryData ? queryData.slice(0, 5) : [];
    
    // 5. Claude APIで提案生成
    const suggestion = callClaudeForKeywordStrategySuggestion(
      pageUrl, 
      pageInfo, 
      keywordData, 
      keywordAnalysis,
      topQueries
    );
    
    return {
      success: true,
      pageUrl: pageUrl,
      pageTitle: pageInfo.title || '',
      targetKeyword: keywordData.target_keyword,
      keywordData: keywordData,
      analysis: keywordAnalysis,
      suggestion: suggestion,
      generatedAt: new Date().toISOString()
    };
    
  } catch (error) {
    Logger.log(`エラー: ${error.message}`);
    return {
      success: false,
      error: error.message,
      pageUrl: pageUrl
    };
  }
}

/**
 * ターゲットKWデータを取得
 * @param {string} pageUrl - ページURL
 * @returns {Object} ターゲットKWデータ
 */
function getTargetKeywordDataForPage(pageUrl) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('ターゲットKW分析');
  
  if (!sheet) {
    return null;
  }
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const urlIdx = headers.indexOf('page_url');
  const keywordIdx = headers.indexOf('target_keyword');
  const gyronPosIdx = headers.indexOf('gyron_position');
  const gscPosIdx = headers.indexOf('gsc_position');
  const volumeIdx = headers.indexOf('search_volume');
  const competitionIdx = headers.indexOf('competition_level');
  const kwScoreIdx = headers.indexOf('kw_score');
  const statusIdx = headers.indexOf('status');
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowUrl = normalizeUrlForSuggestion(String(row[urlIdx] || ''));
    const targetUrl = normalizeUrlForSuggestion(pageUrl);
    
    if (rowUrl === targetUrl || rowUrl.includes(targetUrl) || targetUrl.includes(rowUrl)) {
      return {
        page_url: row[urlIdx] || pageUrl,
        target_keyword: row[keywordIdx] || '',
        gyron_position: parseFloat(row[gyronPosIdx]) || 0,
        gsc_position: parseFloat(row[gscPosIdx]) || 0,
        search_volume: parseInt(row[volumeIdx]) || 0,
        competition_level: row[competitionIdx] || '',
        kw_score: parseFloat(row[kwScoreIdx]) || 0,
        status: row[statusIdx] || ''
      };
    }
  }
  
  return null;
}

/**
 * キーワード戦略を分析
 * @param {Object} keywordData - ターゲットKWデータ
 * @returns {Object} 分析結果
 */
function analyzeKeywordStrategy(keywordData) {
  const analysis = {
    issues: [],
    opportunities: [],
    priority: 'medium'
  };
  
  // 検索ボリューム分析
  if (keywordData.search_volume >= SUGGESTION_GENERATOR_CONFIG.HIGH_SEARCH_VOLUME) {
    analysis.opportunities.push({
      type: 'high_volume',
      description: `検索ボリュームが${keywordData.search_volume}と高い`,
      impact: '上位表示できれば大きなトラフィックが見込める'
    });
    analysis.priority = 'high';
  } else if (keywordData.search_volume < SUGGESTION_GENERATOR_CONFIG.MEDIUM_SEARCH_VOLUME) {
    analysis.issues.push({
      type: 'low_volume',
      description: `検索ボリュームが${keywordData.search_volume}と低い`,
      impact: 'ターゲットKWの見直しを検討'
    });
  }
  
  // 順位分析
  const position = keywordData.gyron_position || keywordData.gsc_position;
  if (position > 0 && position <= SUGGESTION_GENERATOR_CONFIG.GOOD_POSITION) {
    analysis.opportunities.push({
      type: 'good_position',
      description: `現在${position}位と上位表示`,
      impact: 'CTR改善でさらなる成果が期待できる'
    });
  } else if (position > 20) {
    analysis.issues.push({
      type: 'poor_position',
      description: `現在${position}位と順位が低い`,
      impact: 'コンテンツの大幅な改善が必要'
    });
  } else if (position > 10) {
    analysis.opportunities.push({
      type: 'improvement_potential',
      description: `現在${position}位で1ページ目に近い`,
      impact: '少しの改善で1ページ目入りの可能性'
    });
  }
  
  // Gyron vs GSC順位の乖離
  if (keywordData.gyron_position > 0 && keywordData.gsc_position > 0) {
    const diff = Math.abs(keywordData.gyron_position - keywordData.gsc_position);
    if (diff > 10) {
      analysis.issues.push({
        type: 'position_discrepancy',
        description: `Gyron(${keywordData.gyron_position}位)とGSC(${keywordData.gsc_position}位)で大きな乖離`,
        impact: '計測精度の確認が必要'
      });
    }
  }
  
  // 競合レベル分析
  if (keywordData.competition_level === '激戦' || keywordData.competition_level === '難') {
    analysis.issues.push({
      type: 'high_competition',
      description: `競合レベルが「${keywordData.competition_level}」`,
      impact: '差別化戦略が重要'
    });
  } else if (keywordData.competition_level === '易' || keywordData.competition_level === '超狙い目') {
    analysis.opportunities.push({
      type: 'low_competition',
      description: `競合レベルが「${keywordData.competition_level}」`,
      impact: '積極的にリソースを投下すべき'
    });
    analysis.priority = 'high';
  }
  
  return analysis;
}

/**
 * Claude APIでキーワード戦略提案を生成
 */
function callClaudeForKeywordStrategySuggestion(pageUrl, pageInfo, keywordData, analysis, topQueries) {
  const issuesList = analysis.issues.length > 0
    ? analysis.issues.map((issue, i) => `${i + 1}. 【課題】${issue.description}\n   影響: ${issue.impact}`).join('\n')
    : '特に大きな課題はありません';
  
  const opportunitiesList = analysis.opportunities.length > 0
    ? analysis.opportunities.map((opp, i) => `${i + 1}. 【機会】${opp.description}\n   影響: ${opp.impact}`).join('\n')
    : '特に大きな機会はありません';
  
  const queryList = topQueries.length > 0
    ? topQueries.map((q, i) => `${i + 1}. "${q.query}" - 順位${q.position}位, 表示${q.impressions}回`).join('\n')
    : 'データなし';
  
  const prompt = `以下のページのキーワード戦略を分析し、改善提案をしてください。

【対象ページ】
URL: ${pageUrl}
タイトル: ${pageInfo.title || '不明'}

【ターゲットキーワード情報】
- ターゲットKW: ${keywordData.target_keyword}
- Gyron順位: ${keywordData.gyron_position || 'N/A'}位
- GSC順位: ${keywordData.gsc_position || 'N/A'}位
- 検索ボリューム: ${keywordData.search_volume}/月
- 競合レベル: ${keywordData.competition_level || 'N/A'}
- KWスコア: ${keywordData.kw_score || 'N/A'}点

【5軸スコア】
- キーワード戦略スコア: ${pageInfo.keyword_strategy_score || 'N/A'}点
- 総合優先度スコア: ${pageInfo.total_priority_score || 'N/A'}点

【分析結果 - 課題】
${issuesList}

【分析結果 - 機会】
${opportunitiesList}

【実際の流入クエリ TOP5】
${queryList}

上記データを分析し、以下を提案してください：
1. ターゲットKWの妥当性評価（変更すべきか、維持すべきか）
2. ターゲットKWと実クエリのギャップ分析
3. キーワード戦略の改善案（コンテンツ、タイトル、構造）
4. 追加で狙うべきサブキーワード提案
5. 期待される効果（順位改善、トラフィック増加）`;

  return callClaudeAPI(prompt, getSystemPrompt('keyword'));
}


// ============================================
// メイン関数: 競合分析ベース提案生成（競合難易度スコア対応）★NEW
// ============================================

/**
 * 競合分析ベース提案を生成
 * @param {string} pageUrl - 対象ページURL
 * @returns {Object} 提案結果オブジェクト
 */
function generateCompetitorBasedSuggestions(pageUrl) {
  Logger.log(`=== 競合分析ベース提案生成開始: ${pageUrl} ===`);
  
  try {
    // 1. 競合データを取得
    const competitorData = getCompetitorDataForPage(pageUrl);
    
    if (!competitorData) {
      Logger.log('競合データが見つかりません');
      return {
        success: false,
        error: '競合データが見つかりません',
        pageUrl: pageUrl
      };
    }
    
    // 2. 競合状況を分析
    const competitorAnalysis = analyzeCompetitorSituation(competitorData);
    
    // 3. ページ情報を取得
    const pageInfo = getPageInfo(pageUrl);
    
    // 4. 上位サイト情報を取得
    const topSitesData = getTopSitesData(competitorData.keyword);
    
    // 5. Claude APIで提案生成
    const suggestion = callClaudeForCompetitorSuggestion(
      pageUrl, 
      pageInfo, 
      competitorData, 
      competitorAnalysis,
      topSitesData
    );
    
    return {
      success: true,
      pageUrl: pageUrl,
      pageTitle: pageInfo.title || '',
      targetKeyword: competitorData.keyword,
      winnableScore: competitorData.winnable_score,
      competitionLevel: competitorData.competition_level,
      competitorData: competitorData,
      analysis: competitorAnalysis,
      suggestion: suggestion,
      generatedAt: new Date().toISOString()
    };
    
  } catch (error) {
    Logger.log(`エラー: ${error.message}`);
    return {
      success: false,
      error: error.message,
      pageUrl: pageUrl
    };
  }
}

/**
 * 競合状況を分析
 * @param {Object} competitorData - 競合データ
 * @returns {Object} 分析結果
 */
function analyzeCompetitorSituation(competitorData) {
  const analysis = {
    situation: '',
    strategy: '',
    priority: 'medium',
    recommendations: []
  };
  
  const winnableScore = competitorData.winnable_score || 0;
  const level = competitorData.competition_level || '';
  const daDiff = (competitorData.avg_da_top10 || 0) - (competitorData.own_da || 0);
  
  // 勝算度に基づく状況判定
  if (winnableScore >= SUGGESTION_GENERATOR_CONFIG.HIGH_WINNABLE_SCORE) {
    analysis.situation = '有利な競合環境';
    analysis.strategy = '積極攻勢';
    analysis.priority = 'high';
    analysis.recommendations.push({
      type: 'aggressive',
      action: 'リソースを集中投下してシェア拡大',
      reason: `勝算度${winnableScore}点と高く、上位表示の可能性が高い`
    });
  } else if (winnableScore >= SUGGESTION_GENERATOR_CONFIG.MEDIUM_WINNABLE_SCORE) {
    analysis.situation = '互角の競合環境';
    analysis.strategy = '差別化重視';
    analysis.priority = 'medium';
    analysis.recommendations.push({
      type: 'differentiate',
      action: 'コンテンツの差別化で勝負',
      reason: `勝算度${winnableScore}点、工夫次第で上位表示可能`
    });
  } else {
    analysis.situation = '厳しい競合環境';
    analysis.strategy = 'ニッチ戦略または撤退検討';
    analysis.priority = 'low';
    analysis.recommendations.push({
      type: 'niche_or_retreat',
      action: 'ロングテールKWへのシフトを検討',
      reason: `勝算度${winnableScore}点と低く、正面勝負は困難`
    });
  }
  
  // DA差分に基づく追加分析
  if (daDiff > 20) {
    analysis.recommendations.push({
      type: 'da_gap',
      action: '被リンク獲得施策を並行実施',
      reason: `競合平均DAが自社より${daDiff.toFixed(0)}ポイント高い`
    });
  } else if (daDiff < -10) {
    analysis.recommendations.push({
      type: 'da_advantage',
      action: 'DA優位を活かしてコンテンツ勝負',
      reason: `自社DAが競合平均より${Math.abs(daDiff).toFixed(0)}ポイント高い`
    });
  }
  
  // 競合レベルに基づく追加分析
  if (level === '超狙い目' || level === '易') {
    analysis.recommendations.push({
      type: 'opportunity',
      action: '今すぐリライトを実施',
      reason: `競合レベル「${level}」は絶好のチャンス`
    });
  } else if (level === '激戦') {
    analysis.recommendations.push({
      type: 'caution',
      action: 'ROIを慎重に検討',
      reason: `競合レベル「${level}」は大手が多く難易度高`
    });
  }
  
  return analysis;
}

/**
 * 上位サイトデータを取得
 * @param {string} keyword - ターゲットキーワード
 * @returns {Array} 上位サイトデータ配列
 */
function getTopSitesData(keyword) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('競合分析');
  
  if (!sheet) {
    return [];
  }
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const keywordIdx = headers.indexOf('target_keyword');
  
  // 上位サイト情報の列を探す
  const topSites = [];
  for (let rank = 1; rank <= 5; rank++) {
    const urlIdx = headers.indexOf(`rank_${rank}_url`);
    const daIdx = headers.indexOf(`rank_${rank}_da`);
    
    if (urlIdx !== -1) {
      for (let i = 1; i < data.length; i++) {
        if (data[i][keywordIdx] === keyword) {
          topSites.push({
            rank: rank,
            url: data[i][urlIdx] || '',
            da: data[i][daIdx] || 0
          });
          break;
        }
      }
    }
  }
  
  return topSites;
}

/**
 * Claude APIで競合分析提案を生成
 */
function callClaudeForCompetitorSuggestion(pageUrl, pageInfo, competitorData, analysis, topSitesData) {
  const topSitesList = topSitesData.length > 0
    ? topSitesData.map(site => `${site.rank}位: ${site.url} (DA: ${site.da})`).join('\n')
    : 'データなし';
  
  const recommendationsList = analysis.recommendations
    .map((rec, i) => `${i + 1}. 【${rec.type}】${rec.action}\n   理由: ${rec.reason}`)
    .join('\n');
  
  const prompt = `以下のページの競合分析データを基に、SEO戦略を提案してください。

【対象ページ】
URL: ${pageUrl}
タイトル: ${pageInfo.title || '不明'}

【ターゲットキーワード】
${competitorData.keyword}

【競合分析結果】
- 勝算度スコア: ${competitorData.winnable_score}点/100点
- 競合レベル: ${competitorData.competition_level}
- 自社DA: ${competitorData.own_da}
- 競合平均DA: ${competitorData.avg_da_top10}
- DA差分: ${((competitorData.avg_da_top10 || 0) - (competitorData.own_da || 0)).toFixed(1)}

【競合状況】
${analysis.situation}

【推奨戦略】
${analysis.strategy}

【上位5サイト】
${topSitesList}

【5軸スコア】
- 競合難易度スコア: ${pageInfo.competitor_difficulty_score || 'N/A'}点
- 総合優先度スコア: ${pageInfo.total_priority_score || 'N/A'}点

【分析に基づく推奨アクション】
${recommendationsList}

上記データを分析し、以下を提案してください：
1. 競合に勝つための差別化ポイント
2. コンテンツ面での具体的な改善案
3. 上位サイトから学ぶべきポイント
4. 被リンク・DA向上の施策（必要な場合）
5. 優先度と期待効果`;

  return callClaudeAPI(prompt, getSystemPrompt('competitor'));
}


// ============================================
// メイン関数: 統合リライトレポート生成（5軸統合）
// ============================================

/**
 * 統合リライトレポートを生成（5軸完全対応）
 * @param {string} pageUrl - 対象ページURL
 * @returns {Object} レポート結果オブジェクト
 */
function generateIntegratedRewriteReport(pageUrl) {
  Logger.log(`=== 統合リライトレポート生成開始: ${pageUrl} ===`);
  
  try {
    // 1. 全データを収集
    const pageInfo = getPageInfo(pageUrl);
    const queryData = getQueryDataForPage(pageUrl);
    const uxData = getUXDataForPage(pageUrl);
    const keywordData = getTargetKeywordDataForPage(pageUrl);
    const competitorData = getCompetitorDataForPage(pageUrl);
    
    // 2. 分析結果を統合
    const analyzedQueries = queryData ? analyzeQueryCTRGap(queryData).slice(0, 5) : [];
    const uxProblems = uxData ? analyzeUXProblems(uxData) : [];
    const keywordAnalysis = keywordData ? analyzeKeywordStrategy(keywordData) : null;
    const competitorAnalysis = competitorData ? analyzeCompetitorSituation(competitorData) : null;
    
    // 3. Claude APIで統合レポート生成
    const report = callClaudeForIntegratedReport(
      pageUrl, 
      pageInfo, 
      analyzedQueries, 
      uxData, 
      uxProblems,
      keywordData,
      keywordAnalysis,
      competitorData,
      competitorAnalysis
    );
    
    return {
      success: true,
      pageUrl: pageUrl,
      pageTitle: pageInfo.title || '',
      fiveAxisScores: {
        opportunity: pageInfo.opportunity_score,
        performance: pageInfo.performance_score,
        businessImpact: pageInfo.business_impact_score,
        keywordStrategy: pageInfo.keyword_strategy_score,
        competitorDifficulty: pageInfo.competitor_difficulty_score,
        total: pageInfo.total_priority_score
      },
      queryCount: analyzedQueries.length,
      uxProblemsCount: uxProblems.length,
      hasKeywordData: !!keywordData,
      hasCompetitorData: !!competitorData,
      report: report,
      generatedAt: new Date().toISOString()
    };
    
  } catch (error) {
    Logger.log(`エラー: ${error.message}`);
    return {
      success: false,
      error: error.message,
      pageUrl: pageUrl
    };
  }
}

/**
 * Claude APIで統合レポートを生成（5軸完全対応版）
 */
function callClaudeForIntegratedReport(pageUrl, pageInfo, topQueries, uxData, uxProblems, keywordData, keywordAnalysis, competitorData, competitorAnalysis) {
  // クエリセクション
  const querySection = topQueries.length > 0
    ? topQueries.map((q, i) => 
        `${i + 1}. "${q.query}" - 順位${q.position.toFixed(1)}位, 表示${q.impressions}回, CTR${(q.ctr * 100).toFixed(1)}%, ギャップ${(q.ctrGap * 100).toFixed(1)}%`
      ).join('\n')
    : 'データなし';
  
  // UXセクション
  const uxSection = uxData
    ? `スクロール深度: ${uxData.avg_scroll_depth}%, デッドクリック: ${uxData.dead_clicks}回, レイジクリック: ${uxData.rage_clicks}回`
    : 'データなし';
  
  const problemSection = uxProblems.length > 0
    ? uxProblems.map(p => `- ${p.description}`).join('\n')
    : '重大な問題なし';
  
  // キーワードセクション
  const keywordSection = keywordData
    ? `ターゲットKW: ${keywordData.target_keyword}, 順位: ${keywordData.gyron_position || keywordData.gsc_position || 'N/A'}位, 検索Vol: ${keywordData.search_volume}/月`
    : 'データなし';
  
  const keywordIssues = keywordAnalysis && keywordAnalysis.issues.length > 0
    ? keywordAnalysis.issues.map(i => `- ${i.description}`).join('\n')
    : '特になし';
  
  // 競合セクション
  const competitorSection = competitorData
    ? `キーワード: ${competitorData.keyword}, 勝算度: ${competitorData.winnable_score}点, 競合レベル: ${competitorData.competition_level}, DA差: ${((competitorData.avg_da_top10 || 0) - (competitorData.own_da || 0)).toFixed(0)}`
    : 'データなし';
  
  const competitorStrategy = competitorAnalysis
    ? `状況: ${competitorAnalysis.situation}, 推奨戦略: ${competitorAnalysis.strategy}`
    : '分析なし';
  
  const prompt = `以下のページの全5軸データを分析し、優先度付きの統合リライトレポートを作成してください。

【ページ情報】
URL: ${pageUrl}
タイトル: ${pageInfo.title || '不明'}

【5軸スコア】
① 機会損失スコア: ${pageInfo.opportunity_score || 0}点
② パフォーマンススコア: ${pageInfo.performance_score || 0}点
③ ビジネスインパクトスコア: ${pageInfo.business_impact_score || 0}点
④ キーワード戦略スコア: ${pageInfo.keyword_strategy_score || 0}点
⑤ 競合難易度スコア: ${pageInfo.competitor_difficulty_score || 0}点
【総合優先度スコア: ${pageInfo.total_priority_score || 0}点】

━━━━━━━━━━━━━━━━━━━━━━━━
【①機会損失分析（GSCデータ）】
- 平均順位: ${pageInfo.avg_position || 'N/A'}位
- 表示回数: ${pageInfo.impressions || 'N/A'}回
- クリック数: ${pageInfo.clicks || 'N/A'}回
- CTR: ${pageInfo.ctr ? (pageInfo.ctr * 100).toFixed(1) : 'N/A'}%

改善余地の大きいクエリ:
${querySection}

━━━━━━━━━━━━━━━━━━━━━━━━
【②パフォーマンス分析（UXデータ）】
${uxSection}

検出された問題:
${problemSection}

━━━━━━━━━━━━━━━━━━━━━━━━
【④キーワード戦略分析】
${keywordSection}

課題:
${keywordIssues}

━━━━━━━━━━━━━━━━━━━━━━━━
【⑤競合分析】
${competitorSection}
${competitorStrategy}

━━━━━━━━━━━━━━━━━━━━━━━━

上記の5軸データを総合的に分析し、以下の形式でレポートを作成してください：

## 1. エグゼクティブサマリー
（3行以内で最重要ポイントを要約）

## 2. 5軸スコア診断
（各軸の問題点と改善方針を簡潔に）

## 3. 優先度付き改善提案 TOP5
（優先度：高/中/低、具体的なアクション）

## 4. 期待効果の定量化
（CTR改善率、順位改善、クリック増加数など）

## 5. 次のアクション（ToDoリスト）
（今すぐやること、今週中、今月中に分けて）`;

  return callClaudeAPI(prompt, getSystemPrompt('integrated'));
}


// ============================================
// データ取得ヘルパー関数
// ============================================

/**
 * ページのクエリデータを取得
 */
function getQueryDataForPage(pageUrl) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // まずクエリ分析シートから取得を試みる
  let sheet = ss.getSheetByName('クエリ分析');
  if (sheet) {
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const pageUrlIdx = headers.indexOf('page_url');
    const queryIdx = headers.indexOf('query');
    const positionIdx = headers.indexOf('position');
    const clicksIdx = headers.indexOf('clicks');
    const impressionsIdx = headers.indexOf('impressions');
    const ctrIdx = headers.indexOf('ctr');
    
    if (pageUrlIdx !== -1) {
      const queries = [];
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const rowUrl = normalizeUrlForSuggestion(String(row[pageUrlIdx] || ''));
        const targetUrl = normalizeUrlForSuggestion(pageUrl);
        
        if (rowUrl === targetUrl || rowUrl.includes(targetUrl) || targetUrl.includes(rowUrl)) {
          queries.push({
            query: row[queryIdx] || '',
            position: parseFloat(row[positionIdx]) || 0,
            clicks: parseInt(row[clicksIdx]) || 0,
            impressions: parseInt(row[impressionsIdx]) || 0,
            ctr: parseFloat(row[ctrIdx]) || 0
          });
        }
      }
      
      if (queries.length > 0) {
        return queries;
      }
    }
  }
  
  // クエリ分析シートになければGSC_RAWから取得
  sheet = ss.getSheetByName('GSC_RAW');
  if (!sheet) {
    return [];
  }
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const pageUrlIdx = headers.indexOf('page_url');
  const queryIdx = headers.indexOf('query');
  const positionIdx = headers.indexOf('position');
  const clicksIdx = headers.indexOf('clicks');
  const impressionsIdx = headers.indexOf('impressions');
  const ctrIdx = headers.indexOf('ctr');
  
  if (pageUrlIdx === -1) {
    return [];
  }
  
  const queryMap = {};
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowUrl = normalizeUrlForSuggestion(String(row[pageUrlIdx] || ''));
    const targetUrl = normalizeUrlForSuggestion(pageUrl);
    
    if (rowUrl === targetUrl || rowUrl.includes(targetUrl) || targetUrl.includes(rowUrl)) {
      const query = row[queryIdx] || '';
      
      if (!queryMap[query]) {
        queryMap[query] = {
          query: query,
          position: parseFloat(row[positionIdx]) || 0,
          clicks: parseInt(row[clicksIdx]) || 0,
          impressions: parseInt(row[impressionsIdx]) || 0,
          ctr: parseFloat(row[ctrIdx]) || 0
        };
      } else {
        queryMap[query].clicks += parseInt(row[clicksIdx]) || 0;
        queryMap[query].impressions += parseInt(row[impressionsIdx]) || 0;
      }
    }
  }
  
  return Object.values(queryMap);
}

/**
 * CTRギャップを分析
 */
function analyzeQueryCTRGap(queryData) {
  return queryData.map(q => {
    const position = Math.round(q.position);
    const expectedCTR = SUGGESTION_GENERATOR_CONFIG.EXPECTED_CTR[position] || 0.01;
    const actualCTR = q.ctr;
    const ctrGap = expectedCTR - actualCTR;
    
    const improvementPotential = Math.max(0, ctrGap) * q.impressions;
    const expectedClickIncrease = Math.round(ctrGap * q.impressions);
    
    return {
      ...q,
      expectedCTR: expectedCTR,
      ctrGap: ctrGap,
      improvementPotential: improvementPotential,
      expectedClickIncrease: Math.max(0, expectedClickIncrease),
      priority: ctrGap > 0.05 ? '高' : (ctrGap > 0.02 ? '中' : '低')
    };
  });
}

/**
 * ページ情報を取得
 */
function getPageInfo(pageUrl) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('統合データ');
  
  if (!sheet) {
    return { url: pageUrl };
  }
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const urlIdx = headers.indexOf('page_url');
  const titleIdx = headers.indexOf('page_title');
  const opportunityIdx = headers.indexOf('opportunity_score');
  const performanceIdx = headers.indexOf('performance_score');
  const businessIdx = headers.indexOf('business_impact_score');
  const keywordIdx = headers.indexOf('keyword_strategy_score');
  const competitorIdx = headers.indexOf('competitor_difficulty_score');
  const totalIdx = headers.indexOf('total_priority_score');
  const positionIdx = headers.indexOf('avg_position');
  const clicksIdx = headers.indexOf('total_clicks_30d');
  const impressionsIdx = headers.indexOf('total_impressions_30d');
  const ctrIdx = headers.indexOf('avg_ctr');
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowUrl = normalizeUrlForSuggestion(String(row[urlIdx] || ''));
    const targetUrl = normalizeUrlForSuggestion(pageUrl);
    
    if (rowUrl === targetUrl || rowUrl.includes(targetUrl) || targetUrl.includes(rowUrl)) {
      return {
        url: row[urlIdx] || pageUrl,
        title: row[titleIdx] || '',
        opportunity_score: row[opportunityIdx] || 0,
        performance_score: row[performanceIdx] || 0,
        business_impact_score: row[businessIdx] || 0,
        keyword_strategy_score: row[keywordIdx] || 0,
        competitor_difficulty_score: row[competitorIdx] || 0,
        total_priority_score: row[totalIdx] || 0,
        avg_position: row[positionIdx] || 0,
        clicks: row[clicksIdx] || 0,
        impressions: row[impressionsIdx] || 0,
        ctr: row[ctrIdx] || 0
      };
    }
  }
  
  return { url: pageUrl };
}

/**
 * ページのUXデータを取得
 */
function getUXDataForPage(pageUrl) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Clarity_RAWから取得
  let sheet = ss.getSheetByName('Clarity_RAW');
  if (sheet) {
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const urlIdx = headers.indexOf('page_url');
    const scrollIdx = headers.indexOf('avg_scroll_depth');
    const deadClicksIdx = headers.indexOf('dead_clicks');
    const rageClicksIdx = headers.indexOf('rage_clicks');
    const quickBacksIdx = headers.indexOf('quick_backs');
    const sessionsIdx = headers.indexOf('sessions');
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowUrl = normalizeUrlForSuggestion(String(row[urlIdx] || ''));
      const targetUrl = normalizeUrlForSuggestion(pageUrl);
      
      if (rowUrl === targetUrl || rowUrl.includes(targetUrl) || targetUrl.includes(rowUrl)) {
        return {
          url: row[urlIdx],
          avg_scroll_depth: parseFloat(row[scrollIdx]) || 0,
          dead_clicks: parseInt(row[deadClicksIdx]) || 0,
          rage_clicks: parseInt(row[rageClicksIdx]) || 0,
          quick_backs: parseInt(row[quickBacksIdx]) || 0,
          sessions: parseInt(row[sessionsIdx]) || 0
        };
      }
    }
  }
  
  // 統合データからも取得を試みる
  sheet = ss.getSheetByName('統合データ');
  if (sheet) {
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const urlIdx = headers.indexOf('page_url');
    const scrollIdx = headers.indexOf('clarity_avg_scroll_depth');
    const deadClicksIdx = headers.indexOf('clarity_dead_clicks');
    const rageClicksIdx = headers.indexOf('clarity_rage_clicks');
    const quickBacksIdx = headers.indexOf('clarity_quick_backs');
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowUrl = normalizeUrlForSuggestion(String(row[urlIdx] || ''));
      const targetUrl = normalizeUrlForSuggestion(pageUrl);
      
      if (rowUrl === targetUrl || rowUrl.includes(targetUrl) || targetUrl.includes(rowUrl)) {
        return {
          url: row[urlIdx],
          avg_scroll_depth: parseFloat(row[scrollIdx]) || 0,
          dead_clicks: parseInt(row[deadClicksIdx]) || 0,
          rage_clicks: parseInt(row[rageClicksIdx]) || 0,
          quick_backs: parseInt(row[quickBacksIdx]) || 0,
          sessions: 0
        };
      }
    }
  }
  
  return null;
}

/**
 * UX問題を分析
 */
function analyzeUXProblems(uxData) {
  const problems = [];
  
  if (uxData.avg_scroll_depth < SUGGESTION_GENERATOR_CONFIG.LOW_SCROLL_DEPTH) {
    problems.push({
      type: 'scroll_depth',
      severity: 'high',
      value: uxData.avg_scroll_depth,
      threshold: SUGGESTION_GENERATOR_CONFIG.LOW_SCROLL_DEPTH,
      description: `スクロール深度が${uxData.avg_scroll_depth}%と非常に浅い`,
      impact: 'ファーストビューで離脱している可能性が高い'
    });
  } else if (uxData.avg_scroll_depth < 50) {
    problems.push({
      type: 'scroll_depth',
      severity: 'medium',
      value: uxData.avg_scroll_depth,
      threshold: 50,
      description: `スクロール深度が${uxData.avg_scroll_depth}%で改善余地あり`,
      impact: 'コンテンツ後半が読まれていない'
    });
  }
  
  if (uxData.dead_clicks >= SUGGESTION_GENERATOR_CONFIG.HIGH_DEAD_CLICKS) {
    problems.push({
      type: 'dead_clicks',
      severity: uxData.dead_clicks >= 10 ? 'high' : 'medium',
      value: uxData.dead_clicks,
      threshold: SUGGESTION_GENERATOR_CONFIG.HIGH_DEAD_CLICKS,
      description: `デッドクリックが${uxData.dead_clicks}回発生`,
      impact: 'クリックできると誤解される要素がある'
    });
  }
  
  if (uxData.rage_clicks >= SUGGESTION_GENERATOR_CONFIG.HIGH_RAGE_CLICKS) {
    problems.push({
      type: 'rage_clicks',
      severity: 'high',
      value: uxData.rage_clicks,
      threshold: SUGGESTION_GENERATOR_CONFIG.HIGH_RAGE_CLICKS,
      description: `レイジクリックが${uxData.rage_clicks}回発生`,
      impact: 'ユーザーがフラストレーションを感じている'
    });
  }
  
  if (uxData.quick_backs >= SUGGESTION_GENERATOR_CONFIG.HIGH_QUICK_BACKS) {
    problems.push({
      type: 'quick_backs',
      severity: 'high',
      value: uxData.quick_backs,
      threshold: SUGGESTION_GENERATOR_CONFIG.HIGH_QUICK_BACKS,
      description: `クイックバックが${uxData.quick_backs}回発生`,
      impact: 'コンテンツが検索意図と合っていない可能性'
    });
  }
  
  return problems;
}

/**
 * イベントデータを取得
 */
function getEventData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('イベント分析');
  
  if (!sheet) {
    return [];
  }
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const nameIdx = headers.indexOf('event_name');
  const categoryIdx = headers.indexOf('event_category');
  const countIdx = headers.indexOf('event_count');
  const cvContribIdx = headers.indexOf('cv_contribution');
  const importanceIdx = headers.indexOf('importance');
  
  const events = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[nameIdx]) {
      events.push({
        name: row[nameIdx],
        category: row[categoryIdx] || 'その他',
        count: parseInt(row[countIdx]) || 0,
        cvContribution: parseFloat(row[cvContribIdx]) || 0,
        importance: row[importanceIdx] || '中'
      });
    }
  }
  
  return events.sort((a, b) => b.count - a.count);
}

/**
 * ページの競合データを取得
 */
function getCompetitorDataForPage(pageUrl) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('競合分析');
  
  if (!sheet) {
    return null;
  }
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const urlIdx = headers.indexOf('page_url');
  const keywordIdx = headers.indexOf('target_keyword');
  const winnableIdx = headers.indexOf('winnable_score');
  const levelIdx = headers.indexOf('competition_level');
  const ownDaIdx = headers.indexOf('own_site_da');
  const avgDaIdx = headers.indexOf('avg_da_top10');
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowUrl = normalizeUrlForSuggestion(String(row[urlIdx] || ''));
    const targetUrl = normalizeUrlForSuggestion(pageUrl);
    
    if (rowUrl === targetUrl || rowUrl.includes(targetUrl) || targetUrl.includes(rowUrl)) {
      return {
        keyword: row[keywordIdx] || '',
        winnable_score: parseFloat(row[winnableIdx]) || 0,
        competition_level: row[levelIdx] || '',
        own_da: parseInt(row[ownDaIdx]) || 0,
        avg_da_top10: parseFloat(row[avgDaIdx]) || 0
      };
    }
  }
  
  return null;
}


// ============================================
// Claude API呼び出しヘルパー関数
// ============================================

/**
 * Claude APIでクエリベース提案を生成
 */
function callClaudeForQuerySuggestion(pageUrl, pageInfo, topQueries) {
  const queryList = topQueries.map((q, i) => 
    `${i + 1}. "${q.query}"
   - 順位: ${q.position.toFixed(1)}位
   - 表示回数: ${q.impressions}回
   - クリック: ${q.clicks}回
   - CTR: ${(q.ctr * 100).toFixed(2)}%
   - 期待CTR: ${(q.expectedCTR * 100).toFixed(1)}%
   - CTRギャップ: ${(q.ctrGap * 100).toFixed(2)}%
   - 改善優先度: ${q.priority}
   - 期待クリック増: +${q.expectedClickIncrease}クリック/月`
  ).join('\n\n');
  
  const prompt = `以下のページのGSCクエリデータを分析し、タイトル・メタディスクリプションの改善案を提案してください。

【対象ページ】
URL: ${pageUrl}
タイトル: ${pageInfo.title || '不明'}
現在の平均順位: ${pageInfo.avg_position ? pageInfo.avg_position.toFixed(1) : '不明'}位
月間表示回数: ${pageInfo.impressions || '不明'}回
月間クリック数: ${pageInfo.clicks || '不明'}回

【5軸スコア】
- 機会損失スコア: ${pageInfo.opportunity_score || 'N/A'}点
- キーワード戦略スコア: ${pageInfo.keyword_strategy_score || 'N/A'}点
- 総合優先度スコア: ${pageInfo.total_priority_score || 'N/A'}点

【改善余地の大きいクエリ TOP ${topQueries.length}】
${queryList}

上記データを分析し、以下を提案してください：
1. 最も優先すべきクエリ3つとその理由
2. 具体的なタイトル改善案（現在のタイトルを改善）
3. 具体的なメタディスクリプション案（120文字程度）
4. 期待される効果（CTR改善率、クリック数増加）`;

  return callClaudeAPI(prompt, getSystemPrompt('query'));
}

/**
 * Claude APIでUXベース提案を生成
 */
function callClaudeForUXSuggestion(pageUrl, pageInfo, uxData, problems) {
  const problemList = problems.length > 0 
    ? problems.map((p, i) => 
        `${i + 1}. 【${p.severity === 'high' ? '🔴 重大' : '🟡 注意'}】${p.description}
   - 影響: ${p.impact}`
      ).join('\n\n')
    : 'UXに大きな問題は検出されませんでした。';
  
  const prompt = `以下のページのClarityデータを分析し、UX改善案を提案してください。

【対象ページ】
URL: ${pageUrl}
タイトル: ${pageInfo.title || '不明'}

【5軸スコア】
- パフォーマンススコア: ${pageInfo.performance_score || 'N/A'}点
- 総合優先度スコア: ${pageInfo.total_priority_score || 'N/A'}点

【UX指標】
- スクロール深度: ${uxData.avg_scroll_depth}%
- デッドクリック: ${uxData.dead_clicks}回
- レイジクリック: ${uxData.rage_clicks}回
- クイックバック: ${uxData.quick_backs}回
- セッション数: ${uxData.sessions}回

【検出された問題】
${problemList}

上記データを分析し、以下を提案してください：
1. 最も優先すべきUX問題とその対策
2. ファーストビューの改善案（スクロール深度向上）
3. UI/UXの具体的な改善案
4. 期待される効果（直帰率改善、滞在時間向上など）`;

  return callClaudeAPI(prompt, getSystemPrompt('ux'));
}

/**
 * Claude APIでイベントベース提案を生成
 */
function callClaudeForEventSuggestion(pageUrl, pageInfo, eventData) {
  const eventList = eventData.slice(0, 10).map((e, i) => 
    `${i + 1}. ${e.name}
   - カテゴリ: ${e.category}
   - 発生回数: ${e.count}回
   - CV貢献度: ${e.cvContribution}
   - 重要度: ${e.importance}`
  ).join('\n\n');
  
  const prompt = `以下のGA4イベントデータを分析し、CV導線の改善案を提案してください。

【対象ページ】
URL: ${pageUrl}
タイトル: ${pageInfo.title || '不明'}

【5軸スコア】
- ビジネスインパクトスコア: ${pageInfo.business_impact_score || 'N/A'}点
- 総合優先度スコア: ${pageInfo.total_priority_score || 'N/A'}点

【主要イベント TOP 10】
${eventList}

上記データを分析し、以下を提案してください：
1. CV貢献度の高いイベントへの導線強化案
2. イベント発生を増やすための施策
3. ファネル改善のための具体的なアクション
4. 期待されるCV改善効果`;

  return callClaudeAPI(prompt, getSystemPrompt('seo'));
}


// ============================================
// ユーティリティ関数
// ============================================

/**
 * URLを正規化（SuggestionGenerator専用）
 */
function normalizeUrlForSuggestion(url) {
  if (!url) return '';
  
  let normalized = url
    .replace(/^https?:\/\//, '')
    .replace(/^www\./, '')
    .replace(/^[^\/]+/, '');
  
  if (!normalized.endsWith('/')) {
    normalized += '/';
  }
  
  return normalized.toLowerCase();
}


// ============================================
// テスト関数
// ============================================

/**
 * クエリベース提案生成のテスト
 */
function testQueryBasedSuggestions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('統合データ');
  
  if (!sheet) {
    Logger.log('統合データシートが見つかりません');
    return;
  }
  
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    Logger.log('データがありません');
    return;
  }
  
  const testUrl = data[1][0];
  Logger.log(`テスト対象URL: ${testUrl}`);
  
  const result = generateQueryBasedSuggestions(testUrl);
  Logger.log('=== 結果 ===');
  Logger.log(JSON.stringify(result, null, 2));
  
  if (result.success) {
    Logger.log('\n=== 提案内容 ===');
    Logger.log(result.suggestion);
  }
}

/**
 * UXベース提案生成のテスト
 */
function testUXBasedSuggestions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('統合データ');
  
  if (!sheet) {
    Logger.log('統合データシートが見つかりません');
    return;
  }
  
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    Logger.log('データがありません');
    return;
  }
  
  const testUrl = data[1][0];
  Logger.log(`テスト対象URL: ${testUrl}`);
  
  const result = generateUXBasedSuggestions(testUrl);
  Logger.log('=== 結果 ===');
  Logger.log(JSON.stringify(result, null, 2));
  
  if (result.success) {
    Logger.log('\n=== 提案内容 ===');
    Logger.log(result.suggestion);
  }
}

/**
 * キーワード戦略ベース提案生成のテスト★NEW
 */
function testKeywordStrategySuggestions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('統合データ');
  
  if (!sheet) {
    Logger.log('統合データシートが見つかりません');
    return;
  }
  
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    Logger.log('データがありません');
    return;
  }
  
  const testUrl = data[1][0];
  Logger.log(`テスト対象URL: ${testUrl}`);
  
  const result = generateKeywordStrategySuggestions(testUrl);
  Logger.log('=== 結果 ===');
  Logger.log(JSON.stringify({
    success: result.success,
    pageUrl: result.pageUrl,
    targetKeyword: result.targetKeyword,
    analysis: result.analysis
  }, null, 2));
  
  if (result.success) {
    Logger.log('\n=== 提案内容 ===');
    Logger.log(result.suggestion);
  }
}

/**
 * 競合分析ベース提案生成のテスト★NEW
 */
function testCompetitorBasedSuggestions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('統合データ');
  
  if (!sheet) {
    Logger.log('統合データシートが見つかりません');
    return;
  }
  
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    Logger.log('データがありません');
    return;
  }
  
  const testUrl = data[1][0];
  Logger.log(`テスト対象URL: ${testUrl}`);
  
  const result = generateCompetitorBasedSuggestions(testUrl);
  Logger.log('=== 結果 ===');
  Logger.log(JSON.stringify({
    success: result.success,
    pageUrl: result.pageUrl,
    targetKeyword: result.targetKeyword,
    winnableScore: result.winnableScore,
    competitionLevel: result.competitionLevel,
    analysis: result.analysis
  }, null, 2));
  
  if (result.success) {
    Logger.log('\n=== 提案内容 ===');
    Logger.log(result.suggestion);
  }
}

/**
 * 統合リライトレポート生成のテスト
 */
function testIntegratedRewriteReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('統合データ');
  
  if (!sheet) {
    Logger.log('統合データシートが見つかりません');
    return;
  }
  
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    Logger.log('データがありません');
    return;
  }
  
  const headers = data[0];
  const urlIdx = headers.indexOf('page_url');
  const scoreIdx = headers.indexOf('total_priority_score');
  
  let testUrl = data[1][urlIdx];
  let maxScore = 0;
  
  for (let i = 1; i < Math.min(data.length, 20); i++) {
    const score = parseFloat(data[i][scoreIdx]) || 0;
    if (score > maxScore) {
      maxScore = score;
      testUrl = data[i][urlIdx];
    }
  }
  
  Logger.log(`テスト対象URL: ${testUrl} (スコア: ${maxScore})`);
  
  const result = generateIntegratedRewriteReport(testUrl);
  Logger.log('=== 結果 ===');
  Logger.log(JSON.stringify({
    success: result.success,
    pageUrl: result.pageUrl,
    fiveAxisScores: result.fiveAxisScores,
    queryCount: result.queryCount,
    uxProblemsCount: result.uxProblemsCount,
    hasKeywordData: result.hasKeywordData,
    hasCompetitorData: result.hasCompetitorData
  }, null, 2));
  
  if (result.success) {
    Logger.log('\n=== レポート ===');
    Logger.log(result.report);
  }
}

/**
 * 全機能の簡易テスト（5軸完全対応版）★UPDATED
 */
function testAllSuggestionGenerators() {
  Logger.log('========================================');
  Logger.log('SuggestionGenerator 全機能テスト（5軸対応版）');
  Logger.log('========================================\n');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('統合データ');
  
  if (!sheet) {
    Logger.log('❌ 統合データシートが見つかりません');
    return;
  }
  
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    Logger.log('❌ データがありません');
    return;
  }
  
  const testUrl = data[1][0];
  Logger.log(`テスト対象URL: ${testUrl}\n`);
  
  // 1. クエリベース提案（機会損失スコア対応）
  Logger.log('--- 1. クエリベース提案テスト（機会損失スコア） ---');
  try {
    const queryResult = generateQueryBasedSuggestions(testUrl);
    if (queryResult.success) {
      Logger.log(`✅ 成功: ${queryResult.queryCount}件のクエリを分析`);
    } else {
      Logger.log(`⚠️ データなし: ${queryResult.error}`);
    }
  } catch (e) {
    Logger.log(`❌ エラー: ${e.message}`);
  }
  
  // 2. UXベース提案（パフォーマンススコア対応）
  Logger.log('\n--- 2. UXベース提案テスト（パフォーマンススコア） ---');
  try {
    const uxResult = generateUXBasedSuggestions(testUrl);
    if (uxResult.success) {
      Logger.log(`✅ 成功: ${uxResult.problems ? uxResult.problems.length : 0}件の問題を検出`);
    } else {
      Logger.log(`⚠️ データなし: ${uxResult.error}`);
    }
  } catch (e) {
    Logger.log(`❌ エラー: ${e.message}`);
  }
  
  // 3. イベントベース提案（ビジネスインパクトスコア対応）
  Logger.log('\n--- 3. イベントベース提案テスト（ビジネスインパクトスコア） ---');
  try {
    const eventResult = generateEventBasedSuggestions(testUrl);
    if (eventResult.success) {
      Logger.log(`✅ 成功: ${eventResult.eventCount || 0}件のイベントを分析`);
    } else {
      Logger.log(`⚠️ データなし: ${eventResult.error}`);
    }
  } catch (e) {
    Logger.log(`❌ エラー: ${e.message}`);
  }
  
  // 4. キーワード戦略ベース提案（キーワード戦略スコア対応）★NEW
  Logger.log('\n--- 4. キーワード戦略ベース提案テスト（キーワード戦略スコア） ---');
  try {
    const keywordResult = generateKeywordStrategySuggestions(testUrl);
    if (keywordResult.success) {
      Logger.log(`✅ 成功: ターゲットKW「${keywordResult.targetKeyword}」を分析`);
    } else {
      Logger.log(`⚠️ データなし: ${keywordResult.error}`);
    }
  } catch (e) {
    Logger.log(`❌ エラー: ${e.message}`);
  }
  
  // 5. 競合分析ベース提案（競合難易度スコア対応）★NEW
  Logger.log('\n--- 5. 競合分析ベース提案テスト（競合難易度スコア） ---');
  try {
    const competitorResult = generateCompetitorBasedSuggestions(testUrl);
    if (competitorResult.success) {
      Logger.log(`✅ 成功: 勝算度${competitorResult.winnableScore}点, 競合レベル「${competitorResult.competitionLevel}」`);
    } else {
      Logger.log(`⚠️ データなし: ${competitorResult.error}`);
    }
  } catch (e) {
    Logger.log(`❌ エラー: ${e.message}`);
  }
  
  // 6. 統合リライトレポート（5軸統合）
  Logger.log('\n--- 6. 統合リライトレポートテスト（5軸統合） ---');
  try {
    const integratedResult = generateIntegratedRewriteReport(testUrl);
    if (integratedResult.success) {
      Logger.log(`✅ 成功: 5軸スコア取得完了`);
      Logger.log(`   - 機会損失: ${integratedResult.fiveAxisScores.opportunity}点`);
      Logger.log(`   - パフォーマンス: ${integratedResult.fiveAxisScores.performance}点`);
      Logger.log(`   - ビジネスインパクト: ${integratedResult.fiveAxisScores.businessImpact}点`);
      Logger.log(`   - キーワード戦略: ${integratedResult.fiveAxisScores.keywordStrategy}点`);
      Logger.log(`   - 競合難易度: ${integratedResult.fiveAxisScores.competitorDifficulty}点`);
      Logger.log(`   - 総合: ${integratedResult.fiveAxisScores.total}点`);
    } else {
      Logger.log(`⚠️ エラー: ${integratedResult.error}`);
    }
  } catch (e) {
    Logger.log(`❌ エラー: ${e.message}`);
  }
  
  Logger.log('\n========================================');
  Logger.log('テスト完了（5軸すべて対応）');
  Logger.log('========================================');
}

// ============================================
// Day 22追加: GSC-ターゲットKWズレ分析機能
// ============================================

/**
 * GSCクエリとターゲットKWのズレを分析（複数KW対応版）
 * @param {string} pageUrl - ページURL
 * @return {Object} ズレ分析結果
 */
function analyzeGSCTargetKWGap(pageUrl) {
  Logger.log('=== GSC-ターゲットKWズレ分析: ' + pageUrl + ' ===');
  
  try {
    // 1. GSCクエリデータを取得
    const queryData = getQueryDataForPage(pageUrl);
    
    if (!queryData || queryData.length === 0) {
      return {
        success: false,
        error: 'GSCクエリデータが見つかりません',
        pageUrl: pageUrl
      };
    }
    
    // 2. GyronSEO_RAWから該当ページの全ターゲットKWを取得（★改善点）
    const allTargetKWs = getAllTargetKeywordsForPage(pageUrl);
    Logger.log('登録済ターゲットKW: ' + allTargetKWs.length + '件');
    allTargetKWs.forEach(kw => Logger.log('  - ' + kw));
    
    // 3. GSCクエリの重複を除去・不要クエリを除外して表示回数順にソート（★修正）
    const deduplicatedQueries = [];
    const seenQueryNames = new Set();
    
    // 除外するパターン（サイト名、ブランド名など）
    const excludePatterns = [
      'すまほたっぷ',
      'すまほしゅうり',
      'smaho-tap',
      'site:',
      'スマホタップ'
    ];
    
    queryData.forEach(q => {
      const queryLower = q.query.toLowerCase().trim();
      
      // 除外パターンに該当するかチェック
      const shouldExclude = excludePatterns.some(pattern => 
        queryLower.includes(pattern.toLowerCase())
      );
      
      if (shouldExclude) return;
      
      if (!seenQueryNames.has(queryLower)) {
        seenQueryNames.add(queryLower);
        deduplicatedQueries.push(q);
      }
    });
    
    const sortedQueries = deduplicatedQueries.sort((a, b) => b.impressions - a.impressions);
    const top10Queries = sortedQueries.slice(0, 10);
    
    // 4. 各クエリがいずれかのターゲットKWにマッチするかチェック（★改善点）
    const analyzedQueries = top10Queries.map((q, index) => {
      const queryLower = q.query.toLowerCase().trim();
      
      // いずれかのターゲットKWにマッチするか
      const matchedKW = allTargetKWs.find(kw => {
        const kwLower = kw.toLowerCase().trim();
        return queryLower.includes(kwLower) || 
               kwLower.includes(queryLower) ||
               queryLower === kwLower;
      });
      
      const isRegistered = !!matchedKW;
      
      return {
        rank: index + 1,
        query: q.query,
        position: q.position,
        impressions: q.impressions,
        clicks: q.clicks,
        ctr: q.ctr,
        isRegistered: isRegistered,
        matchedKW: matchedKW || null,
        registrationStatus: isRegistered ? '✅ 登録済' : '⚪ 未登録'
      };
    });
    
    // 5. ズレ分析
    const registeredCount = analyzedQueries.filter(q => q.isRegistered).length;
    const unregisteredCount = analyzedQueries.filter(q => !q.isRegistered).length;
    const topQueryIsRegistered = analyzedQueries.length > 0 && analyzedQueries[0].isRegistered;
    
    // 6. 提案を生成（★文言改善：「推奨」→「検討候補」）
    const suggestions = [];
    const unregisteredHighImpact = analyzedQueries.filter(q => !q.isRegistered && q.impressions >= 100);
    
    // 重複除去
    const uniqueQueries = [];
    const seenWords = new Set();
    
    unregisteredHighImpact.forEach(q => {
      const words = q.query.toLowerCase().split(/\s+/).sort().join(' ');
      if (!seenWords.has(words)) {
        seenWords.add(words);
        uniqueQueries.push(q);
      }
    });
    
    uniqueQueries.slice(0, 3).forEach(q => {
      suggestions.push({
        type: 'consider_target_kw',
        priority: q.rank <= 3 ? '要検討' : '参考',
        keyword: q.query,
        reason: 'GSC' + q.rank + '位「' + q.query + '」が未登録（表示' + q.impressions + '回）',
        action: '「' + q.query + '」をターゲットKWに追加を検討'
      });
    });
    
    // 7. ズレレベル判定
    let gapLevel;
    if (registeredCount >= 5 || (topQueryIsRegistered && registeredCount >= 3)) {
      gapLevel = 'なし';
    } else if (registeredCount >= 3 || topQueryIsRegistered) {
      gapLevel = '小';
    } else if (registeredCount >= 1) {
      gapLevel = '中';
    } else {
      gapLevel = '要確認';
    }
    
    return {
      success: true,
      pageUrl: pageUrl,
      targetKeywords: allTargetKWs,
      analyzedQueries: analyzedQueries,
      summary: {
        totalQueries: analyzedQueries.length,
        registeredCount: registeredCount,
        unregisteredCount: unregisteredCount,
        topQueryIsRegistered: topQueryIsRegistered,
        gapLevel: gapLevel
      },
      suggestions: suggestions,
      generatedAt: new Date().toISOString()
    };
    
  } catch (error) {
    Logger.log('エラー: ' + error.message);
    return {
      success: false,
      error: error.message,
      pageUrl: pageUrl
    };
  }
}

/**
 * チャット表示用のGSCズレ分析フォーマット（改善版）
 * @param {string} pageUrl - ページURL
 * @return {string} フォーマットされた分析結果
 */
function getGSCTargetKWGapForChat(pageUrl) {
  const analysis = analyzeGSCTargetKWGap(pageUrl);
  
  if (!analysis.success) {
    return '⚠️ GSCズレ分析エラー: ' + analysis.error;
  }
  
  let text = '\n\n📊 **GSC実績 vs ターゲットKW分析**\n\n';
  
  // 登録されている全ターゲットKWを表示（★改善点）
  if (analysis.targetKeywords && analysis.targetKeywords.length > 0) {
    text += '**登録済ターゲットKW** (' + analysis.targetKeywords.length + '件):\n';
    analysis.targetKeywords.forEach(kw => {
      text += '- ' + kw + '\n';
    });
  } else {
    text += '**登録済ターゲットKW**: なし\n';
  }
  
  text += '\n**ズレレベル**: ' + analysis.summary.gapLevel + '\n\n';
  
  // クエリ一覧テーブル
  text += '| # | GSCクエリ | 順位 | 表示回数 | ステータス |\n';
  text += '|:--|:----------|-----:|---------:|:----------:|\n';
  
  analysis.analyzedQueries.forEach(q => {
    const posDisplay = q.position ? q.position.toFixed(1) + '位' : 'N/A';
    text += '| ' + q.rank + ' | ' + q.query + ' | ' + posDisplay + ' | ' + q.impressions.toLocaleString() + ' | ' + q.registrationStatus + ' |\n';
  });
  
  // サマリー
  text += '\n**サマリー**: ' + analysis.summary.registeredCount + '件登録済 / ' + analysis.summary.unregisteredCount + '件未登録\n';
  
  // 検討候補（★文言改善）
  if (analysis.suggestions.length > 0) {
    text += '\n💡 **検討候補**:\n';
    analysis.suggestions.forEach((s, i) => {
      const icon = s.priority === '要検討' ? '🔶' : '⚪';
      text += (i + 1) + '. ' + icon + ' ' + s.action + '\n';
      text += '   理由: ' + s.reason + '\n';
    });
    text += '\n※ 上記は参考情報です。戦略に応じてご判断ください。\n';
  } else if (analysis.summary.registeredCount > 0) {
    text += '\n✅ ターゲットKWとGSC実績が概ね一致しています。\n';
  }
  
  return text;
}

/**
 * リライト提案時に自動でGSCズレ分析を含める
 * @param {string} pageUrl - ページURL
 * @param {Object} pageInfo - ページ情報
 * @return {string} 提案テキスト（GSCズレ分析含む）
 */
function generateSuggestionWithGSCGap(pageUrl, pageInfo) {
  // 1. 既存の提案を生成
  let suggestion = '';
  
  // 2. GSCズレ分析を追加
  const gscGapText = getGSCTargetKWGapForChat(pageUrl);
  
  // 3. トレンド分析を追加
  let trendText = '';
  const targetKW = pageInfo.target_keyword || '';
  if (typeof applyTrendModifier === 'function') {
    const trendResult = applyTrendModifier(pageUrl, targetKW, 50);
    if (trendResult.trend) {
      trendText = `\n\n📈 **順位トレンド（過去4週間）**\n`;
      trendText += `トレンド: **${trendResult.trendLabel}**\n`;
      trendText += `${trendResult.message}\n`;
      if (trendResult.weeklyRanks) {
        const ranks = trendResult.weeklyRanks.map(w => w.rank || '圏外').join(' → ');
        trendText += `推移: ${ranks}\n`;
      }
    }
  }
  
  return suggestion + gscGapText + trendText;
}

/**
 * GSCズレ分析のテスト
 */
function testGSCTargetKWGap() {
  Logger.log('=== GSCズレ分析テスト ===');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('統合データ');
  
  if (!sheet) {
    Logger.log('統合データシートが見つかりません');
    return;
  }
  
  const data = sheet.getDataRange().getValues();
  const testUrl = data[1][0]; // 最初のページ
  
  Logger.log(`テスト対象URL: ${testUrl}\n`);
  
  // 分析実行
  const result = analyzeGSCTargetKWGap(testUrl);
  
  if (result.success) {
    Logger.log('ターゲットKW: ' + result.targetKeyword);
    Logger.log('ズレレベル: ' + result.summary.gapLevel);
    Logger.log('登録済: ' + result.summary.registeredCount);
    Logger.log('未登録: ' + result.summary.unregisteredCount);
    
    Logger.log('\n--- 提案 ---');
    result.suggestions.forEach(s => {
      Logger.log(`[${s.priority}] ${s.action}`);
    });
    
    Logger.log('\n--- チャット表示形式 ---');
    Logger.log(getGSCTargetKWGapForChat(testUrl));
  } else {
    Logger.log('エラー: ' + result.error);
  }
  
  Logger.log('\n=== テスト完了 ===');
}

/**
 * GyronSEO_RAWから該当ページの全ターゲットKWを取得（修正版）
 * @param {string} pageUrl - ページURL（パス形式: /ipad-mini-cheap-buy-methods）
 * @return {Array} ターゲットKW配列
 */
function getAllTargetKeywordsForPage(pageUrl) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('GyronSEO_RAW');
    
    if (!sheet) {
      Logger.log('GyronSEO_RAWシートが見つかりません');
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    const keywords = [];
    
    // 入力URLを正規化（パス部分のみ抽出）
    let targetPath = pageUrl.toLowerCase().trim();
    
    // フルURLからパスを抽出
    if (targetPath.includes('://')) {
      try {
        const urlObj = new URL(targetPath);
        targetPath = urlObj.pathname;
      } catch (e) {
        const match = targetPath.match(/https?:\/\/[^\/]+(\/.*)/);
        if (match) targetPath = match[1];
      }
    }
    
    // 先頭・末尾のスラッシュを除去して比較用に統一
    targetPath = targetPath.replace(/^\/|\/$/g, '');
    
    Logger.log('検索対象パス: ' + targetPath);
    
    // A列: キーワード、B列: URL
    for (let i = 1; i < data.length; i++) {
      const kw = data[i][0];
      const url = (data[i][1] || '').toString().trim();
      
      // URLが空の場合はスキップ（圏外KW）
      if (!url) continue;
      
      // URLからパス部分を抽出
      let rowPath = '';
      if (url.includes('://')) {
        try {
          const urlObj = new URL(url);
          rowPath = urlObj.pathname;
        } catch (e) {
          const match = url.match(/https?:\/\/[^\/]+(\/.*)/);
          if (match) rowPath = match[1];
        }
      } else {
        rowPath = url;
      }
      
      // 先頭・末尾のスラッシュを除去
      rowPath = rowPath.replace(/^\/|\/$/g, '').toLowerCase();
      
      // 完全一致でマッチング
      if (rowPath === targetPath) {
        if (kw && !keywords.includes(kw)) {
          keywords.push(kw);
        }
      }
    }
    
    Logger.log('該当ページのターゲットKW: ' + keywords.length + '件');
    
    return keywords;
    
  } catch (error) {
    Logger.log('ターゲットKW取得エラー: ' + error.message);
    return [];
  }
}

function checkGyronRAWStructure() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('GyronSEO_RAW');
  
  if (!sheet) {
    Logger.log('GyronSEO_RAWシートが見つかりません');
    return;
  }
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  Logger.log('=== GyronSEO_RAW構造 ===');
  Logger.log('列数: ' + headers.length);
  Logger.log('行数: ' + data.length);
  
  Logger.log('\n--- 先頭10列のヘッダー ---');
  for (let i = 0; i < Math.min(10, headers.length); i++) {
    Logger.log('列' + (i+1) + ': ' + headers[i]);
  }
  
  Logger.log('\n--- 最初の3行のデータ ---');
  for (let i = 1; i <= Math.min(3, data.length - 1); i++) {
    Logger.log('行' + i + ': ' + data[i].slice(0, 5).join(' | '));
  }
}