/**
 * スコアリング関数群
 * - 機会損失スコア
 * - パフォーマンススコア
 * - ビジネスインパクトスコア
 * - 総合優先度スコア
 * - AI自動提案生成
 * - 週次自動分析・メール送信
 * - AIO最適化提案
 * 
 * 更新日: 2025/12/02
 * バージョン: 2.2（ターゲットKW追加・順位別リライト戦略実装）
 * 
 * 修正内容:
 * - buildSystemPromptWithSiteInfo(): 現在の日付情報を追加、順位別リスク管理原則を追加
 * - buildSuggestionPrompt(): page_title追加、ターゲットKW追加、順位別警告・制約を追加
 * - buildAIOSystemPrompt(): 現在の日付情報を追加
 * - 順位別リライト戦略:
 *   - 1位: リライト非推奨（現状維持）
 *   - 2-5位: 低リスク施策のみ（タイトル変更禁止）
 *   - 6-10位: 積極的改善OK
 *   - 11位以下: 大幅リライトOK
 */

/**
 * 全ページのスコアを計算して統合データシートに書き込む
 */
function calculateScores() {
  Logger.log('=== スコアリング開始 ===');
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('統合データ');
    
    if (!sheet) {
      throw new Error('統合データシートが見つかりません');
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // 列インデックスを取得
    const indexes = getColumnIndexes(headers);
    Logger.log('列インデックス取得完了');
    
    const {
      urlIndex,
      positionIndex,
      clicksIndex,
      impressionsIndex,
      ctrIndex,
      pageViewsIndex,
      bounceRateIndex,
      sessionDurationIndex,
      conversionsIndex,
      opportunityIndex,
      performanceIndex,
      businessImpactIndex,
      totalScoreIndex
    } = indexes;
    
    Logger.log(`opportunity_score列: ${opportunityIndex + 1}番目`);
    Logger.log(`total_priority_score列: ${totalScoreIndex + 1}番目`);
    
    // サイト平均値を計算
    const siteAverages = calculateSiteAverages(data, indexes);
    Logger.log(`サイト平均値: 直帰率=${siteAverages.avgBounceRate.toFixed(2)}%, セッション時間=${siteAverages.avgSessionDuration.toFixed(2)}秒`);
    
    // 各ページのスコアを計算
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const pageUrl = String(row[urlIndex] || '').trim();
      
      if (!pageUrl) continue;
      
      // 各スコアを計算
      const opportunityScore = calculateOpportunityScore(row, indexes);
      const performanceScore = calculatePerformanceScore(row, indexes, siteAverages);
      const businessImpactScore = calculateBusinessImpactScore(row, indexes);
      const totalScore = calculateTotalPriorityScore(
        opportunityScore,
        performanceScore,
        businessImpactScore
      );
      
      // スコアを書き込み
      sheet.getRange(i + 1, opportunityIndex + 1).setValue(opportunityScore);
      sheet.getRange(i + 1, performanceIndex + 1).setValue(performanceScore);
      sheet.getRange(i + 1, businessImpactIndex + 1).setValue(businessImpactScore);
      sheet.getRange(i + 1, totalScoreIndex + 1).setValue(totalScore);
    }
    
    Logger.log(`スコア更新完了: ${opportunityIndex + 1}列目から4列分`);
    
    // 優先度上位ページを表示
    const topPages = getTopPriorityPages(5);
    Logger.log('=== 優先度上位5ページ ===');
    topPages.forEach((page, index) => {
      Logger.log(`${index + 1}位: ${page.url} (スコア: ${page.score}点)`);
    });
    
    Logger.log('=== スコアリング完了 ===');
    
  } catch (error) {
    Logger.log(`エラー: ${error.message}`);
    throw error;
  }
}

/**
 * 列インデックスを取得
 */
function getColumnIndexes(headers) {
  return {
    urlIndex: headers.indexOf('page_url'),
    positionIndex: headers.indexOf('avg_position'),
    clicksIndex: headers.indexOf('total_clicks_30d'),
    impressionsIndex: headers.indexOf('total_impressions_30d'),
    ctrIndex: headers.indexOf('avg_ctr'),
    pageViewsIndex: headers.indexOf('avg_page_views_30d'),
    bounceRateIndex: headers.indexOf('bounce_rate'),
    sessionDurationIndex: headers.indexOf('avg_session_duration'),
    conversionsIndex: headers.indexOf('conversions_30d'),
    opportunityIndex: headers.indexOf('opportunity_score'),
    performanceIndex: headers.indexOf('performance_score'),
    businessImpactIndex: headers.indexOf('business_impact_score'),
    totalScoreIndex: headers.indexOf('total_priority_score')
  };
}

/**
 * サイト平均値を計算
 */
function calculateSiteAverages(data, indexes) {
  let totalBounceRate = 0;
  let totalSessionDuration = 0;
  let count = 0;
  
  for (let i = 1; i < data.length; i++) {
    const bounceRate = parseFloat(data[i][indexes.bounceRateIndex]) || 0;
    const sessionDuration = parseFloat(data[i][indexes.sessionDurationIndex]) || 0;
    
    if (bounceRate > 0 || sessionDuration > 0) {
      totalBounceRate += bounceRate;
      totalSessionDuration += sessionDuration;
      count++;
    }
  }
  
  return {
    avgBounceRate: count > 0 ? totalBounceRate / count : 0,
    avgSessionDuration: count > 0 ? totalSessionDuration / count : 0
  };
}

/**
 * 機会損失スコア（0-100点）
 * ★修正: gyron_positionを考慮（上位表示中のページは優先度を下げる）
 */
function calculateOpportunityScore(row, indexes) {
  const position = parseFloat(row[indexes.positionIndex]) || 100;
  const impressions = parseFloat(row[indexes.impressionsIndex]) || 0;
  const ctr = parseFloat(row[indexes.ctrIndex]) || 0;
  
  // gyron_position（ターゲットKW順位）を取得
  var headers = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('統合データ').getRange(1, 1, 1, 100).getValues()[0];
  var gyronPositionIndex = headers.indexOf('gyron_position');
  var gyronPosition = gyronPositionIndex >= 0 ? parseFloat(row[gyronPositionIndex]) || 0 : 0;
  
  // ★重要: ターゲットKWで1-3位の場合は優先度を大幅に下げる
  if (gyronPosition >= 1 && gyronPosition <= 3) {
    return 10; // 最低スコア（リライト非推奨）
  }
  
  // ★重要: ターゲットKWで4-5位の場合も優先度を下げる
  if (gyronPosition >= 4 && gyronPosition <= 5) {
    return 30; // 低スコア（慎重な改善のみ）
  }
  
  // 順位スコア（40%）- 6位以下のページのみ通常計算
  let positionScore = 0;
  if (gyronPosition >= 6 && gyronPosition <= 10) {
    positionScore = 100; // TOP5を狙える位置
  } else if (gyronPosition >= 11 && gyronPosition <= 20) {
    positionScore = 80; // 改善の余地大
  } else if (gyronPosition >= 21 && gyronPosition <= 30) {
    positionScore = 60; // 大幅改善が必要
  } else if (gyronPosition > 30 || gyronPosition === 0) {
    // gyron_positionがない場合はavg_positionで判定
    if (position >= 4 && position <= 7) {
      positionScore = 100;
    } else if (position >= 8 && position <= 10) {
      positionScore = 80;
    } else if (position >= 11 && position <= 20) {
      positionScore = 50;
    } else if (position >= 21 && position <= 30) {
      positionScore = 30;
    }
  }
  
  // 表示回数スコア（30%）
  let impressionScore = 0;
  if (impressions >= 1000) {
    impressionScore = 100;
  } else if (impressions >= 500) {
    impressionScore = 70;
  } else if (impressions >= 100) {
    impressionScore = 40;
  } else if (impressions >= 10) {
    impressionScore = 20;
  }
  
  // CTRギャップスコア（30%）
  const expectedCTR = getExpectedCTR(position);
  const ctrGap = expectedCTR - ctr;
  let ctrGapScore = 0;
  
  if (ctrGap >= 0.50) {
    ctrGapScore = 100;
  } else if (ctrGap >= 0.30) {
    ctrGapScore = 70;
  } else if (ctrGap >= 0.10) {
    ctrGapScore = 40;
  }
  
  const totalScore = (positionScore * 0.40) + (impressionScore * 0.30) + (ctrGapScore * 0.30);
  return Math.round(totalScore);
}

/**
 * 順位別の期待CTR（業界平均）
 */
function getExpectedCTR(position) {
  const ctrTable = {
    1: 0.316, 2: 0.158, 3: 0.100, 4: 0.073, 5: 0.057,
    6: 0.045, 7: 0.037, 8: 0.031, 9: 0.026, 10: 0.023
  };
  
  if (position <= 10) {
    return ctrTable[Math.round(position)] || 0.020;
  } else if (position <= 20) {
    return 0.015;
  } else {
    return 0.005;
  }
}

/**
 * パフォーマンススコア（0-100点）
 */
function calculatePerformanceScore(row, indexes, siteAverages) {
  const bounceRate = parseFloat(row[indexes.bounceRateIndex]) || 0;
  const sessionDuration = parseFloat(row[indexes.sessionDurationIndex]) || 0;
  
  // 直帰率スコア（50%）
  const bounceRateDiff = bounceRate - siteAverages.avgBounceRate;
  let bounceRateScore = 0;
  
  if (bounceRateDiff >= 30) {
    bounceRateScore = 100;
  } else if (bounceRateDiff >= 20) {
    bounceRateScore = 70;
  } else if (bounceRateDiff >= 10) {
    bounceRateScore = 40;
  }
  
  // 滞在時間スコア（50%）
  const sessionDurationRatio = siteAverages.avgSessionDuration > 0 
    ? sessionDuration / siteAverages.avgSessionDuration 
    : 1;
  let sessionDurationScore = 0;
  
  if (sessionDurationRatio <= 0.50) {
    sessionDurationScore = 100;
  } else if (sessionDurationRatio <= 0.70) {
    sessionDurationScore = 70;
  } else if (sessionDurationRatio <= 0.90) {
    sessionDurationScore = 40;
  }
  
  const totalScore = (bounceRateScore * 0.50) + (sessionDurationScore * 0.50);
  return Math.round(totalScore);
}

/**
 * ビジネスインパクトスコア（0-100点）
 */
function calculateBusinessImpactScore(row, indexes) {
  const pageViews = parseFloat(row[indexes.pageViewsIndex]) || 0;
  const conversions = parseFloat(row[indexes.conversionsIndex]) || 0;
  
  // トラフィックスコア（40%）
  let trafficScore = 0;
  if (pageViews >= 1000) {
    trafficScore = 100;
  } else if (pageViews >= 500) {
    trafficScore = 80;
  } else if (pageViews >= 100) {
    trafficScore = 50;
  } else if (pageViews >= 10) {
    trafficScore = 20;
  }
  
  // コンバージョンスコア（60%）
  let conversionScore = 20; // デフォルト（トップファネル）
  
  if (conversions > 0) {
    conversionScore = 100; // 直接CV
  } else {
    const pageUrl = String(row[indexes.urlIndex] || '').toLowerCase();
    
    if (pageUrl.includes('comparison') || 
        pageUrl.includes('review') || 
        pageUrl.includes('best') ||
        pageUrl.includes('recommend')) {
      conversionScore = 70; // CVに近い（比較・検討コンテンツ）
    } else if (pageUrl.includes('guide') || 
               pageUrl.includes('how-to') ||
               pageUrl.includes('tips')) {
      conversionScore = 40; // 中間ページ
    }
  }
  
  const totalScore = (trafficScore * 0.40) + (conversionScore * 0.60);
  return Math.round(totalScore);
}

/**
 * 総合優先度スコア（0-100点）
 */
function calculateTotalPriorityScore(opportunityScore, performanceScore, businessImpactScore) {
  const totalScore = (opportunityScore * 0.33) + 
                     (performanceScore * 0.33) + 
                     (businessImpactScore * 0.34);
  return Math.round(totalScore);
}

/**
 * 優先度上位ページを取得（フィルター付き）
 */
function getTopPriorityPagesFiltered(limit) {
  limit = limit || 10;
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('統合データ');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  
  var urlIndex = headers.indexOf('page_url');
  var titleIndex = headers.indexOf('page_title');
  var totalScoreIndex = headers.indexOf('total_priority_score');
  var opportunityIndex = headers.indexOf('opportunity_score');
  var performanceIndex = headers.indexOf('performance_score');
  var businessImpactIndex = headers.indexOf('business_impact_score');
  var targetKWIndex = headers.indexOf('target_keyword');
  var gyronPositionIndex = headers.indexOf('gyron_position');
  
  var pages = [];
  
  for (var i = 1; i < data.length; i++) {
    var url = String(data[i][urlIndex] || '').trim();
    var title = String(data[i][titleIndex] || '').trim();
    var totalScore = parseFloat(data[i][totalScoreIndex]) || 0;
    var opportunityScore = parseFloat(data[i][opportunityIndex]) || 0;
    var performanceScore = parseFloat(data[i][performanceIndex]) || 0;
    var businessImpactScore = parseFloat(data[i][businessImpactIndex]) || 0;
    var targetKW = targetKWIndex >= 0 ? String(data[i][targetKWIndex] || '').trim() : '';
    var gyronPosition = gyronPositionIndex >= 0 ? parseFloat(data[i][gyronPositionIndex]) || null : null;
    
    if (url && totalScore > 0) {
      pages.push({
        url: url,
        title: title,
        score: totalScore,
        totalScore: totalScore,
        opportunityScore: opportunityScore,
        performanceScore: performanceScore,
        businessImpactScore: businessImpactScore,
        targetKeyword: targetKW,
        gyronPosition: gyronPosition
      });
    }
  }
  
  // スコア降順でソート
  pages.sort(function(a, b) { return b.score - a.score; });
  
  return pages.slice(0, limit);
}

/**
 * 優先度上位ページを取得
 */
function getTopPriorityPages(limit = 10) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('統合データ');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const urlIndex = headers.indexOf('page_url');
  const titleIndex = headers.indexOf('page_title');
  const scoreIndex = headers.indexOf('total_priority_score');
  const targetKWIndex = headers.indexOf('target_keyword');
  const gyronPositionIndex = headers.indexOf('gyron_position');
  
  const pages = [];
  
  for (let i = 1; i < data.length; i++) {
    const url = String(data[i][urlIndex] || '').trim();
    const title = String(data[i][titleIndex] || '').trim();
    const score = parseFloat(data[i][scoreIndex]) || 0;
    const targetKW = targetKWIndex >= 0 ? String(data[i][targetKWIndex] || '').trim() : '';
    const gyronPosition = gyronPositionIndex >= 0 ? parseFloat(data[i][gyronPositionIndex]) || null : null;
    
    if (url) {
      pages.push({ 
        url, 
        title, 
        score, 
        rowIndex: i,
        targetKeyword: targetKW,
        gyronPosition: gyronPosition
      });
    }
  }
  
  // スコア降順でソート
  pages.sort((a, b) => b.score - a.score);
  
  return pages.slice(0, limit);
}

/**
 * AIリライト提案を生成
 */
function generateRewriteSuggestions(pageUrl) {
  Logger.log('=== AI提案生成開始: ' + pageUrl + ' ===');
  
  try {
    // ページデータを取得
    var pageData = getPageDataForSuggestion(pageUrl);
    
    if (!pageData) {
      return { success: false, suggestion: 'ページデータが見つかりません' };
    }
    
    // Claude APIで提案生成
    var suggestion = callClaudeForSuggestion(pageData);

    // ★ボタンを追加（フェーズ1追加）
    suggestion = addSuggestionButtons(suggestion, pageUrl);
    
    Logger.log('=== AI提案生成完了 ===');
    return { success: true, suggestion: suggestion };
    
  } catch (error) {
    Logger.log('エラー: ' + error.message);
    return { success: false, suggestion: 'エラー: ' + error.message };
  }
}

/**
 * ページデータを取得
 */
function getPageDataForSuggestion(pageUrl) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('統合データ');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const urlIndex = headers.indexOf('page_url');
  
  for (let i = 1; i < data.length; i++) {
    const url = String(data[i][urlIndex] || '').trim();
    
    if (url === pageUrl) {
      return extractPageDataFromRow(data[i], headers);
    }
  }
  
  return null;
}

/**
 * 行データからページ情報を抽出
 */
function extractPageDataFromRow(row, headers) {
  const data = {};
  
  headers.forEach((header, index) => {
    data[header] = row[index];
  });
  
  return data;
}

/**
 * Claude APIでリライト提案を生成
 */
function callClaudeForSuggestion(pageData) {
  const siteInfo = getSiteInfoFromSettings();
  const systemPrompt = buildSystemPromptWithSiteInfo(siteInfo);
  const userPrompt = buildSuggestionPrompt(pageData);
  
  // Claude API呼び出し（ClaudeAPI.gsのcallClaudeAPI使用）
  const suggestion = callClaudeAPI(userPrompt, systemPrompt);
  
  return suggestion;
}

/**
 * サイト情報を取得
 */
function getSiteInfoFromSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName('設定・マスタ');
  
  if (!settingsSheet) {
    return {
      siteName: '不明',
      siteType: '不明',
      siteGenre: '不明'
    };
  }
  
  const data = settingsSheet.getDataRange().getValues();
  const settings = {};
  
  for (let i = 1; i < data.length; i++) {
    const key = data[i][0];
    const value = data[i][1];
    settings[key] = value;
  }
  
  return {
    siteName: settings['SITE_NAME'] || '不明',
    siteType: settings['SITE_TYPE'] || '不明',
    siteGenre: settings['SITE_GENRE'] || '不明'
  };
}

/**
 * システムプロンプト構築
 * ★v2.2: 現在の日付情報を追加、順位別リスク管理原則を追加
 */
function buildSystemPromptWithSiteInfo(siteInfo) {
  // 現在の日付情報を取得
  var today = new Date();
  var currentYear = today.getFullYear();
  var currentDate = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy年MM月dd日');
  
  return `あなたはSEOとコンテンツマーケティングの専門家です。

【重要：現在の日付】
今日の日付は${currentDate}です。
年号を含む提案をする場合は、必ず${currentYear}年を使用してください。
「2024年最新」などの古い年号は絶対に使用しないでください。

【サイト情報】
- サイト名: ${siteInfo.siteName}
- サイトタイプ: ${siteInfo.siteType}
- ジャンル: ${siteInfo.siteGenre}

【重要：順位別リライト戦略】
リライトは必ずしも順位が上がるとは限らず、下がるリスクもあります。
以下の順位別戦略を必ず遵守してください：

■ 1位の場合（リライト非推奨）
- 現状維持を最優先
- タイトル変更は絶対禁止
- 大幅な構造変更も禁止
- 推奨：内部リンク追加、関連コンテンツ新規作成、表示速度改善のみ

■ 2-5位の場合（低リスク施策のみ）
- タイトル変更は禁止（順位下落リスク大）
- 推奨：メタディスクリプション改善、コンテンツ追記、内部リンク最適化、画像追加
- 既存の構造は維持したまま拡充する

■ 6-10位の場合（積極的改善OK）
- タイトル改善OK
- 構造改善OK
- コンテンツ拡充OK

■ 11位以下の場合（大幅リライトOK）
- タイトル刷新OK
- 記事構成の見直しOK
- 競合分析に基づく大幅改善OK

【あなたの役割】
- データに基づいた客観的な分析
- 順位に応じたリスク管理を徹底した提案
- 具体的で実行可能な提案
- ユーザーの意図を理解した回答

【回答スタイル】
- 簡潔で分かりやすい
- 優先順位を明確にする
- 数値データを活用する
- 期待効果を定量化する
- 順位下落リスクを明示する`;
}

/**
 * 順位に応じた警告メッセージを生成
 * @param {Number} gyronPosition - ターゲットKW順位
 * @return {String} 警告メッセージ
 */
function getPositionWarning(gyronPosition) {
  if (!gyronPosition || gyronPosition <= 0) {
    return '';
  }
  
  if (gyronPosition === 1) {
    return `
【⚠️ 重要警告：1位獲得中 - リライト非推奨】
このページはターゲットKWで1位を獲得しています。
リライトによる順位下落リスクを避けるため、大幅な変更は推奨しません。

✅ 推奨する施策（低リスク）:
- 内部リンクの追加（他ページからの流入強化）
- 関連コンテンツの新規作成
- ページ表示速度の改善
- 誤字脱字の修正

❌ 避けるべき施策（高リスク）:
- タイトルの変更
- 大幅な構造変更
- コンテンツの削除・並び替え
`;
  }
  
  if (gyronPosition >= 2 && gyronPosition <= 5) {
    return `
【⚠️ 注意：上位表示中（${gyronPosition}位） - 低リスク施策のみ推奨】
このページは上位表示されています。順位下落リスクを最小化するため、
低リスクな施策のみを提案してください。

✅ 推奨する施策（低リスク）:
- メタディスクリプションの改善（CTR向上）
- コンテンツの追記・拡充（既存構造は維持）
- 内部リンクの最適化
- 画像・図解の追加
- FAQ追加

❌ 避けるべき施策（高リスク）:
- タイトルの変更
- 見出し構造の大幅変更
- コンテンツの削除
`;
  }
  
  if (gyronPosition >= 6 && gyronPosition <= 10) {
    return `
【📈 改善チャンス：${gyronPosition}位 - 積極的改善OK】
このページは6-10位圏内です。上位3位を目指して積極的に改善できます。

✅ 推奨する施策:
- タイトルの改善（クリック率向上）
- メタディスクリプションの最適化
- コンテンツ構造の改善
- 不足コンテンツの追加
- 競合との差別化
`;
  }
  
  // 11位以下
  return `
【🔧 大幅改善推奨：${gyronPosition}位 - 積極的リライトOK】
このページは11位以下です。大幅なリライトで順位改善を目指しましょう。

✅ 推奨する施策:
- タイトルの刷新
- 記事構成の見直し
- 競合分析に基づく不足コンテンツ追加
- 検索意図に合わせた内容改善
- E-E-A-T要素の強化
`;
}

/**
 * 提案形式を順位に応じて変更
 * @param {Number} gyronPosition - ターゲットKW順位
 * @return {String} 提案形式の指示
 */
function getSuggestionFormat(gyronPosition) {
  if (!gyronPosition || gyronPosition <= 0) {
    // 順位データなしの場合は標準形式
    return `【提案形式】
1. タイトル改善案（現在のタイトルをベースに改善した具体的な文言）
2. メタディスクリプション改善案
3. コンテンツ構造の改善
4. 期待される効果（定量的に）`;
  }
  
  if (gyronPosition === 1) {
    return `【提案形式】※1位獲得中のため低リスク施策のみ
1. 現状維持の推奨理由
2. 内部リンク追加の提案（具体的なリンク先ページ案）
3. 関連コンテンツ新規作成の提案
4. その他の低リスク改善案`;
  }
  
  if (gyronPosition >= 2 && gyronPosition <= 5) {
    return `【提案形式】※上位表示中のため低リスク施策のみ
1. メタディスクリプション改善案（タイトルは変更禁止）
2. コンテンツ追記・拡充案（既存構造を維持）
3. 内部リンク最適化の提案
4. 画像・図解追加の提案
5. 期待される効果（定量的に）`;
  }
  
  if (gyronPosition >= 6 && gyronPosition <= 10) {
    return `【提案形式】
1. タイトル改善案（現在のタイトルをベースに改善した具体的な文言）
2. メタディスクリプション改善案
3. コンテンツ構造の改善
4. 競合との差別化ポイント
5. 期待される効果（定量的に）`;
  }
  
  // 11位以下
  return `【提案形式】
1. タイトル刷新案（大幅な改善OK）
2. メタディスクリプション改善案
3. 記事構成の見直し案
4. 追加すべきコンテンツ
5. 競合分析に基づく改善ポイント
6. 期待される効果（定量的に）`;
}

/**
 * 提案プロンプトを生成（WordPress連携版）
 */
function buildSuggestionPrompt(pageData) {
  var gyronPosition = pageData.gyron_position || pageData.position || 0;
  
  // WordPressからページ情報を取得
  var wpData = null;
  try {
    wpData = getWordPressPageData(pageData.page_url);
  } catch (e) {
    Logger.log('WordPress取得スキップ: ' + e.message);
  }
  
  // 順位別の警告・制約
  var positionWarning = '';
  var positionConstraints = '';
  
  if (gyronPosition >= 1 && gyronPosition <= 3) {
    positionWarning = '⚠️ 重要注意: ターゲットKW「' + (pageData.target_keyword || '') + '」で' + gyronPosition + '位獲得中のため、順位下落リスクを避けて低リスク施策のみ提案します';
    positionConstraints = '【絶対禁止】タイトル変更、大幅な構成変更\n【推奨】メタディスクリプション最適化、コンテンツ追記、内部リンク追加';
  } else if (gyronPosition >= 4 && gyronPosition <= 5) {
    positionWarning = '⚠️ 注意: ' + gyronPosition + '位獲得中のため、慎重な改善を推奨';
    positionConstraints = '【非推奨】タイトル大幅変更\n【推奨】メタディスクリプション、コンテンツ追記、内部リンク';
  } else if (gyronPosition >= 6 && gyronPosition <= 10) {
    positionWarning = '📈 ' + gyronPosition + '位からTOP5を目指す改善を提案';
    positionConstraints = '【可能】タイトル微調整、メタディスクリプション、コンテンツ強化';
  } else {
    positionWarning = '🔧 現在' + gyronPosition + '位のため、積極的な改善が可能';
    positionConstraints = '【可能】タイトル変更、大幅リライト、構成変更';
  }
  
  // WordPress情報セクションを構築
  var wpSection = '';
  var faqInstruction = '';
  
  if (wpData && wpData.success) {
    wpSection = `
【WordPressから取得した実際のページ情報】
- メタディスクリプション: ${wpData.metaDescription || '未設定'}
- 文字数: ${wpData.wordCount}文字
- H2見出し数: ${wpData.h2List.length}個
- H2見出し一覧:
${wpData.h2List.map((h2, i) => '  ' + (i + 1) + '. ' + h2).join('\n')}
- FAQ有無: ${wpData.hasFaq ? 'あり（' + wpData.faqCount + '個）' : 'なし'}
- テーブル有無: ${wpData.hasTable ? 'あり' : 'なし'}
- 画像数: ${wpData.imageCount}枚
- 内部リンク数: ${wpData.internalLinks.length}本
- 既存の内部リンク先:
${wpData.internalLinks.slice(0, 10).map(link => '  - ' + link).join('\n')}
`;

    // FAQ指示を設定
    if (wpData.hasFaq) {
      faqInstruction = '- このページには既にFAQセクションがあります。新規FAQ追加は提案せず、必要に応じて「追加すべきQ&A」や「改善すべきQ&A」を提案してください。';
    } else {
      faqInstruction = '- FAQセクションがないため、ユーザーが検索しそうな質問と回答の追加を検討してください。';
    }
  } else {
    wpSection = `
【ページ情報】
- メタディスクリプション: ${pageData.meta_description || '取得できません'}
`;
    faqInstruction = '- 必要に応じてFAQセクションの追加を検討してください。';
  }
  
  var prompt = `
以下のページのリライト提案をしてください。

${positionWarning}

【ページ基本情報】
- URL: ${pageData.page_url}
- タイトル: ${pageData.page_title || '不明'}
- ターゲットキーワード: ${pageData.target_keyword || '不明'}
- Gyron順位: ${gyronPosition}位
${wpSection}
【パフォーマンスデータ】
- 表示回数: ${pageData.impressions || 0}
- クリック数: ${pageData.clicks || 0}
- CTR: ${pageData.ctr || 0}%
- 平均掲載順位: ${pageData.avg_position || '-'}
- PV: ${pageData.pageviews || 0}
- 滞在時間: ${pageData.avg_time || 0}秒
- 直帰率: ${pageData.bounce_rate || 0}%

【順位別の制約】
${positionConstraints}

【提案の注意事項】
${faqInstruction}
${wpData && wpData.hasTable ? '- このページには既にテーブルがあります。必要に応じて改善提案をしてください。' : '- 比較表やテーブルの追加を検討してください。'}
- 内部リンク追加を提案する場合は、サイト内の具体的なページURLを提案してください
- 既存の内部リンク先と重複しないリンク先を提案してください

【出力形式の注意】
- HTMLタグ（<table>、<tr>、<td>など）は使用しないでください
- マークダウンのコードブロック（\`\`\`）は使用しないでください
- テーブルを提案する場合は、内容を箇条書きで説明してください
- 提案は自然な日本語の文章で記述してください
`;
// 出力形式を追加
  prompt += '\n\n' + getSuggestionFormatV2(gyronPosition);
  return prompt;
}



/**
 * 週次レポート生成
 */
function generateWeeklyReport(topPages) {
  const now = new Date();
  const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy/MM/dd');
  
  let report = `【SEOリライト支援ツール】週次レポート\n`;
  report += `レポート日時: ${dateStr}\n\n`;
  report += `=== 今週リライトすべきページ TOP10 ===\n\n`;
  
  topPages.forEach((page, index) => {
    // 順位に応じたリスクレベルを表示
    let riskLevel = '';
    if (page.gyronPosition === 1) {
      riskLevel = '🔴 1位（リライト非推奨）';
    } else if (page.gyronPosition >= 2 && page.gyronPosition <= 5) {
      riskLevel = '🟠 上位（低リスク施策のみ）';
    } else if (page.gyronPosition >= 6 && page.gyronPosition <= 10) {
      riskLevel = '🟡 中位（積極改善OK）';
    } else if (page.gyronPosition > 10) {
      riskLevel = '🟢 下位（大幅改善OK）';
    } else {
      riskLevel = '⚪ 順位不明';
    }
    
    report += `${index + 1}位: ${page.url}\n`;
    report += `   タイトル: ${page.title || '未取得'}\n`;
    report += `   スコア: ${page.score}点\n`;
    report += `   ターゲットKW: ${page.targetKeyword || '未設定'}\n`;
    report += `   KW順位: ${page.gyronPosition ? page.gyronPosition + '位' : 'N/A'} ${riskLevel}\n\n`;
  });
  
  report += `\n【順位別リライト戦略】\n`;
  report += `🔴 1位: リライト非推奨（現状維持）\n`;
  report += `🟠 2-5位: 低リスク施策のみ（タイトル変更禁止）\n`;
  report += `🟡 6-10位: 積極的改善OK\n`;
  report += `🟢 11位以下: 大幅リライトOK\n\n`;
  report += `詳細は統合データシートをご確認ください。\n`;
  
  return report;
}

/**
 * 週次レポートをメール送信
 */
function sendWeeklyReportEmail(report) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName('設定・マスタ');
  
  if (!settingsSheet) {
    Logger.log('設定シートが見つかりません');
    return;
  }
  
  const data = settingsSheet.getDataRange().getValues();
  let email = '';
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === 'NOTIFICATION_EMAIL') {
      email = data[i][1];
      break;
    }
  }
  
  if (!email) {
    Logger.log('通知先メールアドレスが設定されていません');
    return;
  }
  
  const subject = '【SEOツール】週次レポート';
  
  MailApp.sendEmail(email, subject, report);
  Logger.log(`週次レポート送信完了: ${email}`);
}

/**
 * エラーメール送信
 */
function sendErrorEmail(error) {
  const email = Session.getActiveUser().getEmail();
  const subject = 'SEOツール エラー通知';
  const body = `週次分析でエラーが発生しました:\n\n${error.message}\n\n${error.stack}`;
  
  MailApp.sendEmail(email, subject, body);
}

/**
 * 週次分析ログ記録
 */
function logWeeklyAnalysis(topPages) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName('分析ログ');
  
  if (!logSheet) {
    logSheet = ss.insertSheet('分析ログ');
    logSheet.appendRow(['日時', '処理', '上位ページ数', 'メモ']);
  }
  
  const now = new Date();
  const pagesCount = topPages.length;
  
  logSheet.appendRow([now, '週次分析', pagesCount, '正常完了']);
}

/**
 * 週次トリガーをセットアップ
 */
function setupWeeklyTrigger() {
  // 既存のトリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'runWeeklyAnalysis') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // 新規トリガー作成（毎週月曜9:00）
  ScriptApp.newTrigger('runWeeklyAnalysis')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(9)
    .create();
  
  Logger.log('週次トリガー設定完了（毎週月曜9:00）');
}

/**
 * テスト: 提案生成
 */
function testGenerateSuggestion() {
  const testUrl = '/iphonerepair-screen-line';
  const suggestion = generateRewriteSuggestions(testUrl);
  Logger.log('=== 生成された提案 ===');
  Logger.log(suggestion);
}

/**
 * テスト: フィルター済み上位ページ
 */
function testFilteredTopPages() {
  const topPages = getTopPriorityPagesFiltered(10);
  Logger.log('=== 優先度上位10ページ（フィルター済み） ===');
  topPages.forEach((page, index) => {
    let riskLevel = '';
    if (page.gyronPosition === 1) {
      riskLevel = '🔴 リライト非推奨';
    } else if (page.gyronPosition >= 2 && page.gyronPosition <= 5) {
      riskLevel = '🟠 低リスク施策のみ';
    } else if (page.gyronPosition >= 6 && page.gyronPosition <= 10) {
      riskLevel = '🟡 積極改善OK';
    } else if (page.gyronPosition > 10) {
      riskLevel = '🟢 大幅改善OK';
    } else {
      riskLevel = '⚪ 順位不明';
    }
    
    Logger.log(`${index + 1}位: ${page.url}`);
    Logger.log(`   タイトル: ${page.title}`);
    Logger.log(`   スコア: ${page.score}点`);
    Logger.log(`   ターゲットKW: ${page.targetKeyword || '未設定'}`);
    Logger.log(`   KW順位: ${page.gyronPosition ? page.gyronPosition + '位' : 'N/A'} ${riskLevel}`);
  });
}

/**
 * テスト: 週次分析
 */
function testWeeklyAnalysis() {
  runWeeklyAnalysis();
}

/**
 * テスト: メール送信
 */
function testSendEmail() {
  const report = '【テスト】週次レポート\n\nこれはテストメールです。';
  sendWeeklyReportEmail(report);
}

/**
 * AIOスコアを計算（0-100点）
 */
function calculateAIOScore(pageData) {
  const pageUrl = String(pageData.page_url || '').toLowerCase();
  const pageTitle = String(pageData.page_title || '').toLowerCase();
  
  // 質問系キーワードの有無
  const hasQuestionKeywords = containsQuestionKeywords(pageUrl, pageTitle);
  
  if (!hasQuestionKeywords) {
    return 0; // 質問系でなければAIO対象外
  }
  
  // トラフィックスコア（40%）
  const pageViews = parseFloat(pageData.avg_page_views_30d) || 0;
  let trafficScore = 0;
  
  if (pageViews >= 500) {
    trafficScore = 100;
  } else if (pageViews >= 100) {
    trafficScore = 70;
  } else if (pageViews >= 50) {
    trafficScore = 40;
  } else if (pageViews >= 10) {
    trafficScore = 20;
  }
  
  // 順位スコア（30%）
  const position = parseFloat(pageData.avg_position) || 100;
  let positionScore = 0;
  
  if (position <= 5) {
    positionScore = 100; // すでに上位表示
  } else if (position <= 10) {
    positionScore = 80;
  } else if (position <= 20) {
    positionScore = 50;
  } else {
    positionScore = 20;
  }
  
  // CTRスコア（30%）
  const ctr = parseFloat(pageData.avg_ctr) || 0;
  const expectedCTR = getExpectedCTR(position);
  const ctrGap = expectedCTR - ctr;
  let ctrScore = 0;
  
  if (ctrGap >= 0.30) {
    ctrScore = 100;
  } else if (ctrGap >= 0.20) {
    ctrScore = 70;
  } else if (ctrGap >= 0.10) {
    ctrScore = 40;
  }
  
  const totalScore = (trafficScore * 0.40) + (positionScore * 0.30) + (ctrScore * 0.30);
  return Math.round(totalScore);
}

/**
 * 質問系キーワードを含むか判定
 */
function containsQuestionKeywords(url, title) {
  const questionKeywords = [
    'how-to', 'howto', 'what-is', 'why', 'when', 'where', 'which',
    'できない', 'わからない', 'とは', 'なぜ', 'いつ', 'どこ', 'どれ',
    '方法', 'やり方', '手順', '解決', '対処', '原因'
  ];
  
  return questionKeywords.some(keyword => 
    url.includes(keyword) || title.includes(keyword)
  );
}

/**
 * AIO最適化提案を生成
 */
function generateAIOSuggestion(pageUrl) {
  Logger.log(`=== AIO提案生成開始: ${pageUrl} ===`);
  
  try {
    // ページデータを取得
    const pageData = getPageDataForAIO(pageUrl);
    
    if (!pageData) {
      throw new Error('ページデータが見つかりません');
    }
    
    // AIOスコアを計算
    const aioScore = calculateAIOScore(pageData);
    
    if (aioScore === 0) {
      return 'このページはAIO最適化の対象ではありません（質問系キーワードを含まない）';
    }
    
    // GyronSEOデータを取得
    const gyronData = getGyronSEODataForUrl(pageUrl);
    
    // Claude APIで提案生成
    const suggestion = callClaudeForAIOSuggestion(pageData, aioScore, gyronData);
    
    Logger.log('=== AIO提案生成完了 ===');
    return suggestion;
    
  } catch (error) {
    Logger.log(`エラー: ${error.message}`);
    throw error;
  }
}

/**
 * AIO用のページデータを取得
 */
function getPageDataForAIO(pageUrl) {
  return getPageDataForSuggestion(pageUrl);
}

/**
 * GyronSEOデータを取得
 */
function getGyronSEODataForUrl(pageUrl) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gyronSheet = ss.getSheetByName('GyronSEO_RAW');
  
  if (!gyronSheet) {
    return null;
  }
  
  const data = gyronSheet.getDataRange().getValues();
  const headers = data[0];
  const urlIndex = headers.indexOf('page_url');
  const keywordIndex = headers.indexOf('keyword');
  const positionIndex = headers.indexOf('position');
  
  for (let i = 1; i < data.length; i++) {
    const url = String(data[i][urlIndex] || '').trim();
    
    if (url === pageUrl) {
      return {
        targetKeyword: data[i][keywordIndex],
        gyronPosition: data[i][positionIndex]
      };
    }
  }
  
  return null;
}

/**
 * Claude APIでAIO提案を生成
 */
function callClaudeForAIOSuggestion(pageData, aioScore, gyronData) {
  const systemPrompt = buildAIOSystemPrompt();
  const userPrompt = buildAIOUserPrompt(pageData, aioScore, gyronData);
  
  // Claude API呼び出し
  const suggestion = callClaudeAPI(userPrompt, systemPrompt);
  
  return suggestion;
}

/**
 * AIO用システムプロンプト
 * ★v2.2: 現在の日付情報を追加
 */
function buildAIOSystemPrompt() {
  // 現在の日付情報を取得
  var today = new Date();
  var currentYear = today.getFullYear();
  var currentDate = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy年MM月dd日');
  
  return `あなたはSEOとAIO（AI Overviews）最適化の専門家です。

【重要：現在の日付】
今日の日付は${currentDate}です。
年号を含む提案をする場合は、必ず${currentYear}年を使用してください。
「2024年最新」などの古い年号は絶対に使用しないでください。

【AIOとは】
GoogleのAI Overviewsは、検索結果の上部に表示されるAI生成の回答です。

【AIO最適化のポイント】
1. 質問に対する明確な回答を冒頭に配置
2. 構造化データ（FAQ、HowTo等）を活用
3. 簡潔で読みやすい文章構造
4. 信頼性の高い情報源への言及
5. ステップバイステップの説明（How-to記事の場合）

【あなたの役割】
- データに基づいた具体的な提案
- AIO表示を意識したコンテンツ最適化案
- 実装しやすい形での提案`;
}

/**
 * AIO用ユーザープロンプト
 */
function buildAIOUserPrompt(pageData, aioScore, gyronData) {
  let prompt = `以下のページをAIO（AI Overviews）最適化するための提案をお願いします。

【ページURL】
${pageData.page_url}

【ページタイトル】
${pageData.page_title || '取得できません'}

【AIOスコア】
${aioScore}/100点

【現在のパフォーマンス】
- 検索順位: ${pageData.avg_position}位
- CTR: ${((parseFloat(pageData.avg_ctr) || 0) * 100).toFixed(2)}%
- 月間PV: ${pageData.avg_page_views_30d}`;

  if (gyronData) {
    prompt += `\n\n【ターゲットキーワード】
${gyronData.targetKeyword}（Gyron順位: ${gyronData.gyronPosition}位）`;
  }

  prompt += `\n\n以下の形式でAIO最適化の提案をしてください：

【提案形式】
1. 冒頭の回答文（質問に対する明確な答え）
2. 構造化データの追加提案（FAQ、HowTo等）
3. コンテンツ構造の改善案
4. AIO表示される可能性を高めるための施策`;

  return prompt;
}

/**
 * テスト: AIO提案生成
 */
function testGenerateAIOSuggestion() {
  const testUrl = '/iphonerepair-screen-line';
  const suggestion = generateAIOSuggestion(testUrl);
  Logger.log('=== 生成されたAIO提案 ===');
  Logger.log(suggestion);
}

/**
 * Scoring.gs - パフォーマンススコア更新版（Day 9-10）
 * 
 * 変更点:
 * - calculatePerformanceScore()にUX要素を統合（50%比重）
 * - 直帰率スコア: 50% → 25%
 * - 滞在時間スコア: 50% → 25%
 * - UXスコア: 0% → 50%（新規追加）
 * 
 * 更新日: 2025/11/26
 * バージョン: 2.0（UX統合版）
 */

/**
 * パフォーマンススコア計算（UX統合版）
 * 
 * 計算式:
 * performance_score = 
 *   (直帰率スコア × 0.25) + 
 *   (滞在時間スコア × 0.25) + 
 *   (UXスコア × 0.50)
 * 
 * 【従来の指標（50%）】
 * - 直帰率スコア（25%）: サイト平均と比較
 * - 滞在時間スコア（25%）: サイト平均と比較
 * 
 * 【新規UX指標（50%）】
 * - UXスコア（50%）: Clarity + GTMスクロールの統合スコア
 *   - スクロール深度（40%）
 *   - デッドクリック（25%）
 *   - レイジクリック（20%）
 *   - クイックバック（15%）
 * 
 * @param {Object} pageData - ページデータ
 * @return {Number} パフォーマンススコア（0-100）
 */
function calculatePerformanceScoreV2(pageData) {
  // 【従来の指標（50%）】
  const bounceRateScore = calculateBounceRateScore(pageData.bounce_rate) * 0.25;
  const durationScore = calculateDurationScore(pageData.avg_session_duration) * 0.25;
  
  // 【新規UX指標（50%）】
  // clarity_ux_scoreは既にClarityIntegration.gsで計算済み
  const uxScore = pageData.clarity_ux_score || 0;
  const uxWeight = uxScore * 0.50;
  
  // 合計スコア
  const totalScore = bounceRateScore + durationScore + uxWeight;
  
  return Math.round(totalScore);
}

/**
 * 直帰率スコア計算
 * サイト平均との差分で評価
 * 
 * スコアリングロジック:
 * - サイト平均+30%以上: 100点（深刻な問題）
 * - サイト平均+20-30%: 70点（問題あり）
 * - サイト平均+10-20%: 40点（改善余地あり）
 * - サイト平均並み: 0点（問題なし）
 * 
 * @param {Number} bounceRate - 直帰率（%）
 * @return {Number} スコア（0-100）
 */
function calculateBounceRateScore(bounceRate) {
  // サイト平均値を取得（統合データシートから計算）
  const siteAvg = getSiteBounceRateAverage();
  
  // 平均値との差分
  const diff = bounceRate - siteAvg;
  
  if (diff >= 30) return 100;  // 深刻な問題
  if (diff >= 20) return 70;   // 問題あり
  if (diff >= 10) return 40;   // 改善余地あり
  return 0;  // 問題なし
}

/**
 * 滞在時間スコア計算
 * サイト平均との差分で評価
 * 
 * スコアリングロジック:
 * - サイト平均の-50%以上短い: 100点（深刻な問題）
 * - サイト平均の-30-50%短い: 70点（問題あり）
 * - サイト平均の-10-30%短い: 40点（改善余地あり）
 * - サイト平均並み: 0点（問題なし）
 * 
 * @param {Number} duration - 平均セッション時間（秒）
 * @return {Number} スコア（0-100）
 */
function calculateDurationScore(duration) {
  // サイト平均値を取得（統合データシートから計算）
  const siteAvg = getSiteDurationAverage();
  
  // 平均値との差分率
  const diffRate = (siteAvg - duration) / siteAvg;
  
  if (diffRate >= 0.5) return 100;  // 深刻な問題
  if (diffRate >= 0.3) return 70;   // 問題あり
  if (diffRate >= 0.1) return 40;   // 改善余地あり
  return 0;  // 問題なし
}

/**
 * サイト全体の直帰率平均を取得
 * @return {Number} サイト平均直帰率（%）
 */
function getSiteBounceRateAverage() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('統合データ');
  
  if (!sheet) return 60; // デフォルト値
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 60;
  
  // ヘッダー行から列番号を取得
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colBounceRate = headers.indexOf('bounce_rate') + 1;
  
  if (colBounceRate === 0) return 60;
  
  // 全ページの直帰率を取得
  const bounceRates = sheet.getRange(2, colBounceRate, lastRow - 1, 1).getValues().flat();
  
  // 平均値を計算（0を除く）
  const validRates = bounceRates.filter(rate => rate > 0);
  if (validRates.length === 0) return 60;
  
  const sum = validRates.reduce((a, b) => a + b, 0);
  const avg = sum / validRates.length;
  
  return Math.round(avg);
}

/**
 * サイト全体の滞在時間平均を取得
 * @return {Number} サイト平均滞在時間（秒）
 */
function getSiteDurationAverage() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('統合データ');
  
  if (!sheet) return 120; // デフォルト値（2分）
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 120;
  
  // ヘッダー行から列番号を取得
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colDuration = headers.indexOf('avg_session_duration') + 1;
  
  if (colDuration === 0) return 120;
  
  // 全ページの滞在時間を取得
  const durations = sheet.getRange(2, colDuration, lastRow - 1, 1).getValues().flat();
  
  // 平均値を計算（0を除く）
  const validDurations = durations.filter(duration => duration > 0);
  if (validDurations.length === 0) return 120;
  
  const sum = validDurations.reduce((a, b) => a + b, 0);
  const avg = sum / validDurations.length;
  
  return Math.round(avg);
}

/**
 * パフォーマンススコアテスト
 * 手動で実行してテストします
 */
function testPerformanceScoreCalculation() {
  Logger.log('=== パフォーマンススコア計算テスト ===');
  
  // テストケース1: 直帰率高い、滞在時間短い、UX悪い
  const test1 = {
    bounce_rate: 85,
    avg_session_duration: 45,
    clarity_ux_score: 70
  };
  const score1 = calculatePerformanceScoreV2(test1);
  Logger.log('テスト1（全体的に悪い）: ' + score1 + '点');
  
  // テストケース2: 直帰率低い、滞在時間長い、UX良い
  const test2 = {
    bounce_rate: 40,
    avg_session_duration: 180,
    clarity_ux_score: 10
  };
  const score2 = calculatePerformanceScoreV2(test2);
  Logger.log('テスト2（全体的に良い）: ' + score2 + '点');
  
  // テストケース3: 直帰率は良いが、UX悪い
  const test3 = {
    bounce_rate: 45,
    avg_session_duration: 150,
    clarity_ux_score: 85
  };
  const score3 = calculatePerformanceScoreV2(test3);
  Logger.log('テスト3（直帰率は良いがUX悪い）: ' + score3 + '点');
  
  Logger.log('=== テスト完了 ===');
  Logger.log('サイト平均直帰率: ' + getSiteBounceRateAverage() + '%');
  Logger.log('サイト平均滞在時間: ' + getSiteDurationAverage() + '秒');
}

/**
 * テスト: タイトル取得確認
 */
function testTitleExtraction() {
  Logger.log('=== タイトル取得テスト ===');
  
  const topPages = getTopPriorityPagesFiltered(5);
  
  topPages.forEach((page, index) => {
    Logger.log(`${index + 1}位:`);
    Logger.log(`  URL: ${page.url}`);
    Logger.log(`  タイトル: ${page.title}`);
    Logger.log(`  スコア: ${page.score}点`);
    Logger.log(`  ターゲットKW: ${page.targetKeyword || '未設定'}`);
    Logger.log(`  KW順位: ${page.gyronPosition ? page.gyronPosition + '位' : 'N/A'}`);
  });
  
  Logger.log('=== テスト完了 ===');
}

/**
 * テスト: 年号確認
 */
function testYearInPrompt() {
  Logger.log('=== 年号テスト ===');
  
  const siteInfo = getSiteInfoFromSettings();
  const systemPrompt = buildSystemPromptWithSiteInfo(siteInfo);
  
  Logger.log('システムプロンプト（冒頭500文字）:');
  Logger.log(systemPrompt.substring(0, 500));
  
  Logger.log('=== テスト完了 ===');
}

/**
 * テスト: 順位別警告メッセージ確認
 */
function testPositionWarnings() {
  Logger.log('=== 順位別警告メッセージテスト ===');
  
  Logger.log('--- 1位の場合 ---');
  Logger.log(getPositionWarning(1));
  
  Logger.log('--- 3位の場合 ---');
  Logger.log(getPositionWarning(3));
  
  Logger.log('--- 8位の場合 ---');
  Logger.log(getPositionWarning(8));
  
  Logger.log('--- 15位の場合 ---');
  Logger.log(getPositionWarning(15));
  
  Logger.log('=== テスト完了 ===');
}
// ============================================
// フェーズ1追加: 優先度順提案フォーマット
// 追加日: 2025年12月8日
// ============================================

/**
 * 提案形式を優先度順フォーマットに変更
 */
function getSuggestionFormatV2(gyronPosition) {
  var baseFormat = `
【提案形式】※必ずこの形式で出力してください

## 🎯 リライト提案（優先度順）

以下の形式で、優先度の高い順に3〜5個の提案を出力してください。

### 🥇 優先度1: [提案タイトル]
**種別**: [タイトル変更/メタディスクリプション/H2追加/本文追加/Q&A追加/画像追加/内部リンク追加]
**現状**: [現在の状態を簡潔に]
**改善案**: [具体的な改善内容]
**理由**: [この提案を優先する理由と期待効果]

### 🥈 優先度2: [提案タイトル]
**種別**: [種別]
**現状**: [現状]
**改善案**: [改善案]
**理由**: [理由]

### 🥉 優先度3: [提案タイトル]
**種別**: [種別]
**現状**: [現状]
**改善案**: [改善案]
**理由**: [理由]

（必要に応じて優先度4、5も追加）
`;

  if (!gyronPosition || gyronPosition <= 0) {
    return baseFormat;
  }
  
  if (gyronPosition === 1) {
    return baseFormat + `
【特別指示：1位獲得中】
- タイトル変更は絶対に提案しないでください
- 低リスク施策（内部リンク追加、関連コンテンツ新規作成）を優先してください`;
  }
  
  if (gyronPosition >= 2 && gyronPosition <= 5) {
    return baseFormat + `
【特別指示：上位表示中（${gyronPosition}位）】
- タイトル変更は提案しないでください（リスク大）
- メタディスクリプション改善、コンテンツ追記を優先してください`;
  }
  
  if (gyronPosition >= 6 && gyronPosition <= 10) {
    return baseFormat + `
【特別指示：中位（${gyronPosition}位）】
- 積極的な改善を提案してください
- タイトル改善もOKです`;
  }
  
  return baseFormat + `
【特別指示：下位（${gyronPosition}位）】
- 大幅な改善を提案してください
- タイトル刷新、記事構成の見直しも検討してください`;
}
/**
 * 提案テキストに優先度別ボタンを追加
 */
function addSuggestionButtons(suggestion, pageUrl) {
  var sections = parsePrioritySuggestions(suggestion);
  
  Logger.log('検出された優先度提案数: ' + sections.length);
  
  if (sections.length === 0) {
    return suggestion;
  }
  
  var modifiedSuggestion = suggestion;
  
  for (var i = sections.length - 1; i >= 0; i--) {
    var section = sections[i];
    
    var buttonHtml = '\n\n<div class="suggestion-buttons" data-priority="' + section.priority + '">' +
                     '<button class="generate-outline-btn" ' +
                     'data-page-url="' + escapeHtmlAttr(pageUrl) + '" ' +
                     'data-suggestion-title="' + escapeHtmlAttr(section.title) + '" ' +
                     'data-suggestion-type="' + escapeHtmlAttr(section.type) + '" ' +
                     'data-suggestion-content="' + escapeHtmlAttr(section.content) + '">' +
                     '📝 アウトラインを生成</button> ' +
                     '<button class="add-task-btn" ' +
                     'data-page-url="' + escapeHtmlAttr(pageUrl) + '" ' +
                     'data-task-type="' + escapeHtmlAttr(section.type) + '" ' +
                     'data-task-content="' + escapeHtmlAttr(section.content) + '" ' +
                     'data-priority="' + section.priority + '">' +
                     '➕ タスクに追加</button>' +
                     '</div>\n\n---\n';
    
    if (section.endIndex > 0 && section.endIndex <= modifiedSuggestion.length) {
      modifiedSuggestion = modifiedSuggestion.substring(0, section.endIndex) + 
                           buttonHtml + 
                           modifiedSuggestion.substring(section.endIndex);
    }
  }
  
  return modifiedSuggestion;
}


/**
 * 優先度付き提案を解析
 */
function parsePrioritySuggestions(suggestion) {
  var sections = [];
  var priorityPattern = /###\s*\S*\s*優先度(\d+)[：:]\s*(.+?)(?=\n)/g;
  var match;
  
  while ((match = priorityPattern.exec(suggestion)) !== null) {
    var priority = parseInt(match[1]);
    var title = match[2].trim();
    var startIndex = match.index;
    
    if (sections.length > 0) {
      sections[sections.length - 1].endIndex = startIndex;
    }
    
    var contentStart = match.index + match[0].length;
    var nextSection = suggestion.indexOf('### ', contentStart);
    var contentEnd = nextSection > 0 ? nextSection : suggestion.length;
    var sectionContent = suggestion.substring(contentStart, contentEnd).trim();
    
    var typeMatch = sectionContent.match(/\*\*種別\*\*[：:]\s*(.+?)(?:\n|$)/);
    var type = typeMatch ? typeMatch[1].trim() : title;
    
    var improvementMatch = sectionContent.match(/\*\*改善案\*\*[：:]\s*([\s\S]+?)(?=\*\*|$)/);
    var improvement = improvementMatch ? improvementMatch[1].trim() : sectionContent.substring(0, 200);
    
    sections.push({
      priority: priority,
      title: title,
      type: type,
      content: improvement,
      startIndex: startIndex,
      endIndex: contentEnd
    });
  }
  
  if (sections.length > 0 && !sections[sections.length - 1].endIndex) {
    sections[sections.length - 1].endIndex = suggestion.length;
  }
  
  return sections;
}


/**
 * HTML属性用エスケープ
 */
function escapeHtmlAttr(text) {
  if (!text) return '';
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;')
    .replace(/\n/g, ' ');
}
/**
 * 軽量アウトライン生成
 */
function generateOutline(pageUrl, suggestionTitle, suggestionType, suggestionContent) {
  Logger.log('=== アウトライン生成開始 ===');
  
  try {
    var siteInfo = getSiteInfoFromSettings();
    var today = new Date();
    var currentYear = today.getFullYear();
    
    var systemPrompt = 'あなたはSEOコンテンツのアウトライン作成の専門家です。\n\n' +
      '【重要】\n' +
      '- 現在は' + currentYear + '年です\n' +
      '- 軽量で実行しやすいアウトラインを作成してください\n' +
      '- 詳細な本文は書かず、構成案のみを出力してください\n\n' +
      '【サイト情報】\n' +
      '- サイト名: ' + (siteInfo.siteName || '') + '\n' +
      '- ジャンル: ' + (siteInfo.siteGenre || '');
    
    var userPrompt = '以下の提案に基づいて、コンテンツアウトラインを作成してください。\n\n' +
      '【ページURL】\n' + pageUrl + '\n\n' +
      '【提案種別】\n' + suggestionType + '\n\n' +
      '【提案タイトル】\n' + suggestionTitle + '\n\n' +
      '【提案内容】\n' + suggestionContent + '\n\n' +
      '【出力形式】\n' +
      '## 📝 コンテンツアウトライン: [セクション名]\n\n' +
      '【H2案】\n[具体的な見出し案]\n\n' +
      '【含めるべき内容】\n- 項目1\n- 項目2\n- 項目3\n- 項目4\n\n' +
      '【参考データ】\n- 参照すべき情報源1\n- 参照すべき情報源2\n\n' +
      '【想定文字数】\n[推奨文字数]';
    
    var outline = callClaudeAPI(userPrompt, systemPrompt);
    
    return { success: true, outline: outline, pageUrl: pageUrl, suggestionType: suggestionType };
    
  } catch (error) {
    Logger.log('アウトライン生成エラー: ' + error.message);
    return { success: false, error: error.message };
  }
}


/**
 * 提案をタスク管理シートに登録
 */
function registerTaskFromSuggestion(pageUrl, taskType, taskContent, priority) {
  Logger.log('=== タスク登録開始 ===');
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('タスク管理');
    
    // シートがなければ作成
    if (!sheet) {
      sheet = ss.insertSheet('タスク管理');
      sheet.appendRow([
        'task_id', 'page_url', 'page_title', 'task_type', 'task_detail',
        'source', 'priority_rank', 'expected_effect', 'status',
        'created_date', 'completed_date', 'actual_change', 'cooling_days', 'notes'
      ]);
      sheet.getRange(1, 1, 1, 14).setFontWeight('bold').setBackground('#4a90d9').setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }
    
    var now = new Date();
    var dateStr = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
    var taskId = 'TASK_' + dateStr;
    
    // ページタイトルを取得
    var pageTitle = getPageTitleFromUrl(pageUrl);
    
    // 期待効果を種別から推定
    var expectedEffect = getExpectedEffectFromType(taskType);
    
    var data = sheet.getDataRange().getValues();
    
    // 重複チェック（同じページ・タスク種別で未完了のもの）
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] === pageUrl && data[i][3] === taskType && data[i][8] !== '完了') {
        return { success: false, error: '同じタスクが既に存在します', existingTaskId: data[i][0] };
      }
    }
    
    // 新しいタスクを追加
    sheet.appendRow([
      taskId,           // task_id
      pageUrl,          // page_url
      pageTitle,        // page_title
      taskType,         // task_type
      taskContent,      // task_detail
      'AI提案',         // source
      priority,         // priority_rank
      expectedEffect,   // expected_effect
      '未着手',         // status
      Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'), // created_date
      '',               // completed_date
      '',               // actual_change
      '',               // cooling_days
      ''                // notes
    ]);
    
    return { success: true, taskId: taskId, row: sheet.getLastRow() };
    
  } catch (error) {
    Logger.log('タスク登録エラー: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * URLからページタイトルを取得
 */
function getPageTitleFromUrl(pageUrl) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('統合データ');
    if (!sheet) return '';
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var urlIndex = headers.indexOf('page_url');
    var titleIndex = headers.indexOf('page_title');
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][urlIndex] === pageUrl) {
        return data[i][titleIndex] || '';
      }
    }
    return '';
  } catch (e) {
    return '';
  }
}

/**
 * タスク種別から期待効果を推定
 */
function getExpectedEffectFromType(taskType) {
  var effectMap = {
    'タイトル変更': 'CTR改善',
    'メタディスクリプション': 'CTR改善',
    'メタディスクリプション改善': 'CTR改善',
    'H2追加': '検索順位向上',
    '本文追加': '滞在時間改善',
    'Q&A追加': '検索順位向上',
    'FAQ追加': '検索順位向上',
    '画像追加': '滞在時間改善',
    '内部リンク追加': '回遊率向上',
    '動画追加': '滞在時間改善'
  };
  
  return effectMap[taskType] || '';
}