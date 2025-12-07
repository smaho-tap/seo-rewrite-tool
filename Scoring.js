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
    positionIndex: headers.indexOf('gyron_position'),
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
 */
function calculateOpportunityScore(row, indexes) {
  const position = parseFloat(row[indexes.positionIndex]) || 100;
  const impressions = parseFloat(row[indexes.impressionsIndex]) || 0;
  const ctr = parseFloat(row[indexes.ctrIndex]) || 0;
  
  // 順位スコア（40%）★Day 22修正: 11-20位を最優先に
  let positionScore = 0;
  if (position >= 1 && position <= 3) {
    positionScore = 10;   // 現状維持（リスク回避）
  } else if (position >= 4 && position <= 10) {
    positionScore = 75;   // TOP3入りを狙える
  } else if (position >= 11 && position <= 20) {
    positionScore = 100;  // ★最優先：1ページ目入り直前
  } else if (position >= 21 && position <= 30) {
    positionScore = 95;   // ★高優先：1ページ目入り射程圏内
  } else if (position >= 31 && position <= 50) {
    positionScore = 40;   // 中優先
  } else {
    positionScore = 20;   // 低優先
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
  const opportunityIndex = headers.indexOf('opportunity_score');
  const performanceIndex = headers.indexOf('performance_score');
  const businessImpactIndex = headers.indexOf('business_impact_score');
  const targetKWIndex = headers.indexOf('target_keyword');
  const gyronPositionIndex = headers.indexOf('gyron_position');
  
  const pages = [];
  
  for (let i = 1; i < data.length; i++) {
    const url = String(data[i][urlIndex] || '').trim();
    const title = String(data[i][titleIndex] || '').trim();
    const score = parseFloat(data[i][scoreIndex]) || 0;
    const opportunityScore = opportunityIndex >= 0 ? parseFloat(data[i][opportunityIndex]) || 0 : 0;
    const performanceScore = performanceIndex >= 0 ? parseFloat(data[i][performanceIndex]) || 0 : 0;
    const businessImpactScore = businessImpactIndex >= 0 ? parseFloat(data[i][businessImpactIndex]) || 0 : 0;
    const targetKW = targetKWIndex >= 0 ? String(data[i][targetKWIndex] || '').trim() : '';
    const gyronPosition = gyronPositionIndex >= 0 ? parseFloat(data[i][gyronPositionIndex]) || null : null;
    
    if (url) {
      pages.push({ 
        url, 
        title, 
        score,
        totalScore: score,
        opportunityScore: opportunityScore,
        performanceScore: performanceScore,
        businessImpactScore: businessImpactScore,
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
 * ユーザープロンプト構築
 * ★v2.2: page_title追加、ターゲットKW追加、順位別警告・制約を追加
 */
function buildSuggestionPrompt(pageData) {
  // CTRの安全な処理
  var ctrValue = parseFloat(pageData.avg_ctr) || 0;
  var ctrPercent = (ctrValue * 100).toFixed(2);
  
  // ターゲットKW順位を取得
  var gyronPosition = parseFloat(pageData.gyron_position) || null;
  var targetKeyword = pageData.target_keyword || '';
  
  // 順位に応じた警告メッセージを取得
  var positionWarning = getPositionWarning(gyronPosition);
  
  // 順位に応じた提案形式を取得
  var suggestionFormat = getSuggestionFormat(gyronPosition);
  
  var prompt = `以下のページのリライト提案をお願いします。

【ページURL】
${pageData.page_url}

【現在のタイトル】
${pageData.page_title || '取得できません'}

【ターゲットキーワード情報】
- ターゲットKW: ${targetKeyword || '未設定'}
- ターゲットKW順位: ${gyronPosition ? gyronPosition + '位' : 'N/A'}
- GSC平均順位: ${pageData.avg_position || 'N/A'}位（全クエリ平均）
${positionWarning}
【現在のパフォーマンス】
- CTR: ${ctrPercent}%
- 月間クリック数: ${pageData.total_clicks_30d || 0}回
- 月間表示回数: ${pageData.total_impressions_30d || 0}回
- ページビュー: ${pageData.avg_page_views_30d || 0}
- 直帰率: ${pageData.bounce_rate || 0}%
- 平均滞在時間: ${pageData.avg_session_duration || 0}秒

【スコア】
- 機会損失スコア: ${pageData.opportunity_score || 0}/100
- パフォーマンススコア: ${pageData.performance_score || 0}/100
- ビジネスインパクトスコア: ${pageData.business_impact_score || 0}/100
- 総合優先度スコア: ${pageData.total_priority_score || 0}/100

【主要検索クエリ（上位5つ）】
${pageData.top_queries || 'データなし'}

このページを改善して検索順位とCTRを向上させたいです。
上記の順位に応じた警告・制約を必ず遵守して、以下の形式で提案してください：

${suggestionFormat}`;

  return prompt;
}

/**
 * 週次自動分析を実行
 */
function runWeeklyAnalysis() {
  Logger.log('=== 週次自動分析開始 ===');
  
  try {
    // スコアリング実行
    calculateScores();
    
    // 優先度上位10ページを取得
    const topPages = getTopPriorityPagesFiltered(10);
    
    // レポート生成
    const report = generateWeeklyReport(topPages);
    
    // メール送信
    sendWeeklyReportEmail(report);
    
    // ログ記録
    logWeeklyAnalysis(topPages);
    
    Logger.log('=== 週次自動分析完了 ===');
    
  } catch (error) {
    Logger.log(`エラー: ${error.message}`);
    sendErrorEmail(error);
    throw error;
  }
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
// getTopPriorityPages() の修正版
// ============================================

/**
 * 優先度の高いページを取得（冷却期間考慮版）
 * 既存の関数を置き換えるか、新しい関数として追加
 * 
 * @param {number} limit - 取得件数
 * @param {boolean} includeCooling - 冷却中ページも含めるか（デフォルト: false）
 * @return {Array} ページ一覧
 */
function getTopPriorityPagesWithCoolingFilter(limit, includeCooling = false) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('統合データ');
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // 必要な列のインデックスを取得
    const urlCol = headers.indexOf('page_url') !== -1 ? headers.indexOf('page_url') : headers.indexOf('url');
    const titleCol = headers.indexOf('page_title') !== -1 ? headers.indexOf('page_title') : headers.indexOf('title');
    const scoreCol = headers.indexOf('total_priority_score');
    const exclusionCol = headers.indexOf('exclusion_reason');
    
    // データを配列に変換
    let pages = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      // 除外理由がある場合はスキップ
      if (exclusionCol !== -1 && row[exclusionCol]) {
        continue;
      }
      
      const pageUrl = row[urlCol];
      const score = scoreCol !== -1 ? row[scoreCol] : 0;
      
      // 冷却状態をチェック
      const coolingStatus = checkCoolingStatus(pageUrl);
      
      // 冷却中ページを除外する場合
      if (!includeCooling && coolingStatus.isCooling) {
        continue;
      }
      
      const page = {
        url: pageUrl,
        title: titleCol !== -1 ? row[titleCol] : '',
        score: score,
        coolingStatus: coolingStatus,
        isCooling: coolingStatus.isCooling
      };
      
      // 他のスコア情報も追加
      headers.forEach((header, idx) => {
        if (!page[header]) {
          page[header] = row[idx];
        }
      });
      
      pages.push(page);
    }
    
    // スコア降順でソート
    pages.sort((a, b) => (b.score || 0) - (a.score || 0));
    
    return pages.slice(0, limit);
    
  } catch (error) {
    Logger.log(`優先ページ取得エラー: ${error.message}`);
    return [];
  }
}


// ============================================
// 週次分析での冷却期間レポート
// ============================================

/**
 * 冷却中ページのサマリーを生成
 * runWeeklyAnalysis()から呼び出し
 * 
 * @return {Object} サマリー情報
 */
function getCoolingPagesSummary() {
  try {
    const taskSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('タスク管理');
    if (!taskSheet) {
      return { count: 0, pages: [], message: 'タスク管理シートがありません' };
    }
    
    const data = taskSheet.getDataRange().getValues();
    const headers = data[0];
    const urlCol = headers.indexOf('page_url');
    const typeCol = headers.indexOf('task_type');
    const statusCol = headers.indexOf('status');
    const completedDateCol = headers.indexOf('completed_date');
    const coolingDaysCol = headers.indexOf('cooling_days');
    
    const today = new Date();
    const coolingPages = new Map(); // URL -> 冷却情報
    
    for (let i = 1; i < data.length; i++) {
      const status = data[i][statusCol];
      const completedDate = data[i][completedDateCol];
      
      if (status === '完了' && completedDate) {
        const url = data[i][urlCol];
        const taskType = data[i][typeCol];
        const coolingDays = data[i][coolingDaysCol] || 30;
        
        const endDate = new Date(completedDate);
        endDate.setDate(endDate.getDate() + coolingDays);
        
        if (today < endDate) {
          const remainingDays = Math.ceil((endDate - today) / (1000 * 60 * 60 * 24));
          
          if (!coolingPages.has(url)) {
            coolingPages.set(url, {
              url: url,
              tasks: [],
              maxRemainingDays: 0
            });
          }
          
          const pageInfo = coolingPages.get(url);
          pageInfo.tasks.push({
            taskType: taskType,
            remainingDays: remainingDays,
            endDate: endDate
          });
          
          if (remainingDays > pageInfo.maxRemainingDays) {
            pageInfo.maxRemainingDays = remainingDays;
          }
        }
      }
    }
    
    const coolingList = Array.from(coolingPages.values());
    
    // 残り日数でソート
    coolingList.sort((a, b) => a.maxRemainingDays - b.maxRemainingDays);
    
    return {
      count: coolingList.length,
      pages: coolingList,
      nearExpiry: coolingList.filter(p => p.maxRemainingDays <= 7), // 1週間以内に解除
      longCooling: coolingList.filter(p => p.maxRemainingDays > 60) // 2ヶ月以上残り
    };
    
  } catch (error) {
    Logger.log(`冷却サマリー取得エラー: ${error.message}`);
    return { count: 0, pages: [], error: error.message };
  }
}


/**
 * 週次メールに冷却情報を追加
 * @param {string} emailBody - 既存のメール本文
 * @return {string} 冷却情報を追加したメール本文
 */
function addCoolingInfoToWeeklyEmail(emailBody) {
  const coolingSummary = getCoolingPagesSummary();
  
  if (coolingSummary.count === 0) {
    return emailBody;
  }
  
  let coolingSection = '\n\n━━━━━━━━━━━━━━━━━━━━\n';
  coolingSection += '⏳ 冷却期間中のページ\n';
  coolingSection += '━━━━━━━━━━━━━━━━━━━━\n\n';
  
  coolingSection += `現在 ${coolingSummary.count} ページが冷却期間中です。\n\n`;
  
  // もうすぐ解除されるページ
  if (coolingSummary.nearExpiry.length > 0) {
    coolingSection += '【もうすぐ解除（1週間以内）】\n';
    coolingSummary.nearExpiry.forEach(page => {
      const taskTypes = page.tasks.map(t => t.taskType).join(', ');
      coolingSection += `• ${page.url}\n`;
      coolingSection += `  → ${taskTypes}（あと${page.maxRemainingDays}日）\n`;
    });
    coolingSection += '\n';
  }
  
  // 長期冷却中のページ（タイトル変更など）
  if (coolingSummary.longCooling.length > 0) {
    coolingSection += '【長期冷却中（60日以上）】\n';
    coolingSummary.longCooling.forEach(page => {
      const taskTypes = page.tasks.map(t => t.taskType).join(', ');
      coolingSection += `• ${page.url}\n`;
      coolingSection += `  → ${taskTypes}（あと${page.maxRemainingDays}日）\n`;
    });
    coolingSection += '\n';
  }
  
  return emailBody + coolingSection;
}


// ============================================
// AI提案生成時の冷却期間フィルタリング
// ============================================

/**
 * ページに対するAI提案を生成（冷却期間考慮版）
 * SuggestionGenerator.gsから呼び出し
 * 
 * @param {string} pageUrl - ページURL
 * @param {Object} pageData - ページデータ
 * @return {Object} 提案情報（冷却情報含む）
 */
function generateSuggestionsWithCooling(pageUrl, pageData) {
  // 全タスク種別の冷却状態を取得
  const coolingStatus = checkCoolingStatus(pageUrl);
  
  // 冷却中でないタスク種別のみを提案対象に
  const availableTaskTypes = coolingStatus.availableTasks;
  const excludedTaskTypes = coolingStatus.coolingTasks.map(t => ({
    taskType: t.taskType,
    remainingDays: t.remainingDays,
    endDate: t.endDate
  }));
  
  return {
    pageUrl: pageUrl,
    pageData: pageData,
    availableTaskTypes: availableTaskTypes,
    excludedTaskTypes: excludedTaskTypes,
    coolingMessage: excludedTaskTypes.length > 0 
      ? `※ ${excludedTaskTypes.map(t => `${t.taskType}(あと${t.remainingDays}日)`).join(', ')} は冷却期間中のため除外`
      : ''
  };
}


// ============================================
// リライト提案の優先順位付けロジック
// ============================================

/**
 * 提案に推奨順位を付与
 * @param {Array} suggestions - 提案リスト
 * @return {Array} 推奨順位付き提案リスト
 */
function assignPriorityRank(suggestions) {
  // 効果の優先度でソート
  const priorityOrder = {
    'タイトル変更': 1,
    'メタディスクリプション': 2,
    'H1変更': 3,
    'H2追加': 4,
    'H2変更': 5,
    'Q&A追加': 6,
    '本文追加': 7,
    '画像追加': 8,
    '動画追加': 9,
    '内部リンク追加': 10,
    'その他': 99
  };
  
  // ソート
  suggestions.sort((a, b) => {
    const orderA = priorityOrder[a.taskType] || 99;
    const orderB = priorityOrder[b.taskType] || 99;
    return orderA - orderB;
  });
  
  // 推奨順位を付与
  suggestions.forEach((suggestion, index) => {
    suggestion.priorityRank = index + 1;
  });
  
  return suggestions;
}


// ============================================
// テスト関数
// ============================================

/**
 * 冷却期間連携のテスト
 */
function testCoolingIntegration() {
  Logger.log('=== 冷却期間連携テスト開始 ===');
  
  // 1. 冷却サマリー取得テスト
  Logger.log('1. 冷却サマリー取得テスト');
  const summary = getCoolingPagesSummary();
  Logger.log(`冷却中ページ数: ${summary.count}`);
  Logger.log(`もうすぐ解除: ${summary.nearExpiry?.length || 0}`);
  Logger.log(`長期冷却中: ${summary.longCooling?.length || 0}`);
  
  // 2. 優先ページ取得テスト（冷却フィルター付き）
  Logger.log('2. 優先ページ取得テスト');
  const pagesWithFilter = getTopPriorityPagesWithCoolingFilter(5, false);
  Logger.log(`冷却除外後: ${pagesWithFilter.length}件`);
  
  const pagesWithoutFilter = getTopPriorityPagesWithCoolingFilter(5, true);
  Logger.log(`冷却含む: ${pagesWithoutFilter.length}件`);
  
  // 3. 推奨順位付けテスト
  Logger.log('3. 推奨順位付けテスト');
  const testSuggestions = [
    { taskType: '本文追加', detail: 'テスト1' },
    { taskType: 'タイトル変更', detail: 'テスト2' },
    { taskType: 'Q&A追加', detail: 'テスト3' }
  ];
  const ranked = assignPriorityRank(testSuggestions);
  ranked.forEach(s => Logger.log(`${s.priorityRank}. ${s.taskType}`));
  
  Logger.log('=== 冷却期間連携テスト完了 ===');
}

// ============================================
// Scoring.gs 追記: 投稿日フィルタリング連携
// 追記場所: Scoring.gsの最下部
// ============================================

/**
 * 3ヶ月未満の記事を除外した優先ページ取得（拡張版）
 * 冷却期間 + 投稿日フィルターの両方を適用
 * @param {number} limit - 取得件数
 * @return {Array} フィルタリング済みページ
 */
function getTopPriorityPagesFiltered(limit = 10) {
  // 既存の優先ページ取得（多めに取得）
  let pages = [];
  
  if (typeof getTopPriorityPagesWithCooling === 'function') {
    pages = getTopPriorityPagesWithCooling(limit * 3);
  } else if (typeof getTopPriorityPages === 'function') {
    pages = getTopPriorityPages(limit * 3);
  } else {
    Logger.log('警告: 優先ページ取得関数が見つかりません');
    return [];
  }
  
  // 投稿日フィルター（3ヶ月未満を除外）
  if (typeof filterPagesByPublishDate === 'function') {
    const result = filterPagesByPublishDate(pages);
    Logger.log('投稿日フィルター: ' + result.message);
    pages = result.filtered;
  }
  
  // 件数制限
  return pages.slice(0, limit);
}


/**
 * リライト提案時の総合チェック
 * @param {string} pageUrl - ページURL
 * @param {string} taskType - タスク種別
 * @return {Object} チェック結果
 */
function canSuggestRewrite(pageUrl, taskType) {
  const result = {
    canSuggest: true,
    reasons: []
  };
  
  // 1. 冷却期間チェック
  if (typeof shouldExcludeFromSuggestion === 'function') {
    if (shouldExcludeFromSuggestion(pageUrl, taskType)) {
      result.canSuggest = false;
      result.reasons.push('冷却期間中（' + taskType + '）');
    }
  }
  
  // 2. 投稿日チェック（3ヶ月未満）
  if (typeof shouldExcludeByPublishDate === 'function') {
    if (shouldExcludeByPublishDate(pageUrl)) {
      result.canSuggest = false;
      result.reasons.push('投稿から3ヶ月未満');
    }
  }
  
  return result;
}


/**
 * 週次分析用: フィルター適用済みサマリー
 * @return {Object} サマリー情報
 */
function getFilteredPagesSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('統合データ');
  
  if (!sheet) {
    return { error: '統合データシートが見つかりません' };
  }
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rewriteOkColIndex = headers.indexOf('リライト可能');
  const lastRow = sheet.getLastRow();
  
  if (rewriteOkColIndex === -1 || lastRow <= 1) {
    return {
      totalPages: lastRow - 1,
      rewriteReady: lastRow - 1,
      tooNew: 0,
      message: '投稿日フィルター未設定'
    };
  }
  
  const data = sheet.getRange(2, rewriteOkColIndex + 1, lastRow - 1, 1).getValues();
  
  let rewriteReady = 0;
  let tooNew = 0;
  
  for (const row of data) {
    if (row[0] === '○') {
      rewriteReady++;
    } else if (row[0] === '×') {
      tooNew++;
    }
  }
  
  return {
    totalPages: lastRow - 1,
    rewriteReady: rewriteReady,
    tooNew: tooNew,
    unknown: (lastRow - 1) - rewriteReady - tooNew,
    message: `${rewriteReady}件がリライト対象、${tooNew}件が3ヶ月未満`
  };
}

// ============================================
// Day 22追加: トレンド判定機能（GyronSEO 4週間分析）
// ============================================

/**
 * GyronSEO_RAWから過去4週間の順位データを取得
 * @param {string} pageUrl - ページURL（パス形式）
 * @param {string} targetKeyword - ターゲットキーワード
 * @return {Object} 4週間の順位データ
 */
function getGyronRankHistory(pageUrl, targetKeyword) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('GyronSEO_RAW');
  
  if (!sheet) {
    return { success: false, error: 'GyronSEO_RAWシートが見つかりません' };
  }
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  // ヘッダーから日付列を抽出（Date型の列）
  const dateCols = [];
  for (let col = 0; col < headers.length; col++) {
    if (headers[col] instanceof Date) {
      dateCols.push({ col: col, date: headers[col] });
    }
  }
  
  // 日付降順でソート（最新が先頭）
  dateCols.sort((a, b) => b.date - a.date);
  
  // 最新4週間分を取得（週1回更新想定）
  const recentDates = dateCols.slice(0, 4);
  
  if (recentDates.length < 2) {
    return { success: false, error: '十分な履歴データがありません（2週間以上必要）' };
  }
  
  // 該当ページ・KWの順位を検索
  const normalizedUrl = normalizeUrlPath(pageUrl);
  const normalizedKW = (targetKeyword || '').toLowerCase().trim();
  
  let matchedRow = null;
  
  for (let i = 1; i < data.length; i++) {
    const rowKW = (data[i][0] || '').toString().toLowerCase().trim();
    const rowUrl = normalizeUrlPath(data[i][1] || '');
    
    // KWとURLの両方でマッチング
    const kwMatch = normalizedKW && rowKW.includes(normalizedKW);
    const urlMatch = normalizedUrl && (rowUrl === normalizedUrl || rowUrl.includes(normalizedUrl));
    
    if (kwMatch || urlMatch) {
      matchedRow = data[i];
      break;
    }
  }
  
  if (!matchedRow) {
    return { success: false, error: 'マッチするデータが見つかりません' };
  }
  
  // 各週の順位を取得
  const weeklyRanks = recentDates.map(d => {
    const rank = matchedRow[d.col];
    let rankNum = null;
    
    if (rank !== '' && rank !== null) {
      if (String(rank).includes('圏外')) {
        rankNum = 101;
      } else {
        rankNum = parseFloat(rank) || null;
      }
    }
    
    return {
      date: d.date,
      rank: rankNum
    };
  });
  
  return {
    success: true,
    keyword: matchedRow[0],
    url: matchedRow[1],
    weeklyRanks: weeklyRanks
  };
}

/**
 * 4週間のトレンドを判定
 * @param {Array} weeklyRanks - 週次順位データ配列
 * @return {Object} トレンド判定結果
 */
function analyzeRankTrend(weeklyRanks) {
  // 有効な順位データのみ抽出
  const validRanks = weeklyRanks.filter(w => w.rank !== null && w.rank > 0);
  
  if (validRanks.length < 2) {
    return {
      trend: 'unknown',
      trendLabel: '不明',
      priorityModifier: 0,
      message: '十分なデータがありません'
    };
  }
  
  // 最新と4週間前を比較
  const latestRank = validRanks[0].rank;
  const oldestRank = validRanks[validRanks.length - 1].rank;
  const rankChange = oldestRank - latestRank; // 正=改善、負=悪化
  
  // 週ごとの変動幅を計算
  let maxWeeklyChange = 0;
  for (let i = 0; i < validRanks.length - 1; i++) {
    const change = Math.abs(validRanks[i].rank - validRanks[i + 1].rank);
    if (change > maxWeeklyChange) {
      maxWeeklyChange = change;
    }
  }
  
  // 全体の変動幅
  const allRanks = validRanks.map(w => w.rank);
  const minRank = Math.min(...allRanks);
  const maxRank = Math.max(...allRanks);
  const totalVariation = maxRank - minRank;
  
  // トレンド判定
  let trend, trendLabel, priorityModifier, message;
  
  if (maxWeeklyChange >= 6) {
    // 不安定: 週ごとに±6位以上の乱高下
    trend = 'unstable';
    trendLabel = '不安定';
    priorityModifier = -20; // 優先度下げ
    message = `週ごとに${maxWeeklyChange}位の変動あり。様子見推奨`;
  } else if (rankChange >= 6) {
    // 上昇傾向: 4週間で6位以上改善
    trend = 'improving';
    trendLabel = '上昇傾向';
    priorityModifier = -15; // 優先度下げ（好調なので触らない）
    message = `4週間で${rankChange}位改善中。現状維持推奨`;
  } else if (rankChange <= -6) {
    // 下降傾向: 4週間で6位以上悪化
    trend = 'declining';
    trendLabel = '下降傾向';
    priorityModifier = 15; // 優先度上げ
    message = `4週間で${Math.abs(rankChange)}位悪化。要リライト`;
  } else if (totalVariation <= 5) {
    // 安定: 4週間の変動幅±5位以内
    trend = 'stable';
    trendLabel = '安定';
    priorityModifier = 0; // 通常スコアリング
    message = `順位安定（変動幅${totalVariation}位）`;
  } else {
    // その他（小幅変動）
    trend = 'stable';
    trendLabel = '安定';
    priorityModifier = 0;
    message = `小幅変動（変動幅${totalVariation}位）`;
  }
  
  return {
    trend: trend,
    trendLabel: trendLabel,
    priorityModifier: priorityModifier,
    message: message,
    latestRank: latestRank,
    oldestRank: oldestRank,
    rankChange: rankChange,
    maxWeeklyChange: maxWeeklyChange,
    totalVariation: totalVariation
  };
}

/**
 * URLをパス形式に正規化
 */
function normalizeUrlPath(url) {
  if (!url) return '';
  
  let path = String(url).toLowerCase();
  
  // フルURLからパスを抽出
  if (path.includes('://')) {
    try {
      const urlObj = new URL(path);
      path = urlObj.pathname;
    } catch (e) {
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
 * ページのトレンドを取得してスコア修正を適用
 * @param {string} pageUrl - ページURL
 * @param {string} targetKeyword - ターゲットキーワード
 * @param {number} baseScore - 基本スコア
 * @return {Object} トレンド情報と修正後スコア
 */
function applyTrendModifier(pageUrl, targetKeyword, baseScore) {
  const history = getGyronRankHistory(pageUrl, targetKeyword);
  
  if (!history.success) {
    return {
      finalScore: baseScore,
      trend: null,
      message: history.error
    };
  }
  
  const trendAnalysis = analyzeRankTrend(history.weeklyRanks);
  
  // スコア修正を適用（0-100の範囲内に収める）
  let finalScore = baseScore + trendAnalysis.priorityModifier;
  finalScore = Math.max(0, Math.min(100, finalScore));
  
  return {
    finalScore: finalScore,
    baseScore: baseScore,
    modifier: trendAnalysis.priorityModifier,
    trend: trendAnalysis.trend,
    trendLabel: trendAnalysis.trendLabel,
    message: trendAnalysis.message,
    weeklyRanks: history.weeklyRanks
  };
}

/**
 * トレンド判定のテスト
 */
function testTrendAnalysis() {
  Logger.log('=== トレンド判定テスト ===');
  
  // テスト用URL（統合データから取得）
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('統合データ');
  
  if (!sheet) {
    Logger.log('統合データシートが見つかりません');
    return;
  }
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const urlIdx = headers.indexOf('page_url');
  const kwIdx = headers.indexOf('target_keyword');
  
  // 最初の5ページでテスト
  for (let i = 1; i <= Math.min(5, data.length - 1); i++) {
    const url = data[i][urlIdx];
    const kw = data[i][kwIdx];
    
    Logger.log(`\n--- テスト${i}: ${url} ---`);
    Logger.log(`ターゲットKW: ${kw}`);
    
    const result = applyTrendModifier(url, kw, 50);
    
    if (result.trend) {
      Logger.log(`トレンド: ${result.trendLabel}`);
      Logger.log(`基本スコア: ${result.baseScore} → 修正後: ${result.finalScore} (${result.modifier >= 0 ? '+' : ''}${result.modifier})`);
      Logger.log(`メッセージ: ${result.message}`);
      
      if (result.weeklyRanks) {
        Logger.log('週次順位: ' + result.weeklyRanks.map(w => w.rank || 'N/A').join(' → '));
      }
    } else {
      Logger.log(`エラー: ${result.message}`);
    }
  }
  
  Logger.log('\n=== テスト完了 ===');
}

function checkPositionDistribution() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('統合データ');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  // 順位列を探す
  const posIdx = headers.indexOf('gyron_position') !== -1 ? 
                 headers.indexOf('gyron_position') : 
                 headers.indexOf('avg_position');
  
  Logger.log('使用する順位列: ' + headers[posIdx]);
  
  const distribution = { 
    '1-3位': 0, 
    '4-10位': 0, 
    '11-20位': 0, 
    '21-30位': 0, 
    '31-50位': 0, 
    '51位以上': 0,
    '順位なし': 0
  };
  
  for (let i = 1; i < data.length; i++) {
    const pos = parseFloat(data[i][posIdx]) || 0;
    
    if (pos === 0 || isNaN(pos)) distribution['順位なし']++;
    else if (pos <= 3) distribution['1-3位']++;
    else if (pos <= 10) distribution['4-10位']++;
    else if (pos <= 20) distribution['11-20位']++;
    else if (pos <= 30) distribution['21-30位']++;
    else if (pos <= 50) distribution['31-50位']++;
    else distribution['51位以上']++;
  }
  
  Logger.log('=== 順位分布 ===');
  Object.keys(distribution).forEach(k => {
    Logger.log(k + ': ' + distribution[k] + 'ページ');
  });
}

function testPositionScoreChange() {
  Logger.log('=== 順位スコア修正確認テスト ===');
  
  // 各順位帯のスコアを確認
  const testPositions = [1, 3, 5, 10, 11, 15, 20, 25, 30];
  
  testPositions.forEach(pos => {
    let score = 0;
    
    // ★ここが修正後のロジックと一致しているか確認
    if (pos >= 1 && pos <= 3) {
      score = 10;
    } else if (pos >= 4 && pos <= 10) {
      score = 75;
    } else if (pos >= 11 && pos <= 20) {
      score = 100;
    } else if (pos >= 21 && pos <= 30) {
      score = 95;
    } else if (pos >= 31 && pos <= 50) {
      score = 40;
    } else {
      score = 20;
    }
    
    Logger.log(pos + '位 → スコア: ' + score + '点');
  });
  
  Logger.log('\n期待値: 11-20位が100点、21-30位が95点、4-10位が75点、1-3位が10点');
}

function debugTopPagesScore() {
  Logger.log('=== 上位5ページのスコア詳細 ===');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('統合データ');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  // 列インデックス取得
  const urlIdx = headers.indexOf('page_url');
  const posIdx = headers.indexOf('gyron_position');
  const oppIdx = headers.indexOf('opportunity_score');
  const totalIdx = headers.indexOf('total_priority_score');
  
  Logger.log('列: page_url=' + urlIdx + ', gyron_position=' + posIdx + ', opportunity_score=' + oppIdx + ', total_priority_score=' + totalIdx);
  
  // 上位5ページの詳細
  const pages = [];
  for (let i = 1; i < data.length; i++) {
    pages.push({
      url: data[i][urlIdx],
      position: data[i][posIdx],
      opportunityScore: data[i][oppIdx],
      totalScore: data[i][totalIdx]
    });
  }
  
  // totalScoreでソート
  pages.sort((a, b) => (b.totalScore || 0) - (a.totalScore || 0));
  
  Logger.log('\n--- 上位10ページ ---');
  pages.slice(0, 10).forEach((p, i) => {
    Logger.log((i+1) + '位: ' + p.url);
    Logger.log('   順位: ' + p.position + '位');
    Logger.log('   opportunity_score: ' + p.opportunityScore);
    Logger.log('   total_priority_score: ' + p.totalScore);
  });
  
  Logger.log('\n--- 11-20位のページ ---');
  const rank11to20 = pages.filter(p => p.position >= 11 && p.position <= 20);
  rank11to20.slice(0, 5).forEach((p, i) => {
    Logger.log((i+1) + '. ' + p.url);
    Logger.log('   順位: ' + p.position + '位');
    Logger.log('   opportunity_score: ' + p.opportunityScore);
    Logger.log('   total_priority_score: ' + p.totalScore);
  });
}

function debugGSCQueryData() {
  Logger.log('=== GSCクエリデータ確認 ===');
  
  const testUrl = '/ipad-mini-cheap-buy-methods';
  
  // getQueryDataForPage関数を呼び出し
  const queryData = getQueryDataForPage(testUrl);
  
  if (!queryData || queryData.length === 0) {
    Logger.log('データなし');
    return;
  }
  
  Logger.log('取得件数: ' + queryData.length);
  Logger.log('\n--- 先頭10件 ---');
  
  queryData.slice(0, 10).forEach((q, i) => {
    Logger.log((i+1) + '. ' + q.query + ' | 表示: ' + q.impressions + ' | クリック: ' + q.clicks);
  });
}