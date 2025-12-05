/**
 * CompetitorChat.gs
 * 競合分析チャット機能
 * 
 * Day 15実装: 2025年12月1日
 * 
 * 機能:
 * - 外部サイトコンテンツ取得
 * - 競合分析シートからURL取得
 * - ハイブリッド方式（保存データ優先）
 * - 複数競合サイト一括比較
 * - 差分分析
 */

// ============================================================
// 定数
// ============================================================

const COMPETITOR_CHAT_VERSION = '1.0';

// 自社ドメイン
const OWN_DOMAIN = 'smaho-tap.com';

// コンテンツ取得のタイムアウト（ミリ秒）
const FETCH_TIMEOUT = 30000;

// ============================================================
// タスク1: 外部サイトコンテンツ取得
// ============================================================

/**
 * 外部サイトのコンテンツを取得・解析
 * @param {string} url - 取得対象のURL
 * @return {Object} 解析結果
 */
function fetchCompetitorContent(url) {
  const startTime = new Date();
  
  try {
    // URL検証
    if (!url || !url.startsWith('http')) {
      return {
        success: false,
        error: 'Invalid URL',
        url: url
      };
    }
    
    // HTTP リクエスト
    const options = {
      method: 'GET',
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
        'Accept-Language': 'ja,en-US;q=0.9,en;q=0.8'
      },
      muteHttpExceptions: true,
      followRedirects: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode !== 200) {
      return {
        success: false,
        error: `HTTP Error: ${responseCode}`,
        url: url
      };
    }
    
    const html = response.getContentText();
    
    // HTML解析
    const content = parseCompetitorHTML(html, url);
    
    const endTime = new Date();
    content.fetchTime = (endTime - startTime) / 1000;
    content.success = true;
    content.url = url;
    
    Logger.log(`コンテンツ取得成功: ${url} (${content.fetchTime}秒)`);
    
    return content;
    
  } catch (error) {
    Logger.log(`コンテンツ取得エラー: ${url} - ${error.message}`);
    return {
      success: false,
      error: error.message,
      url: url
    };
  }
}

/**
 * HTMLをパースして構造化データに変換
 * @param {string} html - HTMLコンテンツ
 * @param {string} url - 元のURL
 * @return {Object} パース結果
 */
function parseCompetitorHTML(html, url) {
  const result = {
    title: '',
    metaDescription: '',
    h1: [],
    h2: [],
    h3: [],
    headingStructure: [],
    wordCount: 0,
    imageCount: 0,
    internalLinkCount: 0,
    externalLinkCount: 0,
    hasToc: false,
    hasFaq: false,
    hasVideo: false,
    domain: extractDomainFromUrl(url)
  };
  
  try {
    // タイトル抽出
    const titleMatch = html.match(/<title[^>]*>([\s\S]*?)<\/title>/i);
    if (titleMatch) {
      result.title = cleanText(titleMatch[1]);
    }
    
    // メタディスクリプション抽出
    const metaDescMatch = html.match(/<meta[^>]+name=["']description["'][^>]+content=["']([^"']*)["']/i);
    if (!metaDescMatch) {
      const metaDescMatch2 = html.match(/<meta[^>]+content=["']([^"']*)["'][^>]+name=["']description["']/i);
      if (metaDescMatch2) {
        result.metaDescription = cleanText(metaDescMatch2[1]);
      }
    } else {
      result.metaDescription = cleanText(metaDescMatch[1]);
    }
    
    // 見出し抽出（H1, H2, H3）
    const h1Matches = html.match(/<h1[^>]*>([\s\S]*?)<\/h1>/gi) || [];
    const h2Matches = html.match(/<h2[^>]*>([\s\S]*?)<\/h2>/gi) || [];
    const h3Matches = html.match(/<h3[^>]*>([\s\S]*?)<\/h3>/gi) || [];
    
    result.h1 = h1Matches.map(h => cleanText(h.replace(/<[^>]+>/g, '')));
    result.h2 = h2Matches.map(h => cleanText(h.replace(/<[^>]+>/g, '')));
    result.h3 = h3Matches.map(h => cleanText(h.replace(/<[^>]+>/g, '')));
    
    // 見出し構造（順序保持）
    const headingPattern = /<(h[1-3])[^>]*>([\s\S]*?)<\/\1>/gi;
    let match;
    while ((match = headingPattern.exec(html)) !== null) {
      result.headingStructure.push({
        level: match[1].toUpperCase(),
        text: cleanText(match[2].replace(/<[^>]+>/g, ''))
      });
    }
    
    // 本文テキスト抽出（スクリプト・スタイル除去）
    let bodyText = html;
    bodyText = bodyText.replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '');
    bodyText = bodyText.replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '');
    bodyText = bodyText.replace(/<nav[^>]*>[\s\S]*?<\/nav>/gi, '');
    bodyText = bodyText.replace(/<header[^>]*>[\s\S]*?<\/header>/gi, '');
    bodyText = bodyText.replace(/<footer[^>]*>[\s\S]*?<\/footer>/gi, '');
    bodyText = bodyText.replace(/<aside[^>]*>[\s\S]*?<\/aside>/gi, '');
    bodyText = bodyText.replace(/<[^>]+>/g, ' ');
    bodyText = bodyText.replace(/\s+/g, ' ').trim();
    
    result.wordCount = bodyText.length;
    
    // 画像数カウント
    const imgMatches = html.match(/<img[^>]+>/gi) || [];
    result.imageCount = imgMatches.length;
    
    // リンクカウント
    const linkMatches = html.match(/<a[^>]+href=["']([^"']+)["'][^>]*>/gi) || [];
    linkMatches.forEach(link => {
      const hrefMatch = link.match(/href=["']([^"']+)["']/i);
      if (hrefMatch) {
        const href = hrefMatch[1];
        if (href.startsWith('http') && !href.includes(result.domain)) {
          result.externalLinkCount++;
        } else if (!href.startsWith('mailto:') && !href.startsWith('tel:') && !href.startsWith('#')) {
          result.internalLinkCount++;
        }
      }
    });
    
    // 目次の有無（よくあるパターン）
    result.hasToc = /id=["']?toc["']?|class=["'][^"']*toc[^"']*["']|目次|table.of.contents/i.test(html);
    
    // FAQの有無
    result.hasFaq = /faq|よくある質問|q&a|itemtype=["'][^"']*faqpage/i.test(html);
    
    // 動画の有無
    result.hasVideo = /<video|youtube\.com|youtu\.be|vimeo\.com|<iframe[^>]+src=["'][^"']*(?:youtube|vimeo)/i.test(html);
    
  } catch (error) {
    Logger.log(`HTMLパースエラー: ${error.message}`);
  }
  
  return result;
}

/**
 * テキストをクリーンアップ
 * @param {string} text - 元のテキスト
 * @return {string} クリーンなテキスト
 */
function cleanText(text) {
  if (!text) return '';
  return text
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * URLからドメインを抽出
 * @param {string} url - URL
 * @return {string} ドメイン
 */
function extractDomainFromUrl(url) {
  try {
    const match = url.match(/^(?:https?:\/\/)?(?:www\.)?([^\/]+)/i);
    return match ? match[1] : url;
  } catch (e) {
    return url;
  }
}

// ============================================================
// タスク2: 競合分析シートからURL取得
// ============================================================

/**
 * 競合分析シートからキーワードに対応する競合情報を取得（修正版v2）
 * @param {string} keyword - 検索キーワード
 * @return {Object} 競合分析結果
 */
function findCompetitorAnalysis(keyword) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('競合分析');
  
  if (!sheet) {
    return { found: false, message: '競合分析シートが見つかりません' };
  }
  
  const data = sheet.getDataRange().getValues();
  const normalizedKeyword = normalizeKeyword(keyword);
  
  // 自社ドメイン（設定シートから取得、なければデフォルト）
  const ownDomain = getOwnDomain();
  
  // キーワードを検索（列1 = B列 = target_keyword）
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowKeyword = normalizeKeyword(String(row[1] || ''));
    
    if (rowKeyword === normalizedKeyword || rowKeyword.includes(normalizedKeyword) || normalizedKeyword.includes(rowKeyword)) {
      const analysisDate = row[3];
      const ownSiteDA = row[4];
      let ownSiteRank = row[5]; // まずシートの値を確認
      const winnableScore = row[33];
      const competitorLevel = row[34];
      
      // 上位サイト情報を取得（列6から開始、2列ずつ）
      const topSites = [];
      let detectedOwnSiteRank = null;
      
      for (let j = 1; j <= 10; j++) {
        const urlIndex = 6 + (j - 1) * 2;
        const daIndex = urlIndex + 1;
        
        const url = row[urlIndex];
        const da = row[daIndex];
        
        if (url && String(url).trim() !== '') {
          const urlStr = String(url).trim();
          
          // 自社サイトかどうかをチェック
          const isOwnSite = urlStr.includes(ownDomain);
          
          if (isOwnSite && !detectedOwnSiteRank) {
            detectedOwnSiteRank = j;
          }
          
          topSites.push({
            rank: j,
            url: urlStr,
            da: da || 'N/A',
            isOwnSite: isOwnSite
          });
        }
      }
      
      // own_site_current_rankが空の場合、検出した順位を使用
      if (!ownSiteRank && detectedOwnSiteRank) {
        ownSiteRank = detectedOwnSiteRank;
      } else if (!ownSiteRank) {
        ownSiteRank = '圏外';
      }
      
      // データ鮮度チェック
      let freshness = 'unknown';
      let daysSinceAnalysis = null;
      
      if (analysisDate) {
        const analysisDateObj = new Date(analysisDate);
        const now = new Date();
        daysSinceAnalysis = Math.floor((now - analysisDateObj) / (1000 * 60 * 60 * 24));
        
        if (daysSinceAnalysis <= 7) {
          freshness = 'fresh';
        } else if (daysSinceAnalysis <= 30) {
          freshness = 'recent';
        } else {
          freshness = 'stale';
        }
      }
      
      return {
        found: true,
        keyword: row[1],
        pageUrl: row[2],
        analysisDate: analysisDate,
        daysSinceAnalysis: daysSinceAnalysis,
        freshness: freshness,
        ownSiteDA: ownSiteDA,
        ownSiteRank: ownSiteRank,
        winnableScore: winnableScore,
        competitorLevel: competitorLevel,
        topSites: topSites
      };
    }
  }
  
  return { found: false, message: `キーワード「${keyword}」の競合分析データが見つかりません` };
}

/**
 * 自社ドメインを取得
 * @return {string} 自社ドメイン
 */
function getOwnDomain() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = ss.getSheetByName('設定・マスタ');
    
    if (settingsSheet) {
      const data = settingsSheet.getDataRange().getValues();
      for (let i = 0; i < data.length; i++) {
        if (data[i][0] === 'OWN_DOMAIN' || data[i][0] === 'SITE_DOMAIN') {
          return data[i][1];
        }
      }
    }
  } catch (e) {
    Logger.log('設定シートからドメイン取得エラー: ' + e.message);
  }
  
  // デフォルト値
  return 'smaho-tap.com';
}

/**
 * キーワードを正規化（表記ゆれ対応）
 * @param {string} keyword - キーワード
 * @return {string} 正規化されたキーワード
 */
function normalizeKeyword(keyword) {
  if (!keyword) return '';
  return String(keyword)
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .replace(/　/g, ' ')  // 全角スペース→半角
    .trim();
}

/**
 * ページURLから関連するキーワードを検索
 * @param {string} pageUrl - ページURL
 * @return {Array} 関連キーワードリスト
 */
function findKeywordsForPage(pageUrl) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('ターゲットKW分析');
  
  if (!sheet) {
    return [];
  }
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const urlIndex = headers.indexOf('page_url');
  const keywordIndex = headers.indexOf('target_keyword');
  
  if (urlIndex === -1 || keywordIndex === -1) {
    return [];
  }
  
  // URLを正規化
  const normalizedPageUrl = normalizeUrl(pageUrl);
  
  const keywords = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowUrl = normalizeUrl(String(row[urlIndex] || ''));
    
    if (rowUrl === normalizedPageUrl) {
      keywords.push(row[keywordIndex]);
    }
  }
  
  return keywords;
}

/**
 * URLを正規化
 * @param {string} url - URL
 * @return {string} 正規化されたURL
 */
function normalizeUrl(url) {
  if (!url) return '';
  return String(url)
    .replace(/^https?:\/\/[^\/]+/, '')  // ドメイン部分を削除
    .replace(/\/$/, '')  // 末尾スラッシュ削除
    .replace(/#.*$/, '');  // アンカー削除
}

// ============================================================
// タスク3: 意図分析の拡張
// ============================================================

/**
 * 競合分析関連の意図を分析
 * @param {string} message - ユーザーメッセージ
 * @return {Object} 意図分析結果
 */
function analyzeCompetitorIntent(message) {
  const lowerMessage = message.toLowerCase();
  
  // パターン1: URL直接指定
  const urlPattern = /https?:\/\/[^\s]+/g;
  const urlMatches = message.match(urlPattern);
  
  if (urlMatches && urlMatches.length > 0) {
    // 自社サイトのURLかどうかチェック
    const externalUrls = urlMatches.filter(url => !url.includes(OWN_DOMAIN));
    
    if (externalUrls.length > 0) {
      return {
        type: 'url_direct',
        urls: externalUrls,
        action: 'fetch_and_analyze'
      };
    }
  }
  
  // パターン2: キーワード指定で競合分析
  const competitorPatterns = [
    /「([^」]+)」(?:で|の)?(?:競合|上位|ライバル)/,
    /『([^』]+)』(?:で|の)?(?:競合|上位|ライバル)/,
    /(.+?)(?:で|の)(?:競合|上位サイト|ライバル)(?:と|を)?(?:比較|分析)/,
    /(.+?)(?:の)?(?:1位|2位|3位|トップ|上位)(?:サイト)?(?:と|を)?(?:比較|分析)/,
    /競合(?:分析|比較)(?:して|を).*?(?:「|『)?([^」』\s]+)(?:」|』)?/
  ];
  
  for (const pattern of competitorPatterns) {
    const match = message.match(pattern);
    if (match) {
      return {
        type: 'keyword_competitor',
        keyword: match[1].trim(),
        action: 'find_and_compare'
      };
    }
  }
  
  // パターン3: 差分分析リクエスト
  const diffPatterns = [
    /(?:差分|違い|差|ギャップ)(?:を)?(?:分析|比較|教えて)/,
    /(?:何が)?(?:足りない|不足|欠けている)/,
    /(?:上位サイト|競合)(?:に)?(?:あって|ある).*(?:ない|無い)/
  ];
  
  for (const pattern of diffPatterns) {
    if (pattern.test(message)) {
      return {
        type: 'diff_analysis',
        action: 'compare_content'
      };
    }
  }
  
  // パターン4: 一般的な競合分析リクエスト
  if (/競合|ライバル|上位サイト/.test(message)) {
    return {
      type: 'general_competitor',
      action: 'suggest_options'
    };
  }
  
  return {
    type: 'none',
    action: null
  };
}

// ============================================================
// タスク4: ハイブリッド方式実装
// ============================================================

/**
 * 競合分析リクエストを処理（ハイブリッド方式）
 * @param {string} message - ユーザーメッセージ
 * @param {string} currentPageUrl - 現在分析中のページURL（あれば）
 * @return {Object} 処理結果
 */
function handleCompetitorRequest(message, currentPageUrl) {
  const intent = analyzeCompetitorIntent(message);
  
  switch (intent.type) {
    case 'url_direct':
      // パターン1: URL直接指定 → 即座に分析
      return handleDirectUrlAnalysis(intent.urls);
      
    case 'keyword_competitor':
      // パターン2: キーワード指定 → 保存データ優先
      return handleKeywordCompetitorAnalysis(intent.keyword);
      
    case 'diff_analysis':
      // パターン3: 差分分析
      return handleDiffAnalysis(currentPageUrl);
      
    case 'general_competitor':
      // パターン4: 一般的な競合分析
      return {
        type: 'options',
        message: '競合分析の方法を選んでください：\n\n' +
          '1. **キーワード指定**: 「〇〇で上位サイトと比較」\n' +
          '2. **URL直接指定**: 競合サイトのURLを貼り付け\n' +
          '3. **現在のページ**: このページの競合を分析\n\n' +
          'どの方法で分析しますか？'
      };
      
    default:
      return {
        type: 'not_competitor',
        message: null
      };
  }
}

/**
 * URL直接指定の分析
 * @param {Array} urls - 分析対象URL配列
 * @return {Object} 分析結果
 */
function handleDirectUrlAnalysis(urls) {
  const results = [];
  
  for (const url of urls) {
    const content = fetchCompetitorContent(url);
    if (content.success) {
      results.push(content);
    } else {
      results.push({
        url: url,
        error: content.error
      });
    }
  }
  
  return {
    type: 'direct_analysis',
    results: results
  };
}

/**
 * キーワード指定の競合分析（保存データ優先）
 * @param {string} keyword - キーワード
 * @return {Object} 分析結果
 */
function handleKeywordCompetitorAnalysis(keyword) {
  // 保存データを検索
  const savedData = findCompetitorAnalysis(keyword);
  
  if (savedData.found) {
    // データの鮮度をチェック
    if (savedData.isStale) {
      return {
        type: 'stale_data',
        data: savedData,
        message: `「${keyword}」の競合データがありますが、${savedData.daysSinceAnalysis}日前のデータです。\n\n` +
          '選択肢:\n' +
          '1. **保存データを使用**（高速・無料）\n' +
          '2. **最新データを取得**（1-2分・API利用）\n\n' +
          'どちらを使用しますか？'
      };
    }
    
    return {
      type: 'saved_data',
      data: savedData,
      message: null
    };
  } else {
    return {
      type: 'no_data',
      keyword: keyword,
      message: `「${keyword}」の競合データが見つかりません。\n\n` +
        '選択肢:\n' +
        '1. **競合分析を実行**（API利用、$0.006）\n' +
        '2. **類似キーワードで検索**\n' +
        '3. **URLを直接指定**\n\n' +
        'どうしますか？'
    };
  }
}

/**
 * 差分分析
 * @param {string} pageUrl - 自社ページURL
 * @return {Object} 分析結果
 */
function handleDiffAnalysis(pageUrl) {
  if (!pageUrl) {
    return {
      type: 'need_page',
      message: '差分分析を行うページを指定してください。\n' +
        '例: 「/iphone-insurance-comparison の差分を分析」'
    };
  }
  
  // ページに関連するキーワードを取得
  const keywords = findKeywordsForPage(pageUrl);
  
  if (keywords.length === 0) {
    return {
      type: 'no_keywords',
      message: 'このページに関連するターゲットキーワードが見つかりません。'
    };
  }
  
  // 最初のキーワードで競合データを取得
  const competitorData = findCompetitorAnalysis(keywords[0]);
  
  if (!competitorData.found) {
    return {
      type: 'no_competitor_data',
      message: `キーワード「${keywords[0]}」の競合データがありません。\n` +
        '先に競合分析を実行してください。'
    };
  }
  
  return {
    type: 'diff_ready',
    pageUrl: pageUrl,
    keyword: keywords[0],
    competitorData: competitorData
  };
}

// ============================================================
// タスク5: 複数競合サイト一括比較
// ============================================================

/**
 * 複数の競合サイトを一括取得・比較
 * @param {Array} urls - 競合サイトURLの配列
 * @return {Object} 比較結果
 */
function compareMultipleCompetitors(urls) {
  const results = [];
  const errors = [];
  
  // 各URLのコンテンツを取得
  for (const url of urls) {
    try {
      const content = fetchCompetitorContent(url);
      if (content.success) {
        results.push(content);
      } else {
        errors.push({ url: url, error: content.error });
      }
    } catch (e) {
      errors.push({ url: url, error: e.message });
    }
    
    // レート制限対策（1秒待機）
    Utilities.sleep(1000);
  }
  
  if (results.length === 0) {
    return {
      success: false,
      errors: errors,
      message: 'コンテンツを取得できませんでした'
    };
  }
  
  // 共通コンテンツの分析
  const commonContent = analyzeCommonContent(results);
  
  // 統計情報
  const stats = calculateCompetitorStats(results);
  
  return {
    success: true,
    totalFetched: results.length,
    totalErrors: errors.length,
    results: results,
    errors: errors,
    commonContent: commonContent,
    stats: stats
  };
}

/**
 * 複数サイトの共通コンテンツを分析
 * @param {Array} results - 取得結果配列
 * @return {Object} 共通コンテンツ分析
 */
function analyzeCommonContent(results) {
  if (results.length < 2) {
    return { insufficient: true };
  }
  
  // 見出しの共通キーワードを抽出
  const allH2Keywords = {};
  
  results.forEach(result => {
    (result.h2 || []).forEach(heading => {
      // キーワードを抽出（助詞などを除去）
      const keywords = extractKeywords(heading);
      keywords.forEach(kw => {
        allH2Keywords[kw] = (allH2Keywords[kw] || 0) + 1;
      });
    });
  });
  
  // 過半数以上のサイトに存在するキーワード = 必須コンテンツ
  const threshold = Math.ceil(results.length / 2);
  const requiredTopics = Object.entries(allH2Keywords)
    .filter(([kw, count]) => count >= threshold)
    .map(([kw, count]) => ({ keyword: kw, count: count, percentage: Math.round(count / results.length * 100) }))
    .sort((a, b) => b.count - a.count);
  
  // 共通の特徴
  const commonFeatures = {
    hasToc: results.filter(r => r.hasToc).length,
    hasFaq: results.filter(r => r.hasFaq).length,
    hasVideo: results.filter(r => r.hasVideo).length
  };
  
  return {
    requiredTopics: requiredTopics,
    commonFeatures: commonFeatures,
    siteCount: results.length
  };
}

/**
 * テキストからキーワードを抽出
 * @param {string} text - テキスト
 * @return {Array} キーワード配列
 */
function extractKeywords(text) {
  if (!text) return [];
  
  // 日本語の助詞や一般的な単語を除去
  const stopWords = ['の', 'を', 'に', 'は', 'が', 'と', 'で', 'から', 'まで', 'より', 'など', 'とは', 'について', 'とき', 'こと', 'もの', 'ため', 'よう', 'さん', 'ください'];
  
  // スペースで分割し、短すぎる単語や数字のみを除去
  return text
    .replace(/[【】「」『』（）\(\)【】]/g, ' ')
    .split(/[\s　]+/)
    .filter(word => word.length > 1)
    .filter(word => !/^\d+$/.test(word))
    .filter(word => !stopWords.includes(word));
}

/**
 * 競合サイトの統計情報を計算
 * @param {Array} results - 取得結果配列
 * @return {Object} 統計情報
 */
function calculateCompetitorStats(results) {
  if (results.length === 0) {
    return null;
  }
  
  const stats = {
    avgWordCount: 0,
    avgImageCount: 0,
    avgH2Count: 0,
    avgH3Count: 0,
    minWordCount: Infinity,
    maxWordCount: 0
  };
  
  results.forEach(result => {
    stats.avgWordCount += result.wordCount || 0;
    stats.avgImageCount += result.imageCount || 0;
    stats.avgH2Count += (result.h2 || []).length;
    stats.avgH3Count += (result.h3 || []).length;
    
    if (result.wordCount < stats.minWordCount) stats.minWordCount = result.wordCount;
    if (result.wordCount > stats.maxWordCount) stats.maxWordCount = result.wordCount;
  });
  
  stats.avgWordCount = Math.round(stats.avgWordCount / results.length);
  stats.avgImageCount = Math.round(stats.avgImageCount / results.length);
  stats.avgH2Count = Math.round(stats.avgH2Count / results.length);
  stats.avgH3Count = Math.round(stats.avgH3Count / results.length);
  
  return stats;
}

// ============================================================
// 差分分析
// ============================================================

/**
 * 自社サイトと競合サイトの差分を分析
 * @param {Object} ownContent - 自社サイトコンテンツ
 * @param {Array} competitorContents - 競合サイトコンテンツ配列
 * @return {Object} 差分分析結果
 */
function analyzeDifference(ownContent, competitorContents) {
  const comparison = compareMultipleCompetitors(competitorContents.map(c => c.url || c));
  
  if (!comparison.success) {
    return {
      success: false,
      error: comparison.message
    };
  }
  
  const result = {
    success: true,
    ownContent: ownContent,
    competitorStats: comparison.stats,
    commonContent: comparison.commonContent
  };
  
  // 自社に不足しているコンテンツ
  result.missing = {
    features: [],
    topics: []
  };
  
  // 機能の比較
  if (!ownContent.hasToc && comparison.commonContent.commonFeatures.hasToc >= 2) {
    result.missing.features.push('目次');
  }
  if (!ownContent.hasFaq && comparison.commonContent.commonFeatures.hasFaq >= 2) {
    result.missing.features.push('FAQ');
  }
  if (!ownContent.hasVideo && comparison.commonContent.commonFeatures.hasVideo >= 2) {
    result.missing.features.push('動画');
  }
  
  // トピックの比較
  const ownH2Keywords = new Set();
  (ownContent.h2 || []).forEach(heading => {
    extractKeywords(heading).forEach(kw => ownH2Keywords.add(kw));
  });
  
  (comparison.commonContent.requiredTopics || []).forEach(topic => {
    if (!ownH2Keywords.has(topic.keyword)) {
      result.missing.topics.push(topic);
    }
  });
  
  // 数値の比較
  result.numericComparison = {
    wordCount: {
      own: ownContent.wordCount,
      competitorAvg: comparison.stats.avgWordCount,
      diff: ownContent.wordCount - comparison.stats.avgWordCount,
      status: ownContent.wordCount >= comparison.stats.avgWordCount ? 'OK' : '不足'
    },
    imageCount: {
      own: ownContent.imageCount,
      competitorAvg: comparison.stats.avgImageCount,
      diff: ownContent.imageCount - comparison.stats.avgImageCount,
      status: ownContent.imageCount >= comparison.stats.avgImageCount ? 'OK' : '不足'
    },
    h2Count: {
      own: (ownContent.h2 || []).length,
      competitorAvg: comparison.stats.avgH2Count,
      diff: (ownContent.h2 || []).length - comparison.stats.avgH2Count,
      status: (ownContent.h2 || []).length >= comparison.stats.avgH2Count ? 'OK' : '不足'
    }
  };
  
  return result;
}

// ============================================================
// Claude API連携（競合分析用プロンプト）
// ============================================================

/**
 * 競合分析結果をClaudeに送信して分析を依頼
 * @param {Object} analysisData - 分析データ
 * @param {string} userQuestion - ユーザーの質問
 * @return {string} Claudeの回答
 */
function getCompetitorAnalysisFromClaude(analysisData, userQuestion) {
  const prompt = buildCompetitorAnalysisPrompt(analysisData, userQuestion);
  
  // ClaudeAPI.gsの関数を呼び出し
  if (typeof callClaudeAPIWithRetry === 'function') {
    return callClaudeAPIWithRetry(prompt, getCompetitorSystemPrompt());
  } else if (typeof callClaudeAPI === 'function') {
    return callClaudeAPI(prompt, getCompetitorSystemPrompt());
  } else {
    Logger.log('Claude API関数が見つかりません');
    return '競合分析の結果を生成できませんでした。';
  }
}

/**
 * 競合分析用のシステムプロンプトを取得
 * @return {string} システムプロンプト
 */
function getCompetitorSystemPrompt() {
  return `あなたはSEOと競合分析の専門家です。

【役割】
- 競合サイトのコンテンツを分析し、改善提案を行う
- 自社サイトに不足しているコンテンツを特定する
- 具体的で実行可能なアクションを提案する

【分析の視点】
1. コンテンツの網羅性（上位サイトが共通して持つ情報）
2. コンテンツの独自性（差別化ポイント）
3. ユーザー体験（構造、読みやすさ）
4. SEO要素（見出し構造、キーワード配置）

【回答スタイル】
- 具体的な数値を含める
- 優先順位を明確にする
- 実行可能なアクションを提示する
- 期待効果を定量化する`;
}

/**
 * 競合分析用プロンプトを構築
 * @param {Object} data - 分析データ
 * @param {string} question - ユーザーの質問
 * @return {string} プロンプト
 */
function buildCompetitorAnalysisPrompt(data, question) {
  let prompt = `## ユーザーの質問\n${question}\n\n`;
  
  if (data.type === 'saved_data' && data.data) {
    prompt += `## 競合分析データ（保存済み）\n`;
    prompt += `- キーワード: ${data.data.keyword}\n`;
    prompt += `- 分析日: ${data.data.analysisDate}\n`;
    prompt += `- 勝算度スコア: ${data.data.winnableScore}点\n`;
    prompt += `- 競合レベル: ${data.data.competitorLevel}\n`;
    prompt += `- 自社DA: ${data.data.ownSiteDA}\n`;
    prompt += `- 自社順位: ${data.data.ownSiteRank || '圏外'}\n\n`;
    
    prompt += `### 上位サイト\n`;
    data.data.competitors.forEach(comp => {
      prompt += `${comp.rank}位: ${comp.domain} (DA: ${comp.da})\n`;
    });
  }
  
  if (data.type === 'direct_analysis' && data.results) {
    prompt += `## 競合サイトのコンテンツ分析\n\n`;
    data.results.forEach((result, index) => {
      if (result.success) {
        prompt += `### サイト${index + 1}: ${result.domain}\n`;
        prompt += `- タイトル: ${result.title}\n`;
        prompt += `- 文字数: ${result.wordCount}文字\n`;
        prompt += `- 画像数: ${result.imageCount}枚\n`;
        prompt += `- H2見出し数: ${result.h2.length}個\n`;
        prompt += `- 目次: ${result.hasToc ? 'あり' : 'なし'}\n`;
        prompt += `- FAQ: ${result.hasFaq ? 'あり' : 'なし'}\n`;
        prompt += `- 動画: ${result.hasVideo ? 'あり' : 'なし'}\n`;
        prompt += `\nH2見出し:\n`;
        result.h2.forEach(h => prompt += `- ${h}\n`);
        prompt += '\n';
      }
    });
  }
  
  if (data.diffAnalysis) {
    prompt += `## 差分分析結果\n\n`;
    prompt += `### 不足している機能\n`;
    data.diffAnalysis.missing.features.forEach(f => prompt += `- ${f}\n`);
    
    prompt += `\n### 不足しているトピック\n`;
    data.diffAnalysis.missing.topics.forEach(t => 
      prompt += `- ${t.keyword}（競合${t.count}サイトが言及）\n`
    );
    
    prompt += `\n### 数値比較\n`;
    const nc = data.diffAnalysis.numericComparison;
    prompt += `- 文字数: 自社${nc.wordCount.own}字 vs 競合平均${nc.wordCount.competitorAvg}字（${nc.wordCount.status}）\n`;
    prompt += `- 画像数: 自社${nc.imageCount.own}枚 vs 競合平均${nc.imageCount.competitorAvg}枚（${nc.imageCount.status}）\n`;
    prompt += `- H2数: 自社${nc.h2Count.own}個 vs 競合平均${nc.h2Count.competitorAvg}個（${nc.h2Count.status}）\n`;
  }
  
  return prompt;
}

// ============================================================
// テスト関数
// ============================================================

/**
 * 外部サイトコンテンツ取得テスト
 */
function testFetchCompetitorContent() {
  const testUrl = 'https://www.lifenet-seimei.co.jp/';
  const result = fetchCompetitorContent(testUrl);
  
  Logger.log('=== コンテンツ取得テスト ===');
  Logger.log(`URL: ${testUrl}`);
  Logger.log(`成功: ${result.success}`);
  
  if (result.success) {
    Logger.log(`タイトル: ${result.title}`);
    Logger.log(`メタディスクリプション: ${result.metaDescription}`);
    Logger.log(`文字数: ${result.wordCount}`);
    Logger.log(`画像数: ${result.imageCount}`);
    Logger.log(`H2見出し数: ${result.h2.length}`);
    Logger.log(`目次: ${result.hasToc}`);
    Logger.log(`FAQ: ${result.hasFaq}`);
    Logger.log(`取得時間: ${result.fetchTime}秒`);
  } else {
    Logger.log(`エラー: ${result.error}`);
  }
  
  return result;
}

/**
 * 競合分析シート検索テスト
 */
function testFindCompetitorAnalysis() {
  const testKeyword = 'iphone 保険';
  const result = findCompetitorAnalysis(testKeyword);
  
  Logger.log('=== 競合分析シート検索テスト ===');
  Logger.log(`キーワード: ${testKeyword}`);
  Logger.log(`見つかった: ${result.found}`);
  
  if (result.found) {
    Logger.log(`勝算度: ${result.winnableScore}`);
    Logger.log(`競合レベル: ${result.competitorLevel}`);
    Logger.log(`上位サイト数: ${result.competitors.length}`);
    result.competitors.slice(0, 3).forEach(comp => {
      Logger.log(`  ${comp.rank}位: ${comp.domain} (DA: ${comp.da})`);
    });
  }
  
  return result;
}

/**
 * 意図分析テスト
 */
function testAnalyzeCompetitorIntent() {
  const testMessages = [
    'https://example.com/page1 を分析して',
    '「iphone 保険」で上位サイトと比較して',
    '競合サイトとの差分を分析して',
    '競合分析をしたい'
  ];
  
  Logger.log('=== 意図分析テスト ===');
  testMessages.forEach(msg => {
    const result = analyzeCompetitorIntent(msg);
    Logger.log(`メッセージ: "${msg}"`);
    Logger.log(`  タイプ: ${result.type}`);
    Logger.log(`  アクション: ${result.action}`);
  });
}

/**
 * 複数サイト比較テスト
 */
function testCompareMultipleCompetitors() {
  Logger.log('=== 複数サイト比較テスト ===');
  
  const testUrls = [
    'https://penguin-diary.hatenablog.com/entry/amazon-apple',
    'https://www.amazon.co.jp/iphone/s?k=iphone',
    'https://kick-freedom.com/20566/'
  ];
  
  Logger.log(`テストURL数: ${testUrls.length}`);
  
  const result = compareMultipleCompetitors(testUrls);
  
  Logger.log(`成功: ${result.success}`);
  Logger.log(`取得成功数: ${result.totalFetched}`);
  Logger.log(`エラー数: ${result.totalErrors}`);
  
  if (result.results && result.results.length > 0) {
    Logger.log('--- 取得したサイト ---');
    result.results.forEach((site, i) => {
      Logger.log(`${i + 1}. ${site.title} (${site.wordCount}文字)`);
    });
  }
  
  if (result.stats) {
    Logger.log('--- 統計情報 ---');
    Logger.log(`平均文字数: ${result.stats.avgWordCount}`);
    Logger.log(`平均画像数: ${result.stats.avgImageCount}`);
  }
  
  if (result.commonContent) {
    Logger.log('--- 共通コンテンツ ---');
    if (result.commonContent.commonTopics && result.commonContent.commonTopics.length > 0) {
      Logger.log(`共通トピック: ${result.commonContent.commonTopics.join(', ')}`);
    } else {
      Logger.log('共通トピック: なし');
    }
  }
  
  Logger.log('=== テスト完了 ===');
}

/**
 * 競合分析シートの列構造をデバッグ
 */
function debugCompetitorSheetColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('競合分析');
  
  if (!sheet) {
    Logger.log('競合分析シートが見つかりません');
    return;
  }
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  Logger.log('=== 競合分析シート列構造 ===');
  Logger.log(`列数: ${headers.length}`);
  
  headers.forEach((header, index) => {
    Logger.log(`列${index} (${String.fromCharCode(65 + index)}): ${header}`);
  });
  
  // 2行目のデータも確認
  if (sheet.getLastRow() >= 2) {
    const row2 = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
    Logger.log('\n=== 2行目のデータ例 ===');
    headers.forEach((header, index) => {
      Logger.log(`${header}: ${row2[index]}`);
    });
  }
}

/**
 * 競合分析シートのキーワード一覧を確認
 */
function debugListKeywords() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('競合分析');
  
  if (!sheet) {
    Logger.log('競合分析シートが見つかりません');
    return;
  }
  
  const data = sheet.getDataRange().getValues();
  
  Logger.log('=== 競合分析シートのキーワード一覧（先頭20件） ===');
  Logger.log(`総キーワード数: ${data.length - 1}`);
  
  for (let i = 1; i < Math.min(data.length, 21); i++) {
    const keyword = data[i][1]; // target_keyword は列1（B列）
    const winnableScore = data[i][33]; // winnable_score は列33
    const competitorLevel = data[i][34]; // competitor_level は列34
    Logger.log(`${i}: "${keyword}" - 勝算度: ${winnableScore}, レベル: ${competitorLevel}`);
  }
}

/**
 * 存在するキーワードで競合分析シート検索テスト
 */
function testFindCompetitorAnalysisV2() {
  // シートに存在するキーワードでテスト
  const testKeyword = 'iPhone amazon で買う';
  
  Logger.log('=== 競合分析シート検索テスト V2 ===');
  Logger.log(`キーワード: ${testKeyword}`);
  
  const result = findCompetitorAnalysis(testKeyword);
  
  Logger.log(`見つかった: ${result.found}`);
  Logger.log(`勝算度: ${result.winnableScore}`);
  Logger.log(`競合レベル: ${result.competitorLevel}`);
  Logger.log(`自社DA: ${result.ownSiteDA}`);
  Logger.log(`自社順位: ${result.ownSiteRank}`);
  Logger.log(`上位サイト数: ${result.topSites ? result.topSites.length : 0}`);
  Logger.log(`データ鮮度: ${result.freshness}`);
  
  if (result.topSites && result.topSites.length > 0) {
    Logger.log('--- 上位サイト ---');
    result.topSites.forEach((site, i) => {
      Logger.log(`${site.rank}位: ${site.url} (DA: ${site.da})`);
    });
  }
}

/**
 * 自社ドメインを取得
 * @return {string} 自社ドメイン
 */
function getOwnDomain() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = ss.getSheetByName('設定・マスタ');
    
    if (settingsSheet) {
      const data = settingsSheet.getDataRange().getValues();
      for (let i = 0; i < data.length; i++) {
        if (data[i][0] === 'OWN_DOMAIN' || data[i][0] === 'SITE_DOMAIN') {
          return data[i][1];
        }
      }
    }
  } catch (e) {
    Logger.log('設定シートからドメイン取得エラー: ' + e.message);
  }
  
  // デフォルト値
  return 'smaho-tap.com';
}

/**
 * 複数サイト比較テスト（デバッグ版）
 */
function testCompareMultipleCompetitorsDebug() {
  Logger.log('=== 複数サイト比較テスト（デバッグ版） ===');
  
  const testUrls = [
    'https://penguin-diary.hatenablog.com/entry/amazon-apple',
    'https://www.amazon.co.jp/iphone/s?k=iphone'
  ];
  
  Logger.log(`テストURL数: ${testUrls.length}`);
  
  try {
    const result = compareMultipleCompetitors(testUrls);
    
    Logger.log(`結果の型: ${typeof result}`);
    Logger.log(`結果: ${JSON.stringify(result).substring(0, 500)}`);
    
    if (result && result.results) {
      Logger.log(`取得成功数: ${result.results.length}`);
    } else {
      Logger.log('result.results が存在しません');
    }
    
  } catch (error) {
    Logger.log(`エラー: ${error.message}`);
    Logger.log(`スタック: ${error.stack}`);
  }
}