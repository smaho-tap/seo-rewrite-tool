/**
 * WordPress REST API連携
 * 作成日: 2025年12月10日
 */

/**
 * WordPressからページ情報を取得
 * @param {string} pageUrl - ページURL（例: /iphone-battery-danger）
 * @return {Object} ページ情報
 */
function getWordPressPageData(pageUrl) {
  Logger.log('=== WordPress API取得開始: ' + pageUrl + ' ===');
  
  try {
    var slug = pageUrl.replace(/^\//, '').replace(/\/$/, '');
    
    // まず投稿（posts）から検索
    var postData = fetchWPContent('posts', slug);
    
    // 見つからなければ固定ページ（pages）から検索
    if (!postData) {
      postData = fetchWPContent('pages', slug);
    }
    
    if (!postData) {
      Logger.log('ページが見つかりません: ' + pageUrl);
      return null;
    }
    
    // コンテンツを解析
    var content = postData.content.rendered || '';
    var analysis = analyzeWPContent(content);
    
    // メタディスクリプションを取得
    var metaDescription = extractMetaDescription(postData);
    
    var result = {
      success: true,
      id: postData.id,
      title: decodeHtmlEntities(postData.title.rendered || ''),
      slug: postData.slug,
      url: pageUrl,
      metaDescription: metaDescription,
      excerpt: decodeHtmlEntities(postData.excerpt.rendered || '').replace(/<[^>]*>/g, '').trim(),
      content: content,
      wordCount: analysis.wordCount,
      h2List: analysis.h2List,
      h3List: analysis.h3List,
      hasFaq: analysis.hasFaq,
      faqCount: analysis.faqCount,
      hasTable: analysis.hasTable,
      imageCount: analysis.imageCount,
      internalLinks: analysis.internalLinks,
      externalLinks: analysis.externalLinks,
      lastModified: postData.modified
    };
    
    Logger.log('WordPress取得成功: ' + result.title);
    Logger.log('文字数: ' + result.wordCount + ', H2数: ' + result.h2List.length + ', FAQ: ' + result.hasFaq);
    
    return result;
    
  } catch (error) {
    Logger.log('WordPress API エラー: ' + error.message);
    return { success: false, error: error.message };
  }
}

/**
 * WordPress REST APIからコンテンツを取得
 * URLパスまたは検索で取得
 */
function fetchWPContent(type, slug) {
  // 方法1: slug検索
  var url = 'https://smaho-tap.com/wp-json/wp/v2/' + type + '?slug=' + encodeURIComponent(slug);
  
  try {
    var response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      headers: {
        'User-Agent': 'SEO-Rewrite-Tool/1.0'
      }
    });
    
    var code = response.getResponseCode();
    if (code === 200) {
      var data = JSON.parse(response.getContentText());
      if (data && data.length > 0) {
        Logger.log('slug検索でマッチ: ' + data[0].link);
        return data[0];
      }
    }
    
    // 方法2: search検索（slugで見つからない場合）
    Logger.log('slug検索で見つからないため、search検索を試行');
    url = 'https://smaho-tap.com/wp-json/wp/v2/' + type + '?search=' + encodeURIComponent(slug.replace(/-/g, ' ')) + '&per_page=20';
    
    response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      headers: {
        'User-Agent': 'SEO-Rewrite-Tool/1.0'
      }
    });
    
    if (response.getResponseCode() === 200) {
      var data = JSON.parse(response.getContentText());
      
      // linkにslugが含まれる投稿を探す
      for (var i = 0; i < data.length; i++) {
        var post = data[i];
        if (post.link && post.link.indexOf(slug) !== -1) {
          Logger.log('search検索でマッチ: ' + post.link);
          return post;
        }
      }
    }
    
    // 方法3: 全投稿から検索（最後の手段）
    Logger.log('search検索でも見つからないため、link照合を試行');
    url = 'https://smaho-tap.com/wp-json/wp/v2/' + type + '?per_page=100';
    
    response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      headers: {
        'User-Agent': 'SEO-Rewrite-Tool/1.0'
      }
    });
    
    if (response.getResponseCode() === 200) {
      var data = JSON.parse(response.getContentText());
      var targetUrl = 'https://smaho-tap.com/' + slug;
      
      for (var i = 0; i < data.length; i++) {
        var post = data[i];
        if (post.link === targetUrl || post.link === targetUrl + '/') {
          Logger.log('link照合でマッチ: ' + post.link);
          return post;
        }
      }
    }
    
    return null;
    
  } catch (e) {
    Logger.log(type + ' 取得エラー: ' + e.message);
    return null;
  }
}

/**
 * WordPressコンテンツを解析
 */
function analyzeWPContent(html) {
  var result = {
    wordCount: 0,
    h2List: [],
    h3List: [],
    hasFaq: false,
    faqCount: 0,
    hasTable: false,
    imageCount: 0,
    internalLinks: [],
    externalLinks: []
  };
  
  if (!html) return result;
  
  // テキストのみ抽出して文字数カウント
  var textOnly = html.replace(/<[^>]*>/g, '').replace(/\s+/g, '');
  result.wordCount = textOnly.length;
  
  // H2見出しを抽出
  var h2Matches = html.match(/<h2[^>]*>(.*?)<\/h2>/gi) || [];
  result.h2List = h2Matches.map(function(h2) {
    return h2.replace(/<[^>]*>/g, '').trim();
  });
  
  // H3見出しを抽出
  var h3Matches = html.match(/<h3[^>]*>(.*?)<\/h3>/gi) || [];
  result.h3List = h3Matches.map(function(h3) {
    return h3.replace(/<[^>]*>/g, '').trim();
  });
  
  // FAQ検出（複数パターン）
  var faqPatterns = [
    /class="[^"]*faq[^"]*"/i,
    /id="[^"]*faq[^"]*"/i,
    /よくある質問/,
    /FAQ/i,
    /Q\s*[:：]\s*/,
    /<dt[^>]*>.*?<\/dt>/i,
    /itemtype=".*FAQPage.*"/i,
    /wp-block-yoast-faq/i
  ];
  
  for (var i = 0; i < faqPatterns.length; i++) {
    if (faqPatterns[i].test(html)) {
      result.hasFaq = true;
      break;
    }
  }
  
  // FAQ数をカウント（Q:パターン）
  var qMatches = html.match(/Q\s*[:：]/g) || [];
  result.faqCount = qMatches.length;
  
  // テーブル検出
  result.hasTable = /<table/i.test(html);
  
  // 画像数
  var imgMatches = html.match(/<img[^>]*>/gi) || [];
  result.imageCount = imgMatches.length;
  
  // 内部リンク抽出
  var linkPattern = /<a[^>]*href="([^"]*)"[^>]*>/gi;
  var match;
  while ((match = linkPattern.exec(html)) !== null) {
    var href = match[1];
    if (href.indexOf('smaho-tap.com') !== -1 || href.indexOf('/') === 0) {
      // 内部リンク
      var path = href.replace('https://smaho-tap.com', '').replace('http://smaho-tap.com', '');
      if (path && path !== '#' && result.internalLinks.indexOf(path) === -1) {
        result.internalLinks.push(path);
      }
    } else if (href.indexOf('http') === 0) {
      // 外部リンク
      if (result.externalLinks.indexOf(href) === -1) {
        result.externalLinks.push(href);
      }
    }
  }
  
  return result;
}

/**
 * メタディスクリプションを抽出
 */
function extractMetaDescription(postData) {
  // SEO SIMPLE PACKのメタデータを確認
  if (postData.meta && postData.meta._ssp_meta_description) {
    return postData.meta._ssp_meta_description;
  }
  
  // Yoast SEOのメタデータを確認
  if (postData.yoast_head_json && postData.yoast_head_json.description) {
    return postData.yoast_head_json.description;
  }
  
  // excerptをフォールバックとして使用
  if (postData.excerpt && postData.excerpt.rendered) {
    var excerpt = postData.excerpt.rendered.replace(/<[^>]*>/g, '').trim();
    if (excerpt.length > 0) {
      return excerpt.substring(0, 160);
    }
  }
  
  return '';
}

/**
 * HTMLエンティティをデコード
 */
function decodeHtmlEntities(text) {
  if (!text) return '';
  return text
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#039;/g, "'")
    .replace(/&nbsp;/g, ' ')
    .replace(/<[^>]*>/g, '');
}

/**
 * テスト: WordPress API
 */
function testWordPressAPI() {
  var testUrls = [
    '/iphone7-still-using',
    '/ipad-refurbished-restock-timing'
  ];
  
  for (var i = 0; i < testUrls.length; i++) {
    Logger.log('\n=== テスト ' + (i + 1) + ': ' + testUrls[i] + ' ===');
    var result = getWordPressPageData(testUrls[i]);
    
    if (result && result.success) {
      Logger.log('タイトル: ' + result.title);
      Logger.log('メタディスクリプション: ' + (result.metaDescription || '未設定'));
      Logger.log('文字数: ' + result.wordCount);
      Logger.log('H2数: ' + result.h2List.length);
      Logger.log('H2一覧: ' + result.h2List.slice(0, 5).join(' / '));
      Logger.log('FAQ有無: ' + result.hasFaq);
      Logger.log('画像数: ' + result.imageCount);
      Logger.log('内部リンク数: ' + result.internalLinks.length);
    } else {
      Logger.log('取得失敗');
    }
  }
}

/**
 * テスト: 投稿一覧からslugを確認
 */
function testListPosts() {
  var url = 'https://smaho-tap.com/wp-json/wp/v2/posts?per_page=10';
  
  try {
    var response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true
    });
    
    var data = JSON.parse(response.getContentText());
    
    Logger.log('=== 最新10件の投稿 ===');
    for (var i = 0; i < data.length; i++) {
      var post = data[i];
      Logger.log((i + 1) + '. slug: ' + post.slug);
      Logger.log('   title: ' + post.title.rendered);
      Logger.log('   link: ' + post.link);
    }
    
  } catch (e) {
    Logger.log('エラー: ' + e.message);
  }
}