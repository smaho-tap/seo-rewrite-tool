/**
 * SEOリライト支援ツール - WordPressIntegration.gs
 * WordPress REST API連携
 * 作成日: 2025年12月3日
 * バージョン: 1.0
 */

// ============================================
// 設定・定数
// ============================================

/**
 * WordPress接続設定を取得
 * @return {Object} 設定オブジェクト
 */
function getWordPressConfig() {
  const props = PropertiesService.getScriptProperties();
  
  return {
    siteUrl: props.getProperty('WP_SITE_URL') || 'https://smaho-tap.com',
    username: props.getProperty('WP_USERNAME') || '',
    applicationPassword: props.getProperty('WP_APP_PASSWORD') || '',
    restApiBase: '/wp-json/wp/v2'
  };
}


/**
 * WordPress接続設定を保存
 * @param {string} siteUrl - WordPressサイトURL
 * @param {string} username - ユーザー名
 * @param {string} applicationPassword - アプリケーションパスワード
 */
function setWordPressConfig(siteUrl, username, applicationPassword) {
  const props = PropertiesService.getScriptProperties();
  
  props.setProperty('WP_SITE_URL', siteUrl);
  props.setProperty('WP_USERNAME', username);
  props.setProperty('WP_APP_PASSWORD', applicationPassword);
  
  Logger.log('WordPress設定を保存しました');
}


// ============================================
// 認証・接続テスト
// ============================================

/**
 * Basic認証ヘッダーを生成
 * @return {string} Base64エンコードされた認証文字列
 */
function getAuthHeader() {
  const config = getWordPressConfig();
  
  if (!config.username || !config.applicationPassword) {
    throw new Error('WordPress認証情報が設定されていません。setWordPressConfig()で設定してください。');
  }
  
  const credentials = config.username + ':' + config.applicationPassword;
  const encoded = Utilities.base64Encode(credentials);
  
  return 'Basic ' + encoded;
}


/**
 * WordPress REST API接続テスト
 * @return {Object} テスト結果
 */
function testWordPressConnection() {
  try {
    const config = getWordPressConfig();
    const url = config.siteUrl + config.restApiBase + '/users/me';
    
    const options = {
      method: 'get',
      headers: {
        'Authorization': getAuthHeader(),
        'User-Agent': 'SEO-Rewrite-Tool/1.0'
      },
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();
    
    if (code === 200) {
      const user = JSON.parse(response.getContentText());
      Logger.log('✅ WordPress接続成功');
      Logger.log('ユーザー名: ' + user.name);
      Logger.log('メール: ' + user.email);
      
      return {
        success: true,
        user: user.name,
        message: 'WordPress REST API接続成功'
      };
    } else {
      Logger.log('❌ WordPress接続失敗: ' + code);
      return {
        success: false,
        error: 'HTTP ' + code,
        message: response.getContentText()
      };
    }
    
  } catch (error) {
    Logger.log('❌ WordPress接続エラー: ' + error.message);
    return {
      success: false,
      error: error.message
    };
  }
}


// ============================================
// 投稿取得
// ============================================

/**
 * スラッグから投稿を取得
 * @param {string} slug - 投稿のスラッグ（例: iphone-insurance）
 * @return {Object} 投稿データ
 */
function getPostBySlug(slug) {
  try {
    const config = getWordPressConfig();
    
    // スラッグからスラッシュを除去
    const cleanSlug = slug.replace(/^\/|\/$/g, '');
    
    const url = config.siteUrl + config.restApiBase + '/posts?slug=' + encodeURIComponent(cleanSlug);
    
    const options = {
      method: 'get',
      headers: {
        'Authorization': getAuthHeader(),
        'User-Agent': 'SEO-Rewrite-Tool/1.0'
      },
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();
    
    if (code === 200) {
      const posts = JSON.parse(response.getContentText());
      
      if (posts.length === 0) {
        // 固定ページも検索
        const pageUrl = config.siteUrl + config.restApiBase + '/pages?slug=' + encodeURIComponent(cleanSlug);
        const pageResponse = UrlFetchApp.fetch(pageUrl, options);
        
        if (pageResponse.getResponseCode() === 200) {
          const pages = JSON.parse(pageResponse.getContentText());
          if (pages.length > 0) {
            return {
              success: true,
              post: pages[0],
              type: 'page'
            };
          }
        }
        
        return {
          success: false,
          error: '投稿が見つかりません: ' + slug
        };
      }
      
      return {
        success: true,
        post: posts[0],
        type: 'post'
      };
    }
    
    return {
      success: false,
      error: 'HTTP ' + code
    };
    
  } catch (error) {
    Logger.log('投稿取得エラー: ' + error.message);
    return {
      success: false,
      error: error.message
    };
  }
}


/**
 * 投稿IDから投稿を取得
 * @param {number} postId - 投稿ID
 * @return {Object} 投稿データ
 */
function getPostById(postId) {
  try {
    const config = getWordPressConfig();
    const url = config.siteUrl + config.restApiBase + '/posts/' + postId;
    
    const options = {
      method: 'get',
      headers: {
        'Authorization': getAuthHeader(),
        'User-Agent': 'SEO-Rewrite-Tool/1.0'
      },
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();
    
    if (code === 200) {
      return {
        success: true,
        post: JSON.parse(response.getContentText()),
        type: 'post'
      };
    }
    
    // 固定ページとして再試行
    const pageUrl = config.siteUrl + config.restApiBase + '/pages/' + postId;
    const pageResponse = UrlFetchApp.fetch(pageUrl, options);
    
    if (pageResponse.getResponseCode() === 200) {
      return {
        success: true,
        post: JSON.parse(pageResponse.getContentText()),
        type: 'page'
      };
    }
    
    return {
      success: false,
      error: 'HTTP ' + code
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}


/**
 * URLから投稿を取得
 * @param {string} pageUrl - ページURL（例: /iphone-insurance/ または https://smaho-tap.com/iphone-insurance/）
 * @return {Object} 投稿データ
 */
function getPostByUrl(pageUrl) {
  // URLからスラッグを抽出
  let slug = pageUrl;
  
  // フルURLの場合はパスを抽出
  if (pageUrl.includes('://')) {
    const url = new URL(pageUrl);
    slug = url.pathname;
  }
  
  // スラッシュを除去してスラッグに変換
  slug = slug.replace(/^\/|\/$/g, '');
  
  // 階層構造がある場合は最後の部分を使用
  if (slug.includes('/')) {
    slug = slug.split('/').pop();
  }
  
  return getPostBySlug(slug);
}


// ============================================
// 投稿更新
// ============================================

/**
 * 投稿のタイトルを更新
 * @param {number} postId - 投稿ID
 * @param {string} newTitle - 新しいタイトル
 * @param {string} postType - 投稿タイプ（post/page）
 * @return {Object} 更新結果
 */
function updatePostTitle(postId, newTitle, postType) {
  postType = postType || 'post';
  
  try {
    const config = getWordPressConfig();
    const endpoint = postType === 'page' ? '/pages/' : '/posts/';
    const url = config.siteUrl + config.restApiBase + endpoint + postId;
    
    const options = {
      method: 'post',
      headers: {
        'Authorization': getAuthHeader(),
        'User-Agent': 'SEO-Rewrite-Tool/1.0',
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({
        title: newTitle
      }),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();
    
    if (code === 200) {
      const updated = JSON.parse(response.getContentText());
      Logger.log('✅ タイトル更新成功: ' + updated.title.rendered);
      
      return {
        success: true,
        post: updated,
        message: 'タイトルを更新しました'
      };
    }
    
    return {
      success: false,
      error: 'HTTP ' + code,
      message: response.getContentText()
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}


/**
 * 投稿のコンテンツを更新
 * @param {number} postId - 投稿ID
 * @param {string} newContent - 新しいコンテンツ（HTML）
 * @param {string} postType - 投稿タイプ（post/page）
 * @return {Object} 更新結果
 */
function updatePostContent(postId, newContent, postType) {
  postType = postType || 'post';
  
  try {
    const config = getWordPressConfig();
    const endpoint = postType === 'page' ? '/pages/' : '/posts/';
    const url = config.siteUrl + config.restApiBase + endpoint + postId;
    
    const options = {
      method: 'post',
      headers: {
        'Authorization': getAuthHeader(),
        'User-Agent': 'SEO-Rewrite-Tool/1.0',
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({
        content: newContent
      }),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();
    
    if (code === 200) {
      const updated = JSON.parse(response.getContentText());
      Logger.log('✅ コンテンツ更新成功');
      
      return {
        success: true,
        post: updated,
        message: 'コンテンツを更新しました'
      };
    }
    
    return {
      success: false,
      error: 'HTTP ' + code,
      message: response.getContentText()
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}


/**
 * 投稿のメタデータを更新
 * @param {number} postId - 投稿ID
 * @param {Object} metaData - メタデータ（キー: 値のオブジェクト）
 * @param {string} postType - 投稿タイプ
 * @return {Object} 更新結果
 */
function updatePostMeta(postId, metaData, postType) {
  postType = postType || 'post';
  
  try {
    const config = getWordPressConfig();
    const endpoint = postType === 'page' ? '/pages/' : '/posts/';
    const url = config.siteUrl + config.restApiBase + endpoint + postId;
    
    const options = {
      method: 'post',
      headers: {
        'Authorization': getAuthHeader(),
        'User-Agent': 'SEO-Rewrite-Tool/1.0',
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({
        meta: metaData
      }),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();
    
    if (code === 200) {
      return {
        success: true,
        message: 'メタデータを更新しました'
      };
    }
    
    return {
      success: false,
      error: 'HTTP ' + code,
      message: response.getContentText()
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}


// ============================================
// SEO関連（Yoast/RankMath連携）
// ============================================

/**
 * SEOタイトル（メタタイトル）を更新
 * Yoast SEOまたはRankMath使用時
 * @param {number} postId - 投稿ID
 * @param {string} seoTitle - SEOタイトル
 * @param {string} postType - 投稿タイプ
 * @return {Object} 更新結果
 */
function updateSEOTitle(postId, seoTitle, postType) {
  // Yoast SEOの場合
  const yoastResult = updatePostMeta(postId, {
    '_yoast_wpseo_title': seoTitle
  }, postType);
  
  // RankMathの場合も試行
  const rankMathResult = updatePostMeta(postId, {
    'rank_math_title': seoTitle
  }, postType);
  
  return yoastResult.success || rankMathResult.success ? 
    { success: true, message: 'SEOタイトルを更新しました' } :
    { success: false, error: 'SEOプラグインのメタデータ更新に失敗しました' };
}


/**
 * メタディスクリプションを更新
 * @param {number} postId - 投稿ID
 * @param {string} description - メタディスクリプション
 * @param {string} postType - 投稿タイプ
 * @return {Object} 更新結果
 */
function updateMetaDescription(postId, description, postType) {
  // Yoast SEOの場合
  const yoastResult = updatePostMeta(postId, {
    '_yoast_wpseo_metadesc': description
  }, postType);
  
  // RankMathの場合も試行
  const rankMathResult = updatePostMeta(postId, {
    'rank_math_description': description
  }, postType);
  
  return yoastResult.success || rankMathResult.success ? 
    { success: true, message: 'メタディスクリプションを更新しました' } :
    { success: false, error: 'SEOプラグインのメタデータ更新に失敗しました' };
}


// ============================================
// コンテンツ解析
// ============================================

/**
 * 投稿のコンテンツを解析
 * @param {Object} post - 投稿オブジェクト
 * @return {Object} 解析結果
 */
function analyzePostContent(post) {
  const content = post.content.rendered || '';
  
  // H2見出しを抽出
  const h2Matches = content.match(/<h2[^>]*>(.*?)<\/h2>/gi) || [];
  const h2s = h2Matches.map(h => h.replace(/<[^>]+>/g, '').trim());
  
  // H3見出しを抽出
  const h3Matches = content.match(/<h3[^>]*>(.*?)<\/h3>/gi) || [];
  const h3s = h3Matches.map(h => h.replace(/<[^>]+>/g, '').trim());
  
  // 画像を抽出
  const imgMatches = content.match(/<img[^>]+>/gi) || [];
  
  // テキスト文字数
  const textContent = content.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
  const wordCount = textContent.length;
  
  // FAQ/Q&Aを検出
  const hasFaq = content.includes('faq') || 
                 content.includes('よくある質問') || 
                 content.includes('Q&amp;A') ||
                 content.includes('Q&A');
  
  // 動画を検出
  const hasVideo = content.includes('youtube') || 
                   content.includes('vimeo') || 
                   content.includes('<video');
  
  // 目次を検出
  const hasToc = content.includes('toc') || 
                 content.includes('目次') || 
                 content.includes('table-of-contents');
  
  return {
    title: post.title.rendered,
    wordCount: wordCount,
    h2Count: h2s.length,
    h3Count: h3s.length,
    imageCount: imgMatches.length,
    h2s: h2s,
    h3s: h3s,
    hasFaq: hasFaq,
    hasVideo: hasVideo,
    hasToc: hasToc,
    lastModified: post.modified
  };
}


// ============================================
// リライト支援機能
// ============================================

/**
 * ページURLからWordPress投稿を取得し、内容を解析
 * タスク管理システムとの連携用
 * 
 * @param {string} pageUrl - ページURL
 * @return {Object} 投稿情報と解析結果
 */
function getPostInfoForRewrite(pageUrl) {
  try {
    // 投稿を取得
    const postResult = getPostByUrl(pageUrl);
    
    if (!postResult.success) {
      return {
        success: false,
        error: postResult.error,
        wpConnected: true
      };
    }
    
    // コンテンツを解析
    const analysis = analyzePostContent(postResult.post);
    
    return {
      success: true,
      wpConnected: true,
      postId: postResult.post.id,
      postType: postResult.type,
      title: postResult.post.title.rendered,
      slug: postResult.post.slug,
      status: postResult.post.status,
      link: postResult.post.link,
      analysis: analysis,
      rawContent: postResult.post.content.rendered
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.message,
      wpConnected: false
    };
  }
}


/**
 * リライトタスクの変更をWordPressに適用
 * 
 * @param {string} taskId - タスクID
 * @param {Object} changes - 適用する変更内容
 * @return {Object} 適用結果
 */
function applyRewriteToWordPress(taskId, changes) {
  try {
    // タスク情報を取得
    const task = getTaskById(taskId);
    if (!task) {
      return { success: false, error: 'タスクが見つかりません' };
    }
    
    // 投稿を取得
    const postResult = getPostByUrl(task.page_url);
    if (!postResult.success) {
      return { success: false, error: '投稿が見つかりません: ' + postResult.error };
    }
    
    const postId = postResult.post.id;
    const postType = postResult.type;
    
    const results = [];
    
    // タイトル変更
    if (changes.title) {
      const titleResult = updatePostTitle(postId, changes.title, postType);
      results.push({ type: 'title', ...titleResult });
    }
    
    // コンテンツ変更
    if (changes.content) {
      const contentResult = updatePostContent(postId, changes.content, postType);
      results.push({ type: 'content', ...contentResult });
    }
    
    // SEOタイトル変更
    if (changes.seoTitle) {
      const seoTitleResult = updateSEOTitle(postId, changes.seoTitle, postType);
      results.push({ type: 'seoTitle', ...seoTitleResult });
    }
    
    // メタディスクリプション変更
    if (changes.metaDescription) {
      const metaDescResult = updateMetaDescription(postId, changes.metaDescription, postType);
      results.push({ type: 'metaDescription', ...metaDescResult });
    }
    
    // 結果を集計
    const allSuccess = results.every(r => r.success);
    const failedItems = results.filter(r => !r.success);
    
    // タスク管理シートにWordPress更新を記録
    if (allSuccess) {
      updateTaskNotes(taskId, 'WordPress更新完了: ' + new Date().toLocaleString('ja-JP'));
    }
    
    return {
      success: allSuccess,
      results: results,
      failedCount: failedItems.length,
      message: allSuccess ? 
        'WordPressへの変更を適用しました' : 
        failedItems.length + '件の更新に失敗しました'
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}


/**
 * タスクIDからタスクを取得（TaskManagement.gsとの連携）
 * @param {string} taskId - タスクID
 * @return {Object|null} タスクデータ
 */
function getTaskById(taskId) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('タスク管理');
    if (!sheet) return null;
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const taskIdCol = headers.indexOf('task_id');
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][taskIdCol] === taskId) {
        const task = {};
        headers.forEach((header, idx) => {
          task[header] = data[i][idx];
        });
        return task;
      }
    }
    
    return null;
  } catch (error) {
    Logger.log('タスク取得エラー: ' + error.message);
    return null;
  }
}


/**
 * タスクのメモを更新
 * @param {string} taskId - タスクID
 * @param {string} note - 追加するメモ
 */
function updateTaskNotes(taskId, note) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('タスク管理');
    if (!sheet) return;
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const taskIdCol = headers.indexOf('task_id');
    const notesCol = headers.indexOf('notes');
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][taskIdCol] === taskId) {
        const currentNotes = data[i][notesCol] || '';
        const newNotes = currentNotes ? currentNotes + '\n' + note : note;
        sheet.getRange(i + 1, notesCol + 1).setValue(newNotes);
        break;
      }
    }
  } catch (error) {
    Logger.log('メモ更新エラー: ' + error.message);
  }
}


// ============================================
// 一括処理
// ============================================

/**
 * 複数ページの情報を一括取得
 * @param {Array} pageUrls - ページURLの配列
 * @return {Array} 各ページの情報
 */
function getMultiplePostsInfo(pageUrls) {
  const results = [];
  
  for (const url of pageUrls) {
    const info = getPostInfoForRewrite(url);
    results.push({
      url: url,
      ...info
    });
    
    // API制限対策
    Utilities.sleep(500);
  }
  
  return results;
}


// ============================================
// テスト関数
// ============================================

/**
 * WordPress連携機能のテスト
 */
function testWordPressIntegration() {
  Logger.log('=== WordPress連携テスト開始 ===');
  
  // 1. 接続テスト
  Logger.log('\n--- 1. 接続テスト ---');
  const connResult = testWordPressConnection();
  Logger.log('接続結果: ' + JSON.stringify(connResult));
  
  if (!connResult.success) {
    Logger.log('⚠️ 接続に失敗しました。WordPress設定を確認してください。');
    Logger.log('設定方法: setWordPressConfig(siteUrl, username, applicationPassword) を実行');
    return;
  }
  
  // 2. 投稿取得テスト
  Logger.log('\n--- 2. 投稿取得テスト ---');
  const testSlug = 'iphone-insurance'; // テスト用スラッグ（実際のスラッグに変更）
  const postResult = getPostBySlug(testSlug);
  
  if (postResult.success) {
    Logger.log('✅ 投稿取得成功');
    Logger.log('タイトル: ' + postResult.post.title.rendered);
    Logger.log('ID: ' + postResult.post.id);
    Logger.log('タイプ: ' + postResult.type);
    
    // 3. コンテンツ解析テスト
    Logger.log('\n--- 3. コンテンツ解析テスト ---');
    const analysis = analyzePostContent(postResult.post);
    Logger.log('文字数: ' + analysis.wordCount);
    Logger.log('H2数: ' + analysis.h2Count);
    Logger.log('H3数: ' + analysis.h3Count);
    Logger.log('画像数: ' + analysis.imageCount);
    Logger.log('FAQ: ' + analysis.hasFaq);
    Logger.log('動画: ' + analysis.hasVideo);
    Logger.log('目次: ' + analysis.hasToc);
    
    if (analysis.h2s.length > 0) {
      Logger.log('H2見出し:');
      analysis.h2s.slice(0, 5).forEach((h2, i) => {
        Logger.log('  ' + (i + 1) + '. ' + h2);
      });
    }
  } else {
    Logger.log('⚠️ 投稿取得失敗: ' + postResult.error);
    Logger.log('テスト用スラッグを変更してください');
  }
  
  Logger.log('\n=== WordPress連携テスト完了 ===');
}


/**
 * WordPress設定の初期セットアップ
 * 初回のみ実行してください
 */
function setupWordPressConnection() {
  // ここに実際の値を入力して実行
  const siteUrl = 'https://smaho-tap.com';
  const username = ''; // WordPressユーザー名
  const applicationPassword = ''; // アプリケーションパスワード
  
  if (!username || !applicationPassword) {
    Logger.log('⚠️ ユーザー名とアプリケーションパスワードを設定してください');
    Logger.log('');
    Logger.log('【アプリケーションパスワードの取得方法】');
    Logger.log('1. WordPress管理画面にログイン');
    Logger.log('2. ユーザー → プロフィール を開く');
    Logger.log('3. 「アプリケーションパスワード」セクションで新規追加');
    Logger.log('4. 名前を入力（例: SEOツール）して「新しいアプリケーションパスワードを追加」');
    Logger.log('5. 生成されたパスワードをコピー');
    Logger.log('');
    Logger.log('この関数内の username と applicationPassword を設定して再実行してください');
    return;
  }
  
  setWordPressConfig(siteUrl, username, applicationPassword);
  
  // 接続テスト
  testWordPressConnection();
}

function testRealPageFetch() {
  // 投稿IDでテスト（先ほど取得できたID）
  const testId = 8060;
  
  Logger.log('=== 投稿ID取得テスト ===');
  Logger.log('テストID: ' + testId);
  
  const result = getPostById(testId);
  
  if (result.success) {
    Logger.log('✅ 投稿取得成功');
    Logger.log('投稿ID: ' + result.post.id);
    Logger.log('タイトル: ' + result.post.title.rendered);
    Logger.log('スラッグ: ' + result.post.slug);
    Logger.log('リンク: ' + result.post.link);
    Logger.log('タイプ: ' + result.type);
    
    // コンテンツ解析
    const analysis = analyzePostContent(result.post);
    Logger.log('--- コンテンツ解析 ---');
    Logger.log('文字数: ' + analysis.characterCount);
    Logger.log('H2数: ' + analysis.h2Count);
    Logger.log('H3数: ' + analysis.h3Count);
    Logger.log('画像数: ' + analysis.imageCount);
  } else {
    Logger.log('❌ 取得失敗: ' + result.error);
  }
}

function testSyncWithAllPosts() {
  Logger.log('=== 全投稿取得方式でテスト ===');
  
  const config = getWordPressConfig();
  const url = config.siteUrl + '/wp-json/wp/v2/posts?per_page=100&_fields=id,date,link,title';
  
  const options = {
    method: 'get',
    headers: {
      'Authorization': getAuthHeader(),
      'User-Agent': 'SEO-Rewrite-Tool/1.0'
    },
    muteHttpExceptions: true
  };
  
  const response = UrlFetchApp.fetch(url, options);
  const posts = JSON.parse(response.getContentText());
  
  Logger.log('取得投稿数: ' + posts.length);
  
  // 最初の5件を表示
  for (let i = 0; i < Math.min(5, posts.length); i++) {
    const post = posts[i];
    const check = isArticleTooNew(post.date);
    Logger.log((i+1) + '. ' + post.title.rendered.substring(0, 30) + '...');
    Logger.log('   URL: ' + post.link);
    Logger.log('   投稿日: ' + post.date + ' → ' + check.monthsElapsed + 'ヶ月');
  }
}