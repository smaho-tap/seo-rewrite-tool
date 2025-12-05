/**
 * SEOリライト支援ツール - PublishDateFilter.gs
 * 投稿日による記事フィルタリング（3ヶ月未満除外）
 * 作成日: 2025年12月3日
 * バージョン: 1.0
 */

// ============================================
// 設定
// ============================================

/**
 * 最小経過月数（これ未満の記事はリライト対象外）
 */
const MIN_MONTHS_FOR_REWRITE = 3;


// ============================================
// 投稿日取得
// ============================================

/**
 * スラッグから投稿日を取得
 * @param {string} slug - 投稿のスラッグ
 * @return {Object} 投稿日情報
 */
function getPublishDateBySlug(slug) {
  try {
    const result = getPostBySlug(slug);
    
    if (result.success) {
      return {
        success: true,
        publishDate: result.post.date,
        modifiedDate: result.post.modified,
        postId: result.post.id,
        title: result.post.title.rendered
      };
    }
    
    return {
      success: false,
      error: result.error
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}


/**
 * ページURLからスラッグを抽出
 * @param {string} pageUrl - ページURL
 * @return {string} スラッグ
 */
function extractSlugFromUrl(pageUrl) {
  if (!pageUrl) return '';
  
  // URLからスラッグを抽出
  // https://smaho-tap.com/iphone-insurance/ → iphone-insurance
  try {
    const url = pageUrl.replace(/\/$/, ''); // 末尾のスラッシュを除去
    const parts = url.split('/');
    return parts[parts.length - 1];
  } catch (e) {
    return '';
  }
}


/**
 * ページURLから投稿日を取得
 * @param {string} pageUrl - ページURL
 * @return {Object} 投稿日情報
 */
function getPublishDateByUrl(pageUrl) {
  const slug = extractSlugFromUrl(pageUrl);
  
  if (!slug) {
    return {
      success: false,
      error: 'スラッグを抽出できません'
    };
  }
  
  return getPublishDateBySlug(slug);
}


// ============================================
// 投稿日チェック
// ============================================

/**
 * 記事が新しすぎるかチェック（3ヶ月未満）
 * @param {string} publishDateStr - 投稿日（ISO形式）
 * @param {number} minMonths - 最小経過月数（デフォルト3）
 * @return {Object} チェック結果
 */
function isArticleTooNew(publishDateStr, minMonths = MIN_MONTHS_FOR_REWRITE) {
  if (!publishDateStr) {
    return {
      isTooNew: false,
      reason: '投稿日不明のため除外しない',
      publishDate: null,
      monthsElapsed: null
    };
  }
  
  try {
    const publishDate = new Date(publishDateStr);
    const now = new Date();
    
    // 経過月数を計算
    const monthsElapsed = (now.getFullYear() - publishDate.getFullYear()) * 12 
                        + (now.getMonth() - publishDate.getMonth());
    
    const isTooNew = monthsElapsed < minMonths;
    
    return {
      isTooNew: isTooNew,
      reason: isTooNew 
        ? `投稿から${monthsElapsed}ヶ月（${minMonths}ヶ月未満のため除外）`
        : `投稿から${monthsElapsed}ヶ月（リライト対象）`,
      publishDate: publishDate,
      monthsElapsed: monthsElapsed,
      minMonths: minMonths
    };
    
  } catch (error) {
    return {
      isTooNew: false,
      reason: '日付解析エラー: ' + error.message,
      publishDate: null,
      monthsElapsed: null
    };
  }
}


// ============================================
// 統合データシート連携
// ============================================

/**
 * 統合データシートに投稿日列を追加
 */
function addPublishDateColumn() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('統合データ');
  
  if (!sheet) {
    Logger.log('❌ 統合データシートが見つかりません');
    return;
  }
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // 既に投稿日列があるかチェック
  const existingIndex = headers.indexOf('投稿日');
  if (existingIndex !== -1) {
    Logger.log('✅ 投稿日列は既に存在します（列' + (existingIndex + 1) + '）');
    return existingIndex + 1;
  }
  
  // 経過月数列もチェック
  const monthsIndex = headers.indexOf('経過月数');
  if (monthsIndex !== -1) {
    Logger.log('✅ 経過月数列は既に存在します');
  }
  
  // 新しい列を追加（最後の列の次）
  const lastCol = sheet.getLastColumn();
  sheet.getRange(1, lastCol + 1).setValue('投稿日');
  sheet.getRange(1, lastCol + 2).setValue('経過月数');
  sheet.getRange(1, lastCol + 3).setValue('リライト可能');
  
  Logger.log('✅ 投稿日・経過月数・リライト可能列を追加しました');
  
  return lastCol + 1;
}


/**
 * 統合データシートの投稿日を同期（パスマッチング方式）
 */
function syncPublishDates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('統合データ');
  
  if (!sheet) {
    Logger.log('❌ 統合データシートが見つかりません');
    return;
  }
  
  // 投稿日列を確認/追加
  addPublishDateColumn();
  
  Logger.log('=== 投稿日同期開始（パスマッチング方式） ===');
  
  // WordPressから全投稿を取得
  const allPosts = getAllWordPressPosts();
  Logger.log('WordPress投稿数: ' + allPosts.length);
  
  // パスをキーにしたマップを作成
  const postMap = {};
  for (const post of allPosts) {
    // URLからパス部分を抽出
    // https://smaho-tap.com/ipad-where-to-buy-cheap → /ipad-where-to-buy-cheap
    const path = '/' + post.link.replace(/^https?:\/\/[^\/]+\/?/, '').replace(/\/$/, '');
    postMap[path] = post;
    // スラッシュなし版も登録
    postMap[path.replace(/^\//, '')] = post;
  }
  
  // 統合データシートを更新
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const urlColIndex = headers.indexOf('page_url');
  const publishDateColIndex = headers.indexOf('投稿日');
  const monthsColIndex = headers.indexOf('経過月数');
  const rewriteOkColIndex = headers.indexOf('リライト可能');
  
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  
  let updated = 0;
  let notFound = 0;
  
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const pageUrl = row[urlColIndex];
    
    if (!pageUrl) continue;
    
    // パスを正規化（先頭スラッシュあり、末尾スラッシュなし）
    let normalizedPath = pageUrl.replace(/\/$/, '');
    if (!normalizedPath.startsWith('/')) {
      normalizedPath = '/' + normalizedPath;
    }
    
    // マッチング
    const post = postMap[normalizedPath] || postMap[normalizedPath.replace(/^\//, '')];
    
    if (post) {
      const publishDate = new Date(post.date);
      const check = isArticleTooNew(post.date);
      
      const rowNum = i + 2;
      sheet.getRange(rowNum, publishDateColIndex + 1).setValue(publishDate);
      sheet.getRange(rowNum, monthsColIndex + 1).setValue(check.monthsElapsed);
      sheet.getRange(rowNum, rewriteOkColIndex + 1).setValue(check.isTooNew ? '×' : '○');
      
      updated++;
    } else {
      notFound++;
    }
  }
  
  Logger.log('=== 投稿日同期完了 ===');
  Logger.log('更新: ' + updated + '件 / 未マッチ: ' + notFound + '件');
  
  return { updated: updated, notFound: notFound };
}


/**
 * WordPressから全投稿を取得
 * @return {Array} 全投稿配列
 */
function getAllWordPressPosts() {
  const config = getWordPressConfig();
  const allPosts = [];
  let page = 1;
  const perPage = 100;
  
  const options = {
    method: 'get',
    headers: {
      'Authorization': getAuthHeader(),
      'User-Agent': 'SEO-Rewrite-Tool/1.0'
    },
    muteHttpExceptions: true
  };
  
  // 投稿を取得
  while (true) {
    const url = config.siteUrl + '/wp-json/wp/v2/posts?per_page=' + perPage + '&page=' + page + '&_fields=id,date,link,title';
    const response = UrlFetchApp.fetch(url, options);
    
    if (response.getResponseCode() !== 200) break;
    
    const posts = JSON.parse(response.getContentText());
    if (posts.length === 0) break;
    
    allPosts.push(...posts);
    
    if (posts.length < perPage) break;
    page++;
    Utilities.sleep(300);
  }
  
  // 固定ページも取得
  page = 1;
  while (true) {
    const url = config.siteUrl + '/wp-json/wp/v2/pages?per_page=' + perPage + '&page=' + page + '&_fields=id,date,link,title';
    const response = UrlFetchApp.fetch(url, options);
    
    if (response.getResponseCode() !== 200) break;
    
    const pages = JSON.parse(response.getContentText());
    if (pages.length === 0) break;
    
    allPosts.push(...pages);
    
    if (pages.length < perPage) break;
    page++;
    Utilities.sleep(300);
  }
  
  return allPosts;
}


// ============================================
// フィルタリング関数
// ============================================

/**
 * リライト対象ページをフィルタリング（3ヶ月未満を除外）
 * @param {Array} pages - ページ配列
 * @return {Object} フィルタリング結果
 */
function filterPagesByPublishDate(pages) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('統合データ');
  
  if (!sheet) {
    return {
      filtered: pages,
      excluded: [],
      message: '統合データシートなし'
    };
  }
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const urlColIndex = headers.indexOf('page_url');
  const rewriteOkColIndex = headers.indexOf('リライト可能');
  
  if (rewriteOkColIndex === -1) {
    return {
      filtered: pages,
      excluded: [],
      message: 'リライト可能列なし（投稿日同期未実行）'
    };
  }
  
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  
  // URLとリライト可否のマップを作成
  const rewriteOkMap = {};
  for (const row of data) {
    const url = row[urlColIndex];
    const rewriteOk = row[rewriteOkColIndex];
    if (url) {
      rewriteOkMap[url] = rewriteOk === '○';
    }
  }
  
  const filtered = [];
  const excluded = [];
  
  for (const page of pages) {
    const url = page.page_url || page.url || page.pageUrl;
    
    // リライト可否をチェック
    if (rewriteOkMap[url] === false) {
      excluded.push({
        ...page,
        exclusionReason: '投稿から3ヶ月未満'
      });
    } else {
      filtered.push(page);
    }
  }
  
  return {
    filtered: filtered,
    excluded: excluded,
    message: `${filtered.length}件がリライト対象、${excluded.length}件が3ヶ月未満で除外`
  };
}


/**
 * 3ヶ月チェック付きの優先ページ取得
 * （既存のgetTopPriorityPagesWithCooling()を拡張）
 * @param {number} limit - 取得件数
 * @return {Array} フィルタリング済みページ
 */
function getTopPriorityPagesWithPublishDateFilter(limit = 10) {
  // 既存の冷却期間チェック付き取得
  const pages = getTopPriorityPagesWithCooling ? 
                getTopPriorityPagesWithCooling(limit * 2) : 
                getTopPriorityPages(limit * 2);
  
  // 3ヶ月チェックでフィルタリング
  const result = filterPagesByPublishDate(pages);
  
  Logger.log('投稿日フィルタリング: ' + result.message);
  
  // 指定件数に制限
  return result.filtered.slice(0, limit);
}


// ============================================
// Scoring.gs 追記用関数
// ============================================

/**
 * 3ヶ月未満の記事かチェック（Scoring.gs用）
 * @param {string} pageUrl - ページURL
 * @return {boolean} 3ヶ月未満ならtrue
 */
function shouldExcludeByPublishDate(pageUrl) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('統合データ');
  
  if (!sheet) return false;
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const urlColIndex = headers.indexOf('page_url');
  const rewriteOkColIndex = headers.indexOf('リライト可能');
  
  if (rewriteOkColIndex === -1) return false;
  
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  
  for (const row of data) {
    if (row[urlColIndex] === pageUrl) {
      return row[rewriteOkColIndex] !== '○';
    }
  }
  
  return false;
}


// ============================================
// テスト関数
// ============================================

/**
 * 投稿日フィルタリング機能のテスト
 */
function testPublishDateFilter() {
  Logger.log('=== 投稿日フィルタリングテスト ===');
  
  // 1. 投稿日チェックテスト
  Logger.log('--- 1. 投稿日チェックテスト ---');
  
  // 1ヶ月前
  const oneMonthAgo = new Date();
  oneMonthAgo.setMonth(oneMonthAgo.getMonth() - 1);
  const check1 = isArticleTooNew(oneMonthAgo.toISOString());
  Logger.log('1ヶ月前: ' + (check1.isTooNew ? '❌除外' : '✅対象') + ' - ' + check1.reason);
  
  // 3ヶ月前
  const threeMonthsAgo = new Date();
  threeMonthsAgo.setMonth(threeMonthsAgo.getMonth() - 3);
  const check3 = isArticleTooNew(threeMonthsAgo.toISOString());
  Logger.log('3ヶ月前: ' + (check3.isTooNew ? '❌除外' : '✅対象') + ' - ' + check3.reason);
  
  // 6ヶ月前
  const sixMonthsAgo = new Date();
  sixMonthsAgo.setMonth(sixMonthsAgo.getMonth() - 6);
  const check6 = isArticleTooNew(sixMonthsAgo.toISOString());
  Logger.log('6ヶ月前: ' + (check6.isTooNew ? '❌除外' : '✅対象') + ' - ' + check6.reason);
  
  // 2. 実際のページでテスト
  Logger.log('--- 2. 実際のページ取得テスト ---');
  
  const testSlug = 'ipad-refurbished-restock-timing';
  const result = getPublishDateBySlug(testSlug);
  
  if (result.success) {
    Logger.log('投稿日: ' + result.publishDate);
    const check = isArticleTooNew(result.publishDate);
    Logger.log('判定: ' + (check.isTooNew ? '❌除外' : '✅対象') + ' - ' + check.reason);
  } else {
    Logger.log('取得失敗: ' + result.error);
  }
  
  // 3. 列追加テスト
  Logger.log('--- 3. 列追加テスト ---');
  addPublishDateColumn();
  
  Logger.log('=== テスト完了 ===');
}


/**
 * 投稿日同期のテスト（5件のみ）
 */
function testSyncPublishDates() {
  Logger.log('=== 投稿日同期テスト（5件） ===');
  const result = syncPublishDates(5);
  Logger.log('結果: ' + JSON.stringify(result));
}


/**
 * 初期セットアップ（列追加 + 同期）
 */
function setupPublishDateFilter() {
  Logger.log('=== 投稿日フィルター初期セットアップ ===');
  
  // 1. 列追加
  addPublishDateColumn();
  
  // 2. 投稿日同期（最初は50件）
  Logger.log('投稿日を同期中...');
  const result = syncPublishDates(50);
  
  Logger.log('=== セットアップ完了 ===');
  Logger.log('処理: ' + result.processed + '件 / 更新: ' + result.updated + '件');
  
  return result;
}

function debugUrlMatching() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('統合データ');
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const urlColIndex = headers.indexOf('page_url');
  
  // 統合データの最初の5件のURL
  Logger.log('=== 統合データのURL ===');
  const data = sheet.getRange(2, urlColIndex + 1, 5, 1).getValues();
  for (let i = 0; i < data.length; i++) {
    Logger.log((i+1) + '. ' + data[i][0]);
  }
  
  // WordPressの最初の5件のURL
  Logger.log('=== WordPressのURL ===');
  const posts = getAllWordPressPosts();
  for (let i = 0; i < Math.min(5, posts.length); i++) {
    Logger.log((i+1) + '. ' + posts[i].link);
  }
}