/**
 * ReportGenerator.gs
 * レポート出力機能 - Google Docs/PDF生成
 * 
 * 作成日: 2025年12月4日
 * バージョン: 1.0
 * 
 * 機能:
 * - リライト提案書生成（ページ単位）
 * - 効果測定レポート生成（リライト後）
 * - 週次レポート生成
 * - PDF変換・ダウンロードURL生成
 */

// ========================================
// 設定
// ========================================

const REPORT_CONFIG = {
  // レポート出力先フォルダID（Google Drive）
  // 初回実行時に自動作成されます
  get folderId() {
    let folderId = PropertiesService.getScriptProperties().getProperty('REPORT_FOLDER_ID');
    if (!folderId) {
      folderId = createReportFolder_();
    }
    return folderId;
  },
  
  // サイト名（レポートタイトルに使用）
  siteName: 'スマホタップ',
  
  // サイトURL
  siteUrl: 'https://smaho-tap.com'
};

/**
 * レポート出力用フォルダを作成（内部関数）
 */
function createReportFolder_() {
  const folderName = 'SEOリライト支援ツール_レポート';
  
  // 既存フォルダを検索
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    const folder = folders.next();
    PropertiesService.getScriptProperties().setProperty('REPORT_FOLDER_ID', folder.getId());
    Logger.log('既存フォルダを使用: ' + folder.getName());
    return folder.getId();
  }
  
  // 新規作成
  const folder = DriveApp.createFolder(folderName);
  PropertiesService.getScriptProperties().setProperty('REPORT_FOLDER_ID', folder.getId());
  Logger.log('フォルダを作成: ' + folder.getName());
  return folder.getId();
}

// ========================================
// リライト提案書生成
// ========================================

/**
 * リライト提案書を生成
 * 
 * @param {string} pageUrl - ページURL（例: /iphone-hoken-osusume/）
 * @param {boolean} convertToPdf - PDF変換するか（デフォルト: true）
 * @return {Object} {success, docUrl, pdfUrl, message}
 */
function generateRewriteProposal(pageUrl, convertToPdf = true) {
  Logger.log('=== リライト提案書生成開始 ===');
  Logger.log('対象ページ: ' + pageUrl);
  
  try {
    // 1. ページデータ取得
    const pageData = getPageDataForReport(pageUrl);
    if (!pageData) {
      return { success: false, message: 'ページデータが見つかりません: ' + pageUrl };
    }
    
    // 2. 競合分析データ取得
    const competitorData = getCompetitorDataForReport(pageUrl);
    
    // 3. クエリ分析データ取得
    const queryData = getQueryDataForReport(pageUrl);
    
    // 4. AI提案を取得（既存の提案があれば使用）
    const suggestions = getAISuggestionsForReport(pageUrl, pageData);
    
    // 5. Google Docs作成
    const doc = createProposalDocument(pageUrl, pageData, competitorData, queryData, suggestions);
    
    // 6. PDF変換（オプション）
    let pdfUrl = null;
    if (convertToPdf) {
      pdfUrl = convertDocToPdf(doc.getId());
    }
    
    Logger.log('=== リライト提案書生成完了 ===');
    
    return {
      success: true,
      docUrl: doc.getUrl(),
      pdfUrl: pdfUrl,
      docId: doc.getId(),
      message: '提案書を生成しました'
    };
    
  } catch (error) {
    Logger.log('❌ エラー: ' + error.message);
    return { success: false, message: error.message };
  }
}

/**
 * 提案書ドキュメントを作成
 */
function createProposalDocument(pageUrl, pageData, competitorData, queryData, suggestions) {
  const title = '【リライト提案書】' + (pageData.title || pageUrl) + '_' + formatDate_(new Date());
  
  // Google Docs作成
  const doc = DocumentApp.create(title);
  const body = doc.getBody();
  
  // フォルダに移動（エラー時はスキップ）
  try {
    const file = DriveApp.getFileById(doc.getId());
    const folderId = REPORT_CONFIG.folderId;
    if (folderId) {
      const folder = DriveApp.getFolderById(folderId);
      folder.addFile(file);
      DriveApp.getRootFolder().removeFile(file);
    }
  } catch (folderError) {
    Logger.log('⚠️ フォルダ移動スキップ: ' + folderError.message);
  }
  
  // ========================================
  // ドキュメント内容作成
  // ========================================
  
  // タイトル
  body.appendParagraph('リライト提案書')
    .setHeading(DocumentApp.ParagraphHeading.TITLE)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  
  body.appendParagraph('作成日: ' + formatDate_(new Date()))
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  
  body.appendHorizontalRule();
  
  // 1. ページ情報
  body.appendParagraph('1. ページ情報')
    .setHeading(DocumentApp.ParagraphHeading.HEADING1);
  
  // 投稿日フォーマット
  let publishDateStr = '（不明）';
  if (pageData.publishDate) {
    try {
      publishDateStr = Utilities.formatDate(new Date(pageData.publishDate), 'Asia/Tokyo', 'yyyy/MM/dd');
    } catch (e) {
      publishDateStr = String(pageData.publishDate);
    }
  }
  
  const pageInfoTable = [
    ['項目', '値'],
    ['URL', REPORT_CONFIG.siteUrl + pageUrl],
    ['タイトル', pageData.title || '（取得できません）'],
    ['ターゲットKW', pageData.targetKeyword || '（未設定）'],
    ['現在の順位', pageData.gyronPosition ? pageData.gyronPosition + '位' : '（データなし）'],
    ['月間PV', pageData.avgPageViews || 0],
    ['投稿日', publishDateStr],
    ['経過月数', pageData.monthsElapsed ? pageData.monthsElapsed + 'ヶ月' : '（不明）']
  ];
  appendTable_(body, pageInfoTable);
  
  // 2. 5軸スコア
  body.appendParagraph('2. 5軸スコア分析')
    .setHeading(DocumentApp.ParagraphHeading.HEADING1);
  
  const scoreTable = [
    ['スコア項目', '点数', '評価'],
    ['機会損失スコア', pageData.opportunityScore || 0, getScoreEvaluation_(pageData.opportunityScore)],
    ['パフォーマンススコア', pageData.performanceScore || 0, getScoreEvaluation_(pageData.performanceScore)],
    ['ビジネスインパクト', pageData.businessImpactScore || 0, getScoreEvaluation_(pageData.businessImpactScore)],
    ['キーワード戦略', pageData.keywordStrategyScore || 0, getScoreEvaluation_(pageData.keywordStrategyScore)],
    ['競合難易度', pageData.competitorDifficultyScore || 0, getScoreEvaluation_(pageData.competitorDifficultyScore)],
    ['【総合優先度】', pageData.totalPriorityScore || 0, getPriorityEvaluation_(pageData.totalPriorityScore)]
  ];
  appendTable_(body, scoreTable);
  
  // 3. 競合分析
  body.appendParagraph('3. 競合分析')
    .setHeading(DocumentApp.ParagraphHeading.HEADING1);
  
  if (competitorData && competitorData.length > 0) {
    body.appendParagraph('ターゲットKW「' + (pageData.targetKeyword || '') + '」の上位サイト:');
    
    const compTable = [['順位', 'サイト', 'DA', 'DA差']];
    competitorData.slice(0, 5).forEach((comp, index) => {
      compTable.push([
        (index + 1) + '位',
        comp.domain || comp.url,
        comp.da || '—',
        comp.da_diff || '—'
      ]);
    });
    appendTable_(body, compTable);
    
    body.appendParagraph('勝算度: ' + (pageData.winnableScore || '—') + '点 / 競合レベル: ' + (pageData.competitionLevel || '—'));
  } else {
    body.appendParagraph('（競合分析データなし）');
  }
  
  // 4. 主要クエリ
  body.appendParagraph('4. 主要クエリ分析')
    .setHeading(DocumentApp.ParagraphHeading.HEADING1);
  
  if (queryData && queryData.length > 0) {
    const queryTable = [['クエリ', '表示回数', 'クリック数', 'CTR', '順位']];
    queryData.slice(0, 10).forEach(q => {
      queryTable.push([
        q.query,
        q.impressions || 0,
        q.clicks || 0,
        (q.ctr ? (q.ctr * 100).toFixed(1) + '%' : '—'),
        q.position ? q.position.toFixed(1) : '—'
      ]);
    });
    appendTable_(body, queryTable);
  } else {
    body.appendParagraph('（クエリデータなし）');
  }
  
  // 5. AI提案
  body.appendParagraph('5. リライト提案')
    .setHeading(DocumentApp.ParagraphHeading.HEADING1);
  
  if (suggestions) {
    // タイトル改善
    if (suggestions.title) {
      body.appendParagraph('■ タイトル改善案')
        .setHeading(DocumentApp.ParagraphHeading.HEADING2);
      body.appendParagraph('現在: ' + (pageData.title || ''));
      body.appendParagraph('提案: ' + suggestions.title);
    }
    
    // メタディスクリプション
    if (suggestions.metaDescription) {
      body.appendParagraph('■ メタディスクリプション改善案')
        .setHeading(DocumentApp.ParagraphHeading.HEADING2);
      body.appendParagraph(suggestions.metaDescription);
    }
    
    // 見出し追加
    if (suggestions.headings && suggestions.headings.length > 0) {
      body.appendParagraph('■ 追加すべき見出し')
        .setHeading(DocumentApp.ParagraphHeading.HEADING2);
      suggestions.headings.forEach(h => {
        body.appendListItem(h);
      });
    }
    
    // コンテンツ追加
    if (suggestions.content) {
      body.appendParagraph('■ コンテンツ改善ポイント')
        .setHeading(DocumentApp.ParagraphHeading.HEADING2);
      body.appendParagraph(suggestions.content);
    }
    
    // その他提案
    if (suggestions.other) {
      body.appendParagraph('■ その他の提案')
        .setHeading(DocumentApp.ParagraphHeading.HEADING2);
      body.appendParagraph(suggestions.other);
    }
  } else {
    body.appendParagraph('（AI提案を生成中...後ほど更新されます）');
  }
  
  // 6. タスクチェックリスト
  body.appendParagraph('6. タスクチェックリスト')
    .setHeading(DocumentApp.ParagraphHeading.HEADING1);
  
  const tasks = [
    '□ タイトルを修正する',
    '□ メタディスクリプションを修正する',
    '□ H2見出しを追加/修正する',
    '□ 本文を加筆する',
    '□ 内部リンクを追加する',
    '□ 画像を追加/最適化する',
    '□ 構造化データを確認する'
  ];
  tasks.forEach(task => {
    body.appendParagraph(task);
  });
  
  // フッター
  body.appendHorizontalRule();
  body.appendParagraph('生成元: SEOリライト支援ツール')
    .setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  
  doc.saveAndClose();
  
  Logger.log('ドキュメント作成完了: ' + doc.getUrl());
  return doc;
}

// ========================================
// 効果測定レポート生成
// ========================================

/**
 * 効果測定レポートを生成
 * 
 * @param {string} rewriteId - リライト履歴ID
 * @param {boolean} convertToPdf - PDF変換するか
 * @return {Object} {success, docUrl, pdfUrl, message}
 */
function generateEffectReport(rewriteId, convertToPdf = true) {
  Logger.log('=== 効果測定レポート生成開始 ===');
  Logger.log('リライトID: ' + rewriteId);
  
  try {
    // 1. リライト履歴データ取得
    const rewriteData = getRewriteHistoryData(rewriteId);
    if (!rewriteData) {
      return { success: false, message: 'リライト履歴が見つかりません: ' + rewriteId };
    }
    
    // 2. Google Docs作成
    const doc = createEffectDocument(rewriteData);
    
    // 3. PDF変換（オプション）
    let pdfUrl = null;
    if (convertToPdf) {
      pdfUrl = convertDocToPdf(doc.getId());
    }
    
    Logger.log('=== 効果測定レポート生成完了 ===');
    
    return {
      success: true,
      docUrl: doc.getUrl(),
      pdfUrl: pdfUrl,
      message: '効果測定レポートを生成しました'
    };
    
  } catch (error) {
    Logger.log('❌ エラー: ' + error.message);
    return { success: false, message: error.message };
  }
}

/**
 * 効果測定ドキュメントを作成
 */
function createEffectDocument(rewriteData) {
  const title = '【効果測定】' + (rewriteData.page_url || '') + '_' + formatDate_(new Date());
  
  const doc = DocumentApp.create(title);
  const body = doc.getBody();
  
  // フォルダに移動（エラー時はスキップ）
try {
  const file = DriveApp.getFileById(doc.getId());
  const folderId = REPORT_CONFIG.folderId;
  if (folderId) {
    const folder = DriveApp.getFolderById(folderId);
    folder.addFile(file);
    DriveApp.getRootFolder().removeFile(file);
  }
} catch (folderError) {
  Logger.log('⚠️ フォルダ移動スキップ: ' + folderError.message);
}
  
  // タイトル
  body.appendParagraph('リライト効果測定レポート')
    .setHeading(DocumentApp.ParagraphHeading.TITLE)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  
  body.appendHorizontalRule();
  
  // 1. 基本情報
  body.appendParagraph('1. リライト概要')
    .setHeading(DocumentApp.ParagraphHeading.HEADING1);
  
  const infoTable = [
    ['項目', '値'],
    ['対象URL', rewriteData.page_url],
    ['リライト実施日', formatDate_(rewriteData.rewrite_date)],
    ['測定日', formatDate_(new Date())],
    ['経過日数', rewriteData.days_since_rewrite + '日']
  ];
  appendTable_(body, infoTable);
  
  // 2. Before/After比較
  body.appendParagraph('2. Before/After比較')
    .setHeading(DocumentApp.ParagraphHeading.HEADING1);
  
  const comparisonTable = [
    ['指標', 'Before', 'After', '変化'],
    ['順位', rewriteData.before_position || '—', rewriteData.after_position || '—', 
      calculateChange_(rewriteData.before_position, rewriteData.after_position, true)],
    ['CTR', formatPercent_(rewriteData.before_ctr), formatPercent_(rewriteData.after_ctr),
      calculateChange_(rewriteData.before_ctr, rewriteData.after_ctr)],
    ['PV', rewriteData.before_pv || '—', rewriteData.after_pv || '—',
      calculateChange_(rewriteData.before_pv, rewriteData.after_pv)],
    ['直帰率', formatPercent_(rewriteData.before_bounce), formatPercent_(rewriteData.after_bounce),
      calculateChange_(rewriteData.before_bounce, rewriteData.after_bounce, true)]
  ];
  appendTable_(body, comparisonTable);
  
  // 3. 判定結果
  body.appendParagraph('3. 判定結果')
    .setHeading(DocumentApp.ParagraphHeading.HEADING1);
  
  const result = rewriteData.success_flag || '判定待ち';
  const resultText = result === '成功' ? '✅ 成功（3指標中2つ以上改善）' :
                     result === '失敗' ? '❌ 失敗（改善なし）' :
                     '⏳ ' + result;
  
  body.appendParagraph(resultText)
    .setBold(true);
  
  // 4. 実施内容
  body.appendParagraph('4. 実施内容')
    .setHeading(DocumentApp.ParagraphHeading.HEADING1);
  
  body.appendParagraph(rewriteData.rewrite_summary || '（記録なし）');
  
  doc.saveAndClose();
  
  return doc;
}

// ========================================
// 週次レポート生成
// ========================================

/**
 * 週次レポートを生成
 * 
 * @param {boolean} convertToPdf - PDF変換するか
 * @return {Object} {success, docUrl, pdfUrl, message}
 */
function generateWeeklyReport(convertToPdf = true) {
  Logger.log('=== 週次レポート生成開始 ===');
  
  try {
    // 1. 週次データ収集
    const weeklyData = collectWeeklyData();
    
    // 2. Google Docs作成
    const doc = createWeeklyDocument(weeklyData);
    
    // 3. PDF変換（オプション）
    let pdfUrl = null;
    if (convertToPdf) {
      pdfUrl = convertDocToPdf(doc.getId());
    }
    
    Logger.log('=== 週次レポート生成完了 ===');
    
    return {
      success: true,
      docUrl: doc.getUrl(),
      pdfUrl: pdfUrl,
      message: '週次レポートを生成しました'
    };
    
  } catch (error) {
    Logger.log('❌ エラー: ' + error.message);
    return { success: false, message: error.message };
  }
}

/**
 * 週次データを収集
 */
function collectWeeklyData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 統合データシートから集計
  const integratedSheet = ss.getSheetByName('統合データ');
  const data = integratedSheet.getDataRange().getValues();
  const headers = data[0];
  
  // 列インデックス
  const cols = {
    totalScore: headers.indexOf('total_priority_score'),
    position: headers.indexOf('position'),
    pv: headers.indexOf('pageviews'),
    rewritable: headers.indexOf('リライト可能')
  };
  
  // 集計
  let totalPages = 0;
  let rewritablePages = 0;
  let highPriorityPages = 0;
  let top10Pages = 0;
  let top30Pages = 0;
  
  for (let i = 1; i < data.length; i++) {
    totalPages++;
    
    const score = data[i][cols.totalScore];
    const position = data[i][cols.position];
    const rewritable = data[i][cols.rewritable];
    
    if (rewritable !== '×') rewritablePages++;
    if (score >= 60) highPriorityPages++;
    if (position && position <= 10) top10Pages++;
    if (position && position <= 30) top30Pages++;
  }
  
  // リライト履歴から今週の実績
  const historySheet = ss.getSheetByName('リライト履歴');
  let weeklyRewrites = 0;
  let weeklySuccess = 0;
  
  if (historySheet) {
    const historyData = historySheet.getDataRange().getValues();
    const oneWeekAgo = new Date();
    oneWeekAgo.setDate(oneWeekAgo.getDate() - 7);
    
    for (let i = 1; i < historyData.length; i++) {
      const date = new Date(historyData[i][1]); // rewrite_date列
      if (date >= oneWeekAgo) {
        weeklyRewrites++;
        if (historyData[i][10] === '成功') weeklySuccess++; // success_flag列
      }
    }
  }
  
  // AIOサマリー
  const aioSheet = ss.getSheetByName('AIO順位履歴');
  let aioCount = 0;
  let ourSiteInAio = 0;
  
  if (aioSheet) {
    const aioData = aioSheet.getDataRange().getValues();
    const oneWeekAgo = new Date();
    oneWeekAgo.setDate(oneWeekAgo.getDate() - 7);
    
    for (let i = 1; i < aioData.length; i++) {
      const date = new Date(aioData[i][1]); // recorded_at列
      if (date >= oneWeekAgo) {
        if (aioData[i][2] === true || aioData[i][2] === 'TRUE') aioCount++; // has_aio列
        if (aioData[i][6] === true || aioData[i][6] === 'TRUE') ourSiteInAio++; // our_site_in_aio列
      }
    }
  }
  
  return {
    totalPages,
    rewritablePages,
    highPriorityPages,
    top10Pages,
    top30Pages,
    weeklyRewrites,
    weeklySuccess,
    aioCount,
    ourSiteInAio,
    generatedAt: new Date()
  };
}

/**
 * 週次ドキュメントを作成
 */
function createWeeklyDocument(weeklyData) {
  const title = '【週次レポート】' + REPORT_CONFIG.siteName + '_' + formatDate_(new Date());
  
  const doc = DocumentApp.create(title);
  const body = doc.getBody();
  
  // フォルダに移動（エラー時はスキップ）
try {
  const file = DriveApp.getFileById(doc.getId());
  const folderId = REPORT_CONFIG.folderId;
  if (folderId) {
    const folder = DriveApp.getFolderById(folderId);
    folder.addFile(file);
    DriveApp.getRootFolder().removeFile(file);
  }
} catch (folderError) {
  Logger.log('⚠️ フォルダ移動スキップ: ' + folderError.message);
  // フォルダ移動に失敗してもドキュメント作成は続行
}
  
  // タイトル
  body.appendParagraph('SEO週次レポート')
    .setHeading(DocumentApp.ParagraphHeading.TITLE)
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  
  body.appendParagraph(REPORT_CONFIG.siteName + ' - ' + formatDate_(weeklyData.generatedAt))
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  
  body.appendHorizontalRule();
  
  // 1. サマリー
  body.appendParagraph('1. 今週のサマリー')
    .setHeading(DocumentApp.ParagraphHeading.HEADING1);
  
  const summaryTable = [
    ['指標', '値'],
    ['総ページ数', weeklyData.totalPages],
    ['リライト可能ページ', weeklyData.rewritablePages],
    ['高優先度ページ（60点以上）', weeklyData.highPriorityPages],
    ['TOP10ページ数', weeklyData.top10Pages],
    ['TOP30ページ数', weeklyData.top30Pages]
  ];
  appendTable_(body, summaryTable);
  
  // 2. リライト実績
  body.appendParagraph('2. 今週のリライト実績')
    .setHeading(DocumentApp.ParagraphHeading.HEADING1);
  
  const rewriteTable = [
    ['指標', '値'],
    ['リライト実施数', weeklyData.weeklyRewrites + '件'],
    ['成功数', weeklyData.weeklySuccess + '件'],
    ['成功率', weeklyData.weeklyRewrites > 0 ? 
      Math.round(weeklyData.weeklySuccess / weeklyData.weeklyRewrites * 100) + '%' : '—']
  ];
  appendTable_(body, rewriteTable);
  
  // 3. AIOサマリー
  body.appendParagraph('3. AI Overview状況')
    .setHeading(DocumentApp.ParagraphHeading.HEADING1);
  
  const aioTable = [
    ['指標', '値'],
    ['AIO表示KW数', weeklyData.aioCount + '件'],
    ['自社サイト掲載数', weeklyData.ourSiteInAio + '件']
  ];
  appendTable_(body, aioTable);
  
  // 4. 推奨アクション
  body.appendParagraph('4. 今週の推奨アクション')
    .setHeading(DocumentApp.ParagraphHeading.HEADING1);
  
  if (weeklyData.highPriorityPages > 0) {
    body.appendParagraph('• 高優先度ページが' + weeklyData.highPriorityPages + '件あります。リライトを検討してください。');
  }
  if (weeklyData.aioCount > 0 && weeklyData.ourSiteInAio === 0) {
    body.appendParagraph('• AI Overviewに自社サイトが掲載されていません。AIO対策を検討してください。');
  }
  
  doc.saveAndClose();
  
  return doc;
}

// ========================================
// PDF変換
// ========================================

/**
 * Google DocsをPDFに変換
 * 
 * @param {string} docId - ドキュメントID
 * @return {string} PDFのURL
 */
function convertDocToPdf(docId) {
  try {
    const doc = DriveApp.getFileById(docId);
    const blob = doc.getAs('application/pdf');
    
    // PDF保存
    const folder = DriveApp.getFolderById(REPORT_CONFIG.folderId);
    const pdfName = doc.getName().replace(/\.gdoc$/, '') + '.pdf';
    const pdf = folder.createFile(blob.setName(pdfName));
    
    Logger.log('PDF作成: ' + pdf.getUrl());
    return pdf.getUrl();
    
  } catch (error) {
    Logger.log('PDF変換エラー: ' + error.message);
    return null;
  }
}

// ========================================
// データ取得ヘルパー
// ========================================

/**
 * ページデータを取得
 */
function getPageDataForReport(pageUrl) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('統合データ');
  
  if (!sheet) {
    return null;
  }
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  // 列インデックスを取得
  const colIndex = {};
  headers.forEach((header, index) => {
    colIndex[header] = index;
  });
  
  // URL列を特定
  const urlCol = colIndex['page_url'] !== undefined ? colIndex['page_url'] : colIndex['page_path'];
  
  // URLでページを検索
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const url = row[urlCol];
    
    // 正規化して比較
    if (normalizeUrlForReport_(url) === normalizeUrlForReport_(pageUrl)) {
      return {
        url: row[colIndex['page_url']] || '',
        title: row[colIndex['page_title']] || '（取得できません）',
        targetKeyword: row[colIndex['target_keyword']] || '',
        category: row[colIndex['category']] || '',
        publishDate: row[colIndex['投稿日']] || row[colIndex['publish_date']] || '',
        monthsElapsed: row[colIndex['経過月数']] || '（不明）',
        avgPageViews: row[colIndex['avg_page_views_30d']] || 0,
        avgSessionDuration: row[colIndex['avg_session_duration']] || 0,
        bounceRate: row[colIndex['bounce_rate']] || 0,
        conversions: row[colIndex['conversions_30d']] || 0,
        avgPosition: row[colIndex['avg_position']] || '',
        gyronPosition: row[colIndex['gyron_position']] || '（データなし）',
        totalClicks: row[colIndex['total_clicks_30d']] || 0,
        totalImpressions: row[colIndex['total_impressions_30d']] || 0,
        avgCtr: row[colIndex['avg_ctr']] || 0,
        topQueries: row[colIndex['top_queries']] || '',
        clarityScrollDepth: row[colIndex['clarity_avg_scroll_depth']] || 0,
        clarityDeadClicks: row[colIndex['clarity_dead_clicks']] || 0,
        clarityUxScore: row[colIndex['clarity_ux_score']] || 0,
        opportunityScore: row[colIndex['opportunity_score']] || 0,
        performanceScore: row[colIndex['performance_score']] || 0,
        businessImpactScore: row[colIndex['business_impact_score']] || 0,
        keywordStrategyScore: row[colIndex['keyword_strategy_score']] || 0,
        competitorDifficultyScore: row[colIndex['competitor_difficulty_score']] || 0,
        totalPriorityScore: row[colIndex['total_priority_score']] || 0,
        exclusionReason: row[colIndex['exclusion_reason']] || '',
        rewritable: row[colIndex['リライト可能']] || ''
      };
    }
  }
  
  return null;
}

/**
 * 競合分析データを取得
 */
function getCompetitorDataForReport(pageUrl) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('競合分析');
  
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  // ターゲットKWを取得
  const pageData = getPageDataForReport(pageUrl);
  if (!pageData || !pageData.target_keyword) return [];
  
  const kwCol = headers.indexOf('keyword');
  const results = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][kwCol] === pageData.target_keyword) {
      const row = {};
      headers.forEach((h, j) => {
        row[h] = data[i][j];
      });
      results.push(row);
    }
  }
  
  return results;
}

/**
 * クエリ分析データを取得
 */
function getQueryDataForReport(pageUrl) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('クエリ分析');
  
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const urlCol = headers.indexOf('page_url') !== -1 ? headers.indexOf('page_url') : headers.indexOf('page_path');
  
  const results = [];
  
  for (let i = 1; i < data.length; i++) {
    if (normalizeUrlForReport_(data[i][urlCol]) === normalizeUrlForReport_(pageUrl)) {
      const row = {};
      headers.forEach((h, j) => {
        row[h] = data[i][j];
      });
      results.push(row);
    }
  }
  
  // 表示回数でソート
  results.sort((a, b) => (b.impressions || 0) - (a.impressions || 0));
  
  return results;
}

/**
 * AI提案を取得（簡易版）
 */
function getAISuggestionsForReport(pageUrl, pageData) {
  // タスク管理シートから既存の提案を取得
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const taskSheet = ss.getSheetByName('タスク管理');
  
  if (!taskSheet) {
    return generateBasicSuggestions_(pageData);
  }
  
  const data = taskSheet.getDataRange().getValues();
  const headers = data[0];
  const urlCol = headers.indexOf('page_url');
  const suggestionCol = headers.indexOf('ai_suggestion');
  
  for (let i = 1; i < data.length; i++) {
    if (normalizeUrlForReport_(data[i][urlCol]) === normalizeUrlForReport_(pageUrl)) {
      try {
        return JSON.parse(data[i][suggestionCol]);
      } catch (e) {
        // JSON解析失敗
      }
    }
  }
  
  return generateBasicSuggestions_(pageData);
}

/**
 * 基本的な提案を生成
 */
function generateBasicSuggestions_(pageData) {
  const suggestions = {};
  
  // 順位に応じた提案
  if (pageData.position) {
    if (pageData.position <= 10) {
      suggestions.title = 'CTR向上のため、数字や具体的なベネフィットを追加';
      suggestions.content = 'TOP10入りしているため、CTR改善とコンテンツの鮮度維持に注力';
    } else if (pageData.position <= 30) {
      suggestions.title = 'ターゲットKWを前方に配置し、検索意図に合致させる';
      suggestions.content = '上位表示に向けて、網羅性とオリジナリティを強化';
      suggestions.headings = ['よくある質問（FAQ）を追加', '比較表を追加', '最新情報を追記'];
    } else {
      suggestions.title = '検索意図を再分析し、タイトルを全面改訂';
      suggestions.content = 'コンテンツの大幅リライトを検討。競合上位の構成を参考に';
    }
  }
  
  return suggestions;
}

/**
 * リライト履歴データを取得
 */
function getRewriteHistoryData(rewriteId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('リライト履歴');
  
  if (!sheet) return null;
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('rewrite_id');
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === rewriteId || data[i][idCol] === parseInt(rewriteId)) {
      const result = {};
      headers.forEach((h, j) => {
        result[h] = data[i][j];
      });
      
      // 経過日数を計算
      if (result.rewrite_date) {
        const rewriteDate = new Date(result.rewrite_date);
        const today = new Date();
        result.days_since_rewrite = Math.floor((today - rewriteDate) / (1000 * 60 * 60 * 24));
      }
      
      return result;
    }
  }
  
  return null;
}

// ========================================
// ユーティリティ関数
// ========================================

/**
 * テーブルを追加
 */
function appendTable_(body, data) {
  // すべての値を文字列に変換
  const stringData = data.map(row => 
    row.map(cell => cell === null || cell === undefined ? '' : String(cell))
  );
  
  const table = body.appendTable(stringData);
  
  // ヘッダー行のスタイル
  const headerRow = table.getRow(0);
  for (let i = 0; i < headerRow.getNumCells(); i++) {
    headerRow.getCell(i).setBackgroundColor('#4285f4').getChild(0).asParagraph().editAsText().setBold(true).setForegroundColor('#ffffff');
  }
  
  body.appendParagraph(''); // 余白
}

/**
 * 日付フォーマット
 */
function formatDate_(date) {
  if (!date) return '';
  const d = new Date(date);
  return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM/dd');
}

/**
 * パーセント表示
 */
function formatPercent_(value) {
  if (!value && value !== 0) return '—';
  return (value * 100).toFixed(1) + '%';
}

/**
 * 変化を計算
 */
function calculateChange_(before, after, lowerIsBetter = false) {
  if (!before || !after) return '—';
  
  const diff = after - before;
  const improved = lowerIsBetter ? diff < 0 : diff > 0;
  const sign = diff > 0 ? '+' : '';
  const emoji = improved ? '📈' : (diff === 0 ? '➡️' : '📉');
  
  return emoji + ' ' + sign + diff.toFixed(1);
}

/**
 * スコア評価
 */
function getScoreEvaluation_(score) {
  if (!score && score !== 0) return '—';
  if (score >= 80) return '🔴 要改善';
  if (score >= 60) return '🟠 注意';
  if (score >= 40) return '🟡 普通';
  return '🟢 良好';
}

/**
 * 優先度評価
 */
function getPriorityEvaluation_(score) {
  if (!score && score !== 0) return '—';
  if (score >= 80) return '🔴 最優先';
  if (score >= 60) return '🟠 高優先';
  if (score >= 40) return '🟡 中優先';
  return '🟢 低優先';
}

/**
 * URL正規化
 */
function normalizeUrlForReport_(url) {
  if (!url) return '';
  let path = String(url);
  
  // 絶対URLの場合
  if (path.startsWith('http')) {
    try {
      path = new URL(path).pathname;
    } catch (e) {}
  }
  
  // 末尾スラッシュを除去
  path = path.replace(/\/$/, '');
  
  return path || '/';
}

// ========================================
// テスト関数
// ========================================

/**
 * リライト提案書生成テスト
 */
function testGenerateProposal() {
  // テスト用URL（実際に存在するページURLに変更してください）
  const testUrl = '/ipad-mini-cheap-buy-methods';
  
  Logger.log('=== リライト提案書生成テスト ===');
  const result = generateRewriteProposal(testUrl, false); // PDF変換なし
  
  Logger.log('結果: ' + JSON.stringify(result, null, 2));
  
  if (result.success) {
    Logger.log('✅ 成功: ' + result.docUrl);
  } else {
    Logger.log('❌ 失敗: ' + result.message);
  }
}

/**
 * 週次レポート生成テスト
 */
function testGenerateWeeklyReport() {
  Logger.log('=== 週次レポート生成テスト ===');
  const result = generateWeeklyReport(false); // PDF変換なし
  
  Logger.log('結果: ' + JSON.stringify(result, null, 2));
  
  if (result.success) {
    Logger.log('✅ 成功: ' + result.docUrl);
  } else {
    Logger.log('❌ 失敗: ' + result.message);
  }
}

/**
 * Google Docs権限テスト
 */
function testDocPermission() {
  try {
    // 簡単なドキュメント作成テスト
    const doc = DocumentApp.create('テスト_削除OK_' + new Date().getTime());
    const docUrl = doc.getUrl();
    Logger.log('✅ ドキュメント作成成功: ' + docUrl);
    
    // テストドキュメントを削除
    DriveApp.getFileById(doc.getId()).setTrashed(true);
    Logger.log('✅ テストドキュメント削除完了');
    
    return true;
  } catch (error) {
    Logger.log('❌ エラー: ' + error.message);
    return false;
  }
}

/**
 * 週次レポート生成テスト（シンプル版）
 */
function testGenerateWeeklyReportSimple() {
  Logger.log('=== シンプル版テスト開始 ===');
  
  try {
    // Step 1: データ収集
    Logger.log('Step 1: データ収集');
    const weeklyData = collectWeeklyData();
    Logger.log('データ収集完了: ' + JSON.stringify(weeklyData));
    
    // Step 2: ドキュメント作成
    Logger.log('Step 2: ドキュメント作成');
    const doc = createWeeklyDocument(weeklyData);
    Logger.log('ドキュメント作成完了: ' + doc.getUrl());
    
    Logger.log('=== シンプル版テスト成功 ===');
    return doc.getUrl();
    
  } catch (error) {
    Logger.log('❌ エラー発生箇所: ' + error.message);
    Logger.log('スタック: ' + error.stack);
    return null;
  }
}

/**
 * 週次レポート生成テスト（フォルダ移動なし）
 */
function testCreateDocWithoutFolder() {
  Logger.log('=== フォルダ移動なしテスト ===');
  
  try {
    const title = '【週次レポート】テスト_' + new Date().getTime();
    
    // ドキュメント作成
    Logger.log('ドキュメント作成中...');
    const doc = DocumentApp.create(title);
    const body = doc.getBody();
    
    // 簡単な内容を追加
    body.appendParagraph('SEO週次レポート')
      .setHeading(DocumentApp.ParagraphHeading.TITLE);
    body.appendParagraph('テスト作成日: ' + new Date().toLocaleString('ja-JP'));
    
    doc.saveAndClose();
    
    Logger.log('✅ 成功: ' + doc.getUrl());
    return doc.getUrl();
    
  } catch (error) {
    Logger.log('❌ エラー: ' + error.message);
    Logger.log('スタック: ' + error.stack);
    return null;
  }
}/**
 * 週次レポート生成テスト（フォルダ移動なし）
 */
function testCreateDocWithoutFolder() {
  Logger.log('=== フォルダ移動なしテスト ===');
  
  try {
    const title = '【週次レポート】テスト_' + new Date().getTime();
    
    // ドキュメント作成
    Logger.log('ドキュメント作成中...');
    const doc = DocumentApp.create(title);
    const body = doc.getBody();
    
    // 簡単な内容を追加
    body.appendParagraph('SEO週次レポート')
      .setHeading(DocumentApp.ParagraphHeading.TITLE);
    body.appendParagraph('テスト作成日: ' + new Date().toLocaleString('ja-JP'));
    
    doc.saveAndClose();
    
    Logger.log('✅ 成功: ' + doc.getUrl());
    return doc.getUrl();
    
  } catch (error) {
    Logger.log('❌ エラー: ' + error.message);
    Logger.log('スタック: ' + error.stack);
    return null;
  }
}

/**
 * レポートフォルダの状態を確認
 */
function checkReportFolder() {
  Logger.log('=== レポートフォルダ確認 ===');
  
  // 保存されているフォルダIDを確認
  const savedFolderId = PropertiesService.getScriptProperties().getProperty('REPORT_FOLDER_ID');
  Logger.log('保存されているフォルダID: ' + (savedFolderId || 'なし'));
  
  if (savedFolderId) {
    try {
      const folder = DriveApp.getFolderById(savedFolderId);
      Logger.log('✅ フォルダ存在: ' + folder.getName());
      Logger.log('フォルダURL: ' + folder.getUrl());
    } catch (e) {
      Logger.log('❌ フォルダが見つかりません: ' + e.message);
      Logger.log('→ フォルダIDをリセットします');
      PropertiesService.getScriptProperties().deleteProperty('REPORT_FOLDER_ID');
    }
  } else {
    Logger.log('フォルダIDが保存されていません → 新規作成を試みます');
  }
  
  // フォルダ作成を試みる
  Logger.log('\n--- フォルダ作成テスト ---');
  try {
    const folderId = createReportFolder_();
    Logger.log('✅ フォルダID: ' + folderId);
    
    const folder = DriveApp.getFolderById(folderId);
    Logger.log('✅ フォルダ名: ' + folder.getName());
    Logger.log('✅ フォルダURL: ' + folder.getUrl());
  } catch (e) {
    Logger.log('❌ フォルダ作成エラー: ' + e.message);
  }
}

/**
 * レポートフォルダIDを手動設定
 */
function setReportFolderId() {
  const folderId = '15O1niCKr1kFTdqklP63e4rdGTkQ0nhok';
  
  PropertiesService.getScriptProperties().setProperty('REPORT_FOLDER_ID', folderId);
  
  Logger.log('✅ フォルダIDを保存しました: ' + folderId);
  
  // 確認
  try {
    const folder = DriveApp.getFolderById(folderId);
    Logger.log('✅ フォルダ名: ' + folder.getName());
    Logger.log('✅ フォルダURL: ' + folder.getUrl());
  } catch (e) {
    Logger.log('❌ フォルダ確認エラー: ' + e.message);
  }
}

/**
 * PDF変換テスト
 */
function testConvertToPdf() {
  Logger.log('=== PDF変換テスト ===');
  
  // 先ほど生成したドキュメントIDを使用
  const docId = '1wIuTMf3qesT5mMIo3vDnJYpHNpMFhHqrYQdTjFdG4eo';
  
  try {
    const pdfUrl = convertDocToPdf(docId);
    Logger.log('✅ PDF変換成功: ' + pdfUrl);
  } catch (e) {
    Logger.log('❌ PDF変換エラー: ' + e.message);
  }
}

/**
 * 統合データシートの列名を確認
 */
function debugIntegratedColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('統合データ');
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  Logger.log('=== 統合データシート 列名一覧 ===');
  headers.forEach((header, index) => {
    Logger.log((index + 1) + ': ' + header);
  });
  
  // サンプルデータ（2行目）
  const sampleData = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
  Logger.log('\n=== サンプルデータ（2行目） ===');
  headers.forEach((header, index) => {
    if (sampleData[index]) {
      Logger.log(header + ': ' + sampleData[index]);
    }
  });
}

/**
 * ページデータ取得をデバッグ
 */
function debugGetPageData() {
  const testUrl = '/geo-battery-deterioration';
  
  Logger.log('=== ページデータ取得デバッグ ===');
  Logger.log('検索URL: ' + testUrl);
  
  const data = getPageDataForReport(testUrl);
  
  if (data) {
    Logger.log('✅ データ取得成功');
    Logger.log('title: ' + data.title);
    Logger.log('targetKeyword: ' + data.targetKeyword);
    Logger.log('gyronPosition: ' + data.gyronPosition);
    Logger.log('avgPageViews: ' + data.avgPageViews);
    Logger.log('publishDate: ' + data.publishDate);
    Logger.log('monthsElapsed: ' + data.monthsElapsed);
    Logger.log('opportunityScore: ' + data.opportunityScore);
    Logger.log('performanceScore: ' + data.performanceScore);
    Logger.log('businessImpactScore: ' + data.businessImpactScore);
    Logger.log('totalPriorityScore: ' + data.totalPriorityScore);
  } else {
    Logger.log('❌ データ取得失敗');
  }
}

/**
 * 製品化準備用シートを作成
 */
function createProductizationSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  Logger.log('=== 製品化準備用シート作成 ===');
  
  // ========================================
  // 1. 効果検証詳細シート
  // ========================================
  let sheet1 = ss.getSheetByName('効果検証詳細');
  if (!sheet1) {
    sheet1 = ss.insertSheet('効果検証詳細');
    Logger.log('✅ 効果検証詳細シート作成');
  } else {
    Logger.log('⚠️ 効果検証詳細シートは既存');
  }
  
  const headers1 = [
    'rewrite_id',
    'page_url',
    'target_keyword',
    'rewrite_date',
    'before_title',
    'after_title',
    '変更内容サマリー',
    '変更詳細',
    'before_position',
    'before_pv_30d',
    'before_ctr',
    'position_7d',
    'position_14d',
    'position_30d',
    'pv_30d_after',
    'ctr_30d_after',
    '成功判定',
    '成功要因_失敗理由',
    'LP掲載可否',
    '備考'
  ];
  
  sheet1.getRange(1, 1, 1, headers1.length).setValues([headers1]);
  sheet1.getRange(1, 1, 1, headers1.length)
    .setBackground('#4285f4')
    .setFontColor('white')
    .setFontWeight('bold');
  sheet1.setFrozenRows(1);
  
  // 列幅調整
  sheet1.setColumnWidth(1, 100);  // rewrite_id
  sheet1.setColumnWidth(2, 250);  // page_url
  sheet1.setColumnWidth(3, 200);  // target_keyword
  sheet1.setColumnWidth(4, 100);  // rewrite_date
  sheet1.setColumnWidth(5, 300);  // before_title
  sheet1.setColumnWidth(6, 300);  // after_title
  sheet1.setColumnWidth(7, 250);  // 変更内容サマリー
  sheet1.setColumnWidth(8, 400);  // 変更詳細
  sheet1.setColumnWidth(17, 100); // 成功判定
  sheet1.setColumnWidth(18, 300); // 成功要因_失敗理由
  sheet1.setColumnWidth(19, 100); // LP掲載可否
  sheet1.setColumnWidth(20, 300); // 備考
  
  // 成功判定のドロップダウン
  const successRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['成功', '微妙', '失敗', '判定中'], true)
    .build();
  sheet1.getRange('Q2:Q100').setDataValidation(successRule);
  
  // LP掲載可否のドロップダウン
  const lpRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['○ 掲載OK', '△ 要検討', '× 非掲載'], true)
    .build();
  sheet1.getRange('S2:S100').setDataValidation(lpRule);
  
  Logger.log('✅ 効果検証詳細シート設定完了（20列）');
  
  // ========================================
  // 2. フィードバック・改善点シート
  // ========================================
  let sheet2 = ss.getSheetByName('フィードバック・改善点');
  if (!sheet2) {
    sheet2 = ss.insertSheet('フィードバック・改善点');
    Logger.log('✅ フィードバック・改善点シート作成');
  } else {
    Logger.log('⚠️ フィードバック・改善点シートは既存');
  }
  
  const headers2 = [
    'feedback_id',
    '記録日',
    'カテゴリ',
    '優先度',
    '対象機能',
    '問題・改善点',
    '再現手順',
    '期待する動作',
    '影響度',
    'ステータス',
    '対応日',
    '対応内容',
    '備考'
  ];
  
  sheet2.getRange(1, 1, 1, headers2.length).setValues([headers2]);
  sheet2.getRange(1, 1, 1, headers2.length)
    .setBackground('#ea4335')
    .setFontColor('white')
    .setFontWeight('bold');
  sheet2.setFrozenRows(1);
  
  // 列幅調整
  sheet2.setColumnWidth(1, 100);  // feedback_id
  sheet2.setColumnWidth(2, 100);  // 記録日
  sheet2.setColumnWidth(3, 120);  // カテゴリ
  sheet2.setColumnWidth(4, 80);   // 優先度
  sheet2.setColumnWidth(5, 150);  // 対象機能
  sheet2.setColumnWidth(6, 400);  // 問題・改善点
  sheet2.setColumnWidth(7, 300);  // 再現手順
  sheet2.setColumnWidth(8, 300);  // 期待する動作
  sheet2.setColumnWidth(9, 80);   // 影響度
  sheet2.setColumnWidth(10, 100); // ステータス
  sheet2.setColumnWidth(11, 100); // 対応日
  sheet2.setColumnWidth(12, 300); // 対応内容
  sheet2.setColumnWidth(13, 250); // 備考
  
  // カテゴリのドロップダウン
  const categoryRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['バグ', 'UI/UX', '機能改善', '新機能要望', 'パフォーマンス', 'ドキュメント', 'その他'], true)
    .build();
  sheet2.getRange('C2:C200').setDataValidation(categoryRule);
  
  // 優先度のドロップダウン
  const priorityRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['高', '中', '低'], true)
    .build();
  sheet2.getRange('D2:D200').setDataValidation(priorityRule);
  
  // 対象機能のドロップダウン
  const functionRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['チャット', 'スコアリング', 'レポート出力', '競合分析', 'タスク管理', 'データ収集', 'Clarity連携', 'WordPress連携', 'UI全般', 'その他'], true)
    .build();
  sheet2.getRange('E2:E200').setDataValidation(functionRule);
  
  // 影響度のドロップダウン
  const impactRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['大', '中', '小'], true)
    .build();
  sheet2.getRange('I2:I200').setDataValidation(impactRule);
  
  // ステータスのドロップダウン
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['未対応', '対応中', '完了', '保留', '不採用'], true)
    .build();
  sheet2.getRange('J2:J200').setDataValidation(statusRule);
  
  // 条件付き書式：優先度「高」を赤背景
  const highPriorityRange = sheet2.getRange('D2:D200');
  const highRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('高')
    .setBackground('#ffcdd2')
    .setRanges([highPriorityRange])
    .build();
  
  // 条件付き書式：ステータス「完了」を緑背景
  const statusRange = sheet2.getRange('J2:J200');
  const completeRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('完了')
    .setBackground('#c8e6c9')
    .setRanges([statusRange])
    .build();
  
  const rules = sheet2.getConditionalFormatRules();
  rules.push(highRule);
  rules.push(completeRule);
  sheet2.setConditionalFormatRules(rules);
  
  Logger.log('✅ フィードバック・改善点シート設定完了（13列）');
  
  // ========================================
  // 完了サマリー
  // ========================================
  Logger.log('');
  Logger.log('=== 作成完了 ===');
  Logger.log('1. 効果検証詳細（20列）- リライト効果の詳細記録');
  Logger.log('2. フィードバック・改善点（13列）- MVP改善点の記録');
  Logger.log('');
  Logger.log('スプレッドシート: 22シート構成になりました');
}