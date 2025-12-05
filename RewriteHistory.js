/**
 * リライト履歴管理機能
 * リライトの記録、効果測定、成功パターン分析を行う
 */

/**
 * リライト実施を記録する
 * @param {Object} rewriteData - リライト情報
 * @param {string} rewriteData.page_url - 対象ページURL
 * @param {string} rewriteData.rewrite_type - リライト種別（タイトル/本文/構造/全体）
 * @param {string} rewriteData.changes_summary - 変更内容サマリー
 * @param {string} rewriteData.ai_suggestion - AIの提案内容
 * @return {string} rewriteId - 生成されたリライトID
 */
function recordRewrite(rewriteData) {
  try {
    Logger.log('=== リライト記録開始 ===');
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const historySheet = ss.getSheetByName('リライト履歴');
    
    if (!historySheet) {
      throw new Error('リライト履歴シートが見つかりません');
    }
    
    // リライトID生成（日付 + ランダム文字列）
    const rewriteId = generateRewriteId();
    Logger.log(`生成されたリライトID: ${rewriteId}`);
    
    // リライト前の指標を取得
    const pageUrl = rewriteData.page_url;
    Logger.log(`対象URL: ${pageUrl}`);
    
    const beforeData = getBeforeMetrics(pageUrl);
    
    if (!beforeData) {
      throw new Error('リライト前の指標が取得できませんでした');
    }
    
    Logger.log('リライト前の指標取得成功');
    
    // リライト履歴に記録
    const now = new Date();
    const row = [
      rewriteId,                           // A: rewrite_id
      pageUrl,                             // B: page_url
      now,                                 // C: rewrite_date
      rewriteData.rewrite_type || '全体',   // D: rewrite_type
      rewriteData.changes_summary || '',   // E: changes_summary
      rewriteData.ai_suggestion || '',     // F: ai_suggestion
      // Before指標
      beforeData.position || 0,            // G: before_position
      beforeData.ctr || 0,                 // H: before_ctr
      beforeData.pageViews || 0,           // I: before_pv
      beforeData.bounceRate || 0,          // J: before_bounce_rate
      // After指標（初期値は空）
      '',                                  // K: after_position
      '',                                  // L: after_ctr
      '',                                  // M: after_pv
      '',                                  // N: after_bounce_rate
      // 変化
      '',                                  // O: position_change
      '',                                  // P: ctr_change_rate
      '',                                  // Q: pv_change_rate
      '',                                  // R: success_flag
      '',                                  // S: notes
      ''                                   // T: measured_date
    ];
    
    historySheet.appendRow(row);
    
    Logger.log('リライト履歴シートに記録完了');
    Logger.log(`=== リライト記録完了: ${rewriteId} ===`);
    
    return rewriteId;
    
  } catch (error) {
    Logger.log(`リライト記録エラー: ${error.message}`);
    throw error;
  }
}

/**
 * リライトIDを生成する
 * @return {string} rewriteId
 */
function generateRewriteId() {
  const date = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
  const random = Math.random().toString(36).substring(2, 8).toUpperCase();
  return `RW_${date}_${random}`;
}

/**
 * リライト前の指標を取得
 * @param {string} pageUrl - ページURL
 * @return {Object} beforeData - リライト前の指標
 */
function getBeforeMetrics(pageUrl) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const integratedSheet = ss.getSheetByName('統合データ');
    
    if (!integratedSheet) {
      throw new Error('統合データシートが見つかりません');
    }
    
    const data = integratedSheet.getDataRange().getValues();
    const headers = data[0];
    
    // 列インデックス取得
    const urlIdx = headers.indexOf('page_url');
    const positionIdx = headers.indexOf('avg_position');
    const ctrIdx = headers.indexOf('avg_ctr');
    const pvIdx = headers.indexOf('avg_page_views_30d');
    const bounceRateIdx = headers.indexOf('bounce_rate');
    
    // URLで検索
    for (let i = 1; i < data.length; i++) {
      if (data[i][urlIdx] === pageUrl) {
        return {
          position: data[i][positionIdx],
          ctr: data[i][ctrIdx],
          pageViews: data[i][pvIdx],
          bounceRate: data[i][bounceRateIdx]
        };
      }
    }
    
    return null;
    
  } catch (error) {
    Logger.log(`リライト前指標取得エラー: ${error.message}`);
    return null;
  }
}

/**
 * リライト効果を測定する（リライト後30日経過時に実行）
 * @param {string} rewriteId - リライトID
 * @return {Object} result - 効果測定結果
 */
function measureRewriteEffect(rewriteId) {
  try {
    Logger.log('=== リライト効果測定開始 ===');
    Logger.log(`対象リライトID: ${rewriteId}`);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const historySheet = ss.getSheetByName('リライト履歴');
    
    if (!historySheet) {
      throw new Error('リライト履歴シートが見つかりません');
    }
    
    const data = historySheet.getDataRange().getValues();
    const headers = data[0];
    
    // 列インデックス
    const idIdx = headers.indexOf('rewrite_id');
    const urlIdx = headers.indexOf('page_url');
    const dateIdx = headers.indexOf('rewrite_date');
    const beforePosIdx = headers.indexOf('before_position');
    const beforeCtrIdx = headers.indexOf('before_ctr');
    const beforePvIdx = headers.indexOf('before_pv');
    const beforeBounceIdx = headers.indexOf('before_bounce_rate');
    
    // リライトIDで検索
    let targetRow = -1;
    let targetData = null;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][idIdx] === rewriteId) {
        targetRow = i + 1; // シート行番号（1-indexed）
        targetData = data[i];
        break;
      }
    }
    
    if (!targetData) {
      throw new Error(`リライトID ${rewriteId} が見つかりません`);
    }
    
    const pageUrl = targetData[urlIdx];
    const rewriteDate = new Date(targetData[dateIdx]);
    
    Logger.log(`対象URL: ${pageUrl}`);
    Logger.log(`リライト日: ${Utilities.formatDate(rewriteDate, Session.getScriptTimeZone(), 'yyyy-MM-dd')}`);
    
    // 現在日付との差分チェック
    const now = new Date();
    const daysSinceRewrite = Math.floor((now - rewriteDate) / (1000 * 60 * 60 * 24));
    
    Logger.log(`リライトから ${daysSinceRewrite} 日経過`);
    
    if (daysSinceRewrite < 30) {
      Logger.log('警告: リライトから30日経過していません');
      return {
        success: false,
        message: `リライトから30日経過していません（${daysSinceRewrite}日）`
      };
    }
    
    // リライト後の指標を取得
    const afterData = getAfterMetrics(pageUrl);
    
    if (!afterData) {
      throw new Error('リライト後の指標が取得できませんでした');
    }
    
    Logger.log('リライト後の指標取得成功');
    
    // Before/After比較
    const beforeMetrics = {
      position: targetData[beforePosIdx],
      ctr: targetData[beforeCtrIdx],
      pageViews: targetData[beforePvIdx],
      bounceRate: targetData[beforeBounceIdx]
    };
    
    const comparison = compareMetrics(beforeMetrics, afterData);
    
    Logger.log('Before/After比較完了');
    Logger.log(`順位変化: ${comparison.positionChange}`);
    Logger.log(`CTR変化率: ${comparison.ctrChangeRate}%`);
    Logger.log(`PV変化率: ${comparison.pvChangeRate}%`);
    Logger.log(`成功判定: ${comparison.success ? '成功' : '要改善'}`);
    
    // リライト履歴シートを更新
    historySheet.getRange(targetRow, headers.indexOf('after_position') + 1).setValue(afterData.position);
    historySheet.getRange(targetRow, headers.indexOf('after_ctr') + 1).setValue(afterData.ctr);
    historySheet.getRange(targetRow, headers.indexOf('after_pv') + 1).setValue(afterData.pageViews);
    historySheet.getRange(targetRow, headers.indexOf('after_bounce_rate') + 1).setValue(afterData.bounceRate);
    historySheet.getRange(targetRow, headers.indexOf('position_change') + 1).setValue(comparison.positionChange);
    historySheet.getRange(targetRow, headers.indexOf('ctr_change_rate') + 1).setValue(comparison.ctrChangeRate);
    historySheet.getRange(targetRow, headers.indexOf('pv_change_rate') + 1).setValue(comparison.pvChangeRate);
    historySheet.getRange(targetRow, headers.indexOf('success_flag') + 1).setValue(comparison.success);
    historySheet.getRange(targetRow, headers.indexOf('measured_date') + 1).setValue(now);
    
    Logger.log('リライト履歴シート更新完了');
    Logger.log('=== リライト効果測定完了 ===');
    
    return {
      success: true,
      rewriteId: rewriteId,
      pageUrl: pageUrl,
      before: beforeMetrics,
      after: afterData,
      comparison: comparison
    };
    
  } catch (error) {
    Logger.log(`リライト効果測定エラー: ${error.message}`);
    throw error;
  }
}

/**
 * リライト後の指標を取得
 * @param {string} pageUrl - ページURL
 * @return {Object} afterData - リライト後の指標
 */
function getAfterMetrics(pageUrl) {
  // getBeforeMetrics()と同じロジックで現在の指標を取得
  return getBeforeMetrics(pageUrl);
}

/**
 * Before/After指標を比較
 * @param {Object} before - リライト前指標
 * @param {Object} after - リライト後指標
 * @return {Object} comparison - 比較結果
 */
function compareMetrics(before, after) {
  const positionChange = before.position - after.position; // 順位は低い方が良い
  const ctrChangeRate = ((after.ctr - before.ctr) / before.ctr * 100).toFixed(2);
  const pvChangeRate = ((after.pageViews - before.pageViews) / before.pageViews * 100).toFixed(2);
  
  // 成功判定（3つの指標のうち2つ以上が改善）
  let improvements = 0;
  if (positionChange > 0) improvements++; // 順位上昇
  if (parseFloat(ctrChangeRate) > 0) improvements++; // CTR上昇
  if (parseFloat(pvChangeRate) > 0) improvements++; // PV上昇
  
  const success = improvements >= 2;
  
  return {
    positionChange: positionChange,
    ctrChangeRate: parseFloat(ctrChangeRate),
    pvChangeRate: parseFloat(pvChangeRate),
    improvements: improvements,
    success: success
  };
}

/**
 * リライト履歴を取得（フィルター対応）
 * @param {Object} filters - フィルター条件（オプション）
 * @param {string} filters.page_url - ページURL
 * @param {Date} filters.start_date - 開始日
 * @param {Date} filters.end_date - 終了日
 * @param {boolean} filters.success_only - 成功したもののみ
 * @return {Array} history - リライト履歴
 */
function getRewriteHistory(filters) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const historySheet = ss.getSheetByName('リライト履歴');
    
    if (!historySheet) {
      throw new Error('リライト履歴シートが見つかりません');
    }
    
    const data = historySheet.getDataRange().getValues();
    const headers = data[0];
    
    // ヘッダー行を除く
    const rows = data.slice(1);
    
    // フィルター適用
    let filteredRows = rows;
    
    if (filters) {
      if (filters.page_url) {
        const urlIdx = headers.indexOf('page_url');
        filteredRows = filteredRows.filter(row => row[urlIdx] === filters.page_url);
      }
      
      if (filters.start_date) {
        const dateIdx = headers.indexOf('rewrite_date');
        filteredRows = filteredRows.filter(row => new Date(row[dateIdx]) >= filters.start_date);
      }
      
      if (filters.end_date) {
        const dateIdx = headers.indexOf('rewrite_date');
        filteredRows = filteredRows.filter(row => new Date(row[dateIdx]) <= filters.end_date);
      }
      
      if (filters.success_only) {
        const successIdx = headers.indexOf('success_flag');
        filteredRows = filteredRows.filter(row => row[successIdx] === true);
      }
    }
    
    // オブジェクト配列に変換
    const history = filteredRows.map(row => {
      const obj = {};
      headers.forEach((header, idx) => {
        obj[header] = row[idx];
      });
      return obj;
    });
    
    Logger.log(`リライト履歴取得: ${history.length}件`);
    
    return history;
    
  } catch (error) {
    Logger.log(`リライト履歴取得エラー: ${error.message}`);
    return [];
  }
}

/**
 * 成功パターンを分析（Claude API使用）
 * @return {string} analysis - 分析結果
 */
function analyzeSuccessPatterns() {
  try {
    Logger.log('=== 成功パターン分析開始 ===');
    
    // 成功したリライトのみ取得
    const successHistory = getRewriteHistory({ success_only: true });
    
    if (successHistory.length === 0) {
      Logger.log('成功したリライトがまだありません');
      return '成功したリライトがまだありません。30日後に効果測定を実行してください。';
    }
    
    Logger.log(`成功リライト件数: ${successHistory.length}件`);
    
    // 分析用データ整形
    const analysisData = successHistory.map(h => ({
      page_url: h.page_url,
      rewrite_type: h.rewrite_type,
      changes_summary: h.changes_summary,
      position_change: h.position_change,
      ctr_change_rate: h.ctr_change_rate,
      pv_change_rate: h.pv_change_rate
    }));
    
    // Claude APIで分析
    const prompt = buildSuccessPatternPrompt(analysisData);
    
    // ClaudeAPI.gsの関数を呼び出し
    const analysis = callClaudeAPI(prompt, buildSuccessPatternSystemPrompt());
    
    Logger.log('=== 成功パターン分析完了 ===');
    
    return analysis;
    
  } catch (error) {
    Logger.log(`成功パターン分析エラー: ${error.message}`);
    return `分析エラー: ${error.message}`;
  }
}

/**
 * 成功パターン分析のシステムプロンプト
 * @return {string} systemPrompt
 */
function buildSuccessPatternSystemPrompt() {
  return `あなたはSEOリライトの成功パターンを分析する専門家です。

過去の成功事例から共通点・パターンを抽出し、今後のリライトに活かせる知見を提供してください。

【分析の観点】
1. リライト種別の傾向（どのタイプが成功しやすいか）
2. 変更内容の共通点（どんな改善が効果的か）
3. 効果の大きさ（順位・CTR・PVの変化）
4. 今後の推奨アクション

具体的で実用的な知見を提供してください。`;
}

/**
 * 成功パターン分析プロンプト構築
 * @param {Array} analysisData - 分析データ
 * @return {string} prompt
 */
function buildSuccessPatternPrompt(analysisData) {
  let prompt = `以下は過去に成功したリライトのデータです。共通点やパターンを分析してください。\n\n`;
  
  prompt += `【成功したリライト一覧】\n`;
  
  analysisData.forEach((data, idx) => {
    prompt += `\n${idx + 1}. ${data.page_url}\n`;
    prompt += `   リライト種別: ${data.rewrite_type}\n`;
    prompt += `   変更内容: ${data.changes_summary}\n`;
    prompt += `   順位変化: ${data.position_change > 0 ? '+' : ''}${data.position_change}位\n`;
    prompt += `   CTR変化: ${data.ctr_change_rate > 0 ? '+' : ''}${data.ctr_change_rate}%\n`;
    prompt += `   PV変化: ${data.pv_change_rate > 0 ? '+' : ''}${data.pv_change_rate}%\n`;
  });
  
  prompt += `\n【分析してほしいこと】\n`;
  prompt += `1. どのリライト種別が最も効果的か\n`;
  prompt += `2. 成功した変更内容の共通点\n`;
  prompt += `3. 今後優先すべきリライトタイプ\n`;
  prompt += `4. 具体的な推奨アクション\n`;
  
  return prompt;
}

/**
 * テスト関数: リライト記録
 */
function testRecordRewrite() {
  const testData = {
    page_url: '/master-ipad-app-switcher',  // ← パス部分のみに変更
    rewrite_type: 'タイトル',
    changes_summary: 'タイトルにターゲットKW「マスター」を追加、より具体的な表現に変更',
    ai_suggestion: 'タイトルタグを「iPadのアプリ切り替え完全ガイド」から「iPadアプリ切り替えをマスター！完全ガイド【初心者向け】」に変更することを推奨'
  };
  
  const rewriteId = recordRewrite(testData);
  Logger.log(`テスト成功 - リライトID: ${rewriteId}`);
}

/**
 * テスト関数: 効果測定
 */
function testMeasureRewriteEffect() {
  // 実際のリライトIDを指定してテスト
  const rewriteId = 'RW_20251124_XXXXXX'; // 実際のIDに置き換え
  
  try {
    const result = measureRewriteEffect(rewriteId);
    Logger.log('テスト成功');
    Logger.log(JSON.stringify(result, null, 2));
  } catch (error) {
    Logger.log(`テストエラー: ${error.message}`);
  }
}

/**
 * テスト関数: 履歴取得
 */
function testGetRewriteHistory() {
  const history = getRewriteHistory();
  Logger.log(`全履歴件数: ${history.length}`);
  
  // 成功したもののみ
  const successOnly = getRewriteHistory({ success_only: true });
  Logger.log(`成功履歴件数: ${successOnly.length}`);
}

/**
 * テスト関数: 成功パターン分析
 */
function testAnalyzeSuccessPatterns() {
  const analysis = analyzeSuccessPatterns();
  Logger.log('=== 分析結果 ===');
  Logger.log(analysis);
}
