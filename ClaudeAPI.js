/**
 * SEOリライト支援ツール - ClaudeAPI.gs
 * バージョン: 2.0（5軸スコア対応版）
 * 更新日: 2025年12月1日
 * 
 * 変更履歴:
 * - v2.0: 5軸スコア対応システムプロンプト追加
 * - v2.0: 提案生成用専用関数追加
 * - v1.0: 基本機能（リトライ10回、段階的待機時間）
 */

// ============================================
// 定数定義
// ============================================

/**
 * SEO専用システムプロンプト（5軸スコア対応版）
 */
const SYSTEM_PROMPT_SEO_V2 = `あなたはSEOとコンテンツマーケティングの専門家です。
データに基づいた客観的な分析と、具体的で実行可能な提案を行います。

【5軸スコアリングシステムの理解】
このツールでは、以下の5軸でページを評価しています（各20%、合計100点）：

① 機会損失スコア（opportunity_score）
- 順位4-10位でCTRが低いページを特定
- 表示回数が多いがクリックされていないページ
- CTRギャップ（予想CTR vs 実際のCTR）を評価
- 高スコア = 改善余地が大きい

② パフォーマンススコア（performance_score）
- 直帰率、滞在時間を評価
- UX指標（Clarityデータ）を統合
  - スクロール深度（浅い = 問題あり）
  - デッドクリック（多い = UI問題）
  - レイジクリック（多い = ユーザーフラストレーション）
  - クイックバック（多い = コンテンツ不満足）
- 高スコア = UX改善が必要

③ ビジネスインパクトスコア（business_impact_score）
- トラフィック量を評価
- コンバージョン数・近接度を評価
- 高スコア = ビジネスへの影響が大きい

④ キーワード戦略スコア（keyword_strategy_score）
- キーワード品質（検索ボリューム、競合難易度、順位）
- ターゲットKW達成度（実クエリとの一致度）
- 高スコア = キーワード戦略の見直しが必要

⑤ 競合難易度スコア（competitor_difficulty_score）
- 自サイトDA vs 競合平均DAを比較
- 勝算度スコア（7段階判定）
  - 超狙い目/易/中/やや難/難/厳しい/激戦
- 高スコア = 勝算がある（リライト推奨）
- 低スコア = 競合が強い（別KW検討）

【総合優先度スコアの解釈】
- 80点以上: 最優先（今すぐリライト）
- 60-79点: 高優先（今週中）
- 40-59点: 中優先（今月中）
- 40点未満: 低優先（様子見）

【絶対的足切り条件】
- 順位101位以上 AND 競合レベル「激戦」→ スコア0（除外）

【あなたの役割】
1. 5軸スコアを総合的に判断し、リライト優先度を評価する
2. 各スコアの問題点を特定し、具体的な改善策を提案する
3. データに基づいた定量的な期待効果を示す
4. ユーザーの質問意図を理解し、適切なデータを参照して回答する

【回答スタイル】
- 簡潔で分かりやすい日本語
- 具体的な数値やデータを引用
- 優先順位を明確にする
- 実行可能なアクションを提示
- 期待効果を定量化（可能な場合）`;

/**
 * クエリベース提案用システムプロンプト
 */
const SYSTEM_PROMPT_QUERY_SUGGESTION = `あなたはSEOライティングの専門家です。
Google Search Consoleのクエリデータを分析し、具体的なタイトル・メタディスクリプションの改善案を提案します。

【分析の視点】
1. CTRギャップの大きいクエリ（表示されているがクリックされていない）
2. 順位は良いがCTRが低いクエリ（タイトル・説明文の問題）
3. 表示回数が多い重要クエリ
4. CVに近いクエリ（購入意図、比較検討ワード）

【提案フォーマット】
各提案は以下の形式で出力してください：

🎯 対象クエリ: [クエリ名]
📊 現状: 順位[X]位、表示[Y]回、CTR[Z]%
💡 提案:
  - タイトル案: [具体的なタイトル案]
  - メタディスクリプション案: [具体的な説明文案]
📈 期待効果: CTRが[現在]%→[目標]%に改善で、月間[N]クリック増加見込み

【注意事項】
- タイトルは30-35文字程度
- メタディスクリプションは120文字程度
- ターゲットクエリを自然に含める
- ユーザーの検索意図に応える内容`;

/**
 * UXベース提案用システムプロンプト
 */
const SYSTEM_PROMPT_UX_SUGGESTION = `あなたはUX/UIの専門家です。
Microsoft Clarityのデータを分析し、ユーザー体験の改善案を提案します。

【分析指標】
1. スクロール深度（avg_scroll_depth）
   - 30%未満: 深刻な問題（ファーストビューで離脱）
   - 30-50%: 要改善
   - 50-70%: 普通
   - 70%以上: 良好

2. デッドクリック（dead_clicks）
   - クリックしても何も起きない場所でのクリック
   - 多い = UIの問題、クリッカブル要素の誤解

3. レイジクリック（rage_clicks）
   - 短時間に同じ場所を何度もクリック
   - 多い = ユーザーフラストレーション、ページ遅延

4. クイックバック（quick_backs）
   - すぐに前のページに戻る行動
   - 多い = コンテンツが期待と不一致

【提案フォーマット】
🔍 問題点: [具体的な問題]
📊 データ: [関連する指標値]
💡 改善案:
  1. [具体的なアクション1]
  2. [具体的なアクション2]
📈 期待効果: [改善による期待効果]`;

/**
 * キーワード戦略用システムプロンプト★NEW
 */
const SYSTEM_PROMPT_KEYWORD_STRATEGY = `あなたはキーワード戦略の専門家です。
ターゲットキーワードと実際の流入クエリを分析し、キーワード戦略の最適化案を提案します。

【分析の視点】
1. ターゲットKWの妥当性
   - 検索ボリュームと競合難易度のバランス
   - 現在の順位とポテンシャル
   - ビジネスとの関連性

2. ターゲットKWと実クエリのギャップ
   - ターゲットKWで流入しているか
   - 想定外の有望クエリがあるか
   - 検索意図のズレはないか

3. キーワード最適化の方向性
   - ターゲットKWの変更が必要か
   - コンテンツの調整方針
   - サブキーワードの活用

【提案フォーマット】
📊 キーワード戦略診断:
- ターゲットKW: [KW名]
- 現在順位: [X]位
- 検索ボリューム: [Y]/月
- 競合レベル: [レベル]

🎯 ターゲットKW評価:
- 妥当性: [高/中/低]
- 理由: [具体的な理由]

⚠️ 課題:
1. [課題1]
2. [課題2]

💡 改善提案:
1. [具体的なアクション1]
2. [具体的なアクション2]

📈 期待効果:
- 順位改善: [現在]位 → [目標]位
- トラフィック増加: +[X]%`;

/**
 * 競合分析用システムプロンプト
 */
const SYSTEM_PROMPT_COMPETITOR_ANALYSIS = `あなたは競合分析の専門家です。
競合サイトのコンテンツを分析し、自社サイトとの差分を特定します。

【分析の視点】
1. コンテンツ構造の違い
   - 見出し構成（H1-H3）
   - 情報の網羅性
   - コンテンツの深さ

2. 上位サイトの共通点
   - 必ず含まれている情報（必須コンテンツ）
   - 共通の見出しパターン
   - 文字数・画像数の傾向

3. 自社に足りないコンテンツ
   - 競合にあって自社にない情報
   - 追加すべきセクション

【提案フォーマット】
📊 競合分析サマリー:
- 分析対象: [競合サイト数]サイト
- 平均文字数: [X]文字（自社: [Y]文字）
- 平均見出し数: [X]個（自社: [Y]個）

🏆 上位サイトの共通コンテンツ:
1. [共通セクション1]
2. [共通セクション2]

⚠️ 自社に不足しているコンテンツ:
1. [不足コンテンツ1]
2. [不足コンテンツ2]

💡 改善提案:
1. [具体的なアクション]`;

/**
 * 統合リライトレポート用システムプロンプト
 */
const SYSTEM_PROMPT_INTEGRATED_REPORT = `あなたはSEOコンサルタントです。
複数のデータソースを統合し、優先度付きのリライト提案レポートを作成します。

【入力データ】
- 5軸スコア（機会損失、パフォーマンス、ビジネスインパクト、KW戦略、競合難易度）
- GSCクエリデータ（順位、表示回数、CTR）
- Clarityデータ（UX指標）
- 競合分析データ（DA、勝算度）

【レポート構成】
1. エグゼクティブサマリー（3行以内）
2. 5軸スコア分析
3. 優先度付き改善提案（最大5つ）
4. 期待効果の定量化
5. 次のアクション

【提案の優先度基準】
- 高: 即効性が高く、工数が低い
- 中: 効果は高いが工数もかかる
- 低: 長期的な改善

【出力フォーマット】
## 📋 リライト提案レポート: [ページタイトル]

### 📊 5軸スコア分析
| 軸 | スコア | 判定 | コメント |
|---|---|---|---|
| 機会損失 | XX点 | ⚠️ | ... |

### 🎯 改善提案（優先度順）
#### 1. [提案タイトル]（優先度: 高）
- 内容: ...
- 期待効果: ...
- 工数目安: ...

### 📈 期待効果まとめ
- CTR改善: +X%
- 順位改善: +X位
- 月間クリック増: +X

### ✅ 次のアクション
1. ...`;


// ============================================
// メイン関数
// ============================================

/**
 * Claude APIを呼び出す（リトライ機能付き）
 * @param {string} userMessage - ユーザーメッセージ
 * @param {string} systemPrompt - システムプロンプト（省略時はデフォルト使用）
 * @returns {string} AIの応答テキスト
 */
function callClaudeAPI(userMessage, systemPrompt) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');
  const url = 'https://api.anthropic.com/v1/messages';
  
  if (!apiKey) {
    throw new Error('Claude API Keyが設定されていません');
  }
  
  if (!userMessage || userMessage.trim() === '') {
    userMessage = 'こんにちは';
  }
  
  // システムプロンプトが指定されていない場合はSEO専用プロンプトを使用
  const effectiveSystemPrompt = systemPrompt || SYSTEM_PROMPT_SEO_V2;
  
  const requestBody = {
    model: 'claude-sonnet-4-20250514',
    max_tokens: 4096, // 長い提案レポートに対応
    messages: [
      {
        role: 'user',
        content: userMessage
      }
    ],
    system: effectiveSystemPrompt
  };
  
  const options = {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01'
    },
    payload: JSON.stringify(requestBody),
    muteHttpExceptions: true
  };
  
  const maxRetries = 10;
  let lastError = null;
  
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      Logger.log(`[${attempt}/${maxRetries}] Claude API呼び出し試行中...`);
      
      const response = UrlFetchApp.fetch(url, options);
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();
      
      if (responseCode === 200) {
        const data = JSON.parse(responseText);
        
        if (data.usage) {
          logTokenUsage(data.usage);
        }
        
        if (data.content && data.content[0] && data.content[0].text) {
          Logger.log(`✅ 成功！（試行${attempt}回目で成功）`);
          return data.content[0].text;
        } else {
          throw new Error('応答にテキストが含まれていません');
        }
        
      } else if (responseCode === 529) {
        Logger.log(`❌ 529エラー（試行${attempt}/${maxRetries}）`);
        lastError = new Error('Claude APIが過負荷状態です');
        
        if (attempt < maxRetries) {
          let waitTime;
          if (attempt <= 2) {
            waitTime = 5000;
          } else if (attempt <= 4) {
            waitTime = 10000;
          } else if (attempt <= 6) {
            waitTime = 20000;
          } else if (attempt <= 8) {
            waitTime = 40000;
          } else {
            waitTime = 60000;
          }
          
          Logger.log(`⏳ ${waitTime / 1000}秒待機してリトライします...`);
          Utilities.sleep(waitTime);
        }
        
      } else if (responseCode === 429) {
        Logger.log(`❌ 429エラー: レート制限超過`);
        lastError = new Error('API使用制限に達しました');
        
        if (attempt < maxRetries) {
          const waitTime = 60000;
          Logger.log(`⏳ ${waitTime / 1000}秒待機してリトライします...`);
          Utilities.sleep(waitTime);
        }
        
      } else {
        Logger.log(`❌ APIエラー: ${responseCode}`);
        Logger.log('エラー内容: ' + responseText);
        throw new Error(`Claude API Error: ${responseCode}`);
      }
      
    } catch (error) {
      Logger.log(`❌ 試行${attempt}でエラー: ${error.message}`);
      lastError = error;
      
      if (attempt < maxRetries && 
          (error.message.includes('529') || 
           error.message.includes('過負荷') || 
           error.message.includes('429'))) {
        
        let waitTime = 10000;
        if (attempt > 5) {
          waitTime = 30000;
        }
        
        Logger.log(`⏳ ${waitTime / 1000}秒待機してリトライします...`);
        Utilities.sleep(waitTime);
      } else if (attempt >= maxRetries) {
        break;
      } else {
        throw error;
      }
    }
  }
  
  Logger.log(`❌ ${maxRetries}回の試行全てが失敗しました`);
  
  if (lastError) {
    throw new Error(`Claude APIが現在利用できません。しばらく時間を置いてから再度お試しください。（${maxRetries}回試行）`);
  } else {
    throw new Error('Claude API呼び出しに失敗しました');
  }
}


// ============================================
// 専用プロンプト関数
// ============================================

/**
 * クエリベース提案を生成
 * @param {string} pageUrl - 対象ページURL
 * @param {Array} queryData - クエリデータ配列
 * @returns {string} 提案テキスト
 */
function generateQuerySuggestion(pageUrl, queryData) {
  const prompt = `以下のページのGSCクエリデータを分析し、タイトル・メタディスクリプションの改善案を提案してください。

【対象ページ】
${pageUrl}

【クエリデータ】
${JSON.stringify(queryData, null, 2)}

上位5つの重要クエリについて、具体的な改善案を提案してください。`;

  return callClaudeAPI(prompt, SYSTEM_PROMPT_QUERY_SUGGESTION);
}

/**
 * UXベース提案を生成
 * @param {string} pageUrl - 対象ページURL
 * @param {Object} uxData - UX指標データ
 * @returns {string} 提案テキスト
 */
function generateUXSuggestion(pageUrl, uxData) {
  const prompt = `以下のページのClarityデータを分析し、UX改善案を提案してください。

【対象ページ】
${pageUrl}

【UX指標】
- スクロール深度: ${uxData.avg_scroll_depth || 'N/A'}%
- デッドクリック: ${uxData.dead_clicks || 0}回
- レイジクリック: ${uxData.rage_clicks || 0}回
- クイックバック: ${uxData.quick_backs || 0}回
- セッション数: ${uxData.sessions || 0}

問題点を特定し、具体的な改善案を提案してください。`;

  return callClaudeAPI(prompt, SYSTEM_PROMPT_UX_SUGGESTION);
}

/**
 * 競合分析提案を生成
 * @param {string} keyword - 対象キーワード
 * @param {Object} ownContent - 自社コンテンツデータ
 * @param {Array} competitorContents - 競合コンテンツデータ配列
 * @returns {string} 提案テキスト
 */
function generateCompetitorSuggestion(keyword, ownContent, competitorContents) {
  const prompt = `以下のキーワードについて、競合サイトと自社サイトを比較分析してください。

【対象キーワード】
${keyword}

【自社サイト】
- URL: ${ownContent.url}
- タイトル: ${ownContent.title}
- 文字数: ${ownContent.wordCount}文字
- 見出し数: H1=${ownContent.h1Count}, H2=${ownContent.h2Count}, H3=${ownContent.h3Count}
- 見出し一覧: ${JSON.stringify(ownContent.headings)}

【競合サイト（上位${competitorContents.length}サイト）】
${competitorContents.map((c, i) => `
${i + 1}位: ${c.url}
- タイトル: ${c.title}
- 文字数: ${c.wordCount}文字
- 見出し数: H1=${c.h1Count}, H2=${c.h2Count}, H3=${c.h3Count}
- 見出し一覧: ${JSON.stringify(c.headings)}
`).join('\n')}

上位サイトの共通点と自社に不足しているコンテンツを特定し、具体的な改善案を提案してください。`;

  return callClaudeAPI(prompt, SYSTEM_PROMPT_COMPETITOR_ANALYSIS);
}

/**
 * 統合リライトレポートを生成
 * @param {Object} pageData - ページの全データ
 * @returns {string} レポートテキスト
 */
function generateIntegratedReport(pageData) {
  const prompt = `以下のページデータを分析し、優先度付きのリライト提案レポートを作成してください。

【ページ情報】
- URL: ${pageData.url}
- タイトル: ${pageData.title}

【5軸スコア】
- 機会損失スコア: ${pageData.opportunity_score}点
- パフォーマンススコア: ${pageData.performance_score}点
- ビジネスインパクトスコア: ${pageData.business_impact_score}点
- キーワード戦略スコア: ${pageData.keyword_strategy_score}点
- 競合難易度スコア: ${pageData.competitor_difficulty_score}点
- 総合優先度スコア: ${pageData.total_priority_score}点

【GSCデータ】
- 平均順位: ${pageData.avg_position}位
- 表示回数: ${pageData.impressions}回
- クリック数: ${pageData.clicks}回
- CTR: ${pageData.ctr}%
- 上位クエリ: ${pageData.top_queries}

【UXデータ】
- スクロール深度: ${pageData.avg_scroll_depth || 'N/A'}%
- デッドクリック: ${pageData.dead_clicks || 0}回
- レイジクリック: ${pageData.rage_clicks || 0}回

【競合データ】
- 競合レベル: ${pageData.competition_level}
- 勝算度スコア: ${pageData.winnable_score}点

総合的に分析し、優先度付きの改善提案レポートを作成してください。`;

  return callClaudeAPI(prompt, SYSTEM_PROMPT_INTEGRATED_REPORT);
}


// ============================================
// ユーティリティ関数
// ============================================

/**
 * トークン使用量をログに記録
 * @param {Object} usage - トークン使用量オブジェクト
 */
function logTokenUsage(usage) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('API使用ログ');
    
    if (!sheet) {
      sheet = ss.insertSheet('API使用ログ');
      sheet.appendRow(['日時', '入力トークン', '出力トークン', '合計トークン', 'コスト($)']);
      
      const headerRange = sheet.getRange(1, 1, 1, 5);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#4285f4');
      headerRange.setFontColor('#ffffff');
    }
    
    const inputTokens = usage.input_tokens || 0;
    const outputTokens = usage.output_tokens || 0;
    const totalTokens = inputTokens + outputTokens;
    const cost = (inputTokens / 1000000 * 3) + (outputTokens / 1000000 * 15);
    
    sheet.appendRow([
      new Date(),
      inputTokens,
      outputTokens,
      totalTokens,
      cost.toFixed(6)
    ]);
    
    Logger.log(`📊 トークン使用: 入力=${inputTokens}, 出力=${outputTokens}, コスト=$${cost.toFixed(6)}`);
    
  } catch (error) {
    Logger.log('トークン使用量の記録エラー: ' + error.message);
  }
}

/**
 * システムプロンプトを取得
 * @param {string} type - プロンプトタイプ（'seo', 'query', 'ux', 'competitor', 'integrated'）
 * @returns {string} システムプロンプト
 */
function getSystemPrompt(type) {
  switch (type) {
    case 'seo':
      return SYSTEM_PROMPT_SEO_V2;
    case 'query':
      return SYSTEM_PROMPT_QUERY_SUGGESTION;
    case 'ux':
      return SYSTEM_PROMPT_UX_SUGGESTION;
    case 'competitor':
      return SYSTEM_PROMPT_COMPETITOR_ANALYSIS;
    case 'keyword':
      return SYSTEM_PROMPT_KEYWORD_STRATEGY;
    case 'integrated':
      return SYSTEM_PROMPT_INTEGRATED_REPORT;
    default:
      return SYSTEM_PROMPT_SEO_V2;
  }
}


// ============================================
// テスト関数
// ============================================

/**
 * Claude APIの基本テスト
 */
function testClaudeAPI() {
  const testMessage = "こんにちは！SEOについて簡単に説明してください（2行で）。";
  
  try {
    Logger.log('=== Claude APIテスト開始 ===');
    const response = callClaudeAPI(testMessage, null);
    Logger.log('Claude応答:\n' + response);
    Logger.log('=== テスト成功！ ===');
    
    return 'テスト成功！Claude APIが正常に動作しています。';
    
  } catch (error) {
    Logger.log('=== テスト失敗 ===');
    Logger.log('エラー: ' + error.message);
    throw error;
  }
}

/**
 * 5軸スコア対応のテスト
 */
function testFiveAxisPrompt() {
  const testMessage = `以下のページの5軸スコアを分析してください。

【ページ情報】
URL: /insurance/recommend/
タイトル: スマホ保険おすすめ10選

【5軸スコア】
- 機会損失スコア: 75点
- パフォーマンススコア: 45点
- ビジネスインパクトスコア: 80点
- キーワード戦略スコア: 60点
- 競合難易度スコア: 70点
- 総合優先度スコア: 66点

このページのリライト優先度と改善ポイントを教えてください。`;

  try {
    Logger.log('=== 5軸スコア対応テスト開始 ===');
    const response = callClaudeAPI(testMessage, SYSTEM_PROMPT_SEO_V2);
    Logger.log('Claude応答:\n' + response);
    Logger.log('=== テスト成功！ ===');
    
    return response;
    
  } catch (error) {
    Logger.log('=== テスト失敗 ===');
    Logger.log('エラー: ' + error.message);
    throw error;
  }
}

/**
 * 統合レポート生成のテスト
 */
function testIntegratedReport() {
  const testPageData = {
    url: '/insurance/recommend/',
    title: 'スマホ保険おすすめ10選',
    opportunity_score: 75,
    performance_score: 45,
    business_impact_score: 80,
    keyword_strategy_score: 60,
    competitor_difficulty_score: 70,
    total_priority_score: 66,
    avg_position: 8.5,
    impressions: 15000,
    clicks: 450,
    ctr: 3.0,
    top_queries: 'スマホ保険 おすすめ, スマホ保険 比較, iphone 保険',
    avg_scroll_depth: 45,
    dead_clicks: 12,
    rage_clicks: 3,
    competition_level: '中',
    winnable_score: 70
  };

  try {
    Logger.log('=== 統合レポート生成テスト開始 ===');
    const response = generateIntegratedReport(testPageData);
    Logger.log('Claude応答:\n' + response);
    Logger.log('=== テスト成功！ ===');
    
    return response;
    
  } catch (error) {
    Logger.log('=== テスト失敗 ===');
    Logger.log('エラー: ' + error.message);
    throw error;
  }
}