/**
 * EventAnalysis.gs 拡張版
 * 
 * 変更内容:
 * - generateEventDescriptions()関数を追加
 * - analyzeGA4Events()でdescription自動生成を実行
 * - GA4単体でも「何を計測しているか」が明確になる
 * 
 * バージョン: 1.1（Day 7-8.5拡張版）
 * 最終更新: 2025/11/26
 */

/**
 * メイン実行関数: GA4イベント分析（拡張版）
 * 
 * 実行手順:
 * 1. GA4からイベントデータ取得
 * 2. Claude APIでカテゴリ自動分類
 * 3. ★NEW: Claude APIで各イベントの説明を自動生成
 * 4. イベント分析シートに書き込み
 * 5. 条件付き書式設定
 */
function analyzeGA4Events() {
  const startTime = new Date();
  Logger.log('========================================');
  Logger.log('GA4イベント分析開始（拡張版）');
  Logger.log('========================================');
  
  try {
    // Step 1: GA4からイベントデータ取得
    Logger.log('Step 1: GA4からイベントデータ取得中...');
    const events = fetchGA4Events_();
    Logger.log(`✓ ${events.length}件のイベントを取得しました`);
    
    if (events.length === 0) {
      throw new Error('GA4からイベントデータを取得できませんでした');
    }
    
    // Step 2: Claude APIでカテゴリ自動分類
    Logger.log('Step 2: Claude APIでイベントを分類中...');
    const categorizedEvents = categorizeEvents_(events);
    Logger.log(`✓ ${categorizedEvents.length}件のイベントを分類しました`);
    
    // Step 3: ★NEW Claude APIで各イベントの説明を自動生成
    Logger.log('Step 3: Claude APIで各イベントの説明を生成中...');
    const eventsWithDescription = generateEventDescriptions_(categorizedEvents);
    Logger.log(`✓ ${eventsWithDescription.length}件のイベント説明を生成しました`);
    
    // Step 4: イベント分析シートに書き込み
    Logger.log('Step 4: イベント分析シートに書き込み中...');
    writeEventAnalysisSheet_(eventsWithDescription);
    Logger.log('✓ イベント分析シートに書き込み完了');
    
    // Step 5: 条件付き書式設定
    Logger.log('Step 5: 条件付き書式を設定中...');
    applyConditionalFormatting_();
    Logger.log('✓ 条件付き書式設定完了');
    
    const endTime = new Date();
    const duration = (endTime - startTime) / 1000;
    
    Logger.log('========================================');
    Logger.log(`GA4イベント分析完了（実行時間: ${duration.toFixed(3)}秒）`);
    Logger.log('========================================');
    
    return {
      success: true,
      eventCount: eventsWithDescription.length,
      duration: duration
    };
    
  } catch (error) {
    Logger.log(`❌ エラー: ${error.message}`);
    Logger.log(`スタックトレース: ${error.stack}`);
    throw error;
  }
}

/**
 * GA4からイベントデータを取得
 * @private
 */
function fetchGA4Events_() {
  const propertyId = PropertiesService.getScriptProperties().getProperty('GA4_PROPERTY_ID');
  
  if (!propertyId) {
    throw new Error('GA4_PROPERTY_IDが設定されていません');
  }
  
  try {
    const request = {
      dateRanges: [{
        startDate: '30daysAgo',
        endDate: 'yesterday'
      }],
      dimensions: [
        { name: 'eventName' }
      ],
      metrics: [
        { name: 'eventCount' },
        { name: 'conversions' }
      ],
      orderBys: [{
        metric: { metricName: 'eventCount' },
        desc: true
      }]
    };
    
    const response = AnalyticsData.Properties.runReport(request, propertyId);
    
    if (!response.rows || response.rows.length === 0) {
      Logger.log('⚠️ GA4からイベントデータが取得できませんでした');
      return [];
    }
    
    const events = response.rows.map(row => ({
      event_name: row.dimensionValues[0].value,
      event_count: parseInt(row.metricValues[0].value) || 0,
      conversions: parseInt(row.metricValues[1].value) || 0
    }));
    
    Logger.log(`GA4から${events.length}件のイベントを取得`);
    return events;
    
  } catch (error) {
    Logger.log(`GA4データ取得エラー: ${error.message}`);
    throw error;
  }
}

/**
 * Claude APIでイベントを分類
 * @private
 */
function categorizeEvents_(events) {
  return events.map(event => {
    Logger.log(`分類中: ${event.event_name} (${event.event_count}回)`);
    
    const result = categorizeEventWithClaude_(event.event_name);
    
    return {
      ...event,
      event_category: result.category,
      cv_contribution: result.cv_contribution,
      importance: result.importance
    };
  });
}

/**
 * ★NEW: Claude APIで各イベントの説明を自動生成
 * @private
 */
function generateEventDescriptions_(events) {
  Logger.log('各イベントの説明を生成します...');
  
  return events.map((event, index) => {
    Logger.log(`説明生成中 (${index + 1}/${events.length}): ${event.event_name}`);
    
    // Claude APIで説明生成
    const description = generateSingleEventDescription_(event);
    
    return {
      ...event,
      description: description
    };
  });
}

/**
 * ★NEW: 単一イベントの説明を生成
 * @private
 */
function generateSingleEventDescription_(event) {
  const prompt = `以下のGA4イベントは何を計測していますか？
初心者にも分かるように、1-2文で簡潔に説明してください。
専門用語は避け、「〇〇を計測しています」という形式で答えてください。

イベント名: ${event.event_name}
発火回数: ${event.event_count}回（過去30日）
カテゴリ: ${event.event_category}
CV貢献度: ${event.cv_contribution}
重要度: ${event.importance}

【回答例】
- page_view → ページの閲覧回数を計測しています。サイト訪問者がどのページをどれだけ見たかを分析できます。
- click → ボタンやリンクのクリックを計測しています。ユーザーがどこをクリックしているか追跡できます。
- form_submit → フォーム送信を計測しています。お問い合わせや会員登録などの完了数を分析できます。`;
  
  try {
    const description = callClaudeAPIWithRetry_(prompt, 500);
    Logger.log(`✓ 説明生成完了: ${description.substring(0, 50)}...`);
    return description;
    
  } catch (error) {
    Logger.log(`⚠️ 説明生成エラー: ${error.message}`);
    // エラー時はデフォルトメッセージ
    return `${event.event_name}イベントを計測しています。詳細な説明は後で追加してください。`;
  }
}

/**
 * Claude APIで単一イベントをカテゴリ分類
 * @private
 */
function categorizeEventWithClaude_(eventName) {
  const prompt = `以下のGA4イベント名を分析し、カテゴリ、CV貢献度、重要度を判定してください。

イベント名: ${eventName}

以下のJSON形式で回答してください:
{
  "category": "CV" or "エンゲージメント" or "その他",
  "cv_contribution": 0-100の数値,
  "importance": "高" or "中" or "低"
}

【判定基準】
- CV: purchase, form_submit, sign_up, add_to_cart など、コンバージョンに直結
- エンゲージメント: scroll, click, video_play など、ユーザー行動の計測
- その他: page_view, session_start など、基本的な計測

CV貢献度:
- 90-100: 直接的な収益（purchase, affiliate_click等）
- 60-80: 重要なアクション（form_submit, sign_up等）
- 30-50: エンゲージメント（click, scroll等）
- 0-20: 基本計測（page_view等）

重要度:
- 高: ビジネスに直結、最優先で分析すべき
- 中: ユーザー行動理解に重要
- 低: 参考程度

【重要】JSONのみを出力してください。説明文は不要です。`;

  const response = callClaudeAPIWithRetry_(prompt, 300);
  
  try {
    // JSONをパース
    const jsonMatch = response.match(/\{[\s\S]*\}/);
    if (!jsonMatch) {
      throw new Error('JSON形式の応答が見つかりません');
    }
    
    const result = JSON.parse(jsonMatch[0]);
    
    // バリデーション
    if (!result.category || !result.cv_contribution || !result.importance) {
      throw new Error('必須フィールドが不足しています');
    }
    
    Logger.log(`✓ ${eventName} → ${result.category}, CV:${result.cv_contribution}, 重要度:${result.importance}`);
    return result;
    
  } catch (error) {
    Logger.log(`⚠️ JSON解析エラー: ${error.message}`);
    Logger.log(`レスポンス: ${response}`);
    
    // エラー時はデフォルト値
    return {
      category: 'その他',
      cv_contribution: 10,
      importance: '低'
    };
  }
}

/**
 * イベント分析シートに書き込み
 * @private
 */
function writeEventAnalysisSheet_(events) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('イベント分析');
  
  if (!sheet) {
    throw new Error('イベント分析シートが見つかりません');
  }
  
  // 既存データをクリア（ヘッダー行は残す）
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 9).clearContent();
  }
  
  // データを2次元配列に変換
  const data = events.map(event => [
    Utilities.getUuid(),                    // A: event_id
    event.event_name,                       // B: event_name
    event.event_category,                   // C: event_category
    event.event_count,                      // D: event_count
    event.cv_contribution,                  // E: cv_contribution
    event.importance,                       // F: importance
    event.description,                      // G: description ★NEW
    true,                                   // H: enabled
    new Date()                              // I: last_updated
  ]);
  
  // 一括書き込み
  if (data.length > 0) {
    sheet.getRange(2, 1, data.length, 9).setValues(data);
    Logger.log(`${data.length}件のイベントデータを書き込みました`);
  }
}

/**
 * 条件付き書式を設定
 * @private
 */
function applyConditionalFormatting_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('イベント分析');
  
  if (!sheet) {
    return;
  }
  
  // 既存の条件付き書式をクリア
  sheet.clearConditionalFormatRules();
  
  const rules = [];
  
  // ルール1: カテゴリ列（C列）- CV
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('CV')
    .setBackground('#d4edda')
    .setFontColor('#155724')
    .setRanges([sheet.getRange('C2:C1000')])
    .build());
  
  // ルール2: カテゴリ列（C列）- エンゲージメント
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('エンゲージメント')
    .setBackground('#fff3cd')
    .setFontColor('#856404')
    .setRanges([sheet.getRange('C2:C1000')])
    .build());
  
  // ルール3: カテゴリ列（C列）- その他
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('その他')
    .setBackground('#f8f9fa')
    .setFontColor('#6c757d')
    .setRanges([sheet.getRange('C2:C1000')])
    .build());
  
  // ルール4: CV貢献度列（E列）- 高（80以上）
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(80)
    .setBackground('#d4edda')
    .setFontColor('#155724')
    .setBold(true)
    .setRanges([sheet.getRange('E2:E1000')])
    .build());
  
  // ルール5: CV貢献度列（E列）- 中（50-79）
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(50, 79)
    .setBackground('#fff3cd')
    .setFontColor('#856404')
    .setRanges([sheet.getRange('E2:E1000')])
    .build());
  
  // ルール6: CV貢献度列（E列）- 低（49以下）
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(50)
    .setBackground('#f8f9fa')
    .setFontColor('#6c757d')
    .setRanges([sheet.getRange('E2:E1000')])
    .build());
  
  // ルール7: 重要度列（F列）- 高
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('高')
    .setBackground('#d4edda')
    .setFontColor('#155724')
    .setBold(true)
    .setRanges([sheet.getRange('F2:F1000')])
    .build());
  
  // ルール8: 重要度列（F列）- 中
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('中')
    .setBackground('#fff3cd')
    .setFontColor('#856404')
    .setRanges([sheet.getRange('F2:F1000')])
    .build());
  
  // ルール9: 重要度列（F列）- 低
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('低')
    .setBackground('#f8f9fa')
    .setFontColor('#6c757d')
    .setRanges([sheet.getRange('F2:F1000')])
    .build());
  
  sheet.setConditionalFormatRules(rules);
  Logger.log('条件付き書式を設定しました（9ルール）');
}

/**
 * Claude API呼び出し（リトライ機能付き）
 * @private
 */
function callClaudeAPIWithRetry_(message, maxTokens = 1000) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');
  
  if (!apiKey) {
    throw new Error('CLAUDE_API_KEYが設定されていません');
  }
  
  const url = 'https://api.anthropic.com/v1/messages';
  const maxRetries = 10;
  let attempt = 0;
  
  while (attempt < maxRetries) {
    try {
      const payload = {
        model: 'claude-sonnet-4-20250514',
        max_tokens: maxTokens,
        messages: [{
          role: 'user',
          content: message
        }]
      };
      
      const options = {
        method: 'post',
        headers: {
          'Content-Type': 'application/json',
          'x-api-key': apiKey,
          'anthropic-version': '2023-06-01'
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };
      
      const response = UrlFetchApp.fetch(url, options);
      const responseCode = response.getResponseCode();
      
      if (responseCode === 200) {
        const data = JSON.parse(response.getContentText());
        
        // トークン使用量をログ記録
        if (data.usage) {
          logTokenUsage_(data.usage, 'callClaudeAPIWithRetry_');
        }
        
        return data.content[0].text;
      } else {
        throw new Error(`API Error: ${responseCode} - ${response.getContentText()}`);
      }
      
    } catch (error) {
      attempt++;
      if (attempt >= maxRetries) {
        throw new Error(`Claude API呼び出し失敗（${maxRetries}回試行）: ${error.message}`);
      }
      
      const waitTime = Math.pow(2, attempt) * 1000;
      Logger.log(`リトライ ${attempt}/${maxRetries} (待機: ${waitTime}ms)`);
      Utilities.sleep(waitTime);
    }
  }
}

/**
 * トークン使用量をログ記録
 * @private
 */
function logTokenUsage_(usage, functionName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('API使用ログ');
  
  if (!sheet) {
    sheet = ss.insertSheet('API使用ログ');
    sheet.appendRow(['日時', 'APIの名前', '関数名', '入力トークン', '出力トークン', '合計トークン', 'API呼び出し回数', 'コスト($)']);
  }
  
  const inputTokens = usage.input_tokens || 0;
  const outputTokens = usage.output_tokens || 0;
  const totalTokens = inputTokens + outputTokens;
  
  // コスト計算（claude-sonnet-4-20250514）
  const inputCost = (inputTokens / 1000000) * 3;
  const outputCost = (outputTokens / 1000000) * 15;
  const totalCost = inputCost + outputCost;
  
  sheet.appendRow([
    new Date(),
    'Claude',
    functionName,
    inputTokens,
    outputTokens,
    totalTokens,
    1,
    totalCost.toFixed(6)
  ]);
}

/**
 * イベント分析シート作成（初回のみ）
 */
function createEventAnalysisSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('イベント分析');
  
  if (sheet) {
    Logger.log('イベント分析シートは既に存在します');
    return;
  }
  
  // シート作成
  sheet = ss.insertSheet('イベント分析');
  
  // ヘッダー行作成
  const headers = [
    'event_id',
    'event_name',
    'event_category',
    'event_count',
    'cv_contribution',
    'importance',
    'description',        // ★NEW: 説明列
    'enabled',
    'last_updated'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // ヘッダー行の書式設定
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  
  // 列幅設定
  sheet.setColumnWidth(1, 150);  // event_id
  sheet.setColumnWidth(2, 200);  // event_name
  sheet.setColumnWidth(3, 150);  // event_category
  sheet.setColumnWidth(4, 100);  // event_count
  sheet.setColumnWidth(5, 120);  // cv_contribution
  sheet.setColumnWidth(6, 100);  // importance
  sheet.setColumnWidth(7, 400);  // description ★NEW: 幅広く
  sheet.setColumnWidth(8, 80);   // enabled
  sheet.setColumnWidth(9, 150);  // last_updated
  
  // データ検証（enabled列）
  const enabledValidation = SpreadsheetApp.newDataValidation()
    .requireCheckbox()
    .build();
  sheet.getRange('H2:H1000').setDataValidation(enabledValidation);
  
  // シートを固定
  sheet.setFrozenRows(1);
  
  Logger.log('イベント分析シートを作成しました');
}