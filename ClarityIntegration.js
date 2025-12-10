/**
 * ClarityIntegration.gs
 * Microsoft Clarity Data Export API連携とUXスコアリング
 * 
 * 機能:
 * - Clarity Data Export APIからUX指標を自動取得
 * - GTMスクロールイベントをGA4から取得
 * - Clarityデータを統合データシートにマージ
 * - UX総合スコアを計算（0-100点）
 * - パフォーマンススコアにUX要素を統合
 * - 週次自動実行トリガー設定
 * 
 * 作成日: 2025/11/26
 * バージョン: 1.0
 */

// ===================================
// 1. 初期設定
// ===================================

/**
 * Clarity APIトークンを設定
 * 手動で一度だけ実行してください
 * 
 * 手順:
 * 1. Clarityにログイン
 * 2. プロジェクト ll3jrfi0ba を開く
 * 3. Settings → Data Export → Generate new API token
 * 4. トークン名: "SEO-Tool-API" を入力
 * 5. 生成されたトークンをコピー
 * 6. この関数の token 変数に貼り付け
 * 7. この関数を実行
 * 
 * GTMスクロールイベント設定について:
 * - 推奨: autoDetectGTMSettings() 関数を実行して自動検出（簡単・確実）
 * - 手動設定も可能（以下のコメントを外す）
 * 
 * 自動検出のメリット:
 * - イベント名・パラメータ名・パターンを自動で検出
 * - GTM設定を確認する手間が不要
 * - Day 7-8のイベント分析シートを活用
 */
function setupClarityAPI() {
  const token = "eyJhbGciOiJSUzI1NiIsImtpZCI6IjQ4M0FCMDhFNUYwRDMxNjdEOTRFMTQ3M0FEQTk2RTcyRDkwRUYwRkYiLCJ0eXAiOiJKV1QifQ.eyJqdGkiOiI1OWVmNTNmMC02ZWZiLTQ2MGEtODFjNy1lYTI3YjNjNjNmYTAiLCJzdWIiOiIyMTkyMjgwNTA4MDcxMTkwIiwic2NvcGUiOiJEYXRhLkV4cG9ydCIsIm5iZiI6MTc2NDIwOTE5OSwiZXhwIjo0OTE3ODA5MTk5LCJpYXQiOjE3NjQyMDkxOTksImlzcyI6ImNsYXJpdHkiLCJhdWQiOiJjbGFyaXR5LmRhdGEtZXhwb3J0ZXIifQ.TSp9f0CI2FvhTsgMhksjXa9IJrzEPPdWQrNFY_Sp21YtvgqpZh2qa256IXAsnsGyWAFdxlcA5vUrlivQMUpIUqVfgIdHGHfbCJ1Z7hog-IKA0jje9qKmC7BrLIktyb3km_VdnMpf6Y7BsXDErROUHx-22zOZ_UQh7hXHr8m8r140yeDhILSKFbgADTrRF-Oxdhyh4Xwkz7rExodJYymqOzfXQj0mEQfZEq_nEclbEeU-Lxn9NAplDQegmYHVyEYRp6WjfSKUED3TwiRon-9xipK9UXqL9_IIZzxYSJDTfH7PngdG7z5Ouv0tLRq6AE_SuwnlehYe1GaRSLnOuqBHbA"; // ← ここにトークンを貼り付け
  const projectId = "ll3jrfi0ba";
  
  // GTMスクロールイベント名（手動設定する場合のみ）
  // 推奨: この設定は不要です。autoDetectGTMSettings()を実行してください。
  // 手動で指定する場合は、以下のコメントを外して設定してください
  // const gtmScrollEventName = "scrolled";  // ← GTMで設定したイベント名
  
  PropertiesService.getScriptProperties().setProperty('CLARITY_API_TOKEN', token);
  PropertiesService.getScriptProperties().setProperty('CLARITY_PROJECT_ID', projectId);
  
  // GTMスクロールイベント名を手動設定する場合（非推奨）
  // if (typeof gtmScrollEventName !== 'undefined') {
  //   PropertiesService.getScriptProperties().setProperty('GTM_SCROLL_EVENT_NAME', gtmScrollEventName);
  //   Logger.log('GTMスクロールイベント名: ' + gtmScrollEventName + '（手動設定）');
  // } else {
  //   Logger.log('GTMスクロールイベント名: 未設定');
  // }
  
  Logger.log('✅ Clarity APIトークンを設定しました');
  Logger.log('プロジェクトID: ' + projectId);
  Logger.log('\n次のステップ:');
  Logger.log('  autoDetectGTMSettings() 関数を実行して、GTM設定を自動検出してください');
}

/**
 * Clarity_RAWシートを作成
 * 手動で一度だけ実行してください
 */
function createClarityRawSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 既存シート確認
  let sheet = ss.getSheetByName('Clarity_RAW');
  if (sheet) {
    Logger.log('Clarity_RAWシートは既に存在します');
    return;
  }
  
  // 新規シート作成
  sheet = ss.insertSheet('Clarity_RAW');
  
  // ヘッダー設定
  const headers = [
    'fetch_date',        // データ取得日
    'page_url',          // ページURL
    'sessions',          // セッション数
    'avg_scroll_depth',  // 平均スクロール深度（%）
    'dead_clicks',       // デッドクリック数
    'rage_clicks',       // レイジクリック数
    'quick_backs',       // クイックバック数
    'engagement_time',   // エンゲージメント時間（秒）
    'script_errors'      // JSエラー数
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // ヘッダー行の書式設定
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('#ffffff');
  
  // 列幅調整
  sheet.setColumnWidth(1, 120);  // fetch_date
  sheet.setColumnWidth(2, 300);  // page_url
  sheet.setColumnWidths(3, 7, 100); // その他の列
  
  // シート保護（データ保護のため、編集は関数経由のみ）
  const protection = sheet.protect().setDescription('Clarity_RAW: データ保護');
  protection.setWarningOnly(true);
  
  Logger.log('✅ Clarity_RAWシートを作成しました');
}

// ===================================
// 1.5. GTM設定自動検出（Phase 1）
// ===================================

/**
 * GTM設定を自動検出して保存
 * 
 * 機能:
 * - イベント分析シートからスクロールイベントを検出
 * - イベント名、パラメータ名、パラメータ値のパターンを解析
 * - PropertiesServiceに自動保存
 * 
 * 使い方:
 * 1. Day 7-8でイベント分析シートを作成済みであること
 * 2. この関数を実行
 * 3. setupClarityAPI()は不要になる（トークン設定のみ）
 * 
 * @return {Object} 検出結果
 */
function autoDetectGTMSettings() {
  Logger.log('=== GTM設定自動検出開始 ===');
  
  const result = {
    success: false,
    eventName: null,
    parameterName: null,
    pattern: null,
    message: ''
  };
  
  // ステップ1: イベント分析シートを確認
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const eventSheet = ss.getSheetByName('イベント分析');
  
  if (!eventSheet) {
    result.message = '❌ エラー: イベント分析シートが見つかりません';
    Logger.log(result.message);
    Logger.log('Day 7-8のイベント分析機能を先に実装してください');
    return result;
  }
  
  // ステップ2: スクロールイベントを検出
  Logger.log('ステップ1: スクロールイベントを検出中...');
  const scrollEvent = detectScrollEventFromSheet(eventSheet);
  
  if (!scrollEvent) {
    result.message = '❌ スクロールイベントが見つかりませんでした';
    Logger.log(result.message);
    Logger.log('GTMでスクロールイベントが設定されているか確認してください');
    return result;
  }
  
  Logger.log('✅ スクロールイベント検出: ' + scrollEvent.eventName);
  result.eventName = scrollEvent.eventName;
  
  // ステップ3: パラメータを検出
  Logger.log('ステップ2: パラメータを検出中...');
  const parameter = detectScrollParameter(scrollEvent.eventName);
  
  if (!parameter) {
    result.message = '⚠️ パラメータが検出できませんでした（GA4で確認が必要）';
    Logger.log(result.message);
    // イベント名だけでも保存
    PropertiesService.getScriptProperties().setProperty('GTM_SCROLL_EVENT_NAME', scrollEvent.eventName);
    result.success = true;
    result.message = '✅ イベント名のみ保存しました: ' + scrollEvent.eventName;
    Logger.log(result.message);
    return result;
  }
  
  Logger.log('✅ パラメータ検出: ' + parameter.name);
  result.parameterName = parameter.name;
  
  // ステップ4: パラメータ値のパターンを解析
  Logger.log('ステップ3: パラメータパターンを解析中...');
  const pattern = analyzeScrollParameterPattern(parameter.values);
  
  Logger.log('✅ パターン解析: ' + pattern.type);
  result.pattern = pattern.type;
  
  // ステップ5: 設定を保存
  PropertiesService.getScriptProperties().setProperty('GTM_SCROLL_EVENT_NAME', scrollEvent.eventName);
  PropertiesService.getScriptProperties().setProperty('GTM_SCROLL_PARAMETER_NAME', parameter.name);
  PropertiesService.getScriptProperties().setProperty('GTM_SCROLL_PATTERN', pattern.type);
  
  result.success = true;
  result.message = '✅ GTM設定を自動検出・保存しました';
  
  Logger.log('\n=== 検出結果 ===');
  Logger.log('イベント名: ' + result.eventName);
  Logger.log('パラメータ名: ' + result.parameterName);
  Logger.log('パターン: ' + result.pattern);
  Logger.log('  - 10%刻み: 10, 20, 30, ..., 90');
  Logger.log('  - 4段階: 25, 50, 75, 90');
  Logger.log('================');
  
  return result;
}

/**
 * イベント分析シートからスクロールイベントを検出
 * 
 * @param {Sheet} eventSheet - イベント分析シート
 * @return {Object|null} {eventName, eventCount}
 */
function detectScrollEventFromSheet(eventSheet) {
  const data = eventSheet.getDataRange().getValues();
  const headers = data[0];
  
  // 列のインデックスを取得
  const colEventName = headers.indexOf('event_name');
  const colEventCount = headers.indexOf('event_count');
  
  if (colEventName === -1 || colEventCount === -1) {
    Logger.log('⚠️ イベント分析シートの列構造が不正です');
    return null;
  }
  
  // スクロール関連イベントを検索
  const scrollEvents = [];
  
  for (let i = 1; i < data.length; i++) {
    const eventName = data[i][colEventName];
    const eventCount = data[i][colEventCount];
    
    // イベント名に"scroll"を含む
    if (eventName && typeof eventName === 'string' && eventName.toLowerCase().includes('scroll')) {
      // GA4デフォルトの"scroll"イベント（拡張測定機能）は除外
      // 判定: イベント数が極端に少ない、またはパラメータなしの場合
      if (eventCount > 100) { // 100件以上あればGTMカスタムと判断
        scrollEvents.push({
          eventName: eventName,
          eventCount: eventCount
        });
      }
    }
  }
  
  // イベント数が最も多いものを選択
  if (scrollEvents.length === 0) {
    return null;
  }
  
  scrollEvents.sort((a, b) => b.eventCount - a.eventCount);
  
  Logger.log('検出されたスクロールイベント:');
  scrollEvents.forEach((event, index) => {
    Logger.log('  ' + (index + 1) + '. ' + event.eventName + ' (' + event.eventCount + '件)');
  });
  
  return scrollEvents[0];
}

/**
 * GA4からスクロールイベントのパラメータを検出
 * 
 * @param {String} eventName - イベント名
 * @return {Object|null} {name, values}
 */
function detectScrollParameter(eventName) {
  const propertyId = PropertiesService.getScriptProperties().getProperty('GA4_PROPERTY_ID');
  
  if (!propertyId) {
    Logger.log('⚠️ GA4 Property IDが設定されていません');
    return null;
  }
  
  // よくあるパラメータ名の候補
  const candidates = [
    'scrolled_percentage',
    'percent_scrolled',
    'scroll_depth',
    'scroll_percentage',
    'depth'
  ];
  
  for (const candidate of candidates) {
    try {
      const request = {
        dateRanges: [{
          startDate: '7daysAgo',
          endDate: 'yesterday'
        }],
        dimensions: [
          { name: 'eventName' },
          { name: 'customEvent:' + candidate }
        ],
        metrics: [
          { name: 'eventCount' }
        ],
        dimensionFilter: {
          filter: {
            fieldName: 'eventName',
            stringFilter: {
              value: eventName
            }
          }
        },
        limit: 100
      };
      
      const response = AnalyticsData.Properties.runReport(request, 'properties/' + propertyId);
      
      if (response.rows && response.rows.length > 0) {
        // パラメータ値を収集
        const values = response.rows.map(row => row.dimensionValues[1].value);
        
        Logger.log('✅ パラメータ「' + candidate + '」を検出');
        Logger.log('サンプル値: ' + values.slice(0, 5).join(', '));
        
        return {
          name: candidate,
          values: values
        };
      }
    } catch (error) {
      // このパラメータ名では見つからなかった
      continue;
    }
  }
  
  return null;
}

/**
 * スクロールパラメータ値のパターンを解析
 * 
 * @param {Array} values - パラメータ値の配列
 * @return {Object} {type, description}
 */
function analyzeScrollParameterPattern(values) {
  // %記号を除去して数値化
  const numbers = values.map(v => parseInt(String(v).replace('%', ''))).filter(n => !isNaN(n));
  
  // ユニークな値を抽出
  const uniqueValues = [...new Set(numbers)].sort((a, b) => a - b);
  
  Logger.log('ユニーク値: ' + uniqueValues.join(', '));
  
  // パターン判定
  if (uniqueValues.includes(25) && uniqueValues.includes(75) && !uniqueValues.includes(20) && !uniqueValues.includes(70)) {
    return {
      type: '4段階',
      description: '25, 50, 75, 90',
      values: [25, 50, 75, 90]
    };
  } else if (uniqueValues.includes(10) && uniqueValues.includes(20) && uniqueValues.includes(30)) {
    return {
      type: '10%刻み',
      description: '10, 20, 30, ..., 90',
      values: [10, 20, 30, 40, 50, 60, 70, 80, 90]
    };
  } else {
    return {
      type: 'カスタム',
      description: uniqueValues.join(', '),
      values: uniqueValues
    };
  }
}

/**
 * GTM設定自動検出のテスト
 */
function testAutoDetectGTMSettings() {
  Logger.log('=== GTM設定自動検出テスト ===\n');
  
  const result = autoDetectGTMSettings();
  
  Logger.log('\n=== テスト結果 ===');
  Logger.log('成功: ' + result.success);
  Logger.log('メッセージ: ' + result.message);
  
  if (result.success) {
    Logger.log('\n保存された設定:');
    Logger.log('  - GTM_SCROLL_EVENT_NAME: ' + PropertiesService.getScriptProperties().getProperty('GTM_SCROLL_EVENT_NAME'));
    Logger.log('  - GTM_SCROLL_PARAMETER_NAME: ' + PropertiesService.getScriptProperties().getProperty('GTM_SCROLL_PARAMETER_NAME'));
    Logger.log('  - GTM_SCROLL_PATTERN: ' + PropertiesService.getScriptProperties().getProperty('GTM_SCROLL_PATTERN'));
  }
  
  Logger.log('\n=== テスト完了 ===');
}

// ===================================
// 2. Clarityデータ取得
// ===================================

/**
 * Clarity Data Export APIからデータ取得（キャッシュ対応）
 * 
 * 制限事項:
 * - 1日10リクエスト/プロジェクト
 * - 過去1-3日分のみ取得可能
 * - 最大1,000行/リクエスト
 * 
 * キャッシュ:
 * - 24時間キャッシュ（開発時の再実行でAPI制限を節約）
 * - forceRefresh=trueで強制的に最新データを取得
 * 
 * @param {Boolean} forceRefresh - 強制リフレッシュ（デフォルト: false）
 * @return {Array} Clarityデータ配列
 */
function fetchClarityData(forceRefresh = false) {
  const token = PropertiesService.getScriptProperties().getProperty('CLARITY_API_TOKEN');
  
  if (!token) {
    Logger.log('❌ エラー: Clarity APIトークンが設定されていません');
    Logger.log('setupClarityAPI() 関数を実行してトークンを設定してください');
    return [];
  }
  
  // キャッシュからデータ取得を試みる
  if (!forceRefresh) {
    const cachedData = getClarityCache();
    if (cachedData) {
      Logger.log('✅ キャッシュからClarityデータを取得（API呼び出しなし）');
      Logger.log('キャッシュ件数: ' + cachedData.length + ' ページ');
      return cachedData;
    }
  } else {
    Logger.log('⚠️ 強制リフレッシュ: キャッシュをスキップしてAPI呼び出し');
  }
  
  const apiUrl = 'https://www.clarity.ms/export-data/api/v1/project-live-insights';
  
  // パラメータ: 過去3日分、URLディメンションでブレークダウン
  const params = {
    numOfDays: 3,
    dimension1: 'URL'
  };
  
  const url = apiUrl + '?' + Object.keys(params).map(key => 
    key + '=' + encodeURIComponent(params[key])
  ).join('&');
  
  const options = {
    method: 'GET',
    headers: {
      'Authorization': 'Bearer ' + token,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  };
  
  try {
    Logger.log('Clarity APIにリクエスト中...');
    const response = UrlFetchApp.fetch(url, options);
    const statusCode = response.getResponseCode();
    
    if (statusCode === 200) {
      const data = JSON.parse(response.getContentText());
      Logger.log('✅ Clarityデータ取得成功');
      const parsedData = parseClarityResponse(data);
      
      // キャッシュに保存（24時間）
      setClarityCache(parsedData);
      
      return parsedData;
    } else if (statusCode === 429) {
      Logger.log('⚠️ API制限超過（1日10リクエスト）');
      Logger.log('明日再試行してください');
      
      // 429エラーの場合、期限切れキャッシュがあれば使用
      const cachedData = getClarityCache(true);  // 期限切れでも取得
      if (cachedData) {
        Logger.log('⚠️ 期限切れキャッシュを使用します');
        return cachedData;
      }
      
      return [];
    } else if (statusCode === 401) {
      Logger.log('❌ 認証エラー: APIトークンが無効です');
      return [];
    } else {
      Logger.log('❌ APIエラー: ' + statusCode);
      Logger.log(response.getContentText());
      return [];
    }
  } catch (error) {
    Logger.log('❌ Clarityデータ取得エラー: ' + error.message);
    return [];
  }
}

// ===================================
// 2.5. Clarityキャッシュ管理
// ===================================

/**
 * Clarityキャッシュからデータ取得
 * 
 * @param {Boolean} ignoreExpiry - 期限切れでも取得（デフォルト: false）
 * @return {Array|null} キャッシュデータ、またはnull
 */
function getClarityCache(ignoreExpiry = false) {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'clarity_data';
  const cacheTimeKey = 'clarity_data_time';
  
  try {
    const cachedJson = cache.get(cacheKey);
    const cachedTime = cache.get(cacheTimeKey);
    
    if (!cachedJson) {
      Logger.log('キャッシュなし');
      return null;
    }
    
    // 期限チェック（24時間）
    if (!ignoreExpiry && cachedTime) {
      const cacheAge = Date.now() - parseInt(cachedTime);
      const maxAge = 24 * 60 * 60 * 1000;  // 24時間
      
      if (cacheAge > maxAge) {
        Logger.log('⚠️ キャッシュ期限切れ（' + Math.round(cacheAge / 1000 / 60 / 60) + '時間経過）');
        return null;
      }
      
      Logger.log('キャッシュ有効期限: ' + Math.round((maxAge - cacheAge) / 1000 / 60 / 60) + '時間');
    }
    
    const data = JSON.parse(cachedJson);
    return data;
    
  } catch (error) {
    Logger.log('⚠️ キャッシュ読み込みエラー: ' + error.message);
    return null;
  }
}

/**
 * Clarityデータをキャッシュに保存
 * 
 * @param {Array} data - Clarityデータ配列
 */
function setClarityCache(data) {
  if (!data || data.length === 0) {
    Logger.log('⚠️ 空データのためキャッシュしません');
    return;
  }
  
  const cache = CacheService.getScriptCache();
  const cacheKey = 'clarity_data';
  const cacheTimeKey = 'clarity_data_time';
  
  try {
    const dataJson = JSON.stringify(data);
    
    // データサイズチェック（CacheServiceの制限: 100KB）
    if (dataJson.length > 100000) {
      Logger.log('⚠️ データサイズが大きすぎるため、最初の50件のみキャッシュします');
      const truncatedData = data.slice(0, 50);
      cache.put(cacheKey, JSON.stringify(truncatedData), 86400);  // 24時間
    } else {
      cache.put(cacheKey, dataJson, 86400);  // 24時間
    }
    
    // タイムスタンプを保存
    cache.put(cacheTimeKey, Date.now().toString(), 86400);
    
    Logger.log('✅ Clarityデータをキャッシュに保存しました（24時間有効）');
    Logger.log('キャッシュ件数: ' + data.length + ' ページ');
    
  } catch (error) {
    Logger.log('⚠️ キャッシュ保存エラー: ' + error.message);
  }
}

/**
 * Clarityキャッシュをクリア
 */
function clearClarityCache() {
  const cache = CacheService.getScriptCache();
  cache.remove('clarity_data');
  cache.remove('clarity_data_time');
  Logger.log('✅ Clarityキャッシュをクリアしました');
}

/**
 * Clarityキャッシュの状態を確認
 */
function checkClarityCache() {
  Logger.log('=== Clarityキャッシュ状態確認 ===\n');
  
  const cache = CacheService.getScriptCache();
  const cachedJson = cache.get('clarity_data');
  const cachedTime = cache.get('clarity_data_time');
  
  if (!cachedJson) {
    Logger.log('❌ キャッシュなし');
    return;
  }
  
  try {
    const data = JSON.parse(cachedJson);
    const cacheAge = Date.now() - parseInt(cachedTime);
    const hoursAgo = Math.round(cacheAge / 1000 / 60 / 60 * 10) / 10;  // 小数点1桁
    const maxAge = 24;  // 24時間
    const remaining = maxAge - hoursAgo;
    
    Logger.log('✅ キャッシュあり');
    Logger.log('件数: ' + data.length + ' ページ');
    Logger.log('キャッシュ時刻: ' + new Date(parseInt(cachedTime)).toLocaleString('ja-JP'));
    Logger.log('経過時間: ' + hoursAgo + ' 時間');
    Logger.log('残り有効期限: ' + Math.max(0, remaining).toFixed(1) + ' 時間');
    
    if (remaining <= 0) {
      Logger.log('⚠️ 期限切れ');
    } else if (remaining < 1) {
      Logger.log('⚠️ まもなく期限切れ');
    }
    
    // サンプルデータ
    if (data.length > 0) {
      Logger.log('\nサンプルデータ（最初の1件）:');
      Logger.log('URL: ' + data[0].page_url);
      Logger.log('セッション数: ' + data[0].sessions);
      Logger.log('デッドクリック: ' + data[0].dead_clicks);
    }
    
  } catch (error) {
    Logger.log('❌ キャッシュ読み込みエラー: ' + error.message);
  }
  
  Logger.log('\n=== 確認完了 ===');
}

/**
 * URLを正規化（絶対URL→相対パス）
 * 
 * 例:
 * - "https://smaho-tap.com/" → "/"
 * - "https://smaho-tap.com/page" → "/page"
 * - "/page" → "/page" （既に相対パス）
 * 
 * @param {String} url - URL
 * @return {String} 正規化されたURL（相対パス）
 */
function normalizeUrl(url) {
  if (!url) return '';
  
  // 既に相対パスの場合はそのまま返す
  if (url.startsWith('/')) {
    return url;
  }
  
  // 絶対URLの場合、パス部分のみを抽出
  try {
    // URLオブジェクトとしてパース
    if (url.startsWith('http://') || url.startsWith('https://')) {
      const urlObj = new URL(url);
      const path = urlObj.pathname;
      
      // パスが空の場合は "/"
      return path || '/';
    }
  } catch (error) {
    // URLパースに失敗した場合はそのまま返す
    Logger.log('⚠️ URL正規化エラー: ' + url);
  }
  
  return url;
}

/**
 * Clarityレスポンスをパース
 * 
 * APIレスポンス構造:
 * [
 *   {
 *     "metricName": "Dead Click Count",
 *     "information": [
 *       { "URL": "/page1", "deadClickCount": 5 },
 *       { "URL": "/page2", "deadClickCount": 3 }
 *     ]
 *   },
 *   ...
 * ]
 * 
 * @param {Object} rawData - APIレスポンス
 * @return {Array} パース済みデータ配列
 */
function parseClarityResponse(rawData) {
  const result = {};
  const today = new Date();
  
  // レスポンスが配列かチェック
  if (!Array.isArray(rawData) || rawData.length === 0) {
    Logger.log('⚠️ Clarityレスポンスが空または配列ではありません');
    return [];
  }
  
  Logger.log('Clarityレスポンスをパース中... メトリック数: ' + rawData.length);
  
  // レスポンス構造を解析
  rawData.forEach(metric => {
    const metricName = metric.metricName;
    
    if (!metric.information) return;
    
    Logger.log('  処理中: ' + metricName + ' (' + metric.information.length + ' URL)');
    
    metric.information.forEach(item => {
      // URLフィールドを取得（Url、URL、urlの順に試す）
      const rawUrl = item.Url || item.URL || item.url;
      if (!rawUrl) return;
      
      // URLを正規化（絶対URL→相対パス）
      const url = normalizeUrl(rawUrl);
      
      // 既存データを検索または新規作成
      if (!result[url]) {
        result[url] = {
          fetch_date: today,
          page_url: url,
          sessions: 0,
          avg_scroll_depth: 0,
          dead_clicks: 0,
          rage_clicks: 0,
          quick_backs: 0,
          engagement_time: 0,
          script_errors: 0
        };
      }
      
      // メトリクス値はsubTotalに格納されている
      const value = item.subTotal || item.value || 0;
      
      // メトリクスを格納
      switch (metricName) {
        case 'DeadClickCount':
        case 'Dead Click Count':
          result[url].dead_clicks = parseFloat(value);
          break;
          
        case 'Rage Click Count':
          result[url].rage_clicks = parseFloat(value);
          break;
          
        case 'Quickback Click':
          result[url].quick_backs = parseFloat(value);
          break;
          
        case 'Scroll Depth':
          // スクロール深度は%で返される可能性があるので処理
          const scrollDepth = value.toString().replace('%', '');
          result[url].avg_scroll_depth = parseFloat(scrollDepth);
          break;
          
        case 'Engagement Time':
          result[url].engagement_time = parseFloat(value);
          // セッション数も取得
          if (item.sessionsCount) {
            result[url].sessions = parseInt(item.sessionsCount);
          }
          break;
          
        case 'Script Error Count':
          result[url].script_errors = parseFloat(value);
          break;
          
        case 'Traffic':
          // セッション数を取得
          if (item.sessionsCount) {
            result[url].sessions = parseInt(item.sessionsCount);
          }
          break;
      }
    });
  });
  
  // オブジェクトを配列に変換
  const resultArray = Object.values(result);
  Logger.log('✅ パース完了: ' + resultArray.length + ' ページ');
  
  // サンプルデータをログ出力（デバッグ用）
  if (resultArray.length > 0) {
    Logger.log('サンプルデータ（最初の1件）:');
    Logger.log(JSON.stringify(resultArray[0], null, 2));
  }
  
  return resultArray;
}

/**
 * Clarity_RAWシートにデータ書き込み
 * @param {Array} data - Clarityデータ配列
 */
function writeClarityDataToSheet(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Clarity_RAW');
  
  if (!sheet) {
    Logger.log('❌ エラー: Clarity_RAWシートが見つかりません');
    Logger.log('createClarityRawSheet() 関数を実行してシートを作成してください');
    return;
  }
  
  // ヘッダー確認
  const headers = sheet.getRange(1, 1, 1, 9).getValues()[0];
  if (headers[0] !== 'fetch_date') {
    Logger.log('❌ エラー: Clarity_RAWシートのヘッダーが不正です');
    return;
  }
  
  // データ行作成
  const rows = data.map(item => [
    item.fetch_date,
    item.page_url,
    item.sessions,
    item.avg_scroll_depth,
    item.dead_clicks,
    item.rage_clicks,
    item.quick_backs,
    item.engagement_time,
    item.script_errors
  ]);
  
  // 追記（既存データの下に追加）
  if (rows.length > 0) {
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, rows.length, 9).setValues(rows);
    Logger.log('✅ Clarityデータ ' + rows.length + '行を書き込みました');
  } else {
    Logger.log('⚠️ 書き込むデータがありません');
  }
}

// ===================================
// 3. GTMスクロールイベント取得
// ===================================

/**
 * GA4からGTMスクロールイベントを取得（自動検出対応）
 * 
 * Phase 1機能:
 * - 手動設定を優先（setupClarityAPI()で設定された場合）
 * - 手動設定がない場合は自動検出（複数候補を試す）
 * - 最初に見つかったイベントを使用
 * 
 * GTM設定:
 * - イベント名: 自動検出または手動設定
 * - イベントパラメータ: scrolled_percentage (10%刻み: 10-90 または 25, 50, 75, 90)
 * 
 * 注意:
 * - GA4デフォルトの「Scroll」イベント（拡張測定機能）は使用しない
 * - GTMで設定したカスタムイベントのみ使用（scrolled_percentageパラメータが必要）
 * - 10%刻みのデータは自動的に4段階（25, 50, 75, 90）にマッピングされます
 * 
 * @return {Object} URLごとのスクロールデータ
 */
function fetchGTMScrollEvents() {
  const propertyId = PropertiesService.getScriptProperties().getProperty('GA4_PROPERTY_ID');
  
  if (!propertyId) {
    Logger.log('❌ エラー: GA4 Property IDが設定されていません');
    return {};
  }
  
  // ステップ1: 手動設定を確認
  let scrollEventName = PropertiesService.getScriptProperties().getProperty('GTM_SCROLL_EVENT_NAME');
  
  if (scrollEventName) {
    // 手動設定あり
    Logger.log('GTMスクロールイベント名（手動設定）: ' + scrollEventName);
  } else {
    // ステップ2: 自動検出
    Logger.log('GTMスクロールイベント名が未設定 → 自動検出を開始');
    scrollEventName = tryDetectScrollEvent(propertyId);
    
    if (scrollEventName) {
      Logger.log('✅ スクロールイベント自動検出成功: ' + scrollEventName);
      // 次回のために保存（オプション: コメントアウト可能）
      // PropertiesService.getScriptProperties().setProperty('GTM_SCROLL_EVENT_NAME', scrollEventName);
    } else {
      Logger.log('❌ スクロールイベントが見つかりませんでした');
      Logger.log('以下を確認してください:');
      Logger.log('  1. GTMでスクロールイベントが設定されているか');
      Logger.log('  2. イベントパラメータ「scrolled_percentage」が含まれているか');
      Logger.log('  3. setupClarityAPI()で手動設定することも可能です');
      return {};
    }
  }
  
  // ステップ3: データ取得
  return fetchScrollDataByEventName(scrollEventName, propertyId);
}

/**
 * 複数候補を試してスクロールイベントを検出
 * 
 * @param {String} propertyId - GA4 Property ID
 * @return {String|null} 検出されたイベント名、または null
 */
function tryDetectScrollEvent(propertyId) {
  // よくあるスクロールイベント名の候補
  // 並び順: より具体的なイベント名を優先
  const candidates = [
    'scrolled',      // GTMカスタム（小文字、一般的） ← 優先度UP
    'Scrolled',      // GTMカスタム（大文字S）
    'scroll_depth',  // GTMカスタム
    'page_scroll',   // GTMカスタム
    'scroll'         // GTMカスタム（小文字）← 最後に移動（GA4デフォルトと混同を避ける）
  ];
  
  Logger.log('スクロールイベント候補を試行中: ' + candidates.join(', '));
  
  // 各候補を順番に試す
  for (let i = 0; i < candidates.length; i++) {
    const candidate = candidates[i];
    Logger.log('  試行 ' + (i + 1) + '/' + candidates.length + ': ' + candidate);
    
    // このイベント名でデータが取得できるか確認
    const hasData = checkScrollEventExists(candidate, propertyId);
    
    if (hasData) {
      Logger.log('  → ✅ データあり');
      return candidate;
    } else {
      Logger.log('  → ⚠️ データなし');
    }
  }
  
  return null;
}

/**
 * 指定したイベント名でスクロールデータが存在するか確認
 * 
 * @param {String} eventName - イベント名
 * @param {String} propertyId - GA4 Property ID
 * @return {Boolean} データが存在するか
 */
function checkScrollEventExists(eventName, propertyId) {
  // パラメータ名を取得（自動検出またはデフォルト）
  const parameterName = PropertiesService.getScriptProperties().getProperty('GTM_SCROLL_PARAMETER_NAME') || 'scrolled_percentage';
  
  try {
    const request = {
      dateRanges: [{
        startDate: '7daysAgo',  // 過去7日で確認（軽量）
        endDate: 'yesterday'
      }],
      dimensions: [
        { name: 'eventName' },
        { name: 'customEvent:' + parameterName } // 自動検出されたパラメータ名を使用
      ],
      metrics: [
        { name: 'eventCount' }
      ],
      dimensionFilter: {
        filter: {
          fieldName: 'eventName',
          stringFilter: {
            value: eventName
          }
        }
      },
      limit: 1  // 1件でも存在すればOK
    };
    
    const response = AnalyticsData.Properties.runReport(request, 'properties/' + propertyId);
    
    // データが存在し、scrolled_percentageパラメータがあるか確認
    if (response.rows && response.rows.length > 0) {
      const scrolledPercentage = response.rows[0].dimensionValues[1].value;
      // scrolled_percentageが有効な値か確認
      // 10%刻み(10-90)または25/50/75/90に対応
      const validValues = ['10', '20', '25', '30', '40', '50', '60', '70', '75', '80', '90',
                          '10%', '20%', '25%', '30%', '40%', '50%', '60%', '70%', '75%', '80%', '90%'];
      return validValues.some(val => scrolledPercentage.includes(val));
    }
    
    return false;
  } catch (error) {
    // エラーの場合はデータなしと判断
    return false;
  }
}

/**
 * 指定したイベント名でスクロールデータを取得
 * 
 * @param {String} eventName - イベント名
 * @param {String} propertyId - GA4 Property ID
 * @return {Object} URLごとのスクロールデータ
 */
function fetchScrollDataByEventName(eventName, propertyId) {
  // パラメータ名を取得（自動検出または デフォルト）
  const parameterName = PropertiesService.getScriptProperties().getProperty('GTM_SCROLL_PARAMETER_NAME') || 'scrolled_percentage';
  
  Logger.log('使用するパラメータ名: ' + parameterName);
  
  const request = {
    dateRanges: [{
      startDate: '30daysAgo',
      endDate: 'yesterday'
    }],
    dimensions: [
      { name: 'pagePath' },
      { name: 'customEvent:' + parameterName } // 自動検出されたパラメータ名を使用
    ],
    metrics: [
      { name: 'eventCount' }
    ],
    dimensionFilter: {
      filter: {
        fieldName: 'eventName',
        stringFilter: {
          value: eventName
        }
      }
    }
  };
  
  try {
    Logger.log('GA4からスクロールデータ取得中（イベント名: ' + eventName + '）...');
    const response = AnalyticsData.Properties.runReport(request, 'properties/' + propertyId);
    const result = parseScrollEvents(response);
    Logger.log('✅ スクロールデータ取得成功: ' + Object.keys(result).length + ' URL');
    return result;
  } catch (error) {
    Logger.log('❌ スクロールデータ取得エラー: ' + error.message);
    return {};
  }
}

/**
 * スクロールイベントをパース
 * 10%刻みのGTMデータを、25/50/75/90の4段階にマッピング
 * 
 * マッピング:
 * - 20% → scroll_25（25%の代わり）
 * - 50% → scroll_50（そのまま）
 * - 70% → scroll_75（75%の代わり）
 * - 90% → scroll_90（そのまま）
 * 
 * @param {Object} response - GA4 APIレスポンス
 * @return {Object} URLごとのスクロールデータ
 */
function parseScrollEvents(response) {
  const result = {};
  
  if (!response.rows) {
    Logger.log('⚠️ スクロールイベントデータなし');
    return result;
  }
  
  response.rows.forEach(row => {
    const url = row.dimensionValues[0].value;
    const percentStr = row.dimensionValues[1].value;
    const count = parseInt(row.metricValues[0].value);
    
    // %記号を除去してパース（"25%" → 25, "25" → 25）
    const percent = parseInt(percentStr.replace('%', ''));
    
    if (!result[url]) {
      result[url] = {
        scroll_25: 0,
        scroll_50: 0,
        scroll_75: 0,
        scroll_90: 0
      };
    }
    
    // パーセントごとにカウント（10%刻みのGTMデータを4段階にマッピング）
    if (percent === 20 || percent === 25) {
      // 20%または25% → scroll_25
      result[url].scroll_25 = count;
    } else if (percent === 50) {
      // 50% → scroll_50
      result[url].scroll_50 = count;
    } else if (percent === 70 || percent === 75) {
      // 70%または75% → scroll_75
      result[url].scroll_75 = count;
    } else if (percent === 90) {
      // 90% → scroll_90
      result[url].scroll_90 = count;
    }
    // 10%, 30%, 40%, 60%, 80%は使用しない
  });
  
  return result;
}

// ===================================
// 4. データ統合
// ===================================

/**
 * ClarityとGTMスクロールデータを統合データシートにマージ
 */
function mergeClarityAndScrollToIntegrated() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const claritySheet = ss.getSheetByName('Clarity_RAW');
  const integratedSheet = ss.getSheetByName('統合データ');
  
  if (!claritySheet) {
    Logger.log('❌ エラー: Clarity_RAWシートが見つかりません');
    return;
  }
  
  if (!integratedSheet) {
    Logger.log('❌ エラー: 統合データシートが見つかりません');
    return;
  }
  
  Logger.log('=== データ統合開始 ===');
  
  // 1. Clarity_RAWから最新データ取得（過去7日以内）
  const clarityData = getClarityLatestData(claritySheet);
  const hasClarityData = Object.keys(clarityData).length > 0;
  
  Logger.log('Clarityデータ: ' + Object.keys(clarityData).length + ' URL');
  
  if (!hasClarityData) {
    Logger.log('⚠️ Clarityデータなし → Clarity列の更新をスキップ（既存データを保護）');
  }
  
  // 2. GTMスクロールイベント取得
  const scrollData = fetchGTMScrollEvents();
  Logger.log('GTMスクロールデータ: ' + Object.keys(scrollData).length + ' URL');
  
  // 3. 統合データシート読み込み
  const lastRow = integratedSheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('⚠️ 統合データシートにデータがありません');
    return;
  }
  
  const urls = integratedSheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  Logger.log('統合データシート: ' + urls.length + ' URL');
  
  // 4. 統合データシートの列番号を取得
  const headers = integratedSheet.getRange(1, 1, 1, integratedSheet.getLastColumn()).getValues()[0];
  const colClarity = {
    sessions: headers.indexOf('clarity_sessions') + 1,
    avgScrollDepth: headers.indexOf('clarity_avg_scroll_depth') + 1,
    deadClicks: headers.indexOf('clarity_dead_clicks') + 1,
    rageClicks: headers.indexOf('clarity_rage_clicks') + 1,
    quickBacks: headers.indexOf('clarity_quick_backs') + 1,
    uxScore: headers.indexOf('clarity_ux_score') + 1,
    gtmScroll25: headers.indexOf('gtm_scroll_25_count') + 1,
    gtmScroll50: headers.indexOf('gtm_scroll_50_count') + 1,
    gtmScroll75: headers.indexOf('gtm_scroll_75_count') + 1,
    gtmScroll90: headers.indexOf('gtm_scroll_90_count') + 1
  };
  
  // 列が見つからない場合のエラーチェック
  if (colClarity.sessions === 0) {
    Logger.log('❌ エラー: 統合データシートにClarity関連の列が見つかりません');
    Logger.log('列を追加してください: clarity_sessions, clarity_avg_scroll_depth, etc.');
    return;
  }
  
  // 5. URL別にマッチングしてデータ更新
  let updateCount = 0;
  let clarityUpdateCount = 0;
  let gtmUpdateCount = 0;
  
  urls.forEach((url, index) => {
    const row = index + 2; // ヘッダー行を考慮
    
    // Clarityデータがある場合のみClarity列を更新
    if (hasClarityData) {
      // Clarityデータをマッチング
      const clarity = clarityData[url] || {
        sessions: 0,
        avg_scroll_depth: 0,
        dead_clicks: 0,
        rage_clicks: 0,
        quick_backs: 0
      };
      
      // GTMスクロールデータをマッチング
      const scroll = scrollData[url] || {
        scroll_25: 0,
        scroll_50: 0,
        scroll_75: 0,
        scroll_90: 0
      };
      
      // UXスコア計算
      const uxScore = calculateUXScore(clarity, scroll);
      
      // Clarity列を更新
      integratedSheet.getRange(row, colClarity.sessions).setValue(clarity.sessions);
      integratedSheet.getRange(row, colClarity.avgScrollDepth).setValue(clarity.avg_scroll_depth);
      integratedSheet.getRange(row, colClarity.deadClicks).setValue(clarity.dead_clicks);
      integratedSheet.getRange(row, colClarity.rageClicks).setValue(clarity.rage_clicks);
      integratedSheet.getRange(row, colClarity.quickBacks).setValue(clarity.quick_backs);
      integratedSheet.getRange(row, colClarity.uxScore).setValue(uxScore);
      
      clarityUpdateCount++;
    } else {
      // Clarityデータがない場合、GTMスクロールのみでUXスコア計算
      const scroll = scrollData[url] || {
        scroll_25: 0,
        scroll_50: 0,
        scroll_75: 0,
        scroll_90: 0
      };
      
      // 既存のClarityデータを取得（保護）
      const existingClarity = {
        sessions: integratedSheet.getRange(row, colClarity.sessions).getValue() || 0,
        avg_scroll_depth: integratedSheet.getRange(row, colClarity.avgScrollDepth).getValue() || 0,
        dead_clicks: integratedSheet.getRange(row, colClarity.deadClicks).getValue() || 0,
        rage_clicks: integratedSheet.getRange(row, colClarity.rageClicks).getValue() || 0,
        quick_backs: integratedSheet.getRange(row, colClarity.quickBacks).getValue() || 0
      };
      
      // UXスコア再計算（既存Clarityデータ + 新GTMデータ）
      const uxScore = calculateUXScore(existingClarity, scroll);
      integratedSheet.getRange(row, colClarity.uxScore).setValue(uxScore);
    }
    
    // GTMスクロールデータは常に更新
    const scroll = scrollData[url] || {
      scroll_25: 0,
      scroll_50: 0,
      scroll_75: 0,
      scroll_90: 0
    };
    
    integratedSheet.getRange(row, colClarity.gtmScroll25).setValue(scroll.scroll_25);
    integratedSheet.getRange(row, colClarity.gtmScroll50).setValue(scroll.scroll_50);
    integratedSheet.getRange(row, colClarity.gtmScroll75).setValue(scroll.scroll_75);
    integratedSheet.getRange(row, colClarity.gtmScroll90).setValue(scroll.scroll_90);
    
    gtmUpdateCount++;
    updateCount++;
  });
  
  Logger.log('✅ データ統合完了: ' + updateCount + ' URL更新');
  if (hasClarityData) {
    Logger.log('  - Clarityデータ更新: ' + clarityUpdateCount + ' URL');
  } else {
    Logger.log('  - Clarityデータ更新: スキップ（既存データ保護）');
  }
  Logger.log('  - GTMスクロールデータ更新: ' + gtmUpdateCount + ' URL');
}

/**
 * Clarity_RAWから最新データを取得（URL別に集計）
 * @param {Sheet} sheet - Clarity_RAWシート
 * @return {Object} URL別の最新Clarityデータ
 */
function getClarityLatestData(sheet) {
  const data = sheet.getDataRange().getValues();
  const result = {};
  const sevenDaysAgo = new Date();
  sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7);
  
  // ヘッダー除外、過去7日以内のデータのみ
  for (let i = 1; i < data.length; i++) {
    const fetchDate = new Date(data[i][0]);
    if (fetchDate < sevenDaysAgo) continue;
    
    const rawUrl = data[i][1];
    const url = normalizeUrl(rawUrl);  // URL正規化
    
    if (!result[url]) {
      result[url] = {
        sessions: 0,
        avg_scroll_depth: 0,
        dead_clicks: 0,
        rage_clicks: 0,
        quick_backs: 0,
        count: 0
      };
    }
    
    // 平均を計算するため累積
    result[url].sessions += data[i][2];
    result[url].avg_scroll_depth += data[i][3];
    result[url].dead_clicks += data[i][4];
    result[url].rage_clicks += data[i][5];
    result[url].quick_backs += data[i][6];
    result[url].count++;
  }
  
  // 平均値を計算
  Object.keys(result).forEach(url => {
    const count = result[url].count;
    if (count > 0) {
      result[url].avg_scroll_depth = Math.round(result[url].avg_scroll_depth / count);
      result[url].dead_clicks = Math.round(result[url].dead_clicks / count);
      result[url].rage_clicks = Math.round(result[url].rage_clicks / count);
      result[url].quick_backs = Math.round(result[url].quick_backs / count);
    }
  });
  
  return result;
}

// ===================================
// 5. UXスコアリング
// ===================================

/**
 * UX総合スコアを計算（柔軟版）
 * 
 * 加重平均:
 * - スクロール深度: 40%（最重要）
 * - デッドクリック: 25%
 * - レイジクリック: 20%
 * - クイックバック: 15%
 * 
 * データソース対応:
 * - Clarity + GTM: 両方のデータで統合評価
 * - Clarityのみ: Clarityデータのみで評価
 * - GTMのみ: GTMスクロールデータのみで評価
 * - 両方なし: 0点（評価不可）
 * 
 * @param {Object} clarityData - Clarityデータ
 * @param {Object} scrollData - GTMスクロールデータ
 * @return {Number} UXスコア（0-100）
 */
function calculateUXScore(clarityData, scrollData) {
  // データソースの有無を検出
  const hasClarityData = clarityData.sessions > 0 || clarityData.avg_scroll_depth > 0;
  const hasGTMData = (scrollData.scroll_25 + scrollData.scroll_50 + 
                      scrollData.scroll_75 + scrollData.scroll_90) > 0;
  
  // 両方ともデータなし → 0点（評価不可）
  if (!hasClarityData && !hasGTMData) {
    return 0;
  }
  
  // 各サブスコアを計算
  const scrollScore = calculateScrollScore(clarityData.avg_scroll_depth, scrollData);
  const deadClickScore = calculateDeadClickScore(clarityData.dead_clicks);
  const rageClickScore = calculateRageClickScore(clarityData.rage_clicks);
  const quickBackScore = calculateQuickBackScore(clarityData.quick_backs);
  
  // 加重平均
  const uxScore = 
    (scrollScore * 0.40) +      // スクロール深度（最重要）
    (deadClickScore * 0.25) +   // デッドクリック
    (rageClickScore * 0.20) +   // レイジクリック
    (quickBackScore * 0.15);    // クイックバック
  
  return Math.round(uxScore);
}

/**
 * スクロール深度スコア（Clarity + GTM統合・柔軟版）
 * 
 * スコアリングロジック:
 * - 平均30%未満 or 50%到達率30%未満 → 100点（大問題）
 * - 平均30-50% or 50%到達率30-50% → 70点（問題あり）
 * - 平均50-70% or 75%到達率30%未満 → 40点（改善余地あり）
 * - それ以上 → 0点（問題なし）
 * 
 * データソース対応:
 * - Clarity + GTM: 統合評価（最も厳密）
 * - Clarityのみ: Clarity平均スクロール深度で評価
 * - GTMのみ: GTM到達率で評価（Clarity優先度低いため）
 * - 両方なし: 0点
 * 
 * @param {Number} avgDepth - Clarity平均スクロール深度
 * @param {Object} scrollData - GTMスクロールデータ
 * @return {Number} スコア（0-100）
 */
function calculateScrollScore(avgDepth, scrollData) {
  // GTMデータから実際のスクロール到達率を計算
  const totalScrolls = scrollData.scroll_25 + scrollData.scroll_50 + 
                       scrollData.scroll_75 + scrollData.scroll_90;
  
  // パターン1: GTMデータなし
  if (totalScrolls === 0) {
    // Clarityデータもなし → 0点
    if (avgDepth === 0) return 0;
    
    // Clarityデータのみ使用
    if (avgDepth < 30) return 100;
    if (avgDepth < 50) return 70;
    if (avgDepth < 70) return 40;
    return 0;
  }
  
  // GTM到達率を計算
  const reach50 = (scrollData.scroll_50 + scrollData.scroll_75 + scrollData.scroll_90) / totalScrolls;
  const reach75 = (scrollData.scroll_75 + scrollData.scroll_90) / totalScrolls;
  
  // パターン2: Clarityデータなし、GTMのみ（GTM優先）
  if (avgDepth === 0) {
    if (reach50 < 0.3) return 100;  // 50%到達率が低い
    if (reach50 < 0.5) return 70;
    if (reach75 < 0.3) return 40;   // 75%到達率が低い
    return 0;
  }
  
  // パターン3: 両方あり - 統合評価
  if (avgDepth < 30 || reach50 < 0.3) return 100;  // 大問題
  if (avgDepth < 50 || reach50 < 0.5) return 70;   // 問題あり
  if (avgDepth < 70 || reach75 < 0.3) return 40;   // 改善余地あり
  return 0;  // 問題なし
}

/**
 * デッドクリックスコア
 * 
 * スコアリングロジック:
 * - 10回以上 → 100点（深刻な問題）
 * - 5-9回 → 60点（問題あり）
 * - 1-4回 → 30点（軽微な問題）
 * - 0回 → 0点（問題なし）
 * 
 * @param {Number} deadClicks - デッドクリック数
 * @return {Number} スコア（0-100）
 */
function calculateDeadClickScore(deadClicks) {
  if (deadClicks >= 10) return 100;  // 深刻な問題
  if (deadClicks >= 5) return 60;    // 問題あり
  if (deadClicks >= 1) return 30;    // 軽微な問題
  return 0;  // 問題なし
}

/**
 * レイジクリックスコア
 * 
 * スコアリングロジック:
 * - 5回以上 → 100点（深刻な問題）
 * - 3-4回 → 60点（問題あり）
 * - 1-2回 → 30点（軽微な問題）
 * - 0回 → 0点（問題なし）
 * 
 * @param {Number} rageClicks - レイジクリック数
 * @return {Number} スコア（0-100）
 */
function calculateRageClickScore(rageClicks) {
  if (rageClicks >= 5) return 100;   // 深刻な問題
  if (rageClicks >= 3) return 60;    // 問題あり
  if (rageClicks >= 1) return 30;    // 軽微な問題
  return 0;  // 問題なし
}

/**
 * クイックバックスコア
 * 
 * スコアリングロジック:
 * - 10回以上 → 100点（深刻な問題）
 * - 5-9回 → 60点（問題あり）
 * - 1-4回 → 30点（軽微な問題）
 * - 0回 → 0点（問題なし）
 * 
 * @param {Number} quickBacks - クイックバック数
 * @return {Number} スコア（0-100）
 */
function calculateQuickBackScore(quickBacks) {
  if (quickBacks >= 10) return 100;  // 深刻な問題
  if (quickBacks >= 5) return 60;    // 問題あり
  if (quickBacks >= 1) return 30;    // 軽微な問題
  return 0;  // 問題なし
}

// ===================================
// 6. 週次自動実行
// ===================================

/**
 * 週次トリガー設定
 * 毎週月曜早朝5:30に実行（データ収集の30分後）
 * 
 * 手動で一度だけ実行してください
 */
function setupWeeklyClarityTrigger() {
  // 既存のトリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'weeklyClarityUpdate') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // 新規トリガー作成（毎週月曜5:30）
  ScriptApp.newTrigger('weeklyClarityUpdate')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(5)
    .nearMinute(30)
    .create();
  
  Logger.log('✅ 週次Clarityトリガーを設定しました（毎週月曜5:30）');
}

/**
 * 週次Clarityデータ更新
 * トリガーから自動実行される
 */
function weeklyClarityUpdate() {
  Logger.log('====================================');
  Logger.log('=== 週次Clarity更新開始 ===');
  Logger.log('====================================');
  const startTime = new Date();
  
  try {
    // 0. データソース状況を確認
    Logger.log('Step 0: データソース状況確認');
    logDataSourceStatus();
    
    // 1. Clarityデータ取得
    Logger.log('Step 1: Clarityデータ取得');
    const clarityData = fetchClarityData();
    if (clarityData.length > 0) {
      writeClarityDataToSheet(clarityData);
    } else {
      Logger.log('⚠️ Clarityデータなし（データが存在しないか、API制限）');
    }
    
    // 2. データ統合
    Logger.log('Step 2: データ統合');
    mergeClarityAndScrollToIntegrated();
    
    // 3. パフォーマンススコア再計算
    Logger.log('Step 3: パフォーマンススコア再計算');
    recalculatePerformanceScores();
    
    const endTime = new Date();
    const duration = (endTime - startTime) / 1000;
    
    Logger.log('====================================');
    Logger.log('✅ 週次Clarity更新完了');
    Logger.log('実行時間: ' + duration + '秒');
    Logger.log('====================================');
    
  } catch (error) {
    Logger.log('====================================');
    Logger.log('❌ 週次Clarity更新エラー');
    Logger.log('エラー内容: ' + error.message);
    Logger.log('スタックトレース: ' + error.stack);
    Logger.log('====================================');
  }
}

/**
 * パフォーマンススコア再計算
 * Scoring.gsのcalculatePerformanceScore()を呼び出し
 */
function recalculatePerformanceScores() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('統合データ');
  
  if (!sheet) {
    Logger.log('❌ エラー: 統合データシートが見つかりません');
    return;
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('⚠️ 統合データシートにデータがありません');
    return;
  }
  
  // ヘッダー行から列番号を取得
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colBounceRate = headers.indexOf('bounce_rate') + 1;
  const colDuration = headers.indexOf('avg_session_duration') + 1;
  const colUxScore = headers.indexOf('clarity_ux_score') + 1;
  const colPerfScore = headers.indexOf('performance_score') + 1;
  
  if (colPerfScore === 0) {
    Logger.log('❌ エラー: performance_score列が見つかりません');
    return;
  }
  
  // データ取得
  const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  
  let updateCount = 0;
  
  data.forEach((row, index) => {
    const pageData = {
      bounce_rate: row[colBounceRate - 1],
      avg_session_duration: row[colDuration - 1],
      clarity_ux_score: row[colUxScore - 1]
    };
    
    // Scoring.gsのcalculatePerformanceScore()を呼び出し
    const newScore = calculatePerformanceScore(pageData);
    
    // パフォーマンススコア列を更新
    sheet.getRange(index + 2, colPerfScore).setValue(newScore);
    updateCount++;
  });
  
  Logger.log('✅ パフォーマンススコア再計算完了: ' + updateCount + ' ページ');
}

// ===================================
// 7. テスト・デバッグ
// ===================================

/**
 * データソース状況をログ出力
 * デバッグ・モニタリング用
 */
function logDataSourceStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('統合データ');
  
  if (!sheet) {
    Logger.log('⚠️ 統合データシートが見つかりません');
    return;
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log('⚠️ 統合データシートにデータがありません');
    return;
  }
  
  // ヘッダー行から列番号を取得
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // 列が存在するか確認
  const hasClarityColumns = headers.indexOf('clarity_sessions') !== -1;
  const hasGTMColumns = headers.indexOf('gtm_scroll_25_count') !== -1;
  
  if (!hasClarityColumns || !hasGTMColumns) {
    Logger.log('⚠️ Clarity/GTM列が見つかりません（列追加が必要）');
    return;
  }
  
  // サンプルページでデータソース確認（2行目）
  const sampleRow = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  const claritySessions = sampleRow[headers.indexOf('clarity_sessions')];
  const gtmScroll = sampleRow[headers.indexOf('gtm_scroll_25_count')];
  
  Logger.log('--- データソース状況 ---');
  
  const hasClarityData = claritySessions > 0;
  const hasGTMData = gtmScroll > 0;
  
  if (hasClarityData && hasGTMData) {
    Logger.log('✅ Clarity + GTM: 両方のデータあり（最適）');
    Logger.log('   → スクロール深度、デッドクリック、レイジクリック、クイックバックを統合評価');
  } else if (hasClarityData) {
    Logger.log('⚠️ Clarityのみ: GTMスクロールデータなし');
    Logger.log('   → Clarityデータのみでスクロール評価');
    Logger.log('   推奨: GTMを導入すると、より詳細なスクロール分析が可能です');
  } else if (hasGTMData) {
    Logger.log('⚠️ GTMのみ: Clarityデータなし');
    Logger.log('   → GTMスクロールデータのみで評価（デッドクリック等は評価不可）');
    Logger.log('   推奨: Clarityを導入すると、デッドクリック・レイジクリック分析が可能です');
  } else {
    Logger.log('❌ データなし: ClarityもGTMもデータなし');
    Logger.log('   → UXスコアは0点（評価不可）');
    Logger.log('   推奨: ClarityとGTMを導入してください');
  }
  
  Logger.log('------------------------');
}

/**
 * Day 9-10完全テスト
 * 手動で実行してすべての機能をテストします
 */
function testDay9_10Complete() {
  Logger.log('====================================');
  Logger.log('=== Day 9-10 完全テスト開始 ===');
  Logger.log('====================================');
  
  // 0. データソース状況確認
  Logger.log('\n[Test 0] データソース状況確認');
  logDataSourceStatus();
  
  // 1. APIトークン確認
  Logger.log('\n[Test 1] APIトークン確認');
  const token = PropertiesService.getScriptProperties().getProperty('CLARITY_API_TOKEN');
  const gtmEventName = PropertiesService.getScriptProperties().getProperty('GTM_SCROLL_EVENT_NAME');
  
  if (!token) {
    Logger.log('❌ エラー: Clarity APIトークンが設定されていません');
    Logger.log('setupClarityAPI() 関数を実行してトークンを設定してください');
    return;
  }
  Logger.log('✅ APIトークン: 設定済み');
  
  // GTM設定の確認
  if (gtmEventName) {
    const gtmParamName = PropertiesService.getScriptProperties().getProperty('GTM_SCROLL_PARAMETER_NAME');
    const gtmPattern = PropertiesService.getScriptProperties().getProperty('GTM_SCROLL_PATTERN');
    
    if (gtmParamName && gtmPattern) {
      Logger.log('✅ GTM設定: 自動検出済み');
      Logger.log('  - イベント名: ' + gtmEventName);
      Logger.log('  - パラメータ名: ' + gtmParamName);
      Logger.log('  - パターン: ' + gtmPattern);
    } else {
      Logger.log('⚠️ GTMスクロールイベント名: ' + gtmEventName + '（手動設定）');
    }
  } else {
    Logger.log('⚠️ GTMスクロールイベント名: 未設定（自動検出が必要）');
    Logger.log('推奨: autoDetectGTMSettings() 関数を実行してください');
  }
  
  // 2. Clarityデータ取得テスト
  Logger.log('\n[Test 2] Clarityデータ取得テスト');
  
  // キャッシュ状態確認
  checkClarityCache();
  
  Logger.log('Clarity APIにリクエスト中...');
  const clarityData = fetchClarityData();
  Logger.log('取得件数: ' + clarityData.length + '件');
  if (clarityData.length > 0) {
    Logger.log('サンプルデータ（最初の1件）:');
    Logger.log(JSON.stringify(clarityData[0], null, 2));
    Logger.log('✅ Clarityデータ取得: 成功');
  } else {
    Logger.log('⚠️ Clarityデータなし');
    Logger.log('可能性: (1)過去3日間データなし、(2)API制限、(3)レスポンスパースエラー');
  }
  
  // 3. GTMスクロールイベント取得テスト
  Logger.log('\n[Test 3] GTMスクロールイベント取得テスト');
  const scrollData = fetchGTMScrollEvents();
  const scrollUrls = Object.keys(scrollData);
  Logger.log('取得URL数: ' + scrollUrls.length + '件');
  if (scrollUrls.length > 0) {
    Logger.log('サンプルURL: ' + scrollUrls[0]);
    Logger.log('スクロールデータ: ' + JSON.stringify(scrollData[scrollUrls[0]]));
    Logger.log('✅ GTMスクロールイベント: 成功');
  } else {
    Logger.log('❌ GTMスクロールイベント取得失敗');
  }
  
  // 4. UXスコア計算テスト
  Logger.log('\n[Test 4] UXスコア計算テスト');
  const testClarity = {
    avg_scroll_depth: 45,
    dead_clicks: 3,
    rage_clicks: 1,
    quick_backs: 2
  };
  const testScroll = {
    scroll_25: 100,
    scroll_50: 80,
    scroll_75: 50,
    scroll_90: 20
  };
  const uxScore = calculateUXScore(testClarity, testScroll);
  Logger.log('テストUXスコア: ' + uxScore + '点');
  Logger.log('内訳:');
  Logger.log('  - スクロールスコア: ' + calculateScrollScore(testClarity.avg_scroll_depth, testScroll));
  Logger.log('  - デッドクリックスコア: ' + calculateDeadClickScore(testClarity.dead_clicks));
  Logger.log('  - レイジクリックスコア: ' + calculateRageClickScore(testClarity.rage_clicks));
  Logger.log('  - クイックバックスコア: ' + calculateQuickBackScore(testClarity.quick_backs));
  Logger.log('✅ UXスコア計算: 成功');
  
  // 5. データ統合テスト
  Logger.log('\n[Test 5] データ統合テスト');
  try {
    mergeClarityAndScrollToIntegrated();
    Logger.log('✅ データ統合: 成功');
  } catch (error) {
    Logger.log('❌ データ統合エラー: ' + error.message);
  }
  
  // 6. パフォーマンススコア再計算テスト
  Logger.log('\n[Test 6] パフォーマンススコア再計算テスト');
  try {
    recalculatePerformanceScores();
    Logger.log('✅ パフォーマンススコア再計算: 成功');
  } catch (error) {
    Logger.log('❌ パフォーマンススコア再計算エラー: ' + error.message);
  }
  
  Logger.log('\n====================================');
  Logger.log('=== Day 9-10 完全テスト完了 ===');
  Logger.log('====================================');
}

/**
 * Clarity_RAWから統合データシートにデータをマージ（値を合計バージョン）
 */
function mergeClarityToIntegratedFixed() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const claritySheet = ss.getSheetByName('Clarity_RAW');
  const integratedSheet = ss.getSheetByName('統合データ');
  
  if (!claritySheet || !integratedSheet) {
    Logger.log('シートが見つかりません');
    return;
  }
  
  // Clarity_RAWからデータを読み込み
  const clarityData = claritySheet.getDataRange().getValues();
  Logger.log(`Clarity_RAW: ${clarityData.length - 1}行`);
  
  // URLをキーにしたマップを作成（値を合計）
  const clarityMap = {};
  for (let i = 1; i < clarityData.length; i++) {
    const row = clarityData[i];
    const rawUrl = row[1]; // B列: page_url
    
    if (!rawUrl) continue;
    
    // URL正規化（フルURLからパスを抽出、アンカー除去）
    let normalizedUrl = String(rawUrl);
    
    // https://smaho-tap.com/xxx → /xxx
    normalizedUrl = normalizedUrl.replace(/^https?:\/\/[^\/]+/, '');
    
    // アンカー除去（#以降を削除）
    normalizedUrl = normalizedUrl.split('#')[0];
    
    // 末尾スラッシュ除去
    normalizedUrl = normalizedUrl.replace(/\/$/, '');
    
    // 空の場合はスキップ
    if (!normalizedUrl || normalizedUrl === '') continue;
    
    // 値を合計（既存の値に加算）
    if (!clarityMap[normalizedUrl]) {
      clarityMap[normalizedUrl] = {
        sessions: 0,
        avgScrollDepth: 0,
        deadClicks: 0,
        rageClicks: 0,
        quickBacks: 0
      };
    }
    
    clarityMap[normalizedUrl].sessions += parseFloat(row[2]) || 0;
    clarityMap[normalizedUrl].avgScrollDepth = Math.max(clarityMap[normalizedUrl].avgScrollDepth, parseFloat(row[3]) || 0);
    clarityMap[normalizedUrl].deadClicks += parseFloat(row[4]) || 0;
    clarityMap[normalizedUrl].rageClicks += parseFloat(row[5]) || 0;
    clarityMap[normalizedUrl].quickBacks += parseFloat(row[6]) || 0;
  }
  
  Logger.log(`ClarityMap: ${Object.keys(clarityMap).length}件`);
  
  // サンプルデータを表示
  const sampleUrls = Object.keys(clarityMap).slice(0, 5);
  sampleUrls.forEach(url => {
    Logger.log(`  ${url}: rage_clicks=${clarityMap[url].rageClicks}`);
  });
  
  // 統合データシートを読み込み
  const integratedData = integratedSheet.getDataRange().getValues();
  const headers = integratedData[0];
  
  // 列インデックスを取得
  const urlCol = headers.indexOf('page_url');
  const sessionsCol = headers.indexOf('clarity_sessions');
  const scrollCol = headers.indexOf('clarity_avg_scroll_depth');
  const deadClicksCol = headers.indexOf('clarity_dead_clicks');
  const rageClicksCol = headers.indexOf('clarity_rage_clicks');
  const quickBacksCol = headers.indexOf('clarity_quick_backs');
  
  Logger.log(`列インデックス: url=${urlCol}, sessions=${sessionsCol}, rageClicks=${rageClicksCol}`);
  
  if (sessionsCol === -1 || rageClicksCol === -1) {
    Logger.log('❌ Clarity列が見つかりません');
    return;
  }
  
  // マッチング＆更新
  let matchCount = 0;
  let updateCount = 0;
  
  for (let i = 1; i < integratedData.length; i++) {
    const pageUrl = integratedData[i][urlCol];
    
    if (!pageUrl) continue;
    
    // 統合データのURLを正規化（先頭スラッシュ確認）
    let normalizedPageUrl = String(pageUrl);
    if (!normalizedPageUrl.startsWith('/')) {
      normalizedPageUrl = '/' + normalizedPageUrl;
    }
    // 末尾スラッシュ除去
    normalizedPageUrl = normalizedPageUrl.replace(/\/$/, '');
    
    // マッチング
    if (clarityMap[normalizedPageUrl]) {
      matchCount++;
      const data = clarityMap[normalizedPageUrl];
      
      // 値が0でない場合のみカウント
      if (data.rageClicks > 0 || data.deadClicks > 0) {
        updateCount++;
      }
      
      // 更新
      integratedData[i][sessionsCol] = data.sessions;
      integratedData[i][scrollCol] = data.avgScrollDepth;
      integratedData[i][deadClicksCol] = data.deadClicks;
      integratedData[i][rageClicksCol] = data.rageClicks;
      integratedData[i][quickBacksCol] = data.quickBacks;
    }
  }
  
  Logger.log(`マッチ: ${matchCount}件, 実データあり: ${updateCount}件`);
  
  // 一括書き込み
  integratedSheet.getRange(1, 1, integratedData.length, integratedData[0].length).setValues(integratedData);
  
  Logger.log('✅ Clarityデータ統合完了（修正版）');
}

/**
 * Clarity_RAWの列構造をデバッグ
 */
function debugClarityRawColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Clarity_RAW');
  
  if (!sheet) {
    Logger.log('Clarity_RAWシートが見つかりません');
    return;
  }
  
  // ヘッダー行を取得
  const headers = sheet.getRange(1, 1, 1, 10).getValues()[0];
  Logger.log('=== ヘッダー行 ===');
  headers.forEach((h, i) => {
    Logger.log(`  列${i} (${String.fromCharCode(65+i)}): ${h}`);
  });
  
  // データ行を数行取得
  const dataRows = sheet.getRange(2, 1, 5, 10).getValues();
  Logger.log('=== データ行（最初の5行） ===');
  dataRows.forEach((row, rowIndex) => {
    Logger.log(`行${rowIndex + 2}: ${row.join(' | ')}`);
  });
}