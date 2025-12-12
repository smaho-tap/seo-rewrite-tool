/**
 * ============================================================================
 * Day 11-12: 競合分析API連携 - 認証設定
 * ============================================================================
 * DataForSEO/Moz APIの認証情報を設定するための関数群
 * 
 * 実行手順:
 * 1. DataForSEO/Moz APIアカウントを作成
 * 2. setDataForSEOCredentials() を実行
 * 3. setMozAPIKey() を実行
 * 4. testAPIConnections() で接続確認
 */

/**
 * DataForSEO API認証情報を設定
 * 
 * 実行方法:
 * 1. この関数内のloginとpasswordを実際の値に書き換える
 * 2. Apps Scriptエディタで実行
 * 3. 実行後、login/passwordの値を削除して保存（セキュリティのため）
 */
function setDataForSEOCredentials() {
  // ★★★ ここに実際の値を入力してください ★★★
  const login = 'your-email@example.com';  // DataForSEOのログインID（メールアドレス）
  const password = 'your-password-here';   // DataForSEOのパスワード
  
  // PropertiesServiceに保存
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('DATAFORSEO_LOGIN', login);
  scriptProperties.setProperty('DATAFORSEO_PASSWORD', password);
  
  Logger.log('✓ DataForSEO認証情報を保存しました');
  Logger.log('セキュリティのため、この関数内のlogin/passwordを削除してください');
}

/**
 * Moz API認証情報を設定
 * 
 * 実行方法:
 * 1. この関数内のapiKeyを実際の値に書き換える
 * 2. Apps Scriptエディタで実行
 * 3. 実行後、apiKeyの値を削除して保存（セキュリティのため）
 */
function setMozAPIKey() {
  // ★★★ ここに実際の値を入力してください ★★★
  const apiKey = 'mozscape-your-api-key-here';  // Moz API Key
  
  // PropertiesServiceに保存
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('MOZ_API_KEY', apiKey);
  
  Logger.log('✓ Moz API認証情報を保存しました');
  Logger.log('セキュリティのため、この関数内のapiKeyを削除してください');
}

/**
 * DA許容範囲を設定（デフォルト: +10）
 * 
 * 実務経験に基づき、自社DA + 10までを「勝ち目あり」と判断
 * 将来的に調整する場合は、この値を変更
 */
function setDATolerance() {
  const tolerance = 10;  // 許容範囲（+10まで）
  
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('DA_TOLERANCE', tolerance.toString());
  
  Logger.log(`✓ DA許容範囲を ${tolerance} に設定しました`);
  Logger.log('自社DA + 10までのサイトを「勝ち目あり」と判断します');
}

/**
 * 保存された認証情報を確認（テスト用）
 * 
 * 注意: パスワード/APIキーは表示しません（セキュリティのため）
 */
function checkSavedCredentials() {
  const scriptProperties = PropertiesService.getScriptProperties();
  
  const dataForSeoLogin = scriptProperties.getProperty('DATAFORSEO_LOGIN');
  const dataForSeoPassword = scriptProperties.getProperty('DATAFORSEO_PASSWORD');
  const mozApiKey = scriptProperties.getProperty('MOZ_API_KEY');
  const daTolerance = scriptProperties.getProperty('DA_TOLERANCE');
  
  Logger.log('=== 保存された認証情報 ===');
  Logger.log(`DataForSEO Login: ${dataForSeoLogin ? '✓ 設定済み' : '✗ 未設定'}`);
  Logger.log(`DataForSEO Password: ${dataForSeoPassword ? '✓ 設定済み' : '✗ 未設定'}`);
  Logger.log(`Moz API Key: ${mozApiKey ? '✓ 設定済み' : '✗ 未設定'}`);
  Logger.log(`DA許容範囲: ${daTolerance || 10}`);
  
  if (!dataForSeoLogin || !dataForSeoPassword) {
    Logger.log('⚠ DataForSEO認証情報が未設定です。setDataForSEOCredentials()を実行してください');
  }
  
  if (!mozApiKey) {
    Logger.log('⚠ Moz API認証情報が未設定です。setMozAPIKey()を実行してください');
  }
  
  if (dataForSeoLogin && dataForSeoPassword && mozApiKey) {
    Logger.log('✓ すべての認証情報が設定されています。testAPIConnections()で接続確認できます');
  }
}

/**
 * API接続テスト（DataForSEO）
 * 
 * テストキーワード「iphone 保険」で検索結果を1件取得
 * コスト: $0.006
 */
function testDataForSEOConnection() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const login = scriptProperties.getProperty('DATAFORSEO_LOGIN');
  const password = scriptProperties.getProperty('DATAFORSEO_PASSWORD');
  
  if (!login || !password) {
    Logger.log('✗ DataForSEO認証情報が未設定です');
    return false;
  }
  
  const url = 'https://api.dataforseo.com/v3/serp/google/organic/live/advanced';
  
  const requestBody = [{
    "keyword": "iphone 保険",
    "location_code": 2392,  // 日本
    "language_code": "ja",
    "device": "desktop",
    "depth": 10
  }];
  
  const options = {
    method: 'POST',
    headers: {
      'Authorization': 'Basic ' + Utilities.base64Encode(login + ':' + password),
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(requestBody),
    muteHttpExceptions: true
  };
  
  try {
    Logger.log('DataForSEO APIに接続中...');
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode === 200) {
      const data = JSON.parse(response.getContentText());
      const results = data.tasks[0].result[0].items || [];
      
      Logger.log('✓ DataForSEO API接続成功！');
      Logger.log(`✓ 検索結果 ${results.length} 件取得`);
      Logger.log(`1位: ${results[0]?.url || 'N/A'}`);
      Logger.log(`コスト: $0.006`);
      return true;
    } else {
      Logger.log(`✗ DataForSEO APIエラー: ${responseCode}`);
      Logger.log(response.getContentText());
      return false;
    }
  } catch (error) {
    Logger.log(`✗ DataForSEO API接続エラー: ${error.message}`);
    return false;
  }
}

/**
 * API接続テスト（Moz）
 * 
 * テストURL「example.com」のDAを取得
 * コスト: 無料枠内
 */
function testMozConnection() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const apiKey = scriptProperties.getProperty('MOZ_API_KEY');
  
  if (!apiKey) {
    Logger.log('✗ Moz API認証情報が未設定です');
    return false;
  }
  
  const url = 'https://lsapi.seomoz.com/v2/url_metrics';
  
  const requestBody = {
    "targets": ["example.com"]
  };
  
  const options = {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${apiKey}`,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(requestBody),
    muteHttpExceptions: true
  };
  
  try {
    Logger.log('Moz APIに接続中...');
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode === 200) {
      const data = JSON.parse(response.getContentText());
      const result = data.results[0];
      
      Logger.log('✓ Moz API接続成功！');
      Logger.log(`✓ example.com の DA: ${result.domain_authority || 'N/A'}`);
      Logger.log(`✓ example.com の PA: ${result.page_authority || 'N/A'}`);
      return true;
    } else {
      Logger.log(`✗ Moz APIエラー: ${responseCode}`);
      Logger.log(response.getContentText());
      return false;
    }
  } catch (error) {
    Logger.log(`✗ Moz API接続エラー: ${error.message}`);
    return false;
  }
}

/**
 * 両方のAPI接続をテスト
 * 
 * 実行手順:
 * 1. setDataForSEOCredentials() 実行済み
 * 2. setMozAPIKey() 実行済み
 * 3. この関数を実行して接続確認
 * 
 * コスト: 約$0.006（DataForSEOのテストのみ）
 */
function testAPIConnections() {
  Logger.log('=== API接続テスト開始 ===');
  Logger.log('');
  
  const dataForSeoOk = testDataForSEOConnection();
  Logger.log('');
  
  const mozOk = testMozConnection();
  Logger.log('');
  
  Logger.log('=== テスト結果 ===');
  Logger.log(`DataForSEO API: ${dataForSeoOk ? '✓ OK' : '✗ NG'}`);
  Logger.log(`Moz API: ${mozOk ? '✓ OK' : '✗ NG'}`);
  
  if (dataForSeoOk && mozOk) {
    Logger.log('');
    Logger.log('✓ すべてのAPI接続が正常です！');
    Logger.log('次のステップ: createCompetitorAnalysisSheet() を実行して競合分析シートを作成');
  } else {
    Logger.log('');
    Logger.log('⚠ 一部のAPI接続に失敗しました。認証情報を確認してください');
  }
}

/**
 * 初期設定を一括実行（便利関数）
 * 
 * 注意: 実行前に、この関数内の認証情報を実際の値に書き換えてください
 */
function setupAllCredentials() {
  Logger.log('=== 初期設定開始 ===');
  Logger.log('');
  
  // 認証情報設定
  Logger.log('1. DataForSEO認証情報を設定中...');
  setDataForSEOCredentials();
  Logger.log('');
  
  Logger.log('2. Moz API認証情報を設定中...');
  setMozAPIKey();
  Logger.log('');
  
  Logger.log('3. DA許容範囲を設定中...');
  setDATolerance();
  Logger.log('');
  
  // 設定確認
  Logger.log('4. 設定内容を確認中...');
  checkSavedCredentials();
  Logger.log('');
  
  Logger.log('=== 初期設定完了 ===');
  Logger.log('次のステップ: testAPIConnections() を実行して接続確認');
}
function showAPIKeys() {
  const props = PropertiesService.getScriptProperties();
  Logger.log('DATAFORSEO_LOGIN: ' + props.getProperty('DATAFORSEO_LOGIN'));
  Logger.log('DATAFORSEO_PASSWORD: ' + props.getProperty('DATAFORSEO_PASSWORD'));
  Logger.log('MOZ_API_KEY: ' + props.getProperty('MOZ_API_KEY'));
}