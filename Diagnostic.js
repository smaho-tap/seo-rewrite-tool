/**
 * APIキー診断ツール - 完全版
 * Apps Scriptにコピペして実行してください
 */

// ========================================
// 診断1: PropertiesServiceの確認
// ========================================
function diagnostic1_CheckProperties() {
  Logger.log('=== 診断1: PropertiesService確認 ===');
  
  const key = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');
  
  if (!key) {
    Logger.log('❌ CRITICAL: APIキーが保存されていません！');
    Logger.log('→ setClaudeAPIKey() を実行してください');
    return false;
  }
  
  Logger.log('✅ APIキーは保存されています');
  Logger.log('   先頭15文字: ' + key.substring(0, 15));
  Logger.log('   末尾15文字: ' + key.substring(key.length - 15));
  Logger.log('   キー長: ' + key.length + '文字');
  
  // 形式チェック
  if (!key.startsWith('sk-ant-api03-')) {
    Logger.log('⚠️ WARNING: APIキーの形式が正しくない可能性があります');
    return false;
  }
  
  // 古いキーのチェック
  if (key.includes('E3LatW224') || key.includes('slPFWgAA')) {
    Logger.log('❌ CRITICAL: これは古い無効なAPIキーです！');
    Logger.log('→ 新しいAPIキーで上書きしてください');
    return false;
  }
  
  Logger.log('✅ APIキーの形式は正常です');
  return true;
}

// ========================================
// 診断2: ClaudeSetup.gsの確認
// ========================================
function diagnostic2_CheckSetupFile() {
  Logger.log('=== 診断2: ClaudeSetup.gs確認 ===');
  Logger.log('⚠️ 手動確認が必要:');
  Logger.log('1. ClaudeSetup.gs を開く');
  Logger.log('2. setClaudeAPIKey() 関数内のAPIキーを確認');
  Logger.log('3. 新しいAPIキーになっているか確認');
  Logger.log('');
  Logger.log('もし古いキー（E3LatW224...）が残っている場合:');
  Logger.log('→ 新しいAPIキーに書き換えて、setClaudeAPIKey()を再実行');
}

// ========================================
// 診断3: API接続テスト
// ========================================
function diagnostic3_TestConnection() {
  Logger.log('=== 診断3: API接続テスト ===');
  
  const apiKey = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');
  
  if (!apiKey) {
    Logger.log('❌ APIキーが保存されていないため、テストできません');
    return false;
  }
  
  const url = 'https://api.anthropic.com/v1/messages';
  
  const requestBody = {
    model: 'claude-sonnet-4-20250514',
    max_tokens: 50,
    messages: [{
      role: 'user',
      content: 'Test connection. Reply with OK.'
    }]
  };
  
  const options = {
    method: 'post',
    headers: {
      'Content-Type': 'application/json',
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01'
    },
    payload: JSON.stringify(requestBody),
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    Logger.log('レスポンスコード: ' + responseCode);
    
    if (responseCode === 200) {
      Logger.log('✅ API接続成功！');
      const data = JSON.parse(responseText);
      Logger.log('Claude応答: ' + data.content[0].text);
      return true;
    } else if (responseCode === 401) {
      Logger.log('❌ CRITICAL: 401 認証エラー');
      Logger.log('エラー詳細: ' + responseText);
      Logger.log('');
      Logger.log('原因:');
      Logger.log('1. APIキーが無効（Anthropic側で削除済み）');
      Logger.log('2. APIキーが間違っている');
      Logger.log('3. PropertiesServiceに古いキーが保存されている');
      Logger.log('');
      Logger.log('対処法:');
      Logger.log('1. Anthropic Consoleで新しいAPIキーを発行');
      Logger.log('2. setClaudeAPIKey()の中身を新しいキーに書き換え');
      Logger.log('3. setClaudeAPIKey()を実行');
      Logger.log('4. 再度この診断を実行');
      return false;
    } else {
      Logger.log('❌ エラー: ' + responseCode);
      Logger.log(responseText);
      return false;
    }
  } catch (error) {
    Logger.log('❌ 例外エラー: ' + error.message);
    return false;
  }
}

// ========================================
// 診断4: 全プロパティ一覧
// ========================================
function diagnostic4_ListAllProperties() {
  Logger.log('=== 診断4: 全プロパティ一覧 ===');
  
  const properties = PropertiesService.getScriptProperties().getProperties();
  const keys = Object.keys(properties);
  
  Logger.log('保存されているプロパティ数: ' + keys.length);
  Logger.log('');
  
  if (keys.length === 0) {
    Logger.log('❌ プロパティが1つも保存されていません！');
    return;
  }
  
  keys.forEach(function(key) {
    if (key === 'CLAUDE_API_KEY') {
      const value = properties[key];
      Logger.log('プロパティ: ' + key);
      Logger.log('  → 先頭15文字: ' + value.substring(0, 15));
      Logger.log('  → 末尾15文字: ' + value.substring(value.length - 15));
    } else {
      Logger.log('プロパティ: ' + key + ' = ' + properties[key]);
    }
  });
}

// ========================================
// 総合診断（すべて実行）
// ========================================
function runFullDiagnostic() {
  Logger.log('');
  Logger.log('╔════════════════════════════════════╗');
  Logger.log('║  APIキー完全診断 - 開始           ║');
  Logger.log('╚════════════════════════════════════╝');
  Logger.log('');
  
  const result1 = diagnostic1_CheckProperties();
  Logger.log('');
  
  diagnostic2_CheckSetupFile();
  Logger.log('');
  
  if (result1) {
    diagnostic3_TestConnection();
  } else {
    Logger.log('=== 診断3: スキップ（APIキー未保存のため） ===');
  }
  Logger.log('');
  
  diagnostic4_ListAllProperties();
  Logger.log('');
  
  Logger.log('╔════════════════════════════════════╗');
  Logger.log('║  診断完了                         ║');
  Logger.log('╚════════════════════════════════════╝');
}

// ========================================
// 強制的にAPIキーを再設定
// ========================================
function forceResetAPIKey() {
  Logger.log('=== APIキー強制再設定 ===');
  Logger.log('');
  Logger.log('⚠️ 重要: この関数を実行する前に:');
  Logger.log('1. ClaudeSetup.gs を開く');
  Logger.log('2. setClaudeAPIKey() 関数内のAPIキーを新しいものに書き換える');
  Logger.log('3. この関数ではなく setClaudeAPIKey() を実行する');
  Logger.log('');
  Logger.log('それでも問題が解決しない場合は、以下を確認:');
  Logger.log('- Apps Scriptプロジェクトが複数ある場合、正しいプロジェクトを編集しているか');
  Logger.log('- 実行ログで "Claude API Keyを保存しました" が表示されているか');
}

// ========================================
// 使用方法
// ========================================
function showInstructions() {
  Logger.log('╔════════════════════════════════════════════════╗');
  Logger.log('║  APIキー診断ツール - 使用方法                 ║');
  Logger.log('╚════════════════════════════════════════════════╝');
  Logger.log('');
  Logger.log('【実行手順】');
  Logger.log('1. 関数選択で "runFullDiagnostic" を選択');
  Logger.log('2. 実行ボタンをクリック');
  Logger.log('3. 実行ログを確認');
  Logger.log('');
  Logger.log('【個別診断】');
  Logger.log('- diagnostic1_CheckProperties: PropertiesService確認');
  Logger.log('- diagnostic3_TestConnection: API接続テスト');
  Logger.log('- diagnostic4_ListAllProperties: 全プロパティ表示');
  Logger.log('');
  Logger.log('まずは runFullDiagnostic() を実行してください！');
}