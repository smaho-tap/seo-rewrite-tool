/**
 * Claude API Key確認用
 * APIキーはGASの「プロジェクトの設定」→「スクリプトプロパティ」で設定してください
 */
function checkClaudeAPIKey() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');
  
  if (!apiKey) {
    Logger.log('❌ CLAUDE_API_KEYが設定されていません');
    Logger.log('→ GASの「プロジェクトの設定」→「スクリプトプロパティ」で設定してください');
    return false;
  }
  
  Logger.log('✅ CLAUDE_API_KEYは設定されています');
  Logger.log('キーの先頭: ' + apiKey.substring(0, 20) + '...');
  return true;
}

function testClaudeAPI() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');
  
  if (!apiKey) {
    Logger.log('❌ API Keyが設定されていません');
    return;
  }
  
  const url = 'https://api.anthropic.com/v1/messages';
  
  const requestBody = {
    model: 'claude-sonnet-4-20250514',
    max_tokens: 100,
    messages: [
      {
        role: 'user',
        content: 'Hello! Please respond with "API connection successful!"'
      }
    ]
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
      const data = JSON.parse(responseText);
      Logger.log('✅ Claude API接続成功！');
      Logger.log('Claude応答: ' + data.content[0].text);
      Logger.log('入力トークン: ' + data.usage.input_tokens);
      Logger.log('出力トークン: ' + data.usage.output_tokens);
      
      return '✅ Claude API接続成功';
    } else {
      Logger.log('❌ エラーレスポンス:');
      Logger.log(responseText);
      return '❌ Claude API接続失敗';
    }
    
  } catch (error) {
    Logger.log('❌ エラー: ' + error.message);
    return '❌ Claude API接続失敗: ' + error.message;
  }
}