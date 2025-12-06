/**
 * 保存されているAPIキーを確認（セキュリティチェック用）
 */
function checkStoredAPIKey() {
  const key = PropertiesService.getScriptProperties()
    .getProperty('CLAUDE_API_KEY');
  
  if (!key) {
    Logger.log('❌ APIキーが保存されていません');
    return;
  }
  
  // セキュリティのため、最初と最後の15文字のみ表示
  const keyStart = key.substring(0, 15);
  const keyEnd = key.substring(key.length - 15);
  const keyLength = key.length;
  
  Logger.log('=== APIキー確認 ===');
  Logger.log('先頭15文字: ' + keyStart);
  Logger.log('末尾15文字: ' + keyEnd);
  Logger.log('キー長: ' + keyLength + '文字');
  Logger.log('形式: ' + (keyStart.startsWith('sk-ant-api03-') ? '✅ 正常' : '⚠️ 異常'));
  
  // 新旧キーの判別
  if (keyStart.includes('E3LatW224')) {
    Logger.log('⚠️ 警告: これは古いAPIキーです！新しいキーに更新してください');
  } else {
    Logger.log('✅ 新しいAPIキーが設定されています');
  }
}