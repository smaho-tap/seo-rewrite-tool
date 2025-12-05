function testGSCConnection() {
  const siteUrl = 'https://smaho-tap.com/';
  
  try {
    // OAuth2トークンを取得
    const token = ScriptApp.getOAuthToken();
    
    // 今日から30日前の日付を取得
    const endDate = new Date();
    endDate.setDate(endDate.getDate() - 1); // 昨日
    const startDate = new Date();
    startDate.setDate(startDate.getDate() - 30); // 30日前
    
    const requestBody = {
      startDate: Utilities.formatDate(startDate, 'Asia/Tokyo', 'yyyy-MM-dd'),
      endDate: Utilities.formatDate(endDate, 'Asia/Tokyo', 'yyyy-MM-dd'),
      dimensions: ['page'],
      rowLimit: 10
    };
    
    // API URL（サイトURLをエンコード）
    const encodedSiteUrl = encodeURIComponent(siteUrl);
    const apiUrl = `https://www.googleapis.com/webmasters/v3/sites/${encodedSiteUrl}/searchAnalytics/query`;
    
    const options = {
      method: 'post',
      headers: {
        'Authorization': 'Bearer ' + token,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(requestBody),
      muteHttpExceptions: true
    };
    
    Logger.log('リクエストURL: ' + apiUrl);
    Logger.log('リクエストボディ:');
    Logger.log(requestBody);
    
    const response = UrlFetchApp.fetch(apiUrl, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    Logger.log('レスポンスコード: ' + responseCode);
    
    if (responseCode === 200) {
      const data = JSON.parse(responseText);
      Logger.log('GSC接続成功！');
      Logger.log('取得行数: ' + (data.rows ? data.rows.length : 0));
      
      if (data.rows && data.rows.length > 0) {
        Logger.log('サンプルデータ:');
        Logger.log(data.rows[0]);
      }
      
      return '✅ Search Console API接続成功';
    } else {
      Logger.log('エラーレスポンス:');
      Logger.log(responseText);
      return '❌ Search Console API接続失敗: ' + responseCode;
    }
    
  } catch (error) {
    Logger.log('❌ エラー: ' + error.message);
    Logger.log('エラー詳細:');
    Logger.log(error);
    return '❌ Search Console API接続失敗: ' + error.message;
  }
}
