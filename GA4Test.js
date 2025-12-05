function testGA4Connection() {
  const propertyId = 'properties/388689745';
  
  try {
    const request = {
      dateRanges: [
        {
          startDate: '7daysAgo',
          endDate: 'yesterday'
        }
      ],
      dimensions: [
        { name: 'pagePath' }
      ],
      metrics: [
        { name: 'screenPageViews' }
      ],
      limit: 10
    };
    
    const response = AnalyticsData.Properties.runReport(request, propertyId);
    
    Logger.log('GA4接続成功！');
    Logger.log('取得行数: ' + (response.rows ? response.rows.length : 0));
    
    if (response.rows && response.rows.length > 0) {
      Logger.log('サンプルデータ:');
      Logger.log(response.rows[0]);
    }
    
    return '✅ GA4 API接続成功';
    
  } catch (error) {
    Logger.log('❌ エラー: ' + error.message);
    return '❌ GA4 API接続失敗: ' + error.message;
  }
}
