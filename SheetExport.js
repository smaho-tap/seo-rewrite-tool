/**
 * SheetExport.gs v2
 * スプレッドシート構造を「システム概要」シートに出力
 */

function exportSpreadsheetStructure() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  
  // 出力用シート作成または取得
  let outputSheet = ss.getSheetByName('システム概要');
  if (!outputSheet) {
    outputSheet = ss.insertSheet('システム概要');
  } else {
    outputSheet.clear();
  }
  
  const output = [];
  const timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
  
  // ヘッダー情報
  output.push(['# スプレッドシート構造', '', '', '']);
  output.push(['最終更新', timestamp, '', '']);
  output.push(['シート数', sheets.length, '', '']);
  output.push(['', '', '', '']);
  
  // サマリーテーブル
  output.push(['シート名', '行数', '列数', '用途']);
  
  const sheetPurposes = {
    'GA4_RAW': 'GA4生データ',
    'GSC_RAW': 'Search Console生データ',
    'GyronSEO_RAW': '順位データ',
    '検索ボリューム_RAW': '検索ボリュームデータ',
    'Clarity_RAW': 'UX指標',
    '競合分析': 'キーワード別競合DA',
    'DA履歴': 'ドメインDA履歴管理',
    'ターゲットKW分析': 'ターゲットKW設定',
    'API使用ログ': 'コスト管理',
    '統合データ': 'メインデータ（5軸スコア）',
    'リライト履歴': '効果測定',
    'クエリ分析': 'GSCクエリ分析',
    'KW除外候補': '除外KW管理',
    'イベント分析': 'GA4イベント',
    'GTM分析': 'GTMスクロール分析',
    '議事録': 'AI相談記録',
    '設定・マスタ': '各種設定値',
    '分析ログ': '分析実行履歴',
    'AIO順位履歴': 'AI Overview順位追跡',
    'タスク管理': 'リライトタスク管理',
    'システム概要': 'この出力シート'
  };
  
  for (const sheet of sheets) {
    const name = sheet.getName();
    if (name === 'システム概要') continue;
    const purpose = sheetPurposes[name] || '';
    output.push([name, sheet.getLastRow(), sheet.getLastColumn(), purpose]);
  }
  
  output.push(['', '', '', '']);
  output.push(['--- 各シート詳細 ---', '', '', '']);
  output.push(['', '', '', '']);
  
  // 各シートの詳細
  for (const sheet of sheets) {
    const name = sheet.getName();
    if (name === 'システム概要') continue;
    
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    output.push([`## ${name}`, '', '', '']);
    output.push(['行数', lastRow, '列数', lastCol]);
    
    if (lastRow > 0 && lastCol > 0) {
      // ヘッダー取得
      const headers = sheet.getRange(1, 1, 1, Math.min(lastCol, 20)).getValues()[0];
      output.push(['ヘッダー:', headers.slice(0, 10).join(', '), '', '']);
      
      if (lastCol > 10) {
        output.push(['（続き）:', headers.slice(10, 20).join(', '), '', '']);
      }
    }
    output.push(['', '', '', '']);
  }
  
  // シートに書き込み
  outputSheet.getRange(1, 1, output.length, 4).setValues(output);
  
  // 列幅調整
  outputSheet.setColumnWidth(1, 200);
  outputSheet.setColumnWidth(2, 300);
  outputSheet.setColumnWidth(3, 100);
  outputSheet.setColumnWidth(4, 200);
  
  Logger.log('='.repeat(50));
  Logger.log('エクスポート完了！');
  Logger.log('「システム概要」シートを確認してください');
  Logger.log('='.repeat(50));
}