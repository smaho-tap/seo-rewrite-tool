/**
 * CSV取り込み機能
 * GyronSEO専用（将来的にカラムマッピング機能追加予定）
 */

/**
 * ========================================
 * 設定（将来的にカラムマッピングUIで設定可能にする）
 * ========================================
 */
const CSV_CONFIG = {
  // GyronSEO用の設定
  gyronSEO: {
    name: 'GyronSEO',
    encoding: 'UTF-8',
    columns: {
      keyword: 0,      // A列: キーワード
      url: 1,          // B列: URL
      lpCount: 2,      // C列: LP履歴数
      site: 3,         // D列: サイト
      group: 4,        // E列: グループ
      area: 5,         // F列: 検索エリア名
      tag: 6,          // G列: タグ（記事タイトル）
      rankStart: 7     // H列以降: 日付ごとの順位
    },
    siteUrl: 'https://smaho-tap.com'  // URLからパスを抽出する際に使用
  }
};

/**
 * ========================================
 * メイン関数
 * ========================================
 */

/**
 * Google DriveのCSVファイルをインポート
 * @param {string} fileId - Google DriveのファイルID
 * @param {string} toolType - ツール種別（デフォルト: 'gyronSEO'）
 * @return {Object} result - インポート結果
 */
function importCSVFromDrive(fileId, toolType = 'gyronSEO') {
  try {
    Logger.log('=== CSV取り込み開始 ===');
    Logger.log(`ファイルID: ${fileId}`);
    Logger.log(`ツール種別: ${toolType}`);
    
    // ファイル取得
    const file = DriveApp.getFileById(fileId);
    const fileName = file.getName();
    Logger.log(`ファイル名: ${fileName}`);
    
    // CSV読み込み
    const csvContent = file.getBlob().getDataAsString('UTF-8');
    
    // パース＆インポート実行
    const result = parseAndImportCSV(csvContent, toolType);
    
    Logger.log('=== CSV取り込み完了 ===');
    return result;
    
  } catch (error) {
    Logger.log(`CSV取り込みエラー: ${error.message}`);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * スプレッドシートからデータをインポート
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名（オプション、デフォルトは最初のシート）
 * @param {string} toolType - ツール種別
 * @return {Object} result - インポート結果
 */
function importFromSpreadsheet(spreadsheetId, sheetName = null, toolType = 'gyronSEO') {
  try {
    Logger.log('=== スプレッドシートから取り込み開始 ===');
    Logger.log(`スプレッドシートID: ${spreadsheetId}`);
    
    // スプレッドシート取得
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheet = sheetName ? ss.getSheetByName(sheetName) : ss.getSheets()[0];
    
    if (!sheet) {
      throw new Error('シートが見つかりません');
    }
    
    Logger.log(`シート名: ${sheet.getName()}`);
    
    // データ取得
    const data = sheet.getDataRange().getValues();
    
    // 設定取得
    const config = CSV_CONFIG[toolType];
    if (!config) {
      throw new Error(`未対応のツール種別: ${toolType}`);
    }
    
    Logger.log(`総行数: ${data.length}`);
    
    // ヘッダー行から日付列を抽出
    const headers = data[0];
    const dateColumns = extractDateColumns(headers, config.columns.rankStart);
    Logger.log(`日付列数: ${dateColumns.length}`);
    
    // データ行を処理
    const dataRows = data.slice(1);
    const processedData = [];
    let skippedCount = 0;
    
    for (const row of dataRows) {
      const processed = processGyronSEORow(row, config, dateColumns);
      if (processed) {
        processedData.push(processed);
      } else {
        skippedCount++;
      }
    }
    
    Logger.log(`処理済み: ${processedData.length}件`);
    Logger.log(`スキップ: ${skippedCount}件`);
    
    // シートに書き込み
    writeToGyronSEOSheet(processedData);
    
    Logger.log('=== スプレッドシートから取り込み完了 ===');
    
    return {
      success: true,
      totalRows: data.length - 1,
      importedRows: processedData.length,
      skippedRows: skippedCount,
      message: `${processedData.length}件のキーワードデータをインポートしました`
    };
    
  } catch (error) {
    Logger.log(`取り込みエラー: ${error.message}`);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * CSVテキストを直接インポート（スプレッドシートからコピペする場合）
 * @param {string} csvContent - CSVテキスト
 * @param {string} toolType - ツール種別
 * @return {Object} result - インポート結果
 */
function importCSVFromText(csvContent, toolType = 'gyronSEO') {
  try {
    Logger.log('=== CSV取り込み開始（テキスト）===');
    return parseAndImportCSV(csvContent, toolType);
  } catch (error) {
    Logger.log(`CSV取り込みエラー: ${error.message}`);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * CSVをパースしてインポート
 * @param {string} csvContent - CSVテキスト
 * @param {string} toolType - ツール種別
 * @return {Object} result - インポート結果
 */
function parseAndImportCSV(csvContent, toolType) {
  // 設定取得
  const config = CSV_CONFIG[toolType];
  if (!config) {
    throw new Error(`未対応のツール種別: ${toolType}`);
  }
  
  // CSVパース
  const rows = Utilities.parseCsv(csvContent);
  Logger.log(`総行数: ${rows.length}`);
  
  // ヘッダー行（1行目）から日付列を抽出
  const headers = rows[0];
  const dateColumns = extractDateColumns(headers, config.columns.rankStart);
  Logger.log(`日付列数: ${dateColumns.length}`);
  
  // データ行を処理
  const dataRows = rows.slice(1); // ヘッダー行を除く
  const processedData = [];
  let skippedCount = 0;
  
  for (const row of dataRows) {
    const processed = processGyronSEORow(row, config, dateColumns);
    if (processed) {
      processedData.push(processed);
    } else {
      skippedCount++;
    }
  }
  
  Logger.log(`処理済み: ${processedData.length}件`);
  Logger.log(`スキップ: ${skippedCount}件`);
  
  // シートに書き込み
  writeToGyronSEOSheet(processedData);
  
  return {
    success: true,
    totalRows: rows.length - 1,
    importedRows: processedData.length,
    skippedRows: skippedCount,
    message: `${processedData.length}件のキーワードデータをインポートしました`
  };
}

/**
 * ヘッダー行から日付列を抽出
 * @param {Array} headers - ヘッダー行
 * @param {number} startIndex - 日付列の開始インデックス
 * @return {Array} dateColumns - 日付列情報
 */
function extractDateColumns(headers, startIndex) {
  const dateColumns = [];
  const datePattern = /^\d{4}-\d{2}-\d{2}$/;
  
  for (let i = startIndex; i < headers.length; i++) {
    const header = headers[i];
    let dateStr = '';
    
    // Date型の場合
    if (header instanceof Date) {
      dateStr = Utilities.formatDate(header, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    } 
    // 文字列の場合
    else if (typeof header === 'string') {
      dateStr = header;
    }
    // 数値（シリアル値）の場合
    else if (typeof header === 'number' && header > 40000) {
      // Excelのシリアル値をDateに変換
      const date = new Date((header - 25569) * 86400 * 1000);
      dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
    else {
      dateStr = String(header);
    }
    
    if (datePattern.test(dateStr)) {
      dateColumns.push({
        index: i,
        date: dateStr
      });
    }
  }
  
  // 日付順にソート（降順：最新が先頭）
  dateColumns.sort((a, b) => b.date.localeCompare(a.date));
  
  return dateColumns;
}

/**
 * GyronSEOの1行を処理
 * @param {Array} row - CSVの1行
 * @param {Object} config - 設定
 * @param {Array} dateColumns - 日付列情報
 * @return {Object|null} processedRow - 処理済みデータ（URLがない場合はnull）
 */
function processGyronSEORow(row, config, dateColumns) {
  const cols = config.columns;
  
  const keyword = row[cols.keyword] || '';
  const fullUrl = row[cols.url] || '';
  const group = row[cols.group] || '';
  const tag = row[cols.tag] || '';
  
  // キーワードが空の場合はスキップ
  if (!keyword.toString().trim()) {
    return null;
  }
  
  // URLをパスに変換
  const urlPath = extractUrlPath(fullUrl, config.siteUrl);
  
  // 順位データを取得（最新、7日前、30日前、90日前）
  const latestPosition = getPositionValue(row, dateColumns, 0);
  const position7dAgo = getPositionValue(row, dateColumns, 7);
  const position30dAgo = getPositionValue(row, dateColumns, 30);
  const position90dAgo = getPositionValue(row, dateColumns, 90);
  
  // トレンド計算
  const trend = calculateTrend(latestPosition, position7dAgo);
  
  return {
    keyword: keyword,
    url: urlPath,
    group: group,
    tag: tag,
    latestPosition: latestPosition,
    position7dAgo: position7dAgo,
    position30dAgo: position30dAgo,
    position90dAgo: position90dAgo,
    trend: trend
  };
}

/**
 * 完全URLからパスを抽出
 * @param {string} fullUrl - 完全URL
 * @param {string} siteUrl - サイトURL
 * @return {string} path - パス部分（空の場合は空文字）
 */
function extractUrlPath(fullUrl, siteUrl) {
  if (!fullUrl || fullUrl.toString().trim() === '') {
    return '';  // URL空欄 = 圏外
  }
  
  // サイトURLを除去してパス部分を抽出
  let path = fullUrl.toString().replace(siteUrl, '');
  
  // 先頭に/がなければ追加
  if (path && !path.startsWith('/')) {
    path = '/' + path;
  }
  
  // 末尾の/を除去
  path = path.replace(/\/$/, '');
  
  return path;
}

/**
 * 指定日数前の順位を取得
 * @param {Array} row - CSVの1行
 * @param {Array} dateColumns - 日付列情報（最新順）
 * @param {number} daysAgo - 何日前のデータか
 * @return {number|string} position - 順位（圏外の場合は'圏外'）
 */
function getPositionValue(row, dateColumns, daysAgo) {
  // 対象日付のインデックスを探す
  const targetIndex = Math.min(daysAgo, dateColumns.length - 1);
  
  if (targetIndex < 0 || targetIndex >= dateColumns.length) {
    return '';
  }
  
  const colInfo = dateColumns[targetIndex];
  const value = row[colInfo.index];
  
  return parsePositionValue(value);
}

/**
 * 順位値をパース
 * @param {string} value - 順位値
 * @return {number|string} position
 */
function parsePositionValue(value) {
  if (value === null || value === undefined || value === '') {
    return '';  // データなし
  }
  
  const strValue = value.toString();
  
  if (strValue === '圏外') {
    return 101;  // 圏外は101として扱う
  }
  
  const num = parseInt(strValue, 10);
  if (isNaN(num)) {
    return '';
  }
  
  return num;
}

/**
 * トレンドを計算
 * @param {number|string} latest - 最新順位
 * @param {number|string} before - 比較対象順位
 * @return {string} trend - ↑/↓/→/--
 */
function calculateTrend(latest, before) {
  if (latest === '' || before === '' || latest === 101 || before === 101) {
    return '--';
  }
  
  const diff = before - latest;  // 順位は低い方が良いので、beforeからlatestを引く
  
  if (diff > 2) {
    return '↑';  // 3位以上上昇
  } else if (diff < -2) {
    return '↓';  // 3位以上下降
  } else {
    return '→';  // 横ばい
  }
}

/**
 * GyronSEO_RAWシートに書き込み
 * @param {Array} data - 処理済みデータ
 */
function writeToGyronSEOSheet(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('GyronSEO_RAW');
  
  if (!sheet) {
    throw new Error('GyronSEO_RAWシートが見つかりません');
  }
  
  // 既存データをクリア（ヘッダー行を残す）
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }
  
  if (data.length === 0) {
    Logger.log('書き込むデータがありません');
    return;
  }
  
  // データを2次元配列に変換
  const now = new Date();
  const values = data.map(row => [
    row.keyword,
    row.url,
    row.group,
    row.tag,
    row.latestPosition,
    row.position7dAgo,
    row.position30dAgo,
    row.position90dAgo,
    row.trend,
    now,  // import_date
    now   // last_updated
  ]);
  
  // 一括書き込み
  sheet.getRange(2, 1, values.length, values[0].length).setValues(values);
  
  Logger.log(`GyronSEO_RAWシートに${values.length}行書き込み完了`);
}

/**
 * ========================================
 * ユーティリティ関数
 * ========================================
 */

/**
 * 最新のCSVファイルIDを設定シートから取得
 * @return {string} fileId
 */
function getLatestCSVFileId() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingSheet = ss.getSheetByName('設定・マスタ');
  
  if (!settingSheet) {
    return null;
  }
  
  const data = settingSheet.getDataRange().getValues();
  
  for (const row of data) {
    if (row[0] === 'GYRONSEO_CSV_FILE_ID') {
      return row[1];
    }
  }
  
  return null;
}

/**
 * CSVファイルIDを設定シートに保存
 * @param {string} fileId - ファイルID
 */
function saveCSVFileId(fileId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingSheet = ss.getSheetByName('設定・マスタ');
  
  if (!settingSheet) {
    throw new Error('設定・マスタシートが見つかりません');
  }
  
  const data = settingSheet.getDataRange().getValues();
  let found = false;
  
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === 'GYRONSEO_CSV_FILE_ID') {
      settingSheet.getRange(i + 1, 2).setValue(fileId);
      found = true;
      break;
    }
  }
  
  if (!found) {
    // 新規追加
    settingSheet.appendRow(['GYRONSEO_CSV_FILE_ID', fileId, 'GyronSEO CSVファイルID', new Date()]);
  }
  
  Logger.log(`CSVファイルID保存完了: ${fileId}`);
}

/**
 * ========================================
 * 統合データとの連携
 * ========================================
 */

/**
 * GyronSEOデータを統合データシートにマージ
 * （Day 7で実装予定）
 */
function mergeGyronSEOToIntegrated() {
  Logger.log('=== GyronSEOデータ統合開始 ===');
  
  // TODO: Day 7で実装
  // 1. GyronSEO_RAWからデータ取得
  // 2. 統合データシートのURLとマッチング
  // 3. ターゲットKW、順位情報を追加
  
  Logger.log('=== GyronSEOデータ統合完了 ===');
}

/**
 * ========================================
 * 実行用関数
 * ========================================
 */

/**
 * 今すぐインポート実行（ファイルIDを直接指定）
 */
function runImportNow() {
  const fileId = '1CbYIIv_Slja0XVhW5jYHXfbY-A7xkFIQRAtlAnoDaco';
  const result = importFromSpreadsheet(fileId);
  Logger.log('=== 結果 ===');
  Logger.log(JSON.stringify(result, null, 2));
}

/**
 * テスト: DriveからCSVインポート
 * 使用方法: ファイルIDを指定して実行
 */
function testImportCSVFromDrive() {
  const fileId = '1CbYIIv_Slja0XVhW5jYHXfbY-A7xkFIQRAtlAnoDaco';
  
  const result = importCSVFromDrive(fileId);
  Logger.log('=== テスト結果 ===');
  Logger.log(JSON.stringify(result, null, 2));
}

/**
 * テスト: 設定シートからファイルIDを取得してインポート
 */
function testImportFromSavedFileId() {
  const fileId = getLatestCSVFileId();
  
  if (!fileId) {
    Logger.log('ファイルIDが設定されていません');
    Logger.log('設定・マスタシートに GYRONSEO_CSV_FILE_ID を追加してください');
    return;
  }
  
  const result = importCSVFromDrive(fileId);
  Logger.log('=== テスト結果 ===');
  Logger.log(JSON.stringify(result, null, 2));
}

/**
 * クイックインポート: ファイルIDを指定して即実行
 * @param {string} fileId - Google DriveのファイルID
 */
function quickImport(fileId) {
  if (!fileId) {
    Logger.log('ファイルIDを指定してください');
    Logger.log('例: quickImport("1ABC123xyz...")');
    return;
  }
  
  // ファイルIDを保存
  saveCSVFileId(fileId);
  
  // インポート実行
  const result = importCSVFromDrive(fileId);
  
  Logger.log('=== インポート結果 ===');
  Logger.log(JSON.stringify(result, null, 2));
  
  return result;
}