/**
 * TaskManagement.gs
 * タスク管理システム
 * 
 * 機能:
 * - タスク管理シート作成
 * - AI提案からタスク登録
 * - カスタムタスク追加
 * - タスク完了処理
 * - リライト冷却期間管理
 * - リライト履歴自動連携
 * 
 * @version 1.0
 * @date 2025-12-03 (Day 18)
 */

// ============================================
// 定数定義
// ============================================

/**
 * タスク種別ごとの冷却期間（日数）
 */
const COOLING_PERIODS = {
  'タイトル変更': 90,           // 3ヶ月（頻繁な変更はNG）
  'H1変更': 90,                 // 3ヶ月（タイトルに準じる）
  'メタディスクリプション': 60,  // 2ヶ月
  'H2追加': 30,                 // 1ヶ月
  'H2変更': 30,
  'H3追加': 30,
  'H3変更': 30,
  '本文追加': 30,
  '本文修正': 30,
  'Q&A追加': 30,
  '画像追加': 30,
  '動画追加': 30,
  '内部リンク追加': 30,
  '内部リンク修正': 30,
  '構造化データ追加': 30,
  'その他': 30,
  'default': 30                 // デフォルト
};

/**
 * タスク種別と期待効果のマッピング
 */
const TASK_TYPE_EFFECTS = {
  'タイトル変更': { effect: 'CTR改善', priority: 1 },
  'H1変更': { effect: '順位改善', priority: 2 },
  'メタディスクリプション': { effect: 'CTR改善', priority: 1 },
  'H2追加': { effect: '順位改善', priority: 2 },
  'H2変更': { effect: '順位改善', priority: 2 },
  'H3追加': { effect: '構造改善', priority: 3 },
  'H3変更': { effect: '構造改善', priority: 3 },
  '本文追加': { effect: '滞在時間改善', priority: 3 },
  '本文修正': { effect: '滞在時間改善', priority: 3 },
  'Q&A追加': { effect: '滞在時間改善', priority: 3 },
  '画像追加': { effect: '滞在時間改善', priority: 3 },
  '動画追加': { effect: '滞在時間改善', priority: 3 },
  '内部リンク追加': { effect: '回遊率改善', priority: 4 },
  '内部リンク修正': { effect: '回遊率改善', priority: 4 },
  '構造化データ追加': { effect: 'リッチリザルト', priority: 3 },
  'その他': { effect: '総合改善', priority: 5 }
};

/**
 * ステータス定義
 */
const TASK_STATUS = {
  NOT_STARTED: '未着手',
  IN_PROGRESS: '進行中',
  COMPLETED: '完了',
  ON_HOLD: '保留',
  CANCELLED: 'キャンセル'
};

/**
 * タスクソース定義
 */
const TASK_SOURCE = {
  AI_SUGGESTION: 'AI提案',
  USER_ADDED: 'ユーザー追加'
};


// ============================================
// シート作成関数
// ============================================

/**
 * タスク管理シートを作成
 * 14列構成
 */
function createTaskManagementSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 既存シートがあれば削除確認
  let sheet = ss.getSheetByName('タスク管理');
  if (sheet) {
    Logger.log('タスク管理シートは既に存在します');
    return sheet;
  }
  
  // 新規シート作成
  sheet = ss.insertSheet('タスク管理');
  
  // ヘッダー設定
  const headers = [
    'task_id',           // A: タスクID（自動生成）
    'page_url',          // B: 対象ページURL
    'page_title',        // C: ページタイトル
    'task_type',         // D: タスク種別
    'task_detail',       // E: 具体的内容
    'source',            // F: 提案元（AI提案/ユーザー追加）
    'priority_rank',     // G: 推奨順位（1/2/3/-）
    'expected_effect',   // H: 期待効果
    'status',            // I: ステータス
    'created_date',      // J: 登録日
    'completed_date',    // K: 完了日
    'actual_change',     // L: 実際の変更内容
    'cooling_days',      // M: 冷却日数（自動設定）
    'notes'              // N: メモ
  ];
  
  // ヘッダー書き込み
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // ヘッダー書式設定
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  
  // 列幅設定
  const columnWidths = {
    1: 180,   // task_id
    2: 300,   // page_url
    3: 250,   // page_title
    4: 150,   // task_type
    5: 400,   // task_detail
    6: 100,   // source
    7: 80,    // priority_rank
    8: 120,   // expected_effect
    9: 80,    // status
    10: 100,  // created_date
    11: 100,  // completed_date
    12: 400,  // actual_change
    13: 80,   // cooling_days
    14: 200   // notes
  };
  
  Object.entries(columnWidths).forEach(([col, width]) => {
    sheet.setColumnWidth(parseInt(col), width);
  });
  
  // データ入力規則（ドロップダウン）
  // task_type
  const taskTypes = Object.keys(COOLING_PERIODS).filter(k => k !== 'default');
  const taskTypeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(taskTypes, true)
    .build();
  sheet.getRange('D2:D1000').setDataValidation(taskTypeRule);
  
  // source
  const sourceRule = SpreadsheetApp.newDataValidation()
    .requireValueInList([TASK_SOURCE.AI_SUGGESTION, TASK_SOURCE.USER_ADDED], true)
    .build();
  sheet.getRange('F2:F1000').setDataValidation(sourceRule);
  
  // status
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(Object.values(TASK_STATUS), true)
    .build();
  sheet.getRange('I2:I1000').setDataValidation(statusRule);
  
  // 行を固定
  sheet.setFrozenRows(1);
  
  Logger.log('タスク管理シートを作成しました');
  return sheet;
}


/**
 * 全タスク管理シートを一括作成（メニュー用）
 */
function setupTaskManagementSheets() {
  createTaskManagementSheet();
  Logger.log('タスク管理シートのセットアップが完了しました');
}


// ============================================
// タスク登録関数
// ============================================

/**
 * タスクIDを生成
 * @return {string} TASK_YYYYMMDD_XXX形式のID
 */
function generateTaskId() {
  const now = new Date();
  const dateStr = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyyMMdd');
  const timeStr = Utilities.formatDate(now, 'Asia/Tokyo', 'HHmmss');
  return `TASK_${dateStr}_${timeStr}`;
}


/**
 * AI提案からタスクを登録
 * @param {string} pageUrl - ページURL
 * @param {string} pageTitle - ページタイトル
 * @param {string} taskType - タスク種別
 * @param {string} taskDetail - 具体的内容
 * @param {number} priorityRank - 推奨順位（1-5, または0=なし）
 * @return {Object} 登録結果
 */
function createTaskFromAISuggestion(pageUrl, pageTitle, taskType, taskDetail, priorityRank) {
  return createTask({
    pageUrl: pageUrl,
    pageTitle: pageTitle,
    taskType: taskType,
    taskDetail: taskDetail,
    source: TASK_SOURCE.AI_SUGGESTION,
    priorityRank: priorityRank || 0
  });
}


/**
 * ユーザーが手動でタスクを追加
 * @param {string} pageUrl - ページURL
 * @param {string} pageTitle - ページタイトル
 * @param {string} taskType - タスク種別
 * @param {string} taskDetail - 具体的内容
 * @param {string} notes - メモ（任意）
 * @return {Object} 登録結果
 */
function createCustomTask(pageUrl, pageTitle, taskType, taskDetail, notes) {
  return createTask({
    pageUrl: pageUrl,
    pageTitle: pageTitle,
    taskType: taskType,
    taskDetail: taskDetail,
    source: TASK_SOURCE.USER_ADDED,
    priorityRank: 0,
    notes: notes || ''
  });
}


/**
 * タスクを登録（内部関数）
 * @param {Object} taskData - タスクデータ
 * @return {Object} 登録結果
 */
function createTask(taskData) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('タスク管理');
    if (!sheet) {
      throw new Error('タスク管理シートが見つかりません。先にシートを作成してください。');
    }
    
    const taskId = generateTaskId();
    const now = new Date();
    const coolingDays = COOLING_PERIODS[taskData.taskType] || COOLING_PERIODS.default;
    const effectInfo = TASK_TYPE_EFFECTS[taskData.taskType] || TASK_TYPE_EFFECTS['その他'];
    
    const rowData = [
      taskId,                                    // A: task_id
      taskData.pageUrl,                          // B: page_url
      taskData.pageTitle,                        // C: page_title
      taskData.taskType,                         // D: task_type
      taskData.taskDetail,                       // E: task_detail
      taskData.source,                           // F: source
      taskData.priorityRank || '-',              // G: priority_rank
      effectInfo.effect,                         // H: expected_effect
      TASK_STATUS.NOT_STARTED,                   // I: status
      now,                                       // J: created_date
      '',                                        // K: completed_date
      '',                                        // L: actual_change
      coolingDays,                               // M: cooling_days
      taskData.notes || ''                       // N: notes
    ];
    
    sheet.appendRow(rowData);
    
    Logger.log(`タスク登録完了: ${taskId} - ${taskData.taskType}`);
    
    return {
      success: true,
      taskId: taskId,
      message: `タスクを登録しました: ${taskData.taskType}`,
      coolingDays: coolingDays
    };
    
  } catch (error) {
    Logger.log(`タスク登録エラー: ${error.message}`);
    return {
      success: false,
      error: error.message
    };
  }
}


// ============================================
// タスク完了処理
// ============================================

/**
 * タスクを完了にする
 * - ステータスを「完了」に変更
 * - 完了日を記録
 * - リライト履歴に自動登録
 * - 冷却期間を開始
 * 
 * @param {string} taskId - タスクID
 * @param {string} actualChange - 実際に行った変更内容
 * @return {Object} 完了結果
 */
function completeTask(taskId, actualChange) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('タスク管理');
    if (!sheet) {
      throw new Error('タスク管理シートが見つかりません');
    }
    
    // タスクを検索
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const taskIdCol = headers.indexOf('task_id');
    const statusCol = headers.indexOf('status');
    const completedDateCol = headers.indexOf('completed_date');
    const actualChangeCol = headers.indexOf('actual_change');
    
    let taskRow = -1;
    let taskData = null;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][taskIdCol] === taskId) {
        taskRow = i + 1; // 1-indexed
        taskData = {};
        headers.forEach((header, idx) => {
          taskData[header] = data[i][idx];
        });
        break;
      }
    }
    
    if (taskRow === -1) {
      throw new Error(`タスクが見つかりません: ${taskId}`);
    }
    
    const now = new Date();
    
    // ステータスを完了に更新
    sheet.getRange(taskRow, statusCol + 1).setValue(TASK_STATUS.COMPLETED);
    sheet.getRange(taskRow, completedDateCol + 1).setValue(now);
    sheet.getRange(taskRow, actualChangeCol + 1).setValue(actualChange || taskData.task_detail);
    
    // リライト履歴に自動登録
    const historyResult = addToRewriteHistoryFromTask({
      pageUrl: taskData.page_url,
      taskType: taskData.task_type,
      changesSummary: actualChange || taskData.task_detail,
      aiSuggestion: taskData.source === TASK_SOURCE.AI_SUGGESTION ? taskData.task_detail : '',
      source: taskData.source,
      taskId: taskId
    });
    
    Logger.log(`タスク完了: ${taskId}`);
    
    return {
      success: true,
      taskId: taskId,
      message: 'タスクを完了しました',
      pageUrl: taskData.page_url,
      taskType: taskData.task_type,
      coolingDays: taskData.cooling_days,
      coolingEndDate: new Date(now.getTime() + taskData.cooling_days * 24 * 60 * 60 * 1000),
      historyRegistered: historyResult.success
    };
    
  } catch (error) {
    Logger.log(`タスク完了エラー: ${error.message}`);
    return {
      success: false,
      error: error.message
    };
  }
}


/**
 * リライト履歴シートに登録（動的ヘッダー対応版）
 * 既存シートのヘッダーを読み取り、対応する列にデータを挿入
 * 
 * @param {Object} data - リライトデータ
 * @return {Object} 登録結果
 */
function addToRewriteHistoryFromTask(data) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('リライト履歴');
    if (!sheet) {
      Logger.log('リライト履歴シートが見つかりません');
      return { success: false, error: 'シートが見つかりません' };
    }
    
    const now = new Date();
    const rewriteId = `RW_${Utilities.formatDate(now, 'Asia/Tokyo', 'yyyyMMdd_HHmmss')}`;
    
    // 現在の指標を取得（Before値）
    const beforeMetrics = getCurrentMetrics(data.pageUrl);
    
    // 既存シートのヘッダーを読み取り
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // データマッピング（ヘッダー名 → 値）
    const dataMap = {
      'rewrite_id': rewriteId,
      'page_url': data.pageUrl,
      'rewrite_date': now,
      'rewrite_type': data.taskType,
      'changes_summary': data.changesSummary,
      'changes_detail': '',  // タスク管理では個別記録のため空
      'ai_suggestion': data.aiSuggestion || '',
      'ai_suggested_count': data.source === 'AI提案' ? 1 : 0,
      'user_added_count': data.source === 'ユーザー追加' ? 1 : 0,
      'implemented_count': 1,
      'pending_count': 0,
      'not_needed_count': 0,
      'before_position': beforeMetrics.position || '',
      'before_ctr': beforeMetrics.ctr || '',
      'before_pv': beforeMetrics.pv || '',
      'before_bounce_rate': beforeMetrics.bounceRate || '',
      'before_cv': beforeMetrics.cv || '',
      'after_position': '',
      'after_ctr': '',
      'after_pv': '',
      'after_bounce_rate': '',
      'after_cv': '',
      'position_change': '',
      'ctr_change': '',
      'pv_change': '',
      'success_flag': '',
      'source': data.source || '',
      'task_id': data.taskId || '',
      'notes': data.notes || ''
    };
    
    // ヘッダーに対応した行データを作成
    const rowData = headers.map(header => {
      const normalizedHeader = header.toLowerCase().replace(/\s+/g, '_');
      return dataMap[normalizedHeader] !== undefined ? dataMap[normalizedHeader] : '';
    });
    
    sheet.appendRow(rowData);
    
    Logger.log(`リライト履歴登録: ${rewriteId}`);
    return { success: true, rewriteId: rewriteId };
    
  } catch (error) {
    Logger.log(`リライト履歴登録エラー: ${error.message}`);
    return { success: false, error: error.message };
  }
}


/**
 * 現在の指標を取得（Before値用）
 * @param {string} pageUrl - ページURL
 * @return {Object} 現在の指標
 */
function getCurrentMetrics(pageUrl) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('統合データ');
    if (!sheet) return {};
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const urlCol = headers.indexOf('page_url') !== -1 ? headers.indexOf('page_url') : headers.indexOf('url');
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][urlCol] === pageUrl || data[i][urlCol].includes(pageUrl)) {
        const row = {};
        headers.forEach((header, idx) => {
          row[header] = data[i][idx];
        });
        
        return {
          position: row['avg_position'] || row['position'] || row['average_position'] || '',
          ctr: row['ctr'] || row['click_through_rate'] || '',
          pv: row['pageviews'] || row['pv'] || row['page_views'] || '',
          bounceRate: row['bounce_rate'] || row['bounceRate'] || '',
          cv: row['conversions'] || row['cv'] || row['goal_completions'] || ''
        };
      }
    }
    
    return {};
  } catch (error) {
    Logger.log(`指標取得エラー: ${error.message}`);
    return {};
  }
}


// ============================================
// 冷却期間管理
// ============================================

/**
 * ページの冷却状態をチェック
 * @param {string} pageUrl - ページURL
 * @param {string} taskType - タスク種別（任意、指定時はその種別のみチェック）
 * @return {Object} 冷却状態
 */
function checkCoolingStatus(pageUrl, taskType) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('タスク管理');
    if (!sheet) {
      return { isCooling: false, message: 'タスク管理シートがありません' };
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const urlCol = headers.indexOf('page_url');
    const typeCol = headers.indexOf('task_type');
    const statusCol = headers.indexOf('status');
    const completedDateCol = headers.indexOf('completed_date');
    const coolingDaysCol = headers.indexOf('cooling_days');
    
    const today = new Date();
    const coolingInfo = {
      isCooling: false,
      coolingTasks: [],
      availableTasks: []
    };
    
    // 全タスク種別をチェック対象に
    const taskTypesToCheck = taskType ? [taskType] : Object.keys(COOLING_PERIODS).filter(k => k !== 'default');
    
    for (let i = 1; i < data.length; i++) {
      const rowUrl = data[i][urlCol];
      const rowType = data[i][typeCol];
      const rowStatus = data[i][statusCol];
      const completedDate = data[i][completedDateCol];
      const coolingDays = data[i][coolingDaysCol];
      
      // URLが一致し、完了済みのタスクをチェック
      if (rowUrl === pageUrl && rowStatus === TASK_STATUS.COMPLETED && completedDate) {
        const endDate = new Date(completedDate);
        endDate.setDate(endDate.getDate() + coolingDays);
        
        if (today < endDate) {
          const remainingDays = Math.ceil((endDate - today) / (1000 * 60 * 60 * 24));
          coolingInfo.coolingTasks.push({
            taskType: rowType,
            completedDate: completedDate,
            coolingDays: coolingDays,
            endDate: endDate,
            remainingDays: remainingDays
          });
        }
      }
    }
    
    // 冷却中のタスク種別を特定
    const coolingTypes = coolingInfo.coolingTasks.map(t => t.taskType);
    
    // 冷却中でないタスク種別を特定
    taskTypesToCheck.forEach(type => {
      if (!coolingTypes.includes(type)) {
        coolingInfo.availableTasks.push(type);
      }
    });
    
    coolingInfo.isCooling = coolingInfo.coolingTasks.length > 0;
    
    return coolingInfo;
    
  } catch (error) {
    Logger.log(`冷却状態チェックエラー: ${error.message}`);
    return { isCooling: false, error: error.message };
  }
}


/**
 * AI提案から除外すべきかを判定
 * @param {string} pageUrl - ページURL
 * @param {string} taskType - タスク種別
 * @return {Object} {shouldExclude, reason, remainingDays}
 */
function shouldExcludeFromSuggestion(pageUrl, taskType) {
  const coolingStatus = checkCoolingStatus(pageUrl, taskType);
  
  if (coolingStatus.error) {
    return { shouldExclude: false };
  }
  
  const coolingTask = coolingStatus.coolingTasks.find(t => t.taskType === taskType);
  
  if (coolingTask) {
    return {
      shouldExclude: true,
      reason: `${taskType}は冷却期間中`,
      remainingDays: coolingTask.remainingDays,
      endDate: coolingTask.endDate,
      coolingDays: coolingTask.coolingDays
    };
  }
  
  return { shouldExclude: false };
}


/**
 * ページの最終リライト日を取得
 * @param {string} pageUrl - ページURL
 * @return {Date|null} 最終完了日
 */
function getLastCompletedDate(pageUrl) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('タスク管理');
    if (!sheet) return null;
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const urlCol = headers.indexOf('page_url');
    const statusCol = headers.indexOf('status');
    const completedDateCol = headers.indexOf('completed_date');
    
    let lastDate = null;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][urlCol] === pageUrl && 
          data[i][statusCol] === TASK_STATUS.COMPLETED && 
          data[i][completedDateCol]) {
        const date = new Date(data[i][completedDateCol]);
        if (!lastDate || date > lastDate) {
          lastDate = date;
        }
      }
    }
    
    return lastDate;
    
  } catch (error) {
    Logger.log(`最終完了日取得エラー: ${error.message}`);
    return null;
  }
}


// ============================================
// タスク取得関数
// ============================================

/**
 * ページ別のタスク一覧を取得
 * @param {string} pageUrl - ページURL
 * @return {Array} タスク一覧
 */
function getTasksByPage(pageUrl) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('タスク管理');
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const urlCol = headers.indexOf('page_url');
    
    const tasks = [];
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][urlCol] === pageUrl) {
        const task = {};
        headers.forEach((header, idx) => {
          task[header] = data[i][idx];
        });
        tasks.push(task);
      }
    }
    
    return tasks;
    
  } catch (error) {
    Logger.log(`タスク取得エラー: ${error.message}`);
    return [];
  }
}


/**
 * 未完了タスク一覧を取得
 * @param {string} status - ステータスフィルター（任意）
 * @return {Array} タスク一覧
 */
function getPendingTasks(status) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('タスク管理');
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const statusCol = headers.indexOf('status');
    
    const tasks = [];
    const targetStatuses = status ? [status] : [TASK_STATUS.NOT_STARTED, TASK_STATUS.IN_PROGRESS];
    
    for (let i = 1; i < data.length; i++) {
      if (targetStatuses.includes(data[i][statusCol])) {
        const task = {};
        headers.forEach((header, idx) => {
          task[header] = data[i][idx];
        });
        tasks.push(task);
      }
    }
    
    return tasks;
    
  } catch (error) {
    Logger.log(`未完了タスク取得エラー: ${error.message}`);
    return [];
  }
}


/**
 * タスクのステータスを更新
 * @param {string} taskId - タスクID
 * @param {string} newStatus - 新しいステータス
 * @return {Object} 更新結果
 */
function updateTaskStatus(taskId, newStatus) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('タスク管理');
    if (!sheet) {
      throw new Error('タスク管理シートが見つかりません');
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const taskIdCol = headers.indexOf('task_id');
    const statusCol = headers.indexOf('status');
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][taskIdCol] === taskId) {
        sheet.getRange(i + 1, statusCol + 1).setValue(newStatus);
        Logger.log(`ステータス更新: ${taskId} → ${newStatus}`);
        return { success: true, taskId: taskId, newStatus: newStatus };
      }
    }
    
    throw new Error(`タスクが見つかりません: ${taskId}`);
    
  } catch (error) {
    Logger.log(`ステータス更新エラー: ${error.message}`);
    return { success: false, error: error.message };
  }
}


// ============================================
// AI提案との連携
// ============================================

/**
 * AI提案を冷却期間でフィルタリング
 * @param {string} pageUrl - ページURL
 * @param {Array} suggestions - AI提案リスト（{taskType, taskDetail, priorityRank}の配列）
 * @return {Object} フィルタリング結果
 */
function filterSuggestionsByCooling(pageUrl, suggestions) {
  const result = {
    available: [],
    excluded: []
  };
  
  suggestions.forEach(suggestion => {
    const exclusion = shouldExcludeFromSuggestion(pageUrl, suggestion.taskType);
    
    if (exclusion.shouldExclude) {
      result.excluded.push({
        ...suggestion,
        reason: exclusion.reason,
        remainingDays: exclusion.remainingDays,
        endDate: exclusion.endDate
      });
    } else {
      result.available.push(suggestion);
    }
  });
  
  return result;
}


/**
 * 冷却期間情報を含めた提案メッセージを生成
 * @param {Object} filterResult - filterSuggestionsByCooling()の結果
 * @return {string} ユーザー向けメッセージ
 */
function generateCoolingMessage(filterResult) {
  let message = '';
  
  if (filterResult.excluded.length > 0) {
    message += '\n\n⏳ **冷却期間中の項目**（自動提案から除外）:\n';
    filterResult.excluded.forEach(item => {
      const endDateStr = Utilities.formatDate(item.endDate, 'Asia/Tokyo', 'yyyy/MM/dd');
      message += `- ${item.taskType}: あと${item.remainingDays}日（${endDateStr}まで）\n`;
    });
  }
  
  return message;
}


// ============================================
// ユーティリティ関数
// ============================================

/**
 * タスク種別から冷却日数を取得
 * @param {string} taskType - タスク種別
 * @return {number} 冷却日数
 */
function getCoolingDays(taskType) {
  return COOLING_PERIODS[taskType] || COOLING_PERIODS.default;
}


/**
 * タスク管理のサマリーを取得
 * @return {Object} サマリー情報
 */
function getTaskSummary() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('タスク管理');
    if (!sheet) return { error: 'シートがありません' };
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const statusCol = headers.indexOf('status');
    
    const summary = {
      total: data.length - 1,
      notStarted: 0,
      inProgress: 0,
      completed: 0,
      onHold: 0,
      cancelled: 0
    };
    
    for (let i = 1; i < data.length; i++) {
      const status = data[i][statusCol];
      switch (status) {
        case TASK_STATUS.NOT_STARTED: summary.notStarted++; break;
        case TASK_STATUS.IN_PROGRESS: summary.inProgress++; break;
        case TASK_STATUS.COMPLETED: summary.completed++; break;
        case TASK_STATUS.ON_HOLD: summary.onHold++; break;
        case TASK_STATUS.CANCELLED: summary.cancelled++; break;
      }
    }
    
    return summary;
    
  } catch (error) {
    Logger.log(`サマリー取得エラー: ${error.message}`);
    return { error: error.message };
  }
}


// ============================================
// テスト関数
// ============================================

/**
 * タスク管理機能のテスト
 * Apps Scriptに追加後、この関数を実行してください
 */
function testTaskManagement() {
  Logger.log('=== タスク管理機能テスト開始 ===');
  Logger.log('実行日時: ' + new Date());
  
  let allTestsPassed = true;
  
  // 1. シート作成テスト
  Logger.log('\n--- 1. シート作成テスト ---');
  try {
    const sheet = createTaskManagementSheet();
    if (sheet) {
      Logger.log('✅ タスク管理シート作成成功');
    } else {
      Logger.log('⚠️ シートは既に存在（正常）');
    }
  } catch (error) {
    Logger.log('❌ シート作成失敗: ' + error.message);
    allTestsPassed = false;
  }
  
  // 2. AI提案からタスク登録テスト
  Logger.log('\n--- 2. AI提案タスク登録テスト ---');
  try {
    const aiTaskResult = createTaskFromAISuggestion(
      '/test-page-' + Date.now() + '/',
      'テストページ',
      'タイトル変更',
      '【テスト】タイトル案',
      1
    );
    if (aiTaskResult.success) {
      Logger.log('✅ AI提案タスク登録成功');
      Logger.log('  タスクID: ' + aiTaskResult.taskId);
      Logger.log('  冷却日数: ' + aiTaskResult.coolingDays);
    } else {
      Logger.log('❌ AI提案タスク登録失敗: ' + aiTaskResult.error);
      allTestsPassed = false;
    }
  } catch (error) {
    Logger.log('❌ AI提案タスク登録エラー: ' + error.message);
    allTestsPassed = false;
  }
  
  // 3. カスタムタスク登録テスト
  Logger.log('\n--- 3. カスタムタスク登録テスト ---');
  try {
    const customTaskResult = createCustomTask(
      '/test-page-custom/',
      'カスタムテストページ',
      '動画追加',
      'YouTube動画を埋め込む',
      'テストメモ'
    );
    if (customTaskResult.success) {
      Logger.log('✅ カスタムタスク登録成功');
      Logger.log('  タスクID: ' + customTaskResult.taskId);
    } else {
      Logger.log('❌ カスタムタスク登録失敗: ' + customTaskResult.error);
      allTestsPassed = false;
    }
  } catch (error) {
    Logger.log('❌ カスタムタスク登録エラー: ' + error.message);
    allTestsPassed = false;
  }
  
  // 4. 冷却期間チェックテスト
  Logger.log('\n--- 4. 冷却期間チェックテスト ---');
  try {
    const coolingStatus = checkCoolingStatus('/test-page-custom/');
    Logger.log('✅ 冷却期間チェック成功');
    Logger.log('  冷却中: ' + coolingStatus.isCooling);
    Logger.log('  冷却タスク数: ' + (coolingStatus.coolingTasks?.length || 0));
  } catch (error) {
    Logger.log('❌ 冷却期間チェックエラー: ' + error.message);
    allTestsPassed = false;
  }
  
  // 5. タスク一覧取得テスト
  Logger.log('\n--- 5. タスク一覧取得テスト ---');
  try {
    const tasks = getTasksByPage('/test-page-custom/');
    Logger.log('✅ タスク一覧取得成功');
    Logger.log('  タスク数: ' + tasks.length);
  } catch (error) {
    Logger.log('❌ タスク一覧取得エラー: ' + error.message);
    allTestsPassed = false;
  }
  
  // 6. 未完了タスク取得テスト
  Logger.log('\n--- 6. 未完了タスク取得テスト ---');
  try {
    const pendingTasks = getPendingTasks();
    Logger.log('✅ 未完了タスク取得成功');
    Logger.log('  未完了タスク数: ' + pendingTasks.length);
  } catch (error) {
    Logger.log('❌ 未完了タスク取得エラー: ' + error.message);
    allTestsPassed = false;
  }
  
  // 7. サマリー取得テスト
  Logger.log('\n--- 7. サマリー取得テスト ---');
  try {
    const summary = getTaskSummary();
    Logger.log('✅ サマリー取得成功');
    Logger.log('  総タスク数: ' + summary.total);
    Logger.log('  未着手: ' + summary.notStarted);
    Logger.log('  進行中: ' + summary.inProgress);
    Logger.log('  完了: ' + summary.completed);
  } catch (error) {
    Logger.log('❌ サマリー取得エラー: ' + error.message);
    allTestsPassed = false;
  }
  
  // 8. 冷却日数取得テスト
  Logger.log('\n--- 8. 冷却日数設定テスト ---');
  try {
    const titleCooling = getCoolingDays('タイトル変更');
    const h2Cooling = getCoolingDays('H2追加');
    const defaultCooling = getCoolingDays('unknown_type');
    
    Logger.log('✅ 冷却日数取得成功');
    Logger.log('  タイトル変更: ' + titleCooling + '日');
    Logger.log('  H2追加: ' + h2Cooling + '日');
    Logger.log('  デフォルト: ' + defaultCooling + '日');
    
    if (titleCooling === 90 && h2Cooling === 30) {
      Logger.log('✅ 冷却日数設定は正しい');
    } else {
      Logger.log('⚠️ 冷却日数設定を確認してください');
    }
  } catch (error) {
    Logger.log('❌ 冷却日数取得エラー: ' + error.message);
    allTestsPassed = false;
  }
  
  // 結果サマリー
  Logger.log('\n========================================');
  if (allTestsPassed) {
    Logger.log('🎉 全テスト成功！');
  } else {
    Logger.log('⚠️ 一部テストが失敗しました。上記のログを確認してください。');
  }
  Logger.log('========================================');
  
  return allTestsPassed;
}


/**
 * タスク完了フローのテスト（注意: 実際にデータが変更されます）
 * 本番環境では実行に注意してください
 */
function testCompleteTaskFlow() {
  Logger.log('=== タスク完了フローテスト開始 ===');
  Logger.log('⚠️ このテストはタスクとリライト履歴にデータを追加します');
  
  // 1. テスト用タスクを作成
  const testUrl = '/complete-test-' + Date.now() + '/';
  const taskResult = createCustomTask(
    testUrl,
    'タスク完了テスト',
    '本文追加',
    'テスト用コンテンツ追加',
    ''
  );
  
  if (!taskResult.success) {
    Logger.log('❌ テストタスク作成失敗');
    return false;
  }
  
  Logger.log('✅ テストタスク作成: ' + taskResult.taskId);
  
  // 2. タスクを完了にする
  const completeResult = completeTask(taskResult.taskId, '実際にテストコンテンツを追加しました');
  
  if (completeResult.success) {
    Logger.log('✅ タスク完了成功');
    Logger.log('  冷却期間: ' + completeResult.coolingDays + '日');
    Logger.log('  冷却終了日: ' + completeResult.coolingEndDate);
    Logger.log('  リライト履歴登録: ' + (completeResult.historyRegistered ? '成功' : '失敗'));
  } else {
    Logger.log('❌ タスク完了失敗: ' + completeResult.error);
    return false;
  }
  
  // 3. 冷却状態を確認
  const coolingStatus = checkCoolingStatus(testUrl);
  Logger.log('冷却状態確認:');
  Logger.log('  冷却中: ' + coolingStatus.isCooling);
  
  Logger.log('=== タスク完了フローテスト完了 ===');
  return true;
}


/**
 * リライト履歴シートの構造を確認するテスト
 */
function testRewriteHistoryStructure() {
  Logger.log('=== リライト履歴シート構造確認 ===');
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('リライト履歴');
  
  if (!sheet) {
    Logger.log('❌ リライト履歴シートが見つかりません');
    Logger.log('→ 先にリライト履歴シートを作成してください');
    return;
  }
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  Logger.log('列数: ' + headers.length);
  Logger.log('ヘッダー一覧:');
  headers.forEach((header, idx) => {
    Logger.log('  ' + (idx + 1) + ': ' + header);
  });
  
  // 必要な列が存在するかチェック
  const requiredColumns = ['rewrite_id', 'page_url', 'rewrite_date', 'rewrite_type'];
  const missingColumns = requiredColumns.filter(col => 
    !headers.some(h => h.toLowerCase().replace(/\s+/g, '_') === col)
  );
  
  if (missingColumns.length > 0) {
    Logger.log('⚠️ 不足している列: ' + missingColumns.join(', '));
  } else {
    Logger.log('✅ 必須列はすべて存在します');
  }
}
