/**
 * 議事録自動生成機能
 * チャットセッションの内容を構造化して保存
 */

/**
 * 議事録を生成する
 * @param {Array} conversationHistory - 会話履歴配列
 * @param {string} sessionId - セッションID（オプション、なければ自動生成）
 * @return {Object} result - 生成結果
 */
function generateMinutes(conversationHistory, sessionId) {
  try {
    Logger.log('=== 議事録生成開始 ===');
    
    // セッションID生成（未指定の場合）
    if (!sessionId) {
      sessionId = generateSessionId();
    }
    Logger.log(`セッションID: ${sessionId}`);
    
    // 会話履歴が空の場合
    if (!conversationHistory || conversationHistory.length === 0) {
      throw new Error('会話履歴がありません');
    }
    
    Logger.log(`会話数: ${conversationHistory.length}件`);
    
    // 会話履歴をテキストに変換
    const conversationText = formatConversationForMinutes(conversationHistory);
    
    // Claude APIで議事録生成
    const minutesContent = callClaudeForMinutes(conversationText);
    
    if (!minutesContent) {
      throw new Error('議事録の生成に失敗しました');
    }
    
    Logger.log('Claude APIから議事録取得成功');
    
    // 議事録をパース
    const parsedMinutes = parseMinutesContent(minutesContent);
    
    // 参照ページを抽出
    const referencedPages = extractReferencedPages(conversationHistory);
    
    // 議事録シートに保存
    const now = new Date();
    saveMinutes({
      session_id: sessionId,
      session_date: now,
      session_start: now, // 実際の開始時刻があれば置き換え
      session_end: now,
      summary: parsedMinutes.summary || '',
      findings: parsedMinutes.findings || '',
      decisions: parsedMinutes.decisions || '',
      action_items: parsedMinutes.action_items || '',
      referenced_pages: referencedPages.join(', '),
      referenced_data: '',
      tags: parsedMinutes.tags || '',
      full_conversation: JSON.stringify(conversationHistory)
    });
    
    Logger.log('=== 議事録生成完了 ===');
    
    return {
      success: true,
      sessionId: sessionId,
      minutes: parsedMinutes
    };
    
  } catch (error) {
    Logger.log(`議事録生成エラー: ${error.message}`);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * セッションIDを生成
 * @return {string} sessionId
 */
function generateSessionId() {
  const date = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  const random = Math.random().toString(36).substring(2, 6).toUpperCase();
  return `SESSION_${date}_${random}`;
}

/**
 * 会話履歴を議事録用テキストに変換
 * @param {Array} conversationHistory - 会話履歴
 * @return {string} formattedText
 */
function formatConversationForMinutes(conversationHistory) {
  let text = '';
  
  conversationHistory.forEach((msg, idx) => {
    const role = msg.role === 'user' ? 'ユーザー' : 'AI';
    text += `【${role}】\n${msg.content}\n\n`;
  });
  
  return text;
}

/**
 * Claude APIで議事録を生成
 * @param {string} conversationText - 会話テキスト
 * @return {string} minutesContent
 */
function callClaudeForMinutes(conversationText) {
  const systemPrompt = buildMinutesSystemPrompt();
  const userPrompt = buildMinutesUserPrompt(conversationText);
  
  // ClaudeAPI.gsの関数を呼び出し
  return callClaudeAPI(userPrompt, systemPrompt);
}

/**
 * 議事録生成用システムプロンプト
 * @return {string} systemPrompt
 */
function buildMinutesSystemPrompt() {
  return `あなたはSEO相談セッションの議事録を作成する専門家です。

会話内容を分析し、以下の形式で構造化された議事録を作成してください。

【出力形式】必ず以下のJSON形式で出力してください。

{
  "summary": "相談内容の要約（2-3文）",
  "findings": "発見した問題点・気づき（箇条書き、改行区切り）",
  "decisions": "決定事項（箇条書き、改行区切り）",
  "action_items": "次のアクション（箇条書き、改行区切り）",
  "tags": "関連タグ（カンマ区切り、例: #リライト, #タイトル改善）"
}

【注意事項】
- 必ず有効なJSON形式で出力
- 各項目は簡潔にまとめる
- 具体的なページURLやデータがあれば含める
- アクションアイテムは実行可能な形で記載`;
}

/**
 * 議事録生成用ユーザープロンプト
 * @param {string} conversationText - 会話テキスト
 * @return {string} userPrompt
 */
function buildMinutesUserPrompt(conversationText) {
  return `以下のSEO相談セッションの議事録を作成してください。

【会話ログ】
${conversationText}

上記の会話を分析し、JSON形式で議事録を出力してください。`;
}

/**
 * 議事録内容をパース
 * @param {string} content - Claude APIからの応答
 * @return {Object} parsedMinutes
 */
function parseMinutesContent(content) {
  try {
    // JSONを抽出（コードブロックに囲まれている場合も対応）
    let jsonStr = content;
    
    // ```json ... ``` を除去
    const jsonMatch = content.match(/```json\s*([\s\S]*?)\s*```/);
    if (jsonMatch) {
      jsonStr = jsonMatch[1];
    } else {
      // ``` ... ``` を除去
      const codeMatch = content.match(/```\s*([\s\S]*?)\s*```/);
      if (codeMatch) {
        jsonStr = codeMatch[1];
      }
    }
    
    // { から } までを抽出
    const braceMatch = jsonStr.match(/\{[\s\S]*\}/);
    if (braceMatch) {
      jsonStr = braceMatch[0];
    }
    
    const parsed = JSON.parse(jsonStr);
    
    return {
      summary: parsed.summary || '',
      findings: parsed.findings || '',
      decisions: parsed.decisions || '',
      action_items: parsed.action_items || '',
      tags: parsed.tags || ''
    };
    
  } catch (error) {
    Logger.log(`議事録パースエラー: ${error.message}`);
    Logger.log(`元のコンテンツ: ${content}`);
    
    // パース失敗時はそのまま返す
    return {
      summary: content,
      findings: '',
      decisions: '',
      action_items: '',
      tags: ''
    };
  }
}

/**
 * 会話履歴から参照ページを抽出
 * @param {Array} conversationHistory - 会話履歴
 * @return {Array} pages - 参照されたページURL
 */
function extractReferencedPages(conversationHistory) {
  const pages = [];
  const urlPattern = /\/[a-zA-Z0-9\-_]+(?=[\s,。、]|$)/g;
  
  conversationHistory.forEach(msg => {
    const matches = msg.content.match(urlPattern);
    if (matches) {
      matches.forEach(url => {
        if (!pages.includes(url)) {
          pages.push(url);
        }
      });
    }
  });
  
  return pages;
}

/**
 * 議事録をシートに保存
 * @param {Object} minutesData - 議事録データ
 */
function saveMinutes(minutesData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('議事録');
  
  if (!sheet) {
    throw new Error('議事録シートが見つかりません');
  }
  
  const row = [
    minutesData.session_id,
    minutesData.session_date,
    minutesData.session_start,
    minutesData.session_end,
    minutesData.summary,
    minutesData.findings,
    minutesData.decisions,
    minutesData.action_items,
    minutesData.referenced_pages,
    minutesData.referenced_data,
    minutesData.tags,
    minutesData.full_conversation
  ];
  
  sheet.appendRow(row);
  Logger.log('議事録シートに保存完了');
}

/**
 * 議事録を取得（フィルター対応）
 * @param {Object} filters - フィルター条件（オプション）
 * @return {Array} minutes - 議事録一覧
 */
function getMinutes(filters) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('議事録');
    
    if (!sheet) {
      throw new Error('議事録シートが見つかりません');
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // ヘッダー行を除く
    let rows = data.slice(1).filter(row => row[0] !== ''); // 空行除外
    
    // フィルター適用
    if (filters) {
      if (filters.start_date) {
        const dateIdx = headers.indexOf('session_date');
        rows = rows.filter(row => new Date(row[dateIdx]) >= filters.start_date);
      }
      
      if (filters.end_date) {
        const dateIdx = headers.indexOf('session_date');
        rows = rows.filter(row => new Date(row[dateIdx]) <= filters.end_date);
      }
      
      if (filters.keyword) {
        rows = rows.filter(row => {
          const text = row.join(' ').toLowerCase();
          return text.includes(filters.keyword.toLowerCase());
        });
      }
      
      if (filters.tag) {
        const tagIdx = headers.indexOf('tags');
        rows = rows.filter(row => row[tagIdx] && row[tagIdx].includes(filters.tag));
      }
    }
    
    // オブジェクト配列に変換
    const minutes = rows.map(row => {
      const obj = {};
      headers.forEach((header, idx) => {
        obj[header] = row[idx];
      });
      return obj;
    });
    
    Logger.log(`議事録取得: ${minutes.length}件`);
    
    return minutes;
    
  } catch (error) {
    Logger.log(`議事録取得エラー: ${error.message}`);
    return [];
  }
}

/**
 * 議事録を検索
 * @param {string} keyword - 検索キーワード
 * @return {Array} results - 検索結果
 */
function searchMinutes(keyword) {
  return getMinutes({ keyword: keyword });
}

/**
 * 特定セッションの議事録を取得
 * @param {string} sessionId - セッションID
 * @return {Object} minutes - 議事録
 */
function getMinutesBySessionId(sessionId) {
  const allMinutes = getMinutes();
  return allMinutes.find(m => m.session_id === sessionId) || null;
}

/**
 * 最新の議事録を取得
 * @param {number} count - 取得件数（デフォルト5）
 * @return {Array} minutes - 議事録一覧
 */
function getRecentMinutes(count = 5) {
  const allMinutes = getMinutes();
  
  // 日付順にソート（降順）
  allMinutes.sort((a, b) => new Date(b.session_date) - new Date(a.session_date));
  
  return allMinutes.slice(0, count);
}

/**
 * テスト関数: 議事録生成
 */
function testGenerateMinutes() {
  // テスト用の会話履歴
  const testConversation = [
    {
      role: 'user',
      content: '順位が下がっているページを教えてください'
    },
    {
      role: 'assistant',
      content: '順位が下がっているページを分析しました。\n\n/master-ipad-app-switcher は先週の4.2位から5.8位に下降しています。\n/iphone-hoken-4-recommendation は安定しています。\n\nリライトを検討してみてはいかがでしょうか？'
    },
    {
      role: 'user',
      content: '/master-ipad-app-switcher のリライト提案をお願いします'
    },
    {
      role: 'assistant',
      content: 'リライト提案をします。\n\n【タイトル改善案】\n現在: iPadアプリ切り替え方法\n提案: iPadアプリ切り替え完全ガイド｜マルチタスクを使いこなす\n\n【優先度】高\n\n実施をおすすめします。'
    },
    {
      role: 'user',
      content: 'わかりました。今週中にリライトします'
    }
  ];
  
  const result = generateMinutes(testConversation);
  
  Logger.log('=== テスト結果 ===');
  Logger.log(JSON.stringify(result, null, 2));
}

/**
 * テスト関数: 議事録取得
 */
function testGetMinutes() {
  const minutes = getMinutes();
  Logger.log(`全議事録件数: ${minutes.length}`);
  
  if (minutes.length > 0) {
    Logger.log('最新の議事録:');
    Logger.log(JSON.stringify(minutes[minutes.length - 1], null, 2));
  }
}

/**
 * テスト関数: 議事録検索
 */
function testSearchMinutes() {
  const results = searchMinutes('リライト');
  Logger.log(`検索結果: ${results.length}件`);
}
