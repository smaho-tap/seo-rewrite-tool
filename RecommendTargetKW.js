/**
 * 推奨ターゲットKW選定機能
 * 
 * 複数KWが設定されているページに対して、
 * リライト効果が最も高いKWを自動提案する
 * 
 * 【選定ロジック】
 * 1. 1位のKWは除外（改善余地なし）
 * 2. 検索ボリューム100未満は除外（効果が薄い）
 * 3. スコアリング:
 *    - ボリュームスコア（正規化）
 *    - 順位改善スコア（6-20位が最高）
 * 4. 総合スコア最大のKWを推奨
 * 
 * 作成日: 2025/12/02
 */

/**
 * 指定URLの推奨ターゲットKWを取得
 * @param {String} pageUrl - ページURL（パス形式）
 * @return {Object} 推奨KW情報
 */
function getRecommendedTargetKeyword(pageUrl) {
  Logger.log('=== 推奨ターゲットKW選定開始 ===');
  Logger.log('対象URL: ' + pageUrl);
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var gyronSheet = ss.getSheetByName('GyronSEO_RAW');
  var volumeSheet = ss.getSheetByName('検索ボリューム_RAW');
  
  if (!gyronSheet) {
    return { success: false, error: 'GyronSEO_RAWシートが見つかりません' };
  }
  
  // ========================================
  // 1. 検索ボリュームデータを読み込み
  // ========================================
  var volumeMap = {};
  if (volumeSheet) {
    var volumeData = volumeSheet.getDataRange().getValues();
    for (var v = 1; v < volumeData.length; v++) {
      var kw = String(volumeData[v][0] || '').trim();
      var vol = parseFloat(volumeData[v][1]) || 0;
      if (kw) {
        volumeMap[normalizeKeywordForRecommend(kw)] = vol;
      }
    }
  }
  
  // ========================================
  // 2. GyronSEO_RAWから該当URLのKWを取得
  // ========================================
  var gyronData = gyronSheet.getDataRange().getValues();
  var gyronHeaders = gyronData[0];
  
  // 最新日付列を探す
  var latestDateColIndex = -1;
  var latestDate = null;
  for (var col = 0; col < gyronHeaders.length; col++) {
    if (gyronHeaders[col] instanceof Date) {
      if (!latestDate || gyronHeaders[col] > latestDate) {
        latestDate = gyronHeaders[col];
        latestDateColIndex = col;
      }
    }
  }
  
  if (latestDateColIndex === -1) {
    return { success: false, error: '順位データが見つかりません' };
  }
  
  // 該当URLのKWを収集
  var keywords = [];
  var normalizedPageUrl = pageUrl.startsWith('/') ? pageUrl : '/' + pageUrl;
  
  for (var i = 1; i < gyronData.length; i++) {
    var keyword = String(gyronData[i][0] || '').trim();
    var url = String(gyronData[i][1] || '').trim();
    var position = gyronData[i][latestDateColIndex];
    
    if (!keyword || !url) continue;
    
    // URLを正規化してマッチ確認
    var path = extractPathForRecommend(url);
    if (path !== normalizedPageUrl) continue;
    
    // 順位を数値に変換
    var positionNum = null;
    if (position !== '' && position !== null) {
      if (String(position).includes('圏外')) {
        positionNum = 101;
      } else {
        positionNum = parseFloat(position) || null;
      }
    }
    
    // 検索ボリュームを取得
    var normalizedKW = normalizeKeywordForRecommend(keyword);
    var volume = volumeMap[normalizedKW] || 0;
    
    keywords.push({
      keyword: keyword,
      position: positionNum,
      volume: volume
    });
  }
  
  Logger.log('該当KW数: ' + keywords.length);
  
  if (keywords.length === 0) {
    return { 
      success: false, 
      error: '該当URLのキーワードが見つかりません',
      keywords: []
    };
  }
  
  // ========================================
  // 3. 推奨KWを選定
  // ========================================
  var recommendation = selectRecommendedKeyword(keywords);
  
  Logger.log('=== 推奨ターゲットKW選定完了 ===');
  
  return {
    success: true,
    pageUrl: pageUrl,
    totalKeywords: keywords.length,
    recommendation: recommendation,
    allKeywords: keywords
  };
}

/**
 * KWリストから推奨KWを選定
 * @param {Array} keywords - KWリスト
 * @return {Object} 推奨KW情報
 */
function selectRecommendedKeyword(keywords) {
  // フィルタリング: 1位と100未満を除外
  var candidates = keywords.filter(function(kw) {
    // 1位は除外
    if (kw.position === 1) return false;
    // 検索ボリューム100未満は除外
    if (kw.volume < 100) return false;
    // 圏外は除外
    if (kw.position === 101 || kw.position === null) return false;
    return true;
  });
  
  if (candidates.length === 0) {
    // 候補がない場合、1位以外で最大ボリュームのKWを選ぶ
    var nonRank1 = keywords.filter(function(kw) {
      return kw.position !== 1 && kw.position !== 101 && kw.position !== null;
    });
    
    if (nonRank1.length === 0) {
      return {
        keyword: null,
        reason: '推奨できるキーワードがありません（全て1位または圏外）'
      };
    }
    
    // ボリューム最大を選ぶ
    var maxVolKW = nonRank1.reduce(function(max, kw) {
      return (kw.volume > max.volume) ? kw : max;
    }, { volume: -1 });
    
    return {
      keyword: maxVolKW.keyword,
      position: maxVolKW.position,
      volume: maxVolKW.volume,
      score: 0,
      reason: getRankReason(maxVolKW.position),
      note: '検索ボリューム100以上の候補がないため、ボリューム最大のKWを選定'
    };
  }
  
  // ボリュームの最大値を取得（正規化用）
  var maxVolume = Math.max.apply(null, candidates.map(function(kw) { return kw.volume; }));
  
  // スコアリング
  var scoredCandidates = candidates.map(function(kw) {
    var volumeScore = maxVolume > 0 ? kw.volume / maxVolume : 0;
    var rankScore = getRankScore(kw.position);
    var totalScore = volumeScore * 0.5 + rankScore * 0.5;
    
    return {
      keyword: kw.keyword,
      position: kw.position,
      volume: kw.volume,
      volumeScore: volumeScore,
      rankScore: rankScore,
      totalScore: totalScore
    };
  });
  
  // スコア順にソート
  scoredCandidates.sort(function(a, b) {
    return b.totalScore - a.totalScore;
  });
  
  var best = scoredCandidates[0];
  
  return {
    keyword: best.keyword,
    position: best.position,
    volume: best.volume,
    score: Math.round(best.totalScore * 100),
    reason: getRankReason(best.position),
    alternatives: scoredCandidates.slice(1, 4).map(function(kw) {
      return {
        keyword: kw.keyword,
        position: kw.position,
        volume: kw.volume,
        score: Math.round(kw.totalScore * 100)
      };
    })
  };
}

/**
 * 順位に基づく改善スコアを取得
 * @param {Number} position - 検索順位
 * @return {Number} スコア（0-1）
 */
function getRankScore(position) {
  if (position === null || position === undefined) return 0;
  if (position === 1) return 0;        // 改善不要
  if (position <= 5) return 0.3;       // 2-5位: 低リスク施策のみ
  if (position <= 10) return 0.8;      // 6-10位: TOP5狙える
  if (position <= 20) return 1.0;      // 11-20位: 1ページ目狙える
  if (position <= 50) return 0.5;      // 21-50位: 勝算あり
  return 0.2;                          // 51位以下: 勝算低い
}

/**
 * 順位に基づく理由文を取得
 * @param {Number} position - 検索順位
 * @return {String} 理由文
 */
function getRankReason(position) {
  if (position === null || position === undefined) return '順位不明';
  if (position === 1) return '1位のため改善不要';
  if (position <= 5) return '上位維持のための低リスク施策推奨';
  if (position <= 10) return 'TOP5入りを狙える好位置';
  if (position <= 20) return '1ページ目入りを狙える';
  if (position <= 50) return '改善余地が大きい';
  return '改善に時間がかかる可能性あり';
}

/**
 * キーワードを正規化
 */
function normalizeKeywordForRecommend(keyword) {
  if (!keyword) return '';
  return String(keyword)
    .toLowerCase()
    .replace(/　/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * URLからパス部分を抽出
 */
function extractPathForRecommend(url) {
  if (!url) return '';
  try {
    if (url.startsWith('/')) return url;
    var match = url.match(/https?:\/\/[^\/]+(\/[^\?#]*)?/);
    if (match && match[1]) return match[1];
    if (!url.includes('://')) return '/' + url;
    return '';
  } catch (e) {
    return '';
  }
}

// ========================================
// テスト関数
// ========================================

/**
 * 推奨ターゲットKWのテスト
 */
function testRecommendedTargetKeyword() {
  var testUrl = '/iphonerepair-screen-line';
  var result = getRecommendedTargetKeyword(testUrl);
  
  Logger.log('=== テスト結果 ===');
  Logger.log('URL: ' + testUrl);
  Logger.log('成功: ' + result.success);
  
  if (result.success) {
    Logger.log('総KW数: ' + result.totalKeywords);
    Logger.log('');
    Logger.log('【推奨ターゲットKW】');
    Logger.log('  キーワード: ' + result.recommendation.keyword);
    Logger.log('  現在順位: ' + result.recommendation.position + '位');
    Logger.log('  検索ボリューム: ' + result.recommendation.volume);
    Logger.log('  スコア: ' + result.recommendation.score + '点');
    Logger.log('  理由: ' + result.recommendation.reason);
    
    if (result.recommendation.alternatives && result.recommendation.alternatives.length > 0) {
      Logger.log('');
      Logger.log('【代替候補】');
      result.recommendation.alternatives.forEach(function(alt, i) {
        Logger.log('  ' + (i + 1) + '. ' + alt.keyword + ' (' + alt.position + '位, vol:' + alt.volume + ', スコア:' + alt.score + ')');
      });
    }
    
    Logger.log('');
    Logger.log('【全KW一覧】');
    result.allKeywords.forEach(function(kw) {
      var mark = kw.position === 1 ? '★1位' : kw.position + '位';
      Logger.log('  - ' + kw.keyword + ' (vol:' + kw.volume + ', ' + mark + ')');
    });
  } else {
    Logger.log('エラー: ' + result.error);
  }
}

/**
 * 複数URLで推奨KWをテスト
 */
function testMultipleRecommendations() {
  var testUrls = [
    '/iphonerepair-screen-line',
    '/ipad-mini-cheap-buy-methods',
    '/iphonerepair-screen-color-strange'
  ];
  
  Logger.log('=== 複数URL推奨KWテスト ===');
  
  testUrls.forEach(function(url) {
    var result = getRecommendedTargetKeyword(url);
    
    Logger.log('');
    Logger.log('■ ' + url);
    
    if (result.success && result.recommendation.keyword) {
      Logger.log('  推奨KW: ' + result.recommendation.keyword);
      Logger.log('  順位: ' + result.recommendation.position + '位');
      Logger.log('  ボリューム: ' + result.recommendation.volume);
      Logger.log('  理由: ' + result.recommendation.reason);
    } else {
      Logger.log('  推奨なし: ' + (result.recommendation ? result.recommendation.reason : result.error));
    }
  });
}

/**
 * チャット表示形式のテスト
 */
function testChatFormat() {
  var testUrl = '/iphonerepair-screen-line';
  
  Logger.log('=== チャット表示形式テスト ===');
  Logger.log('URL: ' + testUrl);
  Logger.log('');
  
  var chatText = getRecommendedKeywordForChat(testUrl);
  
  if (chatText) {
    Logger.log('【チャット表示】');
    Logger.log(chatText);
  } else {
    Logger.log('表示なし（KWが1件以下または取得失敗）');
  }
}

/**
 * チャット用の推奨KW情報を取得（フォーマット済み）
 * 全KW一覧 + 推奨KWを表示
 * @param {String} pageUrl - ページURL
 * @return {String} フォーマットされた推奨情報
 */
function getRecommendedKeywordForChat(pageUrl) {
  var result = getRecommendedTargetKeyword(pageUrl);
  
  if (!result.success) {
    return null;
  }
  
  var allKeywords = result.allKeywords || [];
  var rec = result.recommendation;
  
  // KWが1件以下の場合は表示しない
  if (allKeywords.length <= 1) {
    return null;
  }
  
  // ボリューム順にソート
  var sortedKeywords = allKeywords.slice().sort(function(a, b) {
    return b.volume - a.volume;
  });
  
  var text = '\n\n📋 **登録済みターゲットKW**（' + allKeywords.length + '件）:\n\n';
  text += '| KW | 順位 | ボリューム |\n';
  text += '|:---|:----:|----------:|\n';
  
  // 表示件数（5件まで、それ以上は省略）
  var displayCount = Math.min(sortedKeywords.length, 5);
  var recommendedKW = rec.keyword ? normalizeKeywordForRecommend(rec.keyword) : '';
  
  for (var i = 0; i < displayCount; i++) {
    var kw = sortedKeywords[i];
    var isRecommended = normalizeKeywordForRecommend(kw.keyword) === recommendedKW;
    var rankDisplay = kw.position === 1 ? '**1位**' : 
                      kw.position === 101 ? '圏外' : 
                      kw.position + '位';
    
    if (isRecommended) {
      text += '| ★ ' + kw.keyword + ' | ' + rankDisplay + ' | ' + formatNumber(kw.volume) + ' | ← **推奨**\n';
    } else {
      text += '| ' + kw.keyword + ' | ' + rankDisplay + ' | ' + formatNumber(kw.volume) + ' |\n';
    }
  }
  
  // 省略表示
  if (sortedKeywords.length > 5) {
    text += '| ...（他' + (sortedKeywords.length - 5) + '件） | | |\n';
  }
  
  // 推奨KWの理由
  if (rec.keyword) {
    text += '\n📌 **推奨**: ' + rec.keyword + '\n';
    text += '   理由: ' + rec.reason + '（検索ボリューム ' + formatNumber(rec.volume) + '）\n';
  } else if (rec.reason) {
    text += '\n⚠️ ' + rec.reason + '\n';
  }
  
  return text;
}

/**
 * 数値をカンマ区切りでフォーマット
 * @param {Number} num - 数値
 * @return {String} フォーマットされた文字列
 */
function formatNumber(num) {
  if (num === null || num === undefined) return '-';
  return num.toLocaleString();
}