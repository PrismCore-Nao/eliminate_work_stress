// Google Apps Script - 4つの診断フォームのデータを別々のシートに保存

function doPost(e) {
  try {
    // スプレッドシートのIDを設定してください
    const SPREADSHEET_ID = 'https://script.google.com/macros/s/1tUO1DPTCTPJ9-LfXSNDFXGIjk8vg6SI3MtU6j6WxKfs/exec';
    
    // データを取得
    const data = JSON.parse(e.postData.contents);
    const testType = data.test_type;
    
    // スプレッドシートを開く
    const ss = SpreadsheetApp.openById(1tUO1DPTCTPJ9-LfXSNDFXGIjk8vg6SI3MtU6j6WxKfs);
    
    // テストタイプに応じてシート名を決定
    let sheetName = '';
    if (testType === '脳の状態診断') {
      sheetName = '1_脳の状態診断';
    } else if (testType === '適職度診断') {
      sheetName = '2_適職度診断';
    } else if (testType === 'ストレス診断') {
      sheetName = '3_ストレス診断';
    } else if (testType === 'ビッグファイブ診断') {
      sheetName = '4_ビッグファイブ診断';
    } else {
      throw new Error('不明なテストタイプです');
    }
    
    let sheet = ss.getSheetByName(sheetName);
    
    // シートが存在しない場合は作成
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      setupSheet(sheet, testType);
    }
    
    // データを行に変換して追加
    const row = createDataRow(data, testType);
    sheet.appendRow(row);
    
    // 成功レスポンスを返す
    return ContentService.createTextOutput(JSON.stringify({
      result: 'success',
      message: 'データが正常に保存されました',
      sheet: sheetName
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    // エラーレスポンスを返す
    return ContentService.createTextOutput(JSON.stringify({
      result: 'error',
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// シートのセットアップ（ヘッダー行を作成）
function setupSheet(sheet, testType) {
  let headers = [];
  
  if (testType === '脳の状態診断') {
    headers = [
      '送信日',
      '送信時刻',
      '送信日時',
      'Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'Q6', 'Q7', 'Q8', 'Q9', 'Q10',
      '合計点',
      '判定レベル',
      '判定結果'
    ];
  } else if (testType === '適職度診断') {
    headers = [
      '送信日',
      '送信時刻',
      '送信日時',
      'Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'Q6', 'Q7', 'Q8', 'Q9', 'Q10',
      'チェック数',
      '判定レベル',
      '判定結果'
    ];
  } else if (testType === 'ストレス診断') {
    headers = [
      '送信日',
      '送信時刻',
      '送信日時',
      'Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'Q6', 'Q7', 'Q8', 'Q9', 'Q10',
      'Q11', 'Q12', 'Q13', 'Q14', 'Q15',
      '合計点',
      '警報レベル',
      '判定結果'
    ];
  } else if (testType === 'ビッグファイブ診断') {
    headers = [
      '送信日',
      '送信時刻',
      '送信日時',
      'Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'Q6', 'Q7', 'Q8', 'Q9', 'Q10',
      '外向性',
      '神経症傾向',
      '開放性',
      '協調性',
      '誠実性',
      '外向性タイプ',
      '神経症傾向タイプ',
      '開放性タイプ',
      'パーソナリティタイプ'
    ];
  }
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // ヘッダー行の書式設定
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#ffff00');
  headerRange.setFontWeight('bold');
  headerRange.setBorder(true, true, true, true, true, true);
  headerRange.setHorizontalAlignment('center');
  
  // 列幅を自動調整
  for (let i = 1; i <= headers.length; i++) {
    sheet.autoResizeColumn(i);
  }
  
  // 固定行
  sheet.setFrozenRows(1);
}

// データ行を作成
function createDataRow(data, testType) {
  let row = [];
  
  if (testType === '脳の状態診断') {
    row = [
      data.date || '',
      data.time || '',
      data.timestamp || '',
      data.q1 || '', data.q2 || '', data.q3 || '', data.q4 || '', data.q5 || '',
      data.q6 || '', data.q7 || '', data.q8 || '', data.q9 || '', data.q10 || '',
      data.total_score || 0,
      data.result_level || '',
      data.result_text || ''
    ];
  } else if (testType === '適職度診断') {
    row = [
      data.date || '',
      data.time || '',
      data.timestamp || '',
      data.q1 || '0', data.q2 || '0', data.q3 || '0', data.q4 || '0', data.q5 || '0',
      data.q6 || '0', data.q7 || '0', data.q8 || '0', data.q9 || '0', data.q10 || '0',
      data.total_count || 0,
      data.result_level || '',
      data.result_text || ''
    ];
  } else if (testType === 'ストレス診断') {
    row = [
      data.date || '',
      data.time || '',
      data.timestamp || '',
      data.q1 || '', data.q2 || '', data.q3 || '', data.q4 || '', data.q5 || '',
      data.q6 || '', data.q7 || '', data.q8 || '', data.q9 || '', data.q10 || '',
      data.q11 || '', data.q12 || '', data.q13 || '', data.q14 || '', data.q15 || '',
      data.total_score || 0,
      data.result_level || '',
      data.result_text || ''
    ];
  } else if (testType === 'ビッグファイブ診断') {
    row = [
      data.date || '',
      data.time || '',
      data.timestamp || '',
      data.q1 || '', data.q2 || '', data.q3 || '', data.q4 || '', data.q5 || '',
      data.q6 || '', data.q7 || '', data.q8 || '', data.q9 || '', data.q10 || '',
      data.extraversion || 0,
      data.neuroticism || 0,
      data.openness || 0,
      data.agreeableness || 0,
      data.conscientiousness || 0,
      data.extraversion_type || '',
      data.neuroticism_type || '',
      data.openness_type || '',
      data.personality_type || ''
    ];
  }
  
  return row;
}

// テスト用の関数
function testDoPost() {
  // 脳の状態診断のテスト
  const testData1 = {
    test_type: '脳の状態診断',
    date: new Date().toLocaleDateString('ja-JP'),
    time: new Date().toLocaleTimeString('ja-JP'),
    timestamp: new Date().toLocaleString('ja-JP'),
    q1: '2', q2: '1', q3: '0', q4: '1', q5: '2',
    q6: '1', q7: '0', q8: '2', q9: '1', q10: '0',
    total_score: 10,
    result_level: '脳内物質量【低】',
    result_text: 'テスト結果'
  };
  
  const e1 = {
    postData: {
      contents: JSON.stringify(testData1)
    }
  };
  
  const result1 = doPost(e1);
  Logger.log('脳の状態診断:', result1.getContent());
  
  // ビッグファイブ診断のテスト
  const testData2 = {
    test_type: 'ビッグファイブ診断',
    date: new Date().toLocaleDateString('ja-JP'),
    time: new Date().toLocaleTimeString('ja-JP'),
    timestamp: new Date().toLocaleString('ja-JP'),
    q1: '4', q2: '2', q3: '3', q4: '4', q5: '3',
    q6: '1', q7: '3', q8: '2', q9: '1', q10: '2',
    extraversion: 7,
    neuroticism: 3,
    openness: 4,
    agreeableness: 7,
    conscientiousness: 5,
    extraversion_type: '外向型',
    neuroticism_type: '楽観型',
    openness_type: '創造型',
    personality_type: '1. 発想で導くビジョナリー'
  };
  
  const e2 = {
    postData: {
      contents: JSON.stringify(testData2)
    }
  };
  
  const result2 = doPost(e2);
  Logger.log('ビッグファイブ診断:', result2.getContent());
}
