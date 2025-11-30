function doPost(e) {
  try {
    // スプレッドシートIDを設定（ここを自分のスプレッドシートIDに変更）
    const SPREADSHEET_ID = '1tUO1DPTCTPJ9-LfXSNDFXGIjk8vg6SI3MtU6j6WxKfs';
    const ss = SpreadsheetApp.openById(1tUO1DPTCTPJ9-LfXSNDFXGIjk8vg6SI3MtU6j6WxKfs);
    
    // POSTデータを取得
    const data = JSON.parse(e.postData.contents);
    const testType = data.test_type;
    
    // test_typeに基づいてシート名を決定
    let sheetName;
    switch(testType) {
      case '1_brain-state':
        sheetName = '1_brain-state';
        break;
      case '2_job-fit':
        sheetName = '2_job-fit';
        break;
      case '3_stress-check':
        sheetName = '3_stress-check';
        break;
      case '4_bigfive':
        sheetName = '4_bigfive';
        break;
      default:
        sheetName = testType;
    }
    
    // シートを取得または作成
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      
      // ヘッダー行を作成（test_typeに応じて）
      let headers;
      
      if (testType === '1_brain-state') {
        headers = ['送信日', '送信時刻', '送信日時', 'Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'Q6', 'Q7', 'Q8', 'Q9', 'Q10', '合計点数', '判定レベル', '判定結果'];
      } else if (testType === '2_job-fit') {
        headers = ['送信日', '送信時刻', '送信日時', 'Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'Q6', 'Q7', 'Q8', 'Q9', 'Q10', 'チェック数', '判定レベル', '判定結果'];
      } else if (testType === '3_stress-check') {
        headers = ['送信日', '送信時刻', '送信日時', 'Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'Q6', 'Q7', 'Q8', 'Q9', 'Q10', 'Q11', 'Q12', 'Q13', 'Q14', 'Q15', '合計点数', '警報レベル', '脳の状態'];
      } else if (testType === '4_bigfive') {
        headers = ['送信日', '送信時刻', '送信日時', 'Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'Q6', 'Q7', 'Q8', 'Q9', 'Q10', '外向性', '神経症傾向', '開放性', '協調性', '誠実性', 'パーソナリティタイプ'];
      } else {
        headers = ['送信日', '送信時刻', '送信日時', 'データ'];
      }
      
      sheet.appendRow(headers);
      
      // ヘッダー行の書式設定
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground('#ffff00');
      headerRange.setFontWeight('bold');
      headerRange.setHorizontalAlignment('center');
      
      // 列幅を自動調整
      for (let i = 1; i <= headers.length; i++) {
        sheet.autoResizeColumn(i);
      }
    }
    
    // データ行を作成
    let row;
    if (testType === '1_stress-check') {
      row = [
        data.date,
        data.time,
        data.timestamp,
        data.q1, data.q2, data.q3, data.q4, data.q5,
        data.q6, data.q7, data.q8, data.q9, data.q10,
        data.q11, data.q12, data.q13, data.q14, data.q15,
        data.total_score,
        data.result_level,
        data.result_text
      ];
    } else if (testType === '2_brain-state') {
      row = [
        data.date,
        data.time,
        data.timestamp,
        data.q1, data.q2, data.q3, data.q4, data.q5,
        data.q6, data.q7, data.q8, data.q9, data.q10,
        data.total_score,
        data.result_level,
        data.result_text
      ];
    } else if (testType === '3_job-fit') {
      row = [
        data.date,
        data.time,
        data.timestamp,
        data.q1, data.q2, data.q3, data.q4, data.q5,
        data.q6, data.q7, data.q8, data.q9, data.q10,
        data.total_count,
        data.result_level,
        data.result_text
      ];
    
    } else if (testType === '4_bigfive') {
      row = [
        data.date,
        data.time,
        data.timestamp,
        data.q1, data.q2, data.q3, data.q4, data.q5,
        data.q6, data.q7, data.q8, data.q9, data.q10,
        data.extraversion,
        data.neuroticism,
        data.openness,
        data.agreeableness,
        data.conscientiousness,
        data.personality_type
      ];
    } else {
      row = [data.date, data.time, data.timestamp, JSON.stringify(data)];
    }
    
    // データを追加
    sheet.appendRow(row);
    
    // 成功レスポンス
    return ContentService.createTextOutput(JSON.stringify({
      result: 'success',
      message: 'データが正常に保存されました',
      sheet: sheetName,
      row: sheet.getLastRow()
    }))
    .setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    // エラーレスポンス
    return ContentService.createTextOutput(JSON.stringify({
      result: 'error',
      message: error.toString()
    }))
    .setMimeType(ContentService.MimeType.JSON);
  }
}

// テスト用関数
function testDoPost() {
  const testData = {
    postData: {
      contents: JSON.stringify({
        test_type: '1_brain-state',
        date: '2024/11/30',
        time: '14:30:00',
        timestamp: '2024/11/30 14:30:00',
        q1: '2',
        q2: '1',
        q3: '0',
        q4: '1',
        q5: '2',
        q6: '1',
        q7: '0',
        q8: '1',
        q9: '2',
        q10: '1',
        total_score: 11,
        result_level: '脳内物質量【低】',
        result_text: 'ドーパミンやセロトニングが不足傾向。'
      })
    }
  };
  
  const result = doPost(testData);
  Logger.log(result.getContent());
}
