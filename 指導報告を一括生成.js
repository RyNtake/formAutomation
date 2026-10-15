/**
 * メニューを追加
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('★報告書作成')
    .addItem('指導報告を一括生成', 'generateReport')
    .addToUi();
}

/**
 * 指定期間の指導報告を生成して「【仮】指導報告」シートへ転記する
 */
function generateReport() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. シートの取得
  const sourceSheet = ss.getSheetByName('フォーム回答ストレージ');
  const targetSheet = ss.getSheetByName('【仮】指導報告');
  
  if (!sourceSheet || !targetSheet) {
    ui.alert('エラー', '「フォーム回答ストレージ」または「【仮】指導報告」シートが見つかりません。', ui.ButtonSet.OK);
    return;
  }

  // 2. 期間の入力プロンプト
  const startDatePrompt = ui.prompt('開始日の指定', '開始日を入力してください (例: 2026/09/27)', ui.ButtonSet.OK_CANCEL);
  if (startDatePrompt.getSelectedButton() !== ui.Button.OK) return;
  const startDate = parseDate(startDatePrompt.getResponseText().trim());
  if (!startDate) {
    ui.alert('開始日の形式が正しくありません。(例: 2026/09/27)');
    return;
  }

  const endDatePrompt = ui.prompt('終了日の指定', '終了日を入力してください (例: 2026/09/30)', ui.ButtonSet.OK_CANCEL);
  if (endDatePrompt.getSelectedButton() !== ui.Button.OK) return;
  const endDate = parseDate(endDatePrompt.getResponseText().trim());
  if (!endDate) {
    ui.alert('終了日の形式が正しくありません。(例: 2026/09/30)');
    return;
  }
  // 終了日の23:59:59まで含める
  endDate.setHours(23, 59, 59, 999);

  // 3. データ取得
  const data = sourceSheet.getDataRange().getValues();
  if (data.length <= 1) {
    ui.alert('フォーム回答データが存在しません。');
    return;
  }

  const headers = data[0];
  const rows = data.slice(1);

  // ヘッダー列インデックスの特定
  const colDate = findColumnIndex(headers, ['指導日', '授業日', '日付', 'タイムスタンプ']);
  const colStudent = findColumnIndex(headers, ['生徒名', '生徒氏名', '氏名']);
  const colSubject = findColumnIndex(headers, ['教科', '科目']);
  const colTeaching = findColumnIndex(headers, ['教材', '指導項目', '教材名']);
  const colUnit = findColumnIndex(headers, ['単元名', '単元', '範囲メモ']);
  const colEval = findColumnIndex(headers, ['理解', '定着度', '評価']);
  // 「指導内容・ポイント」列
  const colComment = findColumnIndex(headers, ['指導内容・ポイント', '指導内容', 'ポイント', '指導報告', '指導コメント', 'コメント', '詳細']);

  if (colDate === -1 || colStudent === -1) {
    ui.alert('エラー', `必要な列が見つかりませんでした。\n日付列: ${colDate !== -1 ? 'OK' : '見つかりません'}\n生徒名列: ${colStudent !== -1 ? 'OK' : '見つかりません'}`, ui.ButtonSet.OK);
    return;
  }

  // 4. 該当期間のデータをグループ化 (Key: "yyyy/MM/dd___生徒名")
  const groupedData = {};

  rows.forEach(row => {
    const rawDate = row[colDate];
    if (!rawDate) return;
    const targetDate = new Date(rawDate);
    if (isNaN(targetDate.getTime())) return;

    if (targetDate >= startDate && targetDate <= endDate) {
      const dateStr = Utilities.formatDate(targetDate, Session.getScriptTimeZone(), 'yyyy/MM/dd');
      const studentName = String(row[colStudent] || '').trim();
      if (!studentName) return;

      const groupKey = `${dateStr}___${studentName}`;
      if (!groupedData[groupKey]) {
        groupedData[groupKey] = {
          rawDateObj: targetDate,
          dateStr: dateStr,
          studentName: studentName,
          records: []
        };
      }

      groupedData[groupKey].records.push({
        subject: colSubject !== -1 ? String(row[colSubject] || '').trim() : '',
        teaching: colTeaching !== -1 ? String(row[colTeaching] || '').trim() : '',
        unit: colUnit !== -1 ? String(row[colUnit] || '').trim() : '',
        eval: colEval !== -1 ? String(row[colEval] || '').trim() : '',
        comment: colComment !== -1 ? String(row[colComment] || '').trim() : ''
      });
    }
  });

  // 理解度マッピング
  const evalMap = {
    '1': '改めて基礎から確認の必要があります',
    '2': '少し苦手意識が見られます',
    '3': '概ね理解できています',
    '4': '完璧に理解できています'
  };

  // 5. 報告文の生成
  const results = [];

  for (const key in groupedData) {
    const group = groupedData[key];
    const records = group.records;

    // 「M月d日の」表記
    const jpDateText = Utilities.formatDate(group.rawDateObj, Session.getScriptTimeZone(), 'M月d日') + 'の';

    // 宛名：生徒名から名字（半角・全角スペースの前）のみを抽出
    const lastName = group.studentName.split(/[\s ]+/)[0] || group.studentName;
    const studentDisplay = group.studentName;

    // 科目ごとにグループ分け
    const subjects = [...new Set(records.map(r => r.subject).filter(s => s !== ''))];
    if (subjects.length === 0) subjects.push('');

    const subjectTexts = subjects.map(subj => {
      const subjRecords = subj ? records.filter(r => r.subject === subj) : records;
      
      const headerPrefix = records.length >= 2 && subj ? `〜${subj}〜\n` : '';

      // 学習内容まとめ: 「[日付]は[教科]の[教材]の「[単元/範囲]」と...を学習しました。」
      const unitPhrases = subjRecords
        .map(r => {
          const unitName = r.unit || '';
          if (r.teaching && unitName) return `${r.teaching}の「${unitName}」`;
          if (unitName) return `「${unitName}」`;
          return '';
        })
        .filter(t => t !== '');
        
      let studySummary = `${jpDateText}は${subj ? subj + 'の' : ''}${unitPhrases.length > 0 ? unitPhrases.join('と') : '学習内容'}を学習しました。`;

      // 理解度テキストの結合
      const evalTexts = subjRecords
        .map(r => evalMap[r.eval] || '')
        .filter(t => t !== '');
      if (evalTexts.length > 0) {
        const uniqueEvals = [...new Set(evalTexts)];
        studySummary += `${uniqueEvals.join('、')}。`;
      }

      // 指導内容・ポイントの差し込み
      const commentItems = subjRecords
        .map(r => r.comment)
        .filter(c => c !== '');
      
      const details = commentItems.join('\n\n');

      // 本文パートを結合
      const parts = [headerPrefix ? headerPrefix.trimEnd() : '', studySummary];
      if (details) parts.push(details);

      return parts.filter(p => p !== '').join('\n\n');
    });

    const reportBody = subjectTexts.join('\n\n');
    
    // 宛頭: 「〇〇さん、お世話になっております。」（名字のみ＋さん）
    const fullReport = `${lastName}さん、お世話になっております。${jpDateText}${studentDisplay}さんの授業報告です。\n\n${reportBody}\n\n引き続き、精一杯サポートさせていただきます。`;

    results.push([group.dateStr, group.studentName, fullReport]);
  }

  if (results.length === 0) {
    ui.alert('該当する期間のデータが見つかりませんでした。');
    return;
  }

  // 6. 「【仮】指導報告」シートの最終行に追加
  targetSheet.getRange(targetSheet.getLastRow() + 1, 1, results.length, 3).setValues(results);
  ui.alert('完了', `${results.length} 件の指導報告を作成しました。`, ui.ButtonSet.OK);
}

/**
 * 補助関数: 文字列をDate型に変換
 */
function parseDate(str) {
  if (!str) return null;
  const d = new Date(str.replace(/-/g, '/'));
  return isNaN(d.getTime()) ? null : d;
}

/**
 * 補助関数: 該当する列の見出しを柔軟に探索
 */
function findColumnIndex(headers, candidateNames) {
  for (let name of candidateNames) {
    const idx = headers.findIndex(h => String(h).trim().includes(name));
    if (idx !== -1) return idx;
  }
  return -1;
}