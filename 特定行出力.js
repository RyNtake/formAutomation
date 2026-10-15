/**
 * RECORD_SHEET_NAMEから指定した行番号の宿題プリントPDFのみを再生成する
 * @param {number} rowNumber スプレッドシートの行番号（ヘッダーが1行目の場合、データは2行目以降）
 */
function regeneratePdfOnly(rowNumber = 48) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const entrySheet = ss.getSheetByName(CONFIG.RECORD_SHEET_NAME);
  const headers = entrySheet.getRange(1, 1, 1, entrySheet.getLastColumn()).getValues()[0];
  const targetRow = entrySheet.getRange(rowNumber, 1, 1, entrySheet.getLastColumn()).getValues()[0];

  const rawData = {};
  headers.forEach((h, i) => {
    rawData[String(h).trim()] = targetRow[i];
  });

  // 日付の処理
  let timestamp = rawData["タイムスタンプ"];
  let dateObj = new Date();
  if (timestamp) {
    const parsed = new Date(timestamp);
    if (!isNaN(parsed.getTime())) dateObj = parsed;
  }
  const dateStr = Utilities.formatDate(dateObj, "Asia/Tokyo", "yyyyMMdd");
  const displayDate = Utilities.formatDate(dateObj, "Asia/Tokyo", "yyyy年M月d日");
  const dueDateObj = new Date(dateObj.getTime() + (5 * 24 * 60 * 60 * 1000));
  const displayDueDate = Utilities.formatDate(dueDateObj, "Asia/Tokyo", "yyyy年M月d日");

  // 生徒データパース
  const students = [];
  const s1Name = String(rawData["生徒名1"] || rawData["生徒名"] || "").trim();
  if (s1Name) {
    students.push({
      name: s1Name,
      subject: String(rawData["教科1"] || rawData["教科"] || "").trim(),
      division: String(rawData["区分1"] || rawData["区分"] || "通常").trim(),
      textbook: String(rawData["教材1"] || rawData["教材"] || "").trim(),
      exercise: String(rawData["演習1"] || rawData["演習"] || rawData["範囲メモ1"] || "（なし）").trim(),
      homework: String(rawData["宿題1"] || rawData["宿題"] || rawData["次回持越し1"] || "（なし）").trim()
    });
  }

  const s2Name = String(rawData["生徒名2"] || "").trim();
  if (s2Name) {
    students.push({
      name: s2Name,
      subject: String(rawData["教科2"] || "").trim(),
      division: String(rawData["区分2"] || "通常").trim(),
      textbook: String(rawData["教材2"] || "").trim(),
      exercise: String(rawData["演習2"] || rawData["範囲メモ2"] || "（なし）").trim(),
      homework: String(rawData["宿題2"] || rawData["次回持越し2"] || "（なし）").trim()
    });
  }

  if (students.length === 0) {
    Logger.log(`${rowNumber}行目に生徒データが見つかりませんでした。`);
    return;
  }

  // 提出用WebアプリのURL（演習宿題を自動生成.gsと共通）
  const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbxSTQWTFsi4cQCV4_rYBiK4NpzgV3_J3HMaagEc7kOX30_rQoKdvD-TZwrGEjyiwUeH/exec";

  // フォルダ取得 ＆ 提出用WebアプリURLのQRコード生成
  const rootFolder = DriveApp.getFolderById(CONFIG.HOMEWORK_ROOT_FOLDER_ID);
  students.forEach((st, idx) => {
    st.lessonId = `L_${dateStr}_${st.name}`;
    if (students.length > 1 && students.filter(s => s.name === st.name).length > 1) {
      st.lessonId += `_${idx + 1}`;
    }
    const studentFolder = getOrCreateFolder(rootFolder, st.name);
    const dateFolder = getOrCreateFolder(studentFolder, dateStr);
    
    // フォルダIDの取得
    st.folderId = dateFolder.getId();
    st.folderUrl = dateFolder.getUrl();

    // 提出用URL生成（パラメータ: folderId, name）
    const submitUrl = `${WEB_APP_URL}?folderId=${st.folderId}&name=${encodeURIComponent(st.name)}`;

    // 提出用WebアプリのリンクでQRコード生成
    st.qrUrl = `https://quickchart.io/qr?text=${encodeURIComponent(submitUrl)}&size=130`;
  });

  // PDF生成
  const printFolder = DriveApp.getFolderById(CONFIG.PRINT_TMP_FOLDER_ID);
  const combinedHtml = generateCombinedCardHtml(students, displayDate, displayDueDate);
  const blob = HtmlService.createHtmlOutput(combinedHtml).getAs(MimeType.PDF);
  const baseName = students.map(s => s.name).join("・");
  blob.setName(`${baseName}_${dateStr}.pdf`);
  printFolder.createFile(blob);

  Logger.log(`再出力完了: ${baseName}_${dateStr}.pdf を出力しました。（行: ${rowNumber}）`);
}