/**
 * CORE 指導報告・宿題管理 自動化スクリプト
 *
 * 1. フォーム送信時：
 *    - フォーム回答ストレージへ展開（全14列準拠）
 *    - 授業ID発行 & 生徒別提出フォルダ自動作成
 *    - 演習宿題プリント（枠拡張6行 + 提出QRコード付き）をPDF生成
 *    - 提出カンバンシートへのタスク登録
 *
 * 2. 定期実行時：
 *    - 提出フォルダ内の「演習.pdf」「宿題.pdf」を検知してカンバンに打刻
 */

// ================= 設定定数 =================
const CONFIG = {
  // 演習宿題プリント（印刷用PDF）の一時保存先フォルダID
  PRINT_TMP_FOLDER_ID: "1ISUyp1CDBKgfIAzV8BbcAXrA5F8skZZq",
  // 宿題提出用の親フォルダID（URLより取得）
  HOMEWORK_ROOT_FOLDER_ID: "1pENb5PxegKzCI-3RT8EXiBNVTZFwGozg",
  // 各シート名
  ENTRY_SHEET_NAME: "フォームの回答 2",
  RECORD_SHEET_NAME: "フォーム回答ストレージ",
  KANBAN_SHEET_NAME: "提出カンバン",
  DISCORD_WEBHOOK_URL: "https://discord.com/api/webhooks/1556728610121973940/YCqey3O9n5_uRVWU_iXn-GhCF1NNHRmBpZhvdq4DVFFq6mqwU2SacVaNYMRHxC9GtLXs"
};

function getFormText(rawData, keys) {
  for (const key of keys) {
    if (Object.prototype.hasOwnProperty.call(rawData, key)) {
      return String(rawData[key] == null ? "" : rawData[key]).trim();
    }
  }
  return "";
}

function getDueDateOffsetDays(notice) {
  const match = String(notice || "").trim().match(/^dl\s+([+-]\d+)$/i);
  if (!match) return 0;

  const offsetDays = Number(match[1]);
  return Number.isSafeInteger(offsetDays) ? offsetDays : 0;
}

function getDueDate(dateObj, notice) {
  const dueDateOffsetDays = 5 + getDueDateOffsetDays(notice);
  return new Date(dateObj.getTime() + (dueDateOffsetDays * 24 * 60 * 60 * 1000));
}

/**
 * フォーム送信時メインハンドラー
 */
function onFormSubmitHandler(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const entrySheet = ss.getSheetByName(CONFIG.ENTRY_SHEET_NAME);
  const recordSheet = ss.getSheetByName(CONFIG.RECORD_SHEET_NAME);
  const kanbanSheet = getOrCreateKanbanSheet(ss);

  let rawData = {};

  // 1. 回答データの取得
  if (e && e.namedValues) {
    for (const key in e.namedValues) {
      rawData[key.trim()] = e.namedValues[key][0];
    }
  } else {
    // 手動テスト実行時：有効な最新行を探索
    const headers = entrySheet.getRange(1, 1, 1, entrySheet.getLastColumn()).getValues()[0];
    const allValues = entrySheet.getDataRange().getValues();
    let targetRow = null;

    for (let r = allValues.length - 1; r >= 1; r--) {
      const row = allValues[r];
      if (row[0] !== "" || row[1] !== "" || row[2] !== "") {
        targetRow = row;
        break;
      }
    }
    if (!targetRow) {
      Logger.log("データ行が見つかりませんでした。");
      return;
    }
    headers.forEach((h, i) => {
      rawData[String(h).trim()] = targetRow[i];
    });
  }

  // 日付の処理
  let timestamp = rawData["タイムスタンプ"];
  let dateObj = new Date();
  if (timestamp) {
    const parsed = new Date(timestamp);
    if (!isNaN(parsed.getTime())) dateObj = parsed;
  }
  const dateStr = Utilities.formatDate(dateObj, "Asia/Tokyo", "yyyyMMdd");
  const displayDate = Utilities.formatDate(dateObj, "Asia/Tokyo", "yyyy年M月d日");

  const email = String(rawData["メールアドレス"] || "").trim();
  const teacherName = String(rawData["講師名"] || "").trim();

  // 2. 生徒データのパース（生徒1・生徒2）
  const students = [];
  const s1Name = String(rawData["生徒名1"] || rawData["生徒名"] || "").trim();
  if (s1Name) {
    students.push({
      name: s1Name,
      subject: String(rawData["教科1"] || rawData["教科"] || "").trim(),
      division: String(rawData["区分1"] || rawData["区分"] || "通常").trim(),
      textbook: String(rawData["教材1"] || rawData["教材"] || "").trim(),
      rangeMemo: String(rawData["範囲メモ1"] || rawData["範囲メモ"] || "").trim(),
      understanding: String(rawData["理解度1"] || rawData["理解度"] || "").trim(),
      instruction: String(rawData["指導内容・ポイント・解説1"] || rawData["指導内容・ポイント・解説"] || "").trim(),
      carryover: String(rawData["次回持越し1"] || rawData["次回持越し"] || "").trim(),
      notice: String(rawData["連絡事項1"] || rawData["連絡事項"] || "").trim(),
      exercise: getFormText(rawData, ["演習1", "演習"]) || "なし",
      homework: getFormText(rawData, ["宿題1", "宿題"]) || "なし"
    });
  }

  const s2Name = String(rawData["生徒名2"] || "").trim();
  if (s2Name) {
    students.push({
      name: s2Name,
      subject: String(rawData["教科2"] || "").trim(),
      division: String(rawData["区分2"] || "通常").trim(),
      textbook: String(rawData["教材2"] || "").trim(),
      rangeMemo: String(rawData["範囲メモ2"] || "").trim(),
      understanding: String(rawData["理解度2"] || "").trim(),
      instruction: String(rawData["指導内容・ポイント・解説2"] || "").trim(),
      carryover: String(rawData["次回持越し2"] || "").trim(),
      notice: String(rawData["連絡事項2"] || "").trim(),
      exercise: getFormText(rawData, ["演習2"]) || "なし",
      homework: getFormText(rawData, ["宿題2"]) || "なし"
    });
  }

  if (students.length === 0) return;

  students.forEach(st => {
    st.displayDueDate = Utilities.formatDate(
      getDueDate(dateObj, st.notice),
      "Asia/Tokyo",
      "yyyy年M月d日"
    );
  });

  // 3. 各生徒の授業ID採番・フォルダ作成・カンバン登録
  const rootFolder = DriveApp.getFolderById(CONFIG.HOMEWORK_ROOT_FOLDER_ID);

  students.forEach((st, idx) => {
    // 授業ID (例: L_20261004_下田結寿)
    st.lessonId = `L_${dateStr}_${st.name}`;
    if (students.length > 1 && students.filter(s => s.name === st.name).length > 1) {
      st.lessonId += `_${idx + 1}`;
    }

    // フォルダ自動生成：[生徒名] ➔ [日付]
    const studentFolder = getOrCreateFolder(rootFolder, st.name);
    const dateFolder = getOrCreateFolder(studentFolder, dateStr);

    // 生徒がアップロードできるように共有設定（リンクを知る全員が編集可）
    dateFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);

    st.folderId = dateFolder.getId();
    st.folderUrl = dateFolder.getUrl();

    // 1. スキャン提出用WebアプリのURLを指定（先ほどコピーしたURLを貼り付け）
    const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbxSTQWTFsi4cQCV4_rYBiK4NpzgV3_J3HMaagEc7kOX30_rQoKdvD-TZwrGEjyiwUeH/exec";

    // 2. パラメータ（folderId と name）を付与した提出用URLを生成
    const submitUrl = `${WEB_APP_URL}?folderId=${st.folderId}&name=${encodeURIComponent(st.name)}`;

    // 3. その提出用URLをQRコード画像化
    st.qrUrl = `https://quickchart.io/qr?text=${encodeURIComponent(submitUrl)}&size=400`;

    // フォーム回答ストレージへの追記（指定の全14列順）
    if (recordSheet) {
      recordSheet.appendRow([
        dateObj,           // タイムスタンプ
        st.name,           // 生徒名
        st.subject,        // 教科
        st.division,       // 区分
        st.textbook,       // 教材
        st.rangeMemo,      // 範囲メモ
        st.understanding,  // 理解度
        st.instruction,    // 指導内容・ポイント・解説
        st.carryover,      // 次回持越し
        st.notice,         // 連絡事項
        email,             // メールアドレス
        st.exercise,       // 演習
        st.homework,       // 宿題
        teacherName        // 講師名
      ]);
    }

    // カンバン管理シートへの追記
    if (kanbanSheet) {
      kanbanSheet.appendRow([
        st.lessonId,       // A: 授業ID
        dateStr,           // B: 授業日
        st.name,           // C: 生徒名
        st.subject,        // D: 教科
        st.displayDueDate, // E: 提出期限
        st.folderUrl,      // F: 提出フォルダURL
        st.exercise === "なし" ? "免除" : "未", // G: 演習提出
        "",                // H: 演習提出日時
        st.homework === "なし" ? "免除" : "未", // I: 宿題提出
        "",                // J: 宿題提出日時
        "進行中",          // K: ステータス
        st.folderId        // L: フォルダID(管理用)
      ]);
    }
  });

  // 4. 演習宿題プリント（A5横・最大2ページ構成・QRコード付き）をPDF生成
  const printFolder = DriveApp.getFolderById(CONFIG.PRINT_TMP_FOLDER_ID);
  const combinedHtml = generateCombinedCardHtml(students, displayDate);
  const blob = HtmlService.createHtmlOutput(combinedHtml).getAs(MimeType.PDF);

  const baseName = students.map(s => s.name).join("・");
  blob.setName(`${baseName}_${dateStr}.pdf`);
  printFolder.createFile(blob);

  Logger.log(`処理完了: ${baseName}_${dateStr}.pdf を出力しました。`);
}

/**
 * 演習宿題プリント（HTML）生成
 * - 演習・宿題枠を約6行分（205px）に拡張
 * - 下部に「提出期限」＋「提出用QRコード」を配置
 */
function generateCombinedCardHtml(students, displayDate, displayDueDate) {
  const cardsHtml = students.map((st, idx) => {
    const pageBreakStyle = idx < students.length - 1 ? 'style="page-break-after: always;"' : '';
    return `
    <div class="page" ${pageBreakStyle}>
      <div class="card">
        <table class="header-table">
          <tr>
            <td class="title">📝 授業・学習記録カード</td>
            <td class="lesson-id">ID: ${st.lessonId}</td>
            <td class="name-badge">${st.name} さん</td>
          </tr>
        </table>

        <table class="meta-table">
          <tr>
            <td class="label">授業日</td>
            <td class="val">${displayDate}</td>
            <td class="label">区分</td>
            <td class="val">${st.division}</td>
          </tr>
          <tr>
            <td class="label">教科</td>
            <td class="val">${st.subject || "—"}</td>
            <td class="label">教材</td>
            <td class="val">${st.textbook || "教室教材"}</td>
          </tr>
        </table>

        <!-- 演習・宿題枠：6行ゆったり入る高さ(205px) -->
        <div class="columns">
          <div class="col-box">
            <div class="col-title">✏️【本日の演習内容】</div>
            <div class="content-text">${st.exercise}</div>
          </div>
          <div class="col-box" style="border-color: #93c5fd; background: #f8fbff;">
            <div class="col-title" style="color: #1d4ed8;">🏠【次回までの宿題】</div>
            <div class="content-text">${st.homework}</div>
          </div>
        </div>

        <!-- 下部：提出期限 ＋ QRコード案内 -->
        <div class="footer-area">
          <div class="due-box">
            <span class="due-label">🚨 宿題提出期限</span>
            <span class="due-text">${st.displayDueDate || displayDueDate}</span>
          </div>

          <div class="qr-box">
            <img class="qr-img" src="${st.qrUrl}" alt="提出用QR" />
            <div class="qr-desc">
              <span class="qr-title">📱 演習・宿題 提出フォルダ</span>
              <span class="qr-sub">QRを読み取り「演習.pdf」「宿題.pdf」を保存</span>
            </div>
          </div>
        </div>
      </div>
    </div>
    `;
  }).join("");

  return `
  <!DOCTYPE html>
  <html lang="ja">
  <head>
  <meta charset="UTF-8">
  <style>
    @page {
      size: A5 landscape;
      margin: 6mm 8mm;
    }
    body {
      font-family: 'Hiragino Kaku Gothic ProN', 'Yu Gothic', sans-serif;
      color: #1e293b;
      margin: 0;
      padding: 0;
      font-size: 10.5pt;
      line-height: 1.35;
    }
    .page {
      box-sizing: border-box;
      height: 100%;
    }
    .card {
      border: 2px solid #1e3a8a;
      border-radius: 8px;
      padding: 10px 14px;
      box-sizing: border-box;
      height: 96%;
    }
    .header-table {
      width: 100%;
      border-bottom: 2px solid #2563eb;
      padding-bottom: 4px;
      margin-bottom: 8px;
    }
    .title {
      font-size: 15pt;
      font-weight: bold;
      color: #1e3a8a;
    }
    .lesson-id {
      text-align: center;
      font-size: 9pt;
      color: #64748b;
      font-family: monospace;
    }
    .name-badge {
      text-align: right;
      font-size: 13pt;
      font-weight: bold;
      background: #f1f5f9;
      padding: 3px 10px;
      border-radius: 6px;
      border: 1px solid #cbd5e1;
    }
    .meta-table {
      width: 100%;
      border-collapse: collapse;
      margin-bottom: 8px;
      background: #f8fafc;
      border: 1px solid #e2e8f0;
    }
    .meta-table td {
      padding: 4px 8px;
      font-size: 10pt;
      border: 1px solid #e2e8f0;
    }
    .label {
      background: #f1f5f9;
      font-weight: bold;
      color: #475569;
      width: 14%;
    }
    .val {
      font-weight: bold;
      width: 36%;
    }
    /* 6行分を確保するカラム */
    .columns {
      display: flex;
      gap: 10px;
      margin-bottom: 8px;
      height: 205px;
    }
    .col-box {
      flex: 1;
      border: 1.5px solid #cbd5e1;
      border-radius: 6px;
      padding: 6px 10px;
      background: #ffffff;
      overflow: hidden;
    }
    .col-title {
      font-weight: bold;
      font-size: 10.5pt;
      color: #1e40af;
      border-bottom: 1px dashed #cbd5e1;
      padding-bottom: 3px;
      margin-bottom: 5px;
    }
    .content-text {
      font-size: 11pt;
      font-weight: bold;
      color: #0f172a;
      white-space: pre-wrap;
      word-break: break-all;
      line-height: 1.4;
    }
    /* フッター：提出期限＋QR案内 */
    .footer-area {
      display: flex;
      gap: 10px;
      align-items: center;
      height: 160px;
    }
    .due-box {
      flex: 1;
      background: #fef2f2;
      border: 1.5px solid #ef4444;
      border-radius: 6px;
      padding: 6px 10px;
      text-align: center;
      display: flex;
      flex-direction: column;
      justify-content: center;
    }
    .due-label {
      font-size: 8.5pt;
      color: #991b1b;
      font-weight: bold;
    }
    .due-text {
      color: #b91c1c;
      font-size: 12pt;
      font-weight: bold;
    }
    .qr-box {
      flex: 1.3;
      align-self: flex-start;
      border: 1.5px solid #cbd5e1;
      border-radius: 6px;
      padding: 4px 8px;
      background: #ffffff;
      display: flex;
      align-items: center;
      gap: 8px;
    }
    .qr-img {
      width: 160px;
      height: 160px;
    }
    .qr-desc {
      display: flex;
      flex-direction: column;
    }
    .qr-title {
      font-size: 9.5pt;
      font-weight: bold;
      color: #1e3a8a;
    }
    .qr-sub {
      font-size: 8pt;
      color: #64748b;
    }
  </style>
  </head>
  <body>
    ${cardsHtml}
  </body>
  </html>
  `;
}

// ================= カンバン & 提出自動検知 =================

/**
 * 提出カンバンシートの自動初期化
 */
function getOrCreateKanbanSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.KANBAN_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.KANBAN_SHEET_NAME);
    const headers = [
      "授業ID", "授業日", "生徒名", "教科", "提出期限",
      "提出フォルダURL", "演習提出", "演習提出日時",
      "宿題提出", "宿題提出日時", "状況", "フォルダID"
    ];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setBackground("#1e3a8a").setFontColor("#ffffff").setFontWeight("bold");
    sheet.setFrozenRows(1);
    sheet.hideColumns(12); // フォルダID列（システム用）を非表示
  }
  return sheet;
}

/**
 * 【定期実行トリガー用】
 * 提出フォルダをスキャンし、「演習」「宿題」の格納を自動検知して打刻
 * - 演習・宿題の両方が揃ったら「提出済み」に更新
 * - 「丸付け済み」「返却済み」「完了」の行はスキップ（上書き防止）
 * - 初回提出時の日時を保持
 * - 提出フォルダをスキャンし、「演習」「宿題」の格納を自動検知して打刻 ＆ Discord通知
 */
function checkSubmissions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.KANBAN_SHEET_NAME);
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return;

  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const status = row[10];      // K列: 状況

    // すでに丸付け以降のフェーズに進んでいる行はスキップ
    if (status === "丸付け済み" || status === "返却済み" || status === "完了") {
      continue;
    }

    const studentName = row[2];  // C列: 生徒名
    const subject = row[3];      // D列: 教科
    const folderUrl = row[5];    // F列: 提出フォルダURL
    const folderId = row[11];    // L列: フォルダID

    let exerciseStatus = row[6]; // G列: 演習提出
    let homeworkStatus = row[8]; // I列: 宿題提出
    const currentExTime = row[7];// H列: 演習提出日時
    const currentHwTime = row[9];// J列: 宿題提出日時

    if (!folderId) continue;

    try {
      const folder = DriveApp.getFolderById(folderId);
      const files = folder.getFiles();

      while (files.hasNext()) {
        const file = files.next();
        const fileName = file.getName();
        const updatedTime = Utilities.formatDate(file.getLastUpdated(), "Asia/Tokyo", "yyyy/MM/dd HH:mm");

        // 演習ファイルの新規検知
        if (exerciseStatus !== "済" && exerciseStatus !== "免除" && fileName.includes("演習")) {
          sheet.getRange(r + 1, 7).setValue("済");
          if (!currentExTime) {
            sheet.getRange(r + 1, 8).setValue(updatedTime);
          }
          exerciseStatus = "済";

          // Discord通知を送信
          sendDiscordNotification({
            type: "演習",
            studentName: studentName,
            subject: subject,
            fileName: fileName,
            submittedTime: updatedTime,
            folderUrl: folderUrl
          });
        }

        // 宿題ファイルの新規検知
        if (homeworkStatus !== "済" && homeworkStatus !== "免除" && fileName.includes("宿題")) {
          sheet.getRange(r + 1, 9).setValue("済");
          if (!currentHwTime) {
            sheet.getRange(r + 1, 10).setValue(updatedTime);
          }
          homeworkStatus = "済";

          // Discord通知を送信
          sendDiscordNotification({
            type: "宿題",
            studentName: studentName,
            subject: subject,
            fileName: fileName,
            submittedTime: updatedTime,
            folderUrl: folderUrl
          });
        }
      }

      // 両方揃った場合は「提出済み」へ更新
      const exerciseComplete = exerciseStatus === "済" || exerciseStatus === "免除";
      const homeworkComplete = homeworkStatus === "済" || homeworkStatus === "免除";
      if (exerciseComplete && homeworkComplete && status !== "提出済み") {
        sheet.getRange(r + 1, 11).setValue("提出済み");
      }
    } catch (e) {
      Logger.log(`フォルダ走査エラー (ID: ${folderId}): ` + e.message);
    }
  }
}

/**
 * Discordへ提出通知をEmbed形式でPOSTする関数
 */
function sendDiscordNotification(info) {
  const webhookUrl = CONFIG.DISCORD_WEBHOOK_URL;
  if (!webhookUrl || webhookUrl.includes("ここに_コピーした")) {
    Logger.log("Discord Webhook URLが未設定です。");
    return;
  }

  // 演習は青(0x3b82f6)、宿題は緑(0x22c55e)で色分け
  const embedColor = info.type === "宿題" ? 0x22c55e : 0x3b82f6;

  const payload = {
    username: "CORE 提出管理Bot",
    embeds: [
      {
        title: `📥 【${info.type}提出】${info.studentName} さん`,
        color: embedColor,
        fields: [
          { name: "教科", value: info.subject || "—", inline: true },
          { name: "種別", value: info.type, inline: true },
          { name: "検知日時", value: info.submittedTime, inline: true },
          { name: "ファイル名", value: info.fileName, inline: false }
        ],
        description: `📂 [Googleドライブの提出フォルダを開く](${info.folderUrl})`,
        footer: {
          text: "個別指導CORE 提出管理"
        },
        timestamp: new Date().toISOString()
      }
    ]
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    UrlFetchApp.fetch(webhookUrl, options);
  } catch (e) {
    Logger.log("Discord通知エラー: " + e.message);
  }
}

/**
 * フォルダの存在チェック & なければ作成するユーティリティ
 */
function getOrCreateFolder(parentFolder, folderName) {
  const folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return parentFolder.createFolder(folderName);
}

/**
 * 手動テスト実行用
 */
function testSubmit() {
  onFormSubmitHandler(null);
}

function authDrive() {
  DriveApp.getRootFolder();
}
