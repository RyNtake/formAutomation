# Project Context & Coding Guidelines

## Overview
個別指導塾COREの指導報告・宿題提出管理・授業準備自動化システム。
Google Apps Script (GAS), Google Spreadsheet, Google Drive, Google Forms, Discord Webhook, および Make / Notion / LINE API による疎結合システム。

## Architecture Rules
1. **Google Apps Script (GAS) Conventions**:
   - ECMAScript 6 (ES6) 構文を使用する（const/let、アロー関数、テンプレートリテラル等）。
   - 設定値（フォルダID、シート名、Webhook URL等）は `CONFIG` オブジェクトに集約し、ハードコードしない。
   - スプレッドシート操作は `getValues()` / `setValues()` によるバッチ読み書きを行い、ループ内での個別API呼び出しを避ける。
   - `DriveApp` や `SpreadsheetApp` の例外ハンドリング（try-catch、Logger.log）を徹底する。

2. **Web App (GAS HtmlService)**:
   - Webアプリは `doGet(e)` で `HtmlService.createTemplateFromFile('index')` を使用し、`e.parameter`（`folderId`, `name`）をテンプレート変数に展開する。
   - フロントエンド（index.html）からバックエンドの関数呼び出しは `google.script.run.withSuccessHandler(...).withFailureHandler(...)` を使用する。
   - アップロード処理はBase64文字列として受け取り、`Utilities.base64Decode` でBlob化してDrive保存する。

3. **Status Flow & Immutability**:
   - カンバンの状況遷移: `進行中` ➔ `提出済み` ➔ `丸付け済み` ➔ `返却済み` ➔ `完了`。
   - 演習提出・宿題提出の `免除` は提出不要を示し、提出済み相当として扱うこと。
   - `checkSubmissions` 等の自動監視ロジックは、初回提出日時を保持し上書きしないこと。
   - `丸付け済み`, `返却済み`, `完了` のステータスを持つ行は自動検知処理でスキップ（巻き戻し防止）すること。

4. **Output Requirements**:
   - コード修正や新機能提案の際は、変更対象の関数・差分（Diff）を明確にして提示すること。
