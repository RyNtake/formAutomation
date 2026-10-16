const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const test = require("node:test");
const vm = require("node:vm");

const projectRoot = path.resolve(__dirname, "..");

function loadGasScript(fileName, services = {}) {
  const context = vm.createContext({
    console,
    Logger: { log() {} },
    SpreadsheetApp: {
      getActiveSpreadsheet() {
        throw new Error("SpreadsheetApp mock was not configured");
      }
    },
    DriveApp: {
      getFolderById() {
        throw new Error("DriveApp mock was not configured");
      }
    },
    UrlFetchApp: {
      fetch() {
        throw new Error("UrlFetchApp mock was not configured");
      }
    },
    Utilities: {
      formatDate() {
        throw new Error("Utilities mock was not configured");
      }
    },
    ...services
  });

  const source = fs.readFileSync(path.join(projectRoot, fileName), "utf8");
  vm.runInContext(source, context, { filename: fileName });
  return context;
}

class MockSheet {
  constructor(rows) {
    this.rows = rows;
    this.writes = [];
  }

  getDataRange() {
    return { getValues: () => this.rows };
  }

  getRange(row, column) {
    return {
      setValue: (value) => {
        this.rows[row - 1][column - 1] = value;
        this.writes.push({ row, column, value });
      }
    };
  }
}

test("form text uses the first present field and trims it", () => {
  const gas = loadGasScript("演習宿題を自動生成.js");

  assert.equal(gas.getFormText({ 演習1: "  p.12  ", 演習: "p.10" }, ["演習1", "演習"]), "p.12");
  assert.equal(gas.getFormText({ 演習1: null, 演習: " p.10 " }, ["演習1", "演習"]), "");
  assert.equal(gas.getFormText({ 演習: " p.10 " }, ["演習1", "演習"]), "p.10");
});

test("folder helper reuses an existing folder instead of creating a duplicate", () => {
  const gas = loadGasScript("演習宿題を自動生成.js");
  const existingFolder = { id: "existing" };
  let created = false;
  const parent = {
    getFoldersByName(name) {
      assert.equal(name, "生徒A");
      return {
        hasNext: () => true,
        next: () => existingFolder
      };
    },
    createFolder() {
      created = true;
      return {};
    }
  };

  assert.equal(gas.getOrCreateFolder(parent, "生徒A"), existingFolder);
  assert.equal(created, false);
});

test("card HTML includes student details and a page break between students", () => {
  const gas = loadGasScript("演習宿題を自動生成.js");
  const html = gas.generateCombinedCardHtml(
    [
      { lessonId: "L_20261015_生徒A", name: "生徒A", division: "通常", subject: "数学", textbook: "教材A", exercise: "p.10", homework: "p.11", qrUrl: "https://example.test/qr" },
      { lessonId: "L_20261015_生徒B", name: "生徒B", division: "通常", subject: "英語", textbook: "教材B", exercise: "p.20", homework: "p.21", qrUrl: "https://example.test/qr" }
    ],
    "2026年10月15日",
    "2026年10月20日"
  );

  assert.match(html, /生徒A さん/);
  assert.match(html, /生徒B さん/);
  assert.match(html, /page-break-after: always/);
  assert.match(html, /2026年10月20日/);
  assert.match(html, /\.footer-area\s*\{[^}]*height:\s*180px/s);
  assert.match(html, /\.qr-box\s*\{[^}]*align-self:\s*flex-start/s);
  assert.match(html, /\.qr-img\s*\{[^}]*width:\s*160px;[^}]*height:\s*160px/s);
});

test("submission scan records first submissions and advances the kanban status", () => {
  const submittedAt = new Date("2026-10-15T03:04:00.000Z");
  const sheet = new MockSheet([
    ["授業ID", "授業日", "生徒名", "教科", "提出期限", "提出フォルダURL", "演習提出", "演習提出日時", "宿題提出", "宿題提出日時", "状況", "フォルダID"],
    ["L_20261015_生徒A", "20261015", "生徒A", "数学", "", "https://drive.test/folder", "未", "", "未", "", "進行中", "folder-1"]
  ]);
  const files = [
    { getName: () => "演習.pdf", getLastUpdated: () => submittedAt },
    { getName: () => "宿題.pdf", getLastUpdated: () => submittedAt }
  ];
  const notifications = [];
  const gas = loadGasScript("演習宿題を自動生成.js", {
    SpreadsheetApp: { getActiveSpreadsheet: () => ({ getSheetByName: () => sheet }) },
    DriveApp: { getFolderById: () => ({ getFiles: () => filesIterator(files) }) },
    Utilities: {
      formatDate: () => "2026/10/15 12:04"
    },
    UrlFetchApp: {
      fetch(url, options) {
        notifications.push({ url, payload: JSON.parse(options.payload) });
      }
    }
  });

  gas.checkSubmissions();

  assert.equal(sheet.rows[1][6], "済");
  assert.equal(sheet.rows[1][7], "2026/10/15 12:04");
  assert.equal(sheet.rows[1][8], "済");
  assert.equal(sheet.rows[1][9], "2026/10/15 12:04");
  assert.equal(sheet.rows[1][10], "提出済み");
  assert.equal(notifications.length, 2);
  assert.equal(notifications[0].payload.embeds[0].fields[1].value, "演習");
  assert.equal(notifications[1].payload.embeds[0].fields[1].value, "宿題");
});

test("submission scan does not overwrite timestamps or completed statuses", () => {
  const sheet = new MockSheet([
    ["授業ID", "授業日", "生徒名", "教科", "提出期限", "提出フォルダURL", "演習提出", "演習提出日時", "宿題提出", "宿題提出日時", "状況", "フォルダID"],
    ["L_1", "", "生徒A", "", "", "", "済", "初回日時", "済", "初回日時", "提出済み", "folder-1"],
    ["L_2", "", "生徒B", "", "", "", "済", "以前の日時", "済", "以前の日時", "完了", "folder-2"]
  ]);
  let folderLookups = 0;
  const gas = loadGasScript("演習宿題を自動生成.js", {
    SpreadsheetApp: { getActiveSpreadsheet: () => ({ getSheetByName: () => sheet }) },
    DriveApp: {
      getFolderById() {
        folderLookups++;
        return { getFiles: () => filesIterator([]) };
      }
    },
    Utilities: { formatDate: () => "2026/10/15 12:04" }
  });

  gas.checkSubmissions();

  assert.equal(sheet.rows[1][7], "初回日時");
  assert.equal(sheet.rows[1][9], "初回日時");
  assert.equal(sheet.rows[2][10], "完了");
  assert.equal(folderLookups, 1);
});

function filesIterator(files) {
  let index = 0;
  return {
    hasNext: () => index < files.length,
    next: () => files[index++]
  };
}
