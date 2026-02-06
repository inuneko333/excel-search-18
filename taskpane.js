// 設定保存用キー
const SETTINGS_KEY = "pageSearchSettings_v1";

function loadSettings() {
  try {
    const raw = window.localStorage.getItem(SETTINGS_KEY);
    if (!raw) {
      return {
        columnLetter: "A",
        pageSize: 18,
        skipRows: 0,
      };
    }
    const obj = JSON.parse(raw);
    return {
      columnLetter: obj.columnLetter || "A",
      pageSize: Number(obj.pageSize) || 18,
      skipRows: Number(obj.skipRows) || 0,
    };
  } catch (e) {
    console.log("settings load error", e);
    return {
      columnLetter: "A",
      pageSize: 18,
      skipRows: 0,
    };
  }
}

function saveSettings(settings) {
  try {
    window.localStorage.setItem(SETTINGS_KEY, JSON.stringify(settings));
  } catch (e) {
    console.log("settings save error", e);
  }
}

// 列名（A, B, AA...）→ 0始まりの列インデックス
function columnLetterToIndex(letter) {
  if (!letter) return 0;
  const s = letter.toUpperCase().trim();
  let col = 0;
  for (let i = 0; i < s.length; i++) {
    const code = s.charCodeAt(i);
    if (code < 65 || code > 90) {
      // A〜Z以外 → とりあえず0列（A列）扱い
      return 0;
    }
    col = col * 26 + (code - 64); // A=1, B=2 ...
  }
  return col - 1; // 0始まり
}

function initUi() {
  const searchInput = document.getElementById("searchTerm");
  const searchButton = document.getElementById("searchButton");
  const statusMessage = document.getElementById("statusMessage");
  const resultSummary = document.getElementById("resultSummary");
  const resultDetails = document.getElementById("resultDetails");
  const resultPages = document.getElementById("resultPages");

  const toggleSettingsButton = document.getElementById("toggleSettings");
  const settingsPanel = document.getElementById("settingsPanel");
  const columnInput = document.getElementById("columnInput");
  const pageSizeInput = document.getElementById("pageSizeInput");
  const skipRowsInput = document.getElementById("skipRowsInput");

  // 設定読み込み＆UI反映
  const settings = loadSettings();
  columnInput.value = settings.columnLetter;
  pageSizeInput.value = settings.pageSize;
  skipRowsInput.value = settings.skipRows;

  function updateSettingsFromInputs() {
    const newSettings = {
      columnLetter: columnInput.value.trim() || "A",
      pageSize: Number(pageSizeInput.value) || 18,
      skipRows: Number(skipRowsInput.value) || 0,
    };
    saveSettings(newSettings);
    return newSettings;
  }

  // 設定エリア表示/非表示
  let settingsVisible = false;
  toggleSettingsButton.addEventListener("click", () => {
    settingsVisible = !settingsVisible;
    if (settingsVisible) {
      settingsPanel.classList.remove("hidden");
      toggleSettingsButton.textContent = "▼ 設定を隠す";
    } else {
      settingsPanel.classList.add("hidden");
      toggleSettingsButton.textContent = "▶ 設定を表示";
    }
  });

  // 検索実行処理
  async function runSearch() {
    const termRaw = searchInput.value.trim();
    if (!termRaw) {
      statusMessage.textContent = "検索値を入力してください。";
      return;
    }

    const currentSettings = updateSettingsFromInputs();
    const searchValue = termRaw;
    const columnLetter = currentSettings.columnLetter;
    const pageSize = currentSettings.pageSize;
    const skipRows = currentSettings.skipRows;

    statusMessage.textContent = "検索中…";
    resultSummary.textContent = "";
    resultDetails.textContent = "";
    resultPages.textContent = "";

    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const usedRange = sheet.getUsedRange();
        usedRange.load(["values", "rowCount", "columnCount", "rowIndex", "columnIndex"]);
        await context.sync();

        const values = usedRange.values;
        const rowCount = usedRange.rowCount;
        const colCount = usedRange.columnCount;
        const startRowIndex = usedRange.rowIndex;    // 0ベース
        const startColIndex = usedRange.columnIndex; // 0ベース

        const targetColIndex = columnLetterToIndex(columnLetter);
        const colOffset = targetColIndex - startColIndex;

        const hits = [];

        if (colOffset < 0 || colOffset >= colCount) {
          // 使用範囲の外を指定していた場合 → ヒット0
          // 何もしない
        } else {
          for (let r = 0; r < rowCount; r++) {
            const actualRowNumber = startRowIndex + r + 1; // Excel行番号（1始まり）
            const cellValue = values[r][colOffset];

            if (actualRowNumber <= skipRows) {
              continue; // 除外する先頭行
            }

            // 空セルはスキップ
            if (cellValue === null || cellValue === undefined) continue;

            const cellText = String(cellValue).trim();

            if (cellText === searchValue) {
              const logicalRow = actualRowNumber - skipRows; // 有効行番号（1始まり）
              const page = Math.ceil(logicalRow / pageSize);
              hits.push({
                actualRow: actualRowNumber,
                logicalRow,
                page,
              });
            }
          }
        }

        // 結果反映
        if (hits.length === 0) {
          statusMessage.textContent = "";
          resultSummary.textContent = "ヒットなし";
          resultDetails.textContent = "";
          resultPages.textContent = "";
          return;
        }

        statusMessage.textContent = "";

        // サマリ
        resultSummary.textContent = `ヒット: ${hits.length} 件`;

        // 各ヒット一覧
        const detailLines = hits.map((h) => {
          return `行 ${h.actualRow}（有効行 ${h.logicalRow}） → ${h.page} ページ目`;
        });
        resultDetails.innerHTML = detailLines
          .map((line) => `<div class="result-line">${line}</div>`)
          .join("");

        // ページごとの集計
        const byPage = {};
        for (const h of hits) {
          if (!byPage[h.page]) {
            byPage[h.page] = {
              count: 0,
              rows: [],
            };
          }
          byPage[h.page].count++;
          byPage[h.page].rows.push(h.actualRow);
        }

        const sortedPages = Object.keys(byPage)
          .map((p) => Number(p))
          .sort((a, b) => a - b);

        const pageLines = sortedPages.map((p) => {
          const info = byPage[p];
          const rowsText = info.rows.join(", ");
          return `${p} ページ目: ${info.count} 件（行 ${rowsText}）`;
        });

        resultPages.innerHTML = pageLines
          .map((line) => `<div class="result-line">${line}</div>`)
          .join("");
      });
    } catch (error) {
      console.error(error);
      statusMessage.textContent = "エラーが発生しました。コンソールを確認してください。";
    }
  }

  // ボタンクリック
  searchButton.addEventListener("click", () => {
    runSearch();
  });

  // Enter押下で検索
  searchInput.addEventListener("keydown", (ev) => {
    if (ev.key === "Enter") {
      runSearch();
    }
  });

  // 初期フォーカス
  searchInput.focus();
}

// Office 初期化
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    initUi();
  }
});
