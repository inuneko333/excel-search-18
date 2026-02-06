const SETTINGS_KEY = "pageSearchSettings_v2";

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
  } catch {
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
  } catch {
    // 無視してOK
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
      return 0; // 想定外 → A列扱い
    }
    col = col * 26 + (code - 64);
  }
  return col - 1; // 0基準
}

function initUi() {
  const searchInput = document.getElementById("searchTerm");
  const searchButton = document.getElementById("searchButton");
  const statusMessage = document.getElementById("statusMessage");
  const pageHighlight = document.getElementById("pageHighlight");
  const resultMeta = document.getElementById("resultMeta");
  const resultDetails = document.getElementById("resultDetails");
  const resultPages = document.getElementById("resultPages");

  const toggleSettingsButton = document.getElementById("toggleSettings");
  const settingsPanel = document.getElementById("settingsPanel");
  const columnInput = document.getElementById("columnInput");
  const pageSizeInput = document.getElementById("pageSizeInput");
  const skipRowsInput = document.getElementById("skipRowsInput");

  // 設定読み込み
  const settings = loadSettings();
  columnInput.value = settings.columnLetter;
  pageSizeInput.value = settings.pageSize;
  skipRowsInput.value = settings.skipRows;

  function updateSettingsFromInputs() {
    const s = {
      columnLetter: (columnInput.value || "A").trim(),
      pageSize: Number(pageSizeInput.value) || 18,
      skipRows: Number(skipRowsInput.value) || 0,
    };
    saveSettings(s);
    return s;
  }

  // 設定表示/非表示
  let settingsVisible = false;
  toggleSettingsButton.addEventListener("click", () => {
    settingsVisible = !settingsVisible;
    if (settingsVisible) {
      settingsPanel.classList.remove("hidden");
      toggleSettingsButton.textContent = "▼ 詳細設定を隠す";
    } else {
      settingsPanel.classList.add("hidden");
      toggleSettingsButton.textContent = "▶ 詳細設定を表示";
    }
  });

  // ページ強調パネルを更新
  function renderPageHighlight(firstHit, hitCount) {
    const numberEl = pageHighlight.querySelector(".page-number");
    const subinfoEl = pageHighlight.querySelector(".page-subinfo");

    if (!firstHit) {
      pageHighlight.classList.add("page-highlight--empty");
      numberEl.textContent = "–";
      subinfoEl.textContent = "ヒットなし";
      return;
    }

    pageHighlight.classList.remove("page-highlight--empty");
    numberEl.textContent = firstHit.page;
    subinfoEl.textContent =
      `最初のヒット: 行 ${firstHit.actualRow}（有効行 ${firstHit.logicalRow}） / ヒット ${hitCount} 件`;
  }

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
    resultMeta.textContent = "";
    resultDetails.innerHTML = "";
    resultPages.innerHTML = "";
    renderPageHighlight(null, 0);

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

        if (colOffset >= 0 && colOffset < colCount) {
          for (let r = 0; r < rowCount; r++) {
            const actualRowNumber = startRowIndex + r + 1; // Excel行番号（1始まり）
            if (actualRowNumber <= skipRows) continue;

            const cellValue = values[r][colOffset];
            if (cellValue === null || cellValue === undefined) continue;

            const cellText = String(cellValue).trim();
            if (cellText === searchValue) {
              const logicalRow = actualRowNumber - skipRows; // 有効行（1始まり）
              const page = Math.ceil(logicalRow / pageSize);
              hits.push({
                actualRow: actualRowNumber,
                logicalRow,
                page,
              });
            }
          }
        }

        // ヒットなし
        if (hits.length === 0) {
          statusMessage.textContent = "";
          renderPageHighlight(null, 0);
          resultMeta.textContent = "";
          resultDetails.textContent = "";
          resultPages.textContent = "";
          return;
        }

        // 先頭ヒット（いちばん上）
        hits.sort((a, b) => a.actualRow - b.actualRow);
        const firstHit = hits[0];

        // 先頭ヒットのセルへジャンプ
        const upperColumnLetter = (columnLetter || "A").toUpperCase().trim() || "A";
        const firstAddress = `${upperColumnLetter}${firstHit.actualRow}`;
        const firstRange = sheet.getRange(firstAddress);
        firstRange.select();

        // ページ強調
        renderPageHighlight(firstHit, hits.length);

        // メタ情報（小さめ）
        resultMeta.textContent =
          `列: ${upperColumnLetter} / 除外行: ${skipRows} / 1ページ: ${pageSize} 行`;

        // ヒット一覧
        const detailLines = hits.map((h) => {
          return `行 ${h.actualRow}（有効行 ${h.logicalRow}） → ${h.page} ページ目`;
        });
        resultDetails.innerHTML = detailLines
          .map((line) => `<div class="result-line">${line}</div>`)
          .join("");

        // ページ別集計
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
          .map(Number)
          .sort((a, b) => a - b);

        const pageLines = sortedPages.map((p) => {
          const info = byPage[p];
          const rowsText = info.rows.join(", ");
          return `${p} ページ目: ${info.count} 件（行 ${rowsText}）`;
        });

        resultPages.innerHTML = pageLines
          .map((line) => `<div class="result-line">${line}</div>`)
          .join("");

        await context.sync();
        statusMessage.textContent = "";
      });
    } catch (error) {
      console.error(error);
      statusMessage.textContent = "エラーが発生しました。";
    }
  }

  // ボタンクリック
  searchButton.addEventListener("click", () => {
    runSearch();
  });

  // Enterで検索
  searchInput.addEventListener("keydown", (ev) => {
    if (ev.key === "Enter") {
      runSearch();
    }
  });

  // 初期フォーカス
  searchInput.focus();
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    initUi();
  }
});
