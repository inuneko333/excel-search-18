/* global Office, Excel */
let ready = false;

function $(id) { return document.getElementById(id); }

function normalize5(str) {
  const s = String(str ?? "").trim();
  if (s === "") return "";
  // Numeric-ish (e.g., 12345, 12345.0)
  if (/^\d+(\.0+)?$/.test(s)) {
    return String(parseInt(s, 10)).padStart(5, "0");
  }
  // Already looks like 5-digit string
  if (/^\d{1,5}$/.test(s)) return s.padStart(5, "0");
  return s;
}

function setError(msg) {
  const el = $("err");
  if (!msg) { el.style.display = "none"; el.textContent = ""; return; }
  el.textContent = msg;
  el.style.display = "block";
}

function setStatus(msg) {
  $("status").textContent = msg;
}

function setResults(hits) {
  $("hits").textContent = String(hits);
  const val = hits / 18;
  // show with up to 4 decimals but trim trailing zeros
  const s = val.toFixed(4).replace(/\.?(0+)$/, (m, zs) => m.startsWith(".") ? "" : "");
  $("pages").textContent = s;
}

async function countInActiveSheetA(query5) {
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const used = sheet.getUsedRangeOrNullObject();
    used.load(["rowCount", "isNullObject"]);
    await context.sync();

    if (used.isNullObject || used.rowCount === 0) {
      return { hits: 0, rows: 0 };
    }

    const colA = sheet.getRangeByIndexes(0, 0, used.rowCount, 1);
    colA.load("values");
    await context.sync();

    const values = colA.values;
    let hits = 0;
    for (let i = 0; i < values.length; i++) {
      const cell = values[i][0];
      if (cell === null || cell === undefined || cell === "") continue;
      const norm = normalize5(cell);
      if (norm === query5) hits++;
    }
    return { hits, rows: values.length };
  });
}

async function onRun() {
  setError("");
  const q = normalize5($("query").value);
  if (!/^\d{5}$/.test(q)) {
    setError("5桁の数字を入力してください（例：06328）。");
    return;
  }

  $("run").disabled = true;
  setStatus("カウント中…");
  const t0 = performance.now();

  try {
    const { hits, rows } = await countInActiveSheetA(q);
    setResults(hits);
    const ms = Math.round(performance.now() - t0);
    setStatus(`完了：A列の使用範囲 ${rows} 行をチェック（${ms}ms）`);
  } catch (e) {
    console.error(e);
    setError("エラー：Excelとの通信に失敗しました。アドインを閉じて開き直すと直ることがあります。");
    setStatus("—");
  } finally {
    $("run").disabled = false;
  }
}

function onInput() {
  setError("");
  const q = normalize5($("query").value);
  $("run").disabled = !ready || !/^\d{5}$/.test(q);
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    ready = true;
    setStatus("準備OK。番号を入力してください。");
    $("run").disabled = true;
    $("query").addEventListener("input", onInput);
    $("query").addEventListener("keydown", (ev) => {
      if (ev.key === "Enter") onRun();
    });
    $("run").addEventListener("click", onRun);
    $("query").focus();
  } else {
    setStatus("このアドインはExcelでのみ動作します。");
  }
});
