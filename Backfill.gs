/*************************************************
 * Backfill.gs — per-column backfill runner (MMCrawl) with row-by-row logging
 * Menu shows items = Backfill!A2:A (Column_ID values).
 * Each item runs backfill for THAT column only.
 * Reuses helpers from AI_Agent.gs:
 *   - AIA_callOpenAI_(userPrompt)
 *   - AIA_safeUi_()
 *   - getHeaderIndexSmart_(headers, name)
 * Tabs:
 *   - Backfill  (A: Column_ID | B: Prompt_Content | C: Result/logs)
 *   - MMCrawl   (target dataset)
 **************************************************/

/* ===== Menu on open ===== */
function onOpen_Backfill() {
  Backfill_buildMenu_();
}

/* Build the Backfill menu with items taken from Backfill!A2:A */
function Backfill_buildMenu_() {
  const ui =
    typeof AIA_safeUi_ === "function" ? AIA_safeUi_() : SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName("Backfill");
  if (!ui || !sh) return;

  const rows = Backfill_readRows_(); // [{id, prompt, row}]
  const menu = ui.createMenu("Backfill");

  const MAX_ITEMS = 20;
  const n = Math.min(rows.length, MAX_ITEMS);
  for (let i = 0; i < n; i++) {
    const label = rows[i].id;
    menu.addItem(label, "Backfill_run_" + (i + 1));
  }
  if (rows.length === 0) {
    menu.addItem("(No Column_ID rows found)", "Backfill_noop_");
  } else if (rows.length > MAX_ITEMS) {
    menu
      .addSeparator()
      .addItem(`Only first ${MAX_ITEMS} shown`, "Backfill_noop_");
  }

  menu
    .addSeparator()
    .addItem("↻ Refresh Backfill Menu", "Backfill_buildMenu_")
    .addItem("Test OpenAI", "Backfill_testOpenAI_")
    .addToUi();
}

function Backfill_noop_() {}

/* ===== Core runner (by Backfill row index) ===== */
function Backfill_runByIndex_(idx1) {
  const ui =
    typeof AIA_safeUi_ === "function" ? AIA_safeUi_() : SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const back = ss.getSheetByName("Backfill");
  const data = Backfill_readRows_();
  if (idx1 < 1 || idx1 > data.length) {
    if (ui) ui.alert("Invalid menu index.");
    return;
  }

  const { id: columnId, prompt, row: backfillRow } = data[idx1 - 1];

  const mm = ss.getSheetByName("MMCrawl");
  if (!mm) {
    if (ui) ui.alert('Tab "MMCrawl" not found.');
    return;
  }

  const lastRow = mm.getLastRow();
  const lastCol = mm.getLastColumn();
  if (lastRow <= 1) {
    if (ui) ui.alert("MMCrawl has no data rows.");
    return;
  }

  const headers = mm.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const targetColIdx =
    typeof getHeaderIndexSmart_ === "function"
      ? getHeaderIndexSmart_(headers, columnId)
      : headers.findIndex(
          (h) => String(h).toLowerCase() === String(columnId).toLowerCase()
        );
  if (targetColIdx < 0) {
    if (ui) ui.alert(`Header not found in MMCrawl: "${columnId}"`);
    return;
  }

  // Where do we pull URLs/names from for the after-run list?
  const idxUrl = Backfill_headerIndexAny_(headers, [
    "Public Website Homepage URL",
    "Website",
    "URL",
    "Homepage",
    "Company Website URL",
  ]);
  const idxDom = Backfill_headerIndexAny_(headers, [
    "Domain from URL",
    "Domain",
  ]);
  const idxName = Backfill_headerIndexAny_(headers, ["Company Name", "Name"]);

  const values = mm.getRange(2, 1, lastRow - 1, lastCol).getValues();

  // init log cell
  Backfill_writeCell_(
    back,
    backfillRow,
    3,
    `Backfilling "${columnId}" — ${new Date().toLocaleString()}\n`
  );

  let countUpdated = 0;
  let countSkipped = 0;
  let countErrors = 0;
  const updatedUrls = [];

  if (ui) ss.toast(`Backfilling "${columnId}"…`, "Backfill", 5);

  for (let r = 0; r < values.length; r++) {
    const absRow = 2 + r;
    const rowObj = {};
    headers.forEach((h, c) => (rowObj[String(h)] = values[r][c]));

    const urlForRow =
      (idxUrl >= 0 ? values[r][idxUrl] || "" : "") ||
      (idxDom >= 0 ? values[r][idxDom] || "" : "");
    const nameForRow = idxName >= 0 ? values[r][idxName] || "" : "";

    const userPrompt =
      String(prompt || "") +
      "\n\n### Input Row (MMCrawl JSON)\n" +
      JSON.stringify(rowObj, null, 2) +
      `\n\n### Output\n` +
      `Return **one JSON object only** with exactly one key — "${columnId}".\n` +
      `- If you cannot confidently determine a value, return {"${columnId}": ""}.\n` +
      `- For flags like "CNC 3-axis", "CNC 5-axis", "Spares/ Repairs", "Family business", "Medical": return "Yes" or "".\n` +
      `- For "Equipment": return a single-line normalized equipment string per rules, or the literal string "null" if none.\n` +
      `- No prose, no markdown, no code-fences.`;

    try {
      const ans =
        typeof AIA_callOpenAI_ === "function"
          ? AIA_callOpenAI_(userPrompt)
          : Backfill_callOpenAI_Fallback_(userPrompt);

      const obj = Backfill_parseOneJsonObject_(ans);
      let newVal =
        obj && Object.prototype.hasOwnProperty.call(obj, columnId)
          ? obj[columnId]
          : "";

      // normalize flags
      if (
        [
          "CNC 3-axis",
          "CNC 5-axis",
          "Spares/ Repairs",
          "Family business",
          "Medical",
        ].includes(columnId)
      ) {
        newVal = String(newVal || "").trim();
        newVal = /^yes$/i.test(newVal) ? "Yes" : "";
      }

      const curVal = String(values[r][targetColIdx] || "");
      const newStr = newVal == null ? "" : String(newVal);
      if (newStr && newStr !== curVal) {
        mm.getRange(absRow, targetColIdx + 1).setValue(newStr);
        countUpdated++;
        updatedUrls.push(
          (nameForRow ? nameForRow + " — " : "") + (urlForRow || "(no URL)")
        );
        Backfill_appendLog_(
          back,
          backfillRow,
          `Row ${absRow}: OK — updated to "${Backfill_preview_(newStr)}"`
        );
      } else {
        countSkipped++;
        Backfill_appendLog_(
          back,
          backfillRow,
          `Row ${absRow}: Skipped — no change or empty result`
        );
      }
    } catch (err) {
      countErrors++;
      Backfill_appendLog_(
        back,
        backfillRow,
        `Row ${absRow}: ERROR — ${String(err)}`
      );
    }
    Utilities.sleep(180);
  }

  // Final summary
  const summary =
    `\n— Run complete: "${columnId}" — ${new Date().toLocaleString()}` +
    `\nUpdated: ${countUpdated}  |  Skipped: ${countSkipped}  |  Errors: ${countErrors}` +
    (updatedUrls.length
      ? `\nUpdated URLs (${updatedUrls.length}):\n- ` + updatedUrls.join("\n- ")
      : `\nNo rows updated.`);

  Backfill_appendLog_(back, backfillRow, summary);

  if (ui) {
    ss.toast(
      `Backfill "${columnId}" finished — ${countUpdated} updated`,
      "Backfill",
      6
    );
    ui.alert(
      "Backfill complete",
      `Column: ${columnId}\nUpdated: ${countUpdated}\nSkipped: ${countSkipped}\nErrors: ${countErrors}` +
        (updatedUrls.length
          ? `\n\nUpdated URLs (first ${Math.min(
              updatedUrls.length,
              25
            )}):\n- ${updatedUrls.slice(0, 25).join("\n- ")}`
          : ""),
      ui.ButtonSet.OK
    );
  }
}

/* ===== Helpers ===== */

function Backfill_readRows_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName("Backfill");
  if (!sh) return [];
  const last = sh.getLastRow();
  if (last < 2) return [];
  const vals = sh.getRange(2, 1, last - 1, 3).getValues();
  const out = [];
  for (let i = 0; i < vals.length; i++) {
    const id = String(vals[i][0] || "").trim();
    const prompt = String(vals[i][1] || "").trim();
    if (!id) continue;
    out.push({ id, prompt, row: 2 + i });
  }
  return out;
}

function Backfill_parseOneJsonObject_(text) {
  if (!text) return {};
  let t = String(text).trim();
  const fence = t.match(/```json([\s\S]*?)```/i) || t.match(/```([\s\S]*?)```/);
  if (fence) t = fence[1].trim();
  try {
    return JSON.parse(t);
  } catch (_) {}
  const m = t.match(/\{[\s\S]*\}/);
  if (m) {
    try {
      return JSON.parse(m[0]);
    } catch (_) {}
  }
  return {};
}

function Backfill_callOpenAI_Fallback_(userPrompt) {
  const key =
    PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
  if (!key)
    throw new Error(
      "Missing OpenAI API key. Use “AI Integration → Set OpenAI API Key”."
    );
  const payload = {
    model: "gpt-4o",
    temperature: 0.0,
    max_tokens: 1200,
    messages: [
      {
        role: "system",
        content:
          "You are a precise data analyst. Read a single row of MMCrawl JSON and return one JSON object with exactly one key (the requested column). No prose, no markdown.",
      },
      { role: "user", content: userPrompt },
    ],
  };
  const resp = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", {
    method: "post",
    contentType: "application/json",
    muteHttpExceptions: true,
    headers: { Authorization: "Bearer " + key },
    payload: JSON.stringify(payload),
  });
  const code = resp.getResponseCode();
  const text = resp.getContentText();
  if (code < 200 || code >= 300) throw new Error(`HTTP ${code}: ${text}`);
  const data = JSON.parse(text);
  const out = data?.choices?.[0]?.message?.content;
  if (!out) throw new Error("No content from OpenAI.");
  return String(out).trim();
}

/* Write full value to Backfill!C (overwrite) */
function Backfill_writeCell_(backSheet, row, col, text) {
  if (!backSheet) return;
  backSheet.getRange(row, col).setValue(String(text || ""));
}

/* Append one line to Backfill!C with newline */
function Backfill_appendLog_(backSheet, row, line) {
  if (!backSheet) return;
  const cell = backSheet.getRange(row, 3);
  const cur = String(cell.getValue() || "");
  const next = cur ? cur + "\n" + line : line;
  cell.setValue(next);
}

/* Find first matching header by any of the candidate names (case-insensitive) */
function Backfill_headerIndexAny_(headers, names) {
  const lc = headers.map((h) => String(h).toLowerCase());
  for (let i = 0; i < names.length; i++) {
    const j = lc.indexOf(String(names[i]).toLowerCase());
    if (j >= 0) return j;
  }
  return -1;
}

/* Truncate long values for log */
function Backfill_preview_(s, n) {
  n = n || 80;
  s = String(s || "");
  return s.length <= n ? s : s.slice(0, n - 1) + "…";
}

/* Quick sanity check */
function Backfill_testOpenAI_() {
  const ui =
    typeof AIA_safeUi_ === "function" ? AIA_safeUi_() : SpreadsheetApp.getUi();
  try {
    const s = Backfill_callOpenAI_Fallback_('Return {"ok": true}');
    ui.alert("OpenAI OK", s, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert("OpenAI error", String(e), ui.ButtonSet.OK);
  }
}

/* ===== Static wrappers (support first 20 Backfill rows) ===== */
function Backfill_run_1() {
  Backfill_runByIndex_(1);
}
function Backfill_run_2() {
  Backfill_runByIndex_(2);
}
function Backfill_run_3() {
  Backfill_runByIndex_(3);
}
function Backfill_run_4() {
  Backfill_runByIndex_(4);
}
function Backfill_run_5() {
  Backfill_runByIndex_(5);
}
function Backfill_run_6() {
  Backfill_runByIndex_(6);
}
function Backfill_run_7() {
  Backfill_runByIndex_(7);
}
function Backfill_run_8() {
  Backfill_runByIndex_(8);
}
function Backfill_run_9() {
  Backfill_runByIndex_(9);
}
function Backfill_run_10() {
  Backfill_runByIndex_(10);
}
function Backfill_run_11() {
  Backfill_runByIndex_(11);
}
function Backfill_run_12() {
  Backfill_runByIndex_(12);
}
function Backfill_run_13() {
  Backfill_runByIndex_(13);
}
function Backfill_run_14() {
  Backfill_runByIndex_(14);
}
function Backfill_run_15() {
  Backfill_runByIndex_(15);
}
function Backfill_run_16() {
  Backfill_runByIndex_(16);
}
function Backfill_run_17() {
  Backfill_runByIndex_(17);
}
function Backfill_run_18() {
  Backfill_runByIndex_(18);
}
function Backfill_run_19() {
  Backfill_runByIndex_(19);
}
function Backfill_run_20() {
  Backfill_runByIndex_(20);
}
