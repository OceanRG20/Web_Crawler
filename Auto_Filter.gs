/** Auto_Filter.gs
 * Provides the Auto Filter menu and actions.
 * Call AutoFilter.addMenu(SpreadsheetApp.getUi()) from Ai_Agent.onOpen().
 */

/***** CONFIG *****/
const AUTO_FILTER_CONFIG = {
  SHEET_NAME: "MMCrawl",
  MODEL: "gpt-4o-mini",
  TEMPERATURE: 0.1,
  MAX_ROWS_PER_BATCH: 80,
  OPENAI_KEY_PROP: "OPENAI_API_KEY",
  MENU_TITLE: "Auto Filter",
  COL_FILTER_FLAG: "AI Filter Flag",
  COL_FAILED_CRITERIA: "AI Failed Criteria",
  PROP_LAST_PROMPT: "AF_LAST_PROMPT",
};

const AutoFilter = (() => {
  /** PUBLIC: adds the Auto Filter menu into UI */
  function addMenu(ui) {
    (ui || SpreadsheetApp.getUi())
      .createMenu(AUTO_FILTER_CONFIG.MENU_TITLE)
      .addItem("Open Filter Dialog…", "AutoFilter_openFilterDialog")
      .addItem("Run Last Filter", "AutoFilter_runLastFilter")
      .addSeparator()
      .addItem("Clear Filter Flags", "AutoFilter_clearFilterFlags")
      .addItem("Show All Rows (remove filter)", "AutoFilter_removeSheetFilter")
      .addToUi();
  }

  /** UI actions (global for menu wiring) **/
  function AutoFilter_openFilterDialog() {
    const lastPrompt =
      PropertiesService.getScriptProperties().getProperty(
        AUTO_FILTER_CONFIG.PROP_LAST_PROMPT
      ) || "";
    const html = HtmlService.createTemplateFromFile("FilterDialog");
    html.lastPrompt = lastPrompt;
    SpreadsheetApp.getUi().showSidebar(
      html.evaluate().setTitle("AI Auto Filter")
    );
  }

  function AutoFilter_runAutoFilterFromClient(promptText) {
    if (!promptText || !promptText.trim())
      throw new Error("Please enter a filter query.");
    PropertiesService.getScriptProperties().setProperty(
      AUTO_FILTER_CONFIG.PROP_LAST_PROMPT,
      promptText.trim()
    );
    return runAutoFilterInternal(promptText.trim());
  }

  function AutoFilter_runLastFilter() {
    const promptText = PropertiesService.getScriptProperties().getProperty(
      AUTO_FILTER_CONFIG.PROP_LAST_PROMPT
    );
    if (!promptText)
      throw new Error(
        "No previous filter found. Open the dialog to run a new one."
      );
    return runAutoFilterInternal(promptText);
  }

  function AutoFilter_clearFilterFlags() {
    const sh = getSheet();
    const headers = sh
      .getRange(1, 1, 1, sh.getLastColumn())
      .getValues()[0]
      .map(String);
    const idxFlag = headers.indexOf(AUTO_FILTER_CONFIG.COL_FILTER_FLAG);
    const idxFail = headers.indexOf(AUTO_FILTER_CONFIG.COL_FAILED_CRITERIA);
    if (idxFlag === -1 && idxFail === -1) return;

    const rng = sh.getRange(
      2,
      1,
      Math.max(0, sh.getLastRow() - 1),
      sh.getLastColumn()
    );
    if (rng.getNumRows() === 0) return;
    const vals = rng.getValues();
    for (let r = 0; r < vals.length; r++) {
      if (idxFlag !== -1) vals[r][idxFlag] = "";
      if (idxFail !== -1) vals[r][idxFail] = "";
    }
    rng.setValues(vals);
  }

  function AutoFilter_removeSheetFilter() {
    const sh = getSheet();
    const f = sh.getFilter();
    if (f) f.remove();
  }

  /** CORE **/
  function runAutoFilterInternal(promptText) {
    const key = PropertiesService.getScriptProperties().getProperty(
      AUTO_FILTER_CONFIG.OPENAI_KEY_PROP
    );
    if (!key) throw new Error("Missing OPENAI_API_KEY in Script properties.");

    const sh = getSheet();
    const dataRange = sh.getDataRange();
    const values = dataRange.getValues();
    if (values.length < 2) throw new Error("No data rows found.");

    const headers = values[0].map(String);
    ensureColumnExists(sh, headers, AUTO_FILTER_CONFIG.COL_FILTER_FLAG);
    ensureColumnExists(sh, headers, AUTO_FILTER_CONFIG.COL_FAILED_CRITERIA);
  }

  /** === NEW: run filtering batch-by-batch with live progress === **/
  function AutoFilter_runAutoFilterBatch(promptText, batchIndex) {
    const key = PropertiesService.getScriptProperties().getProperty(
      AUTO_FILTER_CONFIG.OPENAI_KEY_PROP
    );
    if (!key) throw new Error("Missing OPENAI_API_KEY in Script properties.");

    const sh = getSheet();
    const hdrs = sh
      .getRange(1, 1, 1, sh.getLastColumn())
      .getValues()[0]
      .map(String);
    const idxFlag = hdrs.indexOf(AUTO_FILTER_CONFIG.COL_FILTER_FLAG);
    const idxFail = hdrs.indexOf(AUTO_FILTER_CONFIG.COL_FAILED_CRITERIA);

    const totalRows = Math.max(0, sh.getLastRow() - 1);
    const rows = sh.getRange(2, 1, totalRows, sh.getLastColumn()).getValues();
    const records = rows.map((r, i) => {
      const obj = { __rowNumber: i + 2 };
      hdrs.forEach((h, c) => (obj[h] = r[c] ?? ""));
      return obj;
    });

    const batches = chunk(records, AUTO_FILTER_CONFIG.MAX_ROWS_PER_BATCH);
    if (batchIndex >= batches.length) {
      applyYesFilter(sh, idxFlag + 1);
      const flagVals = sh
        .getRange(2, idxFlag + 1, totalRows, 1)
        .getValues()
        .flat();
      const yesCount = flagVals.filter(
        (v) => String(v).trim().toLowerCase() === "yes"
      ).length;
      const percent = totalRows ? ((yesCount / totalRows) * 100).toFixed(1) : 0;
      return {
        done: true,
        progress: 100,
        message: `Filtering complete. Showing ${yesCount} of ${totalRows} rows (${percent}%).`,
      };
    }

    const sys = buildSystemPrompt();
    const user = buildUserMessage(promptText, hdrs, batches[batchIndex]);
    const out = callOpenAI(key, sys, user);

    let parsed;
    try {
      parsed = JSON.parse(out);
    } catch (e) {
      throw new Error(`OpenAI returned non-JSON for batch ${batchIndex + 1}`);
    }

    const results = new Map();
    parsed.forEach((rec) => {
      const rn = Number(rec.__rowNumber);
      const flag = String(rec.MeetsCriteria ?? rec.meets_criteria ?? "").trim();
      const failed = String(
        rec.FailedCriteria ?? rec.failed_criteria ?? ""
      ).trim();
      if (!isFinite(rn)) return;
      results.set(rn, {
        flag: /^yes$/i.test(flag) ? "Yes" : "No",
        failed: failed || (/^yes$/i.test(flag) ? "" : "Unspecified"),
      });
    });

    // write results for this batch only
    const startRow = batchIndex * AUTO_FILTER_CONFIG.MAX_ROWS_PER_BATCH + 2;
    const endRow = Math.min(
      startRow + AUTO_FILTER_CONFIG.MAX_ROWS_PER_BATCH - 1,
      sh.getLastRow()
    );
    const batchRange = sh.getRange(
      startRow,
      1,
      endRow - startRow + 1,
      sh.getLastColumn()
    );
    const batchVals = batchRange.getValues();
    for (let r = 0; r < batchVals.length; r++) {
      const rn = startRow + r;
      const res = results.get(rn);
      if (res) {
        batchVals[r][idxFlag] = res.flag;
        batchVals[r][idxFail] = res.failed;
      }
    }
    batchRange.setValues(batchVals);

    const progress = Math.round(((batchIndex + 1) / batches.length) * 100);
    return {
      done: false,
      progress,
      message: `Filtering ${progress}% complete...`,
    };
  }

  /** Helpers **/
  function getSheet() {
    const ss = SpreadsheetApp.getActive();
    return (
      ss.getSheetByName(AUTO_FILTER_CONFIG.SHEET_NAME) || ss.getActiveSheet()
    );
  }

  function ensureColumnExists(sh, headers, name) {
    if (!headers.includes(name)) {
      sh.insertColumnAfter(headers.length);
      sh.getRange(1, headers.length + 1, 1, 1).setValue(name);
    }
  }

  function chunk(arr, size) {
    const out = [];
    for (let i = 0; i < arr.length; i += size) out.push(arr.slice(i, i + size));
    return out;
  }

  function buildSystemPrompt() {
    return `
You are an AI analyst evaluating rows from the MMCrawl sheet.
Many fields are fuzzy (ranges like "5–20M", "~12M"). Do NOT rewrite or harden fuzzy data.

Task: For each row, decide if it meets the client's natural-language criteria.
Use a permissive bias when ambiguity remains but evidence leans toward inclusion.

Return STRICT JSON ONLY: an array of objects with
{
  "__rowNumber": <sheet row number>,
  "MeetsCriteria": "Yes" | "No",
  "FailedCriteria": "<comma-separated reasons, empty if Yes>"
}
No commentary outside JSON.
    `.trim();
  }

  function buildUserMessage(filterPrompt, headers, batch) {
    return JSON.stringify({
      instructions:
        "Evaluate each record against the natural-language filter. Use permissive bias on edge cases.",
      filter_prompt: filterPrompt,
      expected_fields: headers,
      rows: batch,
    });
  }

  function callOpenAI(apiKey, systemPrompt, userMessage) {
    const url = "https://api.openai.com/v1/chat/completions";
    const payload = {
      model: AUTO_FILTER_CONFIG.MODEL,
      temperature: AUTO_FILTER_CONFIG.TEMPERATURE,
      messages: [
        { role: "system", content: systemPrompt },
        { role: "user", content: userMessage },
      ],
      response_format: { type: "json_object" },
    };
    const res = UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      headers: { Authorization: `Bearer ${apiKey}` },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    });
    if (res.getResponseCode() >= 300)
      throw new Error(
        `OpenAI HTTP ${res.getResponseCode()}: ${res.getContentText()}`
      );

    const data = JSON.parse(res.getContentText());
    const content = data.choices?.[0]?.message?.content;
    if (!content) throw new Error("OpenAI returned empty content.");

    // accept {"results":[...]} or [...] or single object
    try {
      const obj = JSON.parse(content);
      if (Array.isArray(obj)) return content;
      if (Array.isArray(obj.results)) return JSON.stringify(obj.results);
      return JSON.stringify([obj]);
    } catch (_) {
      return content;
    }
  }

  function applyYesFilter(sh, colIndex1) {
    const filter =
      sh.getFilter() ||
      sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).createFilter();
    const criteria = SpreadsheetApp.newFilterCriteria()
      .setHiddenValues(["", "No"])
      .build();
    filter.setColumnFilterCriteria(colIndex1, criteria);
  }

  // expose public names for menu bindings
  return {
    addMenu,
    AutoFilter_openFilterDialog,
    AutoFilter_runAutoFilterFromClient,
    AutoFilter_runAutoFilterBatch,
    AutoFilter_runLastFilter,
    AutoFilter_clearFilterFlags,
    AutoFilter_removeSheetFilter,
  };
})();

/** global wrappers so menu items can call them **/
function AutoFilter_openFilterDialog() {
  AutoFilter.AutoFilter_openFilterDialog();
}
function AutoFilter_runAutoFilterFromClient(promptText) {
  return AutoFilter.AutoFilter_runAutoFilterFromClient(promptText);
}
function AutoFilter_runAutoFilterBatch(promptText, batchIndex) {
  return AutoFilter.AutoFilter_runAutoFilterBatch(promptText, batchIndex);
}
function AutoFilter_runLastFilter() {
  return AutoFilter.AutoFilter_runLastFilter();
}
function AutoFilter_clearFilterFlags() {
  return AutoFilter.AutoFilter_clearFilterFlags();
}
function AutoFilter_removeSheetFilter() {
  return AutoFilter.AutoFilter_removeSheetFilter();
}
