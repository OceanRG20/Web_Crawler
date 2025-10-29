/** Auto_Filter.gs — guarded singleton
 * Provides the Auto Filter menu and actions.
 * Safe against duplicate loads (no “already been declared” errors).
 */

/***** CONFIG (guarded) *****/
if (typeof globalThis.AUTO_FILTER_CONFIG === "undefined") {
  globalThis.AUTO_FILTER_CONFIG = {
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
}

/***** MODULE (guarded) *****/
if (typeof globalThis.AutoFilter === "undefined") {
  globalThis.AutoFilter = (() => {
    /** PUBLIC: adds the Auto Filter menu into UI */
    function addMenu(ui) {
      (ui || SpreadsheetApp.getUi())
        .createMenu(AUTO_FILTER_CONFIG.MENU_TITLE)
        .addItem("Open Filter Dialog…", "AutoFilter_openFilterDialog")
        .addItem("Run Last Filter", "AutoFilter_runLastFilter")
        .addSeparator()
        .addItem("Clear Filter Flags", "AutoFilter_clearFilterFlags")
        .addItem(
          "Show All Rows (remove filter)",
          "AutoFilter_removeSheetFilter"
        )
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

      // data
      const dataRange = sh.getDataRange();
      const values = dataRange.getValues();
      if (values.length < 2) throw new Error("No data rows found.");

      const headers = values[0].map(String);
      ensureColumnExists(sh, headers, AUTO_FILTER_CONFIG.COL_FILTER_FLAG);
      ensureColumnExists(sh, headers, AUTO_FILTER_CONFIG.COL_FAILED_CRITERIA);

      // refresh header indices after possible insertion
      const hdrs = sh
        .getRange(1, 1, 1, sh.getLastColumn())
        .getValues()[0]
        .map(String);
      const idxFlag = hdrs.indexOf(AUTO_FILTER_CONFIG.COL_FILTER_FLAG);
      const idxFail = hdrs.indexOf(AUTO_FILTER_CONFIG.COL_FAILED_CRITERIA);

      const rows = sh
        .getRange(2, 1, Math.max(0, sh.getLastRow() - 1), sh.getLastColumn())
        .getValues();
      const records = rows.map((r, i) => {
        const obj = { __rowNumber: i + 2 }; // sheet row
        hdrs.forEach((h, c) => {
          obj[h] = r[c] === undefined ? "" : r[c];
        });
        return obj;
      });

      const batches = chunk(records, AUTO_FILTER_CONFIG.MAX_ROWS_PER_BATCH);
      const results = new Map();

      let yesCount = 0;
      let noCount = 0;

      batches.forEach((batch, bi) => {
        const sys = buildSystemPrompt();
        const user = buildUserMessage(promptText, hdrs, batch);
        const out = callOpenAI(key, sys, user);

        let parsed;
        try {
          parsed = JSON.parse(out);
        } catch (e) {
          throw new Error(
            `OpenAI returned non-JSON for batch ${bi + 1}.\nRaw:\n${out}`
          );
        }

        parsed.forEach((rec) => {
          const rn = Number(rec.__rowNumber);
          const flagRaw =
            rec.MeetsCriteria ?? rec.meets_criteria ?? rec.flag ?? "";
          const flag = /^yes$/i.test(String(flagRaw).trim()) ? "Yes" : "No";
          const failed =
            String(rec.FailedCriteria ?? rec.failed_criteria ?? "").trim() ||
            (flag === "Yes" ? "" : "Unspecified");
          if (!isFinite(rn)) return;
          results.set(rn, { flag, failed });
          if (flag === "Yes") yesCount++;
          else noCount++;
        });
      });

      // write back
      const outRange = sh.getRange(
        2,
        1,
        Math.max(0, sh.getLastRow() - 1),
        sh.getLastColumn()
      );
      if (outRange.getNumRows() > 0) {
        const outVals = outRange.getValues();
        for (let r = 0; r < outVals.length; r++) {
          const rn = r + 2;
          const res = results.get(rn) || { flag: "", failed: "" };
          outVals[r][idxFlag] = res.flag;
          outVals[r][idxFail] = res.failed;
        }
        outRange.setValues(outVals);
      }

      // filter to Yes
      applyYesFilter(sh, idxFlag + 1);
      return {
        ok: true,
        message: `Filtering complete. Showing rows with AI Filter Flag = Yes.  ✅ Yes: ${yesCount}  ❌ No: ${noCount}`,
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
      for (let i = 0; i < arr.length; i += size)
        out.push(arr.slice(i, i + size));
      return out;
    }

    function buildSystemPrompt() {
      return `
You are an AI analyst evaluating rows from the MMCrawl sheet.
Many fields are fuzzy (ranges like "5–20M", "~12M"). Do NOT rewrite or harden fuzzy data.

Task: For each row, decide if it meets the client's natural-language criteria.
Use a permissive bias when ambiguity remains but evidence leans toward inclusion.

Return STRICT JSON ONLY: an array of objects with
[
  {
    "__rowNumber": <sheet row number>,
    "MeetsCriteria": "Yes" | "No",
    "FailedCriteria": "<comma-separated reasons, empty if Yes>"
  },
  ...
]
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
      AutoFilter_runLastFilter,
      AutoFilter_clearFilterFlags,
      AutoFilter_removeSheetFilter,
    };
  })();
}

/***** Global wrappers (guarded) *****/
if (typeof globalThis.AutoFilter_openFilterDialog !== "function") {
  function AutoFilter_openFilterDialog() {
    AutoFilter.AutoFilter_openFilterDialog();
  }
}
if (typeof globalThis.AutoFilter_runAutoFilterFromClient !== "function") {
  function AutoFilter_runAutoFilterFromClient(promptText) {
    return AutoFilter.AutoFilter_runAutoFilterFromClient(promptText);
  }
}
if (typeof globalThis.AutoFilter_runLastFilter !== "function") {
  function AutoFilter_runLastFilter() {
    return AutoFilter.AutoFilter_runLastFilter();
  }
}
if (typeof globalThis.AutoFilter_clearFilterFlags !== "function") {
  function AutoFilter_clearFilterFlags() {
    return AutoFilter.AutoFilter_clearFilterFlags();
  }
}
if (typeof globalThis.AutoFilter_removeSheetFilter !== "function") {
  function AutoFilter_removeSheetFilter() {
    return AutoFilter.AutoFilter_removeSheetFilter();
  }
}
