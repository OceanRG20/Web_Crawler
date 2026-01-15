/*************************************************
 * Backfill_Agent.gs — Column Backfills + News merge
 *
 * CONFIG SHEET: "Backfill"
 *   Row 1 headers:
 *     A: Column_ID      (must match header text in MMCrawl row 1)
 *     B: GPT_Prompt     (template; may contain <<<ROW_DATA_HERE>>>)
 *     C: Result         (optional log for last run)
 *
 * DATA SHEET: "MMCrawl"  (or AIA.MMCRAWL_SHEET if defined)
 *
 * NEWS SHEET: "News Raw"
 *   Must contain:
 *     - Company Website URL
 *     - News Story URL
 *     - Publication Date
 *     - Specific Value
 *
 * Script Properties required (for GPT columns):
 *   OPENAI_API_KEY  – your OpenAI key
 **************************************************/

/** ===== Menu hook (called from main onOpen) ===== */
function onOpen_Backfill(ui) {
  ui = ui || SpreadsheetApp.getUi();

  // --- Columns Backfill (Manual) submenu ---
  const manualSubMenu = ui.createMenu("▶ Columns Backfill (Manual)")
    .addItem("Equipment + CNC 3 & 5-axis", "BF_runBackfill_EquipmentCNCCombo")
    .addItem("Ownership + Family business", "BF_runBackfill_OwnershipFamilyCombo")
    .addItem("Estimated Revenues", "BF_runBackfill_EstimatedRevenues")
    .addItem("Number of employees", "BF_runBackfill_NumberOfEmployees")
    .addItem("Square footage (facility)", "BF_runBackfill_SquareFootage")
    .addItem("Years of operation", "BF_runBackfill_YearsOfOperation")
    .addSeparator()
    .addItem("Equipment", "BF_runBackfill_Equipment")
    .addItem("CNC 3-axis", "BF_runBackfill_CNC3Axis")
    .addItem("CNC 5-axis", "BF_runBackfill_CNC5Axis")
    // .addItem("Ownership", "BF_runBackfill_Ownership")
    .addItem("Spares/ Repairs", "BF_runBackfill_SparesRepairs")
    // .addItem("Family business", "BF_runBackfill_FamilyBusiness")
    .addItem("2nd Address", "BF_runBackfill_SecondAddress")
    .addItem("Region", "BF_runBackfill_Region")
    .addItem("Medical", "BF_runBackfill_Medical");


  // --- Backfill main menu ---
  ui.createMenu("Backfill")
    .addSubMenu(manualSubMenu)
    .addSeparator()
    .addItem("▶ Backfill from News", "BF_runBackfill_FromNews")
    .addSeparator()
    .addItem("▶ Update Rule Backfill…", "BF_showUpdateRuleSidebar") 
    .addToUi();
}

/*************************************************
 * 1) PUBLIC ENTRY FUNCTIONS (Manual GPT backfills)
 **************************************************/

function BF_runBackfill_NumberOfEmployees() {
  BF_runBackfillForMenu_("Number of employees");
}

function BF_runBackfill_EstimatedRevenues() {
  BF_runBackfillForMenu_("Estimated Revenues");
}

function BF_runBackfill_SquareFootage() {
  BF_runBackfillForMenu_("Square footage (facility)");
}

function BF_runBackfill_YearsOfOperation() {
  BF_runBackfillForMenu_("Years of operation");
}

function BF_runBackfill_Equipment() {
  BF_runBackfillForMenu_("Equipment");
}

function BF_runBackfill_CNC3Axis() {
  BF_runBackfillForMenu_("CNC 3-axis");
}

function BF_runBackfill_CNC5Axis() {
  BF_runBackfillForMenu_("CNC 5-axis");
}

function BF_runBackfill_Ownership() {
  BF_runBackfillForMenu_("Ownership");
}

function BF_runBackfill_SparesRepairs() {
  BF_runBackfillForMenu_("Spares/ Repairs");
}

function BF_runBackfill_FamilyBusiness() {
  BF_runBackfillForMenu_("Family business");
}

function BF_runBackfill_SecondAddress() {
  BF_runBackfillForMenu_("2nd Address");
}

function BF_runBackfill_Region() {
  BF_runBackfillForMenu_("Region");
}

function BF_runBackfill_Medical() {
  BF_runBackfillForMenu_("Medical");
}

/**
 * Combined backfill runner:
 * 1. Equipment
 * 2. CNC 3-axis
 * 3. CNC 5-axis
 */
function BF_runBackfill_EquipmentCNCCombo() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();

  const rangeInfo = BF_promptForRowRange_(ui);
  if (!rangeInfo) return;

  const startRow = rangeInfo.startRow;
  const endRow = rangeInfo.endRow;

  try {
    ui.alert("Starting combined backfill:\nEquipment → CNC 3-axis → CNC 5-axis");

    // 1) Equipment
    ss.toast(
      "Running Equipment backfill… (rows " + startRow + "-" + endRow + ")",
      "Backfill progress",
      5
    );
    BF_runBackfillForColumnId_("Equipment", startRow, endRow);

    // 2) CNC 3-axis
    ss.toast(
      "Running CNC 3-axis backfill…",
      "Backfill progress",
      5
    );
    BF_runBackfillForColumnId_("CNC 3-axis", startRow, endRow);

    // 3) CNC 5-axis
    ss.toast(
      "Running CNC 5-axis backfill…",
      "Backfill progress",
      5
    );
    BF_runBackfillForColumnId_("CNC 5-axis", startRow, endRow);

    ui.alert(
      "Combined Backfill Complete.\n" +
      "Processed rows: " + startRow + " - " + endRow + "\n" +
      "Steps completed:\n" +
      " • Equipment\n" +
      " • CNC 3-axis\n" +
      " • CNC 5-axis"
    );

    ss.toast(
      "Equipment + CNC 3/5-axis Backfill finished.",
      "Backfill progress",
      5
    );
  } catch (e) {
    ui.alert("Error during combined backfill:\n" + e);
    throw e;
  }
}

/**
 * Combined backfill runner:
 * 1. Ownership
 * 2. Family business
 */
function BF_runBackfill_OwnershipFamilyCombo() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();

  const rangeInfo = BF_promptForRowRange_(ui);
  if (!rangeInfo) return;

  const startRow = rangeInfo.startRow;
  const endRow = rangeInfo.endRow;

  try {
    ui.alert("Starting combined backfill:\nOwnership → Family business");

    // 1) Ownership
    ss.toast(
      "Running Ownership backfill… (rows " + startRow + "-" + endRow + ")",
      "Backfill progress",
      5
    );
    BF_runBackfillForColumnId_("Ownership", startRow, endRow);

    // 2) Family business
    ss.toast(
      "Running Family business backfill…",
      "Backfill progress",
      5
    );
    BF_runBackfillForColumnId_("Family business", startRow, endRow);

    ui.alert(
      "Combined Backfill Complete.\n" +
      "Processed rows: " + startRow + " - " + endRow + "\n" +
      "Steps completed:\n" +
      " • Ownership\n" +
      " • Family business"
    );

    ss.toast(
      "Ownership + Family business Backfill finished.",
      "Backfill progress",
      5
    );
  } catch (e) {
    ui.alert("Error during combined backfill:\n" + e);
    throw e;
  }
}


/*************************************************
 * 2) Backfill from News (non-GPT, merges "Specific Value")
 **************************************************/

/*************************************************
 * PATCH: Update ONLY "Backfill from News" to match NEW News Raw "Specific Value"
 *
 * New reality:
 * - News Raw column "Specific Value" (I) is NO LONGER a giant "Special Values" JSON.
 * - It now contains either:
 *   (A) a JSON object like {"Medical":"Yes","Ownership":"Sold to PE"}
 *   OR
 *   (B) MMCrawl Updates lines like:
 *       Medical ; "Yes (News: <URL>)", 2020
 *       Ownership ; "Sold to PE (News: <URL>)", 2024
 *
 * This patch parses BOTH formats and applies them to MMCrawl columns.
 **************************************************/

/*************************************************
 * BACKFILL FROM NEWS — PATCH (Boolean columns behavior + JSON-array Specific Value)
 *
 * What this patch does (per your screenshots):
 * 1) For BOOLEAN-style columns (Medical, CNC 3-axis, CNC 5-axis, Family business):
 *    - If MMCrawl already has "Yes" → do NOTHING (do not append anything).
 *    - If MMCrawl has "NI" (or blank / refer to site) and News says "Yes" → set cell to ONLY "Yes"
 *      and hyperlink the "Yes" text to the News Story URL (NO "(News: YEAR)" suffix).
 *    - Never create "Yes ; Yes" chains.
 *
 * 2) Specific Value parsing:
 *    - Supports JSON object, JSON array (of objects/strings), and line format:
 *      Label ; "Value ...", 2020
 **************************************************/

/** ===== Replace your existing BF_runBackfill_FromNews() with this version ===== */
function BF_runBackfill_FromNews() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();

  const mmSheetName = (typeof AIA !== "undefined" && AIA.MMCRAWL_SHEET) || "MMCrawl";
  const newsSheetName = "News Raw";

  const mmSheet = ss.getSheetByName(mmSheetName);
  const newsSheet = ss.getSheetByName(newsSheetName);

  if (!mmSheet) { ui.alert('Data sheet "' + mmSheetName + '" not found.'); return; }
  if (!newsSheet) { ui.alert('News sheet "' + newsSheetName + '" not found.'); return; }

  const rangeInfo = BF_promptForRowRange_(ui);
  if (!rangeInfo) return;

  let startRow = rangeInfo.startRow;
  let endRow = rangeInfo.endRow;

  const mmLastRow = mmSheet.getLastRow();
  if (startRow > mmLastRow) { ui.alert("Start row is beyond MMCrawl data."); return; }
  if (endRow > mmLastRow) endRow = mmLastRow;

  const mmLastCol = mmSheet.getLastColumn();
  const mmHeaders = mmSheet.getRange(1, 1, 1, mmLastCol).getValues()[0];

  const mmUrlCol = BF_findHeaderIndex_(mmHeaders, [
    "Company Website URL",
    "Public Website Homepage URL",
    "Company Website",
    "Website"
  ]);
  if (mmUrlCol === -1) {
    ui.alert('No website column found in MMCrawl (e.g., "Company Website URL").');
    return;
  }

  // Read MMCrawl rows (data starts row 2)
  const mmNumRows = mmLastRow - 1;
  const mmData = mmSheet.getRange(2, 1, mmNumRows, mmLastCol).getValues();

  const targetRowIndexMin = startRow - 2;
  const targetRowIndexMax = endRow - 2;

  // News Raw setup
  const newsLastRow = newsSheet.getLastRow();
  const newsLastCol = newsSheet.getLastColumn();
  if (newsLastRow < 2) { ui.alert("No data rows found in News Raw."); return; }

  const newsHeaders = newsSheet.getRange(1, 1, 1, newsLastCol).getValues()[0];

  const newsUrlCol = BF_findHeaderIndex_(newsHeaders, [
    "Company Website URL",
    "Public Website Homepage URL",
    "Company Website",
    "Website"
  ]);

  const newsSpecificCol = BF_findHeaderIndex_(newsHeaders, [
    "Specific Value",
    "Specific Values"
  ]);

  const newsPubDateCol = BF_findHeaderIndex_(newsHeaders, [
    "Publication Date",
    "Publication date",
    "Pub Date",
    "Pub date"
  ]);

  const newsStoryUrlCol = BF_findHeaderIndex_(newsHeaders, [
    "News Story URL",
    "News URL",
    "Article URL"
  ]);

  if (newsUrlCol === -1) { ui.alert('Column "Company Website URL" not found in News Raw.'); return; }
  if (newsSpecificCol === -1) { ui.alert('Column "Specific Value" not found in News Raw.'); return; }

  const newsNumRows = newsLastRow - 1;
  const newsData = newsSheet.getRange(2, 1, newsNumRows, newsLastCol).getValues();

  // Build map: normalized company URL -> [{ specificText, pubYear, storyUrl }]
  const newsMap = {};
  for (let i = 0; i < newsNumRows; i++) {
    const row = newsData[i];

    const url = (row[newsUrlCol] || "").toString().trim();
    const specificText = (row[newsSpecificCol] || "").toString().trim();
    if (!url || !specificText) continue;

    let pubYear = "";
    if (newsPubDateCol !== -1) {
      const pubStr = (row[newsPubDateCol] || "").toString();
      const ym = pubStr.match(/\b(19|20)\d{2}\b/);
      if (ym) pubYear = ym[0];
    }

    const storyUrl = (newsStoryUrlCol !== -1)
      ? (row[newsStoryUrlCol] || "").toString().trim()
      : "";

    const norm = BF_normalizeUrl_(url);
    if (!norm) continue;

    if (!newsMap[norm]) newsMap[norm] = [];
    newsMap[norm].push({ specificText: specificText, pubYear: pubYear, storyUrl: storyUrl });
  }

  // Label synonyms
  const LABEL_SYNONYM = {
    "Family Ownership": "Family business",
    "Family ownership": "Family business",
    "Employees": "Number of employees",
    "Employee count": "Number of employees",
    "Square footage": "Square footage (facility)",
    "Square Footage": "Square footage (facility)"
  };

  let appliedCount = 0;

  for (let idx = targetRowIndexMin; idx <= targetRowIndexMax; idx++) {
    if (idx < 0 || idx >= mmNumRows) continue;

    const sheetRowNumber = idx + 2;
    const mmRow = mmData[idx];

    const url = (mmRow[mmUrlCol] || "").toString().trim();
    if (!url) continue;

    const norm = BF_normalizeUrl_(url);
    if (!norm) continue;

    const entries = newsMap[norm];
    if (!entries || !entries.length) continue;

    for (let s = 0; s < entries.length; s++) {
      const e = entries[s];

      // Parse Specific Value into updates
      const updates = BF_parseSpecificValueUpdates_(e.specificText, e.pubYear); // [{label,value}]
      for (let u = 0; u < updates.length; u++) {
        let label = (updates[u].label || "").trim();
        const rawValue = (updates[u].value || "").trim();
        if (!label || !rawValue) continue;

        if (LABEL_SYNONYM[label]) label = LABEL_SYNONYM[label];

        const colIndex = BF_findHeaderIndexExact_(mmHeaders, label);
        if (colIndex === -1) continue;

        const cleanedValue = BF_simplifyNewsValue_(rawValue, e.pubYear, label);

        const changed = BF_applyNewsValueToCell_(
          mmSheet,
          sheetRowNumber,
          colIndex + 1,
          cleanedValue,
          e.storyUrl,
          label
        );
        if (changed) appliedCount++;
      }
    }
  }

  ui.alert(
    "Backfill from News complete.\n" +
    "MMCrawl rows processed: " + startRow + "-" + endRow + "\n" +
    "Values applied to MMCrawl cells: " + appliedCount
  );
}

/** ===== Replace your existing BF_parseSpecificValueUpdates_() with this version ===== */
function BF_parseSpecificValueUpdates_(specificText, pubYear) {
  const t = (specificText || "").toString().trim();
  if (!t) return [];

  // A) JSON object
  if (t[0] === "{" && t[t.length - 1] === "}") {
    try {
      const obj = JSON.parse(t);
      if (obj && typeof obj === "object" && !Array.isArray(obj)) {
        const out = [];
        Object.keys(obj).forEach((k) => {
          const v = obj[k];
          if (v === null || v === undefined) return;
          const vs = (typeof v === "string") ? v.trim() : JSON.stringify(v);
          if (!k || !vs) return;
          out.push({ label: String(k).trim(), value: vs });
        });
        return out;
      }
    } catch (_) {}
  }

  // B) JSON array (common cause of your "array symbol" issue in News Raw)
  //    Supports:
  //      - [{"Medical":"Yes"},{"CNC 3-axis":"Yes"}]
  //      - ["Medical ; \"Yes (News: ...)\" , 2020", "Family business ; \"Yes\" , 2023"]
  if (t[0] === "[" && t[t.length - 1] === "]") {
    try {
      const arr = JSON.parse(t);
      if (Array.isArray(arr)) {
        const out = [];
        arr.forEach((item) => {
          if (item === null || item === undefined) return;

          if (typeof item === "string") {
            // treat as line format
            const sub = BF_parseSpecificValueUpdates_(item, pubYear);
            sub.forEach(x => out.push(x));
            return;
          }

          if (typeof item === "object" && !Array.isArray(item)) {
            Object.keys(item).forEach((k) => {
              const v = item[k];
              if (v === null || v === undefined) return;
              const vs = (typeof v === "string") ? v.trim() : JSON.stringify(v);
              if (!k || !vs) return;
              out.push({ label: String(k).trim(), value: vs });
            });
          }
        });
        return out;
      }
    } catch (_) {
      // fall through
    }
  }

  // C) MMCrawl line format: Label ; "Value ...", 2020
  const lines = t.split(/\r?\n/).map(s => s.trim()).filter(Boolean);
  const out2 = [];
  const re = /^([^;]+?)\s*;\s*"(.*)"\s*(?:,\s*(19|20)\d{2})?\s*$/;

  for (let i = 0; i < lines.length; i++) {
    const m = lines[i].match(re);
    if (!m) continue;
    const label = (m[1] || "").trim();
    const val = (m[2] || "").trim();
    if (label && val) out2.push({ label: label, value: val });
  }

  return out2;
}


/** Open the "Update Rule Backfill" dialog */
function BF_showUpdateRuleDialog() {
  const ui = SpreadsheetApp.getUi();

  // Load the HTML file named "Update_Backfill.html"
  const html = HtmlService
    .createHtmlOutputFromFile("Update_Backfill")
    .setWidth(420)
    .setHeight(520);

  ui.showModalDialog(html, "Update Rule Backfill");
}

function BF_showUpdateRuleSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("Update_Backfill")
    .setTitle("Update Rule Backfill");
  SpreadsheetApp.getUi().showSidebar(html);
}


/**
 * Return list of Column_ID values from Backfill sheet.
 */
function BF_getBackfillColumns() {
  const ss = SpreadsheetApp.getActive();
  const backfillSheetName = (typeof AIA !== "undefined" && AIA.BACKFILL_SHEET) || "Backfill";
  const sh = ss.getSheetByName(backfillSheetName);
  if (!sh) throw new Error('Config sheet "' + backfillSheetName + '" not found.');

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const vals = sh.getRange(2, 1, lastRow - 1, 1).getValues();
  return vals
    .map(r => (r[0] || "").toString().trim())
    .filter(s => s);
}

/**
 * Get stored rule for selected Column_ID.
 */
function BF_getBackfillPrompt(columnId) {
  const ss = SpreadsheetApp.getActive();
  const backfillSheetName = (typeof AIA !== "undefined" && AIA.BACKFILL_SHEET) || "Backfill";
  const sh = ss.getSheetByName(backfillSheetName);
  if (!sh) throw new Error('Config sheet "' + backfillSheetName + '" not found.');

  const row = BF_findBackfillConfigRow_(sh, columnId);
  if (row < 2) return "";
  return (sh.getRange(row, 2).getValue() || "").toString();
}

/**
 * Save/update rule.
 * Also saves the previous version of the prompt into Backfill column D.
 *
 * Backfill sheet columns:
 *   A: Column_ID
 *   B: Prompt
 *   C: Result
 *   D: Previous Version of Prompt
 */
function BF_saveBackfillPrompt(columnId, newPrompt) {
  const ss = SpreadsheetApp.getActive();
  const backfillSheetName = (typeof AIA !== "undefined" && AIA.BACKFILL_SHEET) || "Backfill";
  const sh = ss.getSheetByName(backfillSheetName);
  if (!sh) throw new Error('Config sheet "' + backfillSheetName + '" not found.');

  const row = BF_findBackfillConfigRow_(sh, columnId);
  if (row < 2) throw new Error('Column_ID "' + columnId + '" not found in Backfill sheet.');

  const incoming = (newPrompt || "").toString();
  const current = (sh.getRange(row, 2).getValue() || "").toString(); // Column B

  // Save previous prompt to Column D only if it exists AND is changing
  if (current.trim() && current !== incoming) {
    sh.getRange(row, 4).setValue(current); // Column D
  }

  // Save new prompt to Column B
  sh.getRange(row, 2).setValue(incoming);

  return 'Prompt saved for "' + columnId + '".';
}


/**
 * AI rewrite the existing Backfill prompt using a user "plan",
 * then SAVE the updated prompt back into Backfill!B for that Column_ID.
 *
 * Returns: { updatedPrompt: string, message: string }
 */
function BF_aiRewriteAndSavePrompt(columnId, planText) {
  columnId = (columnId || "").toString().trim();
  planText = (planText || "").toString().trim();

  if (!columnId) throw new Error("Missing columnId.");
  if (!planText) throw new Error("Plan text is empty.");

  const ss = SpreadsheetApp.getActive();
  const backfillSheetName = (typeof AIA !== "undefined" && AIA.BACKFILL_SHEET) || "Backfill";
  const backfillSheet = ss.getSheetByName(backfillSheetName);
  if (!backfillSheet) throw new Error('Config sheet "' + backfillSheetName + '" not found.');

  const cfgRow = BF_findBackfillConfigRow_(backfillSheet, columnId);
  if (cfgRow < 2) throw new Error('Column_ID "' + columnId + '" not found in Backfill sheet.');

  const currentPrompt = (backfillSheet.getRange(cfgRow, 2).getValue() || "").toString();

  // Build a strict instruction to rewrite but preserve existing structure.
  const systemPrompt =
    "You are updating a Google Sheets backfill GPT prompt.\n" +
    "Goal: rewrite the CURRENT PROMPT by applying the USER UPDATE PLAN.\n\n" +
    "CRITICAL RULES:\n" +
    "1) Output ONLY the final updated prompt text. No commentary, no markdown.\n" +
    "2) Preserve the overall structure, intent, and column scope.\n" +
    "3) If the prompt includes placeholders like <<<ROW_DATA_HERE>>> you MUST preserve them.\n" +
    "4) Keep it copy-paste ready for the Backfill sheet.\n" +
    "5) Apply the user plan precisely; do not invent unrelated rules.\n";

  const userPrompt =
    "COLUMN_ID:\n" + columnId + "\n\n" +
    "CURRENT PROMPT:\n" +
    "-----BEGIN CURRENT PROMPT-----\n" + currentPrompt + "\n-----END CURRENT PROMPT-----\n\n" +
    "USER UPDATE PLAN:\n" +
    "-----BEGIN PLAN-----\n" + planText + "\n-----END PLAN-----\n\n" +
    "Now produce the UPDATED PROMPT.";

  const updated = BF_callOpenAI_UpdateRule_(systemPrompt, userPrompt).trim();

  // Save
  backfillSheet.getRange(cfgRow, 2).setValue(updated);

  return {
    updatedPrompt: updated,
    message: 'Updated and saved for "' + columnId + '".'
  };
}

function BF_callOpenAI_UpdateRule_(systemPrompt, userPrompt) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
  if (!apiKey) {
    throw new Error(
      "OPENAI_API_KEY not set in Script Properties. " +
      "Set it under: Extensions → Apps Script → Project Settings → Script properties."
    );
  }

  const model = (typeof AIA !== "undefined" && AIA.MODEL) || "gpt-4o";
  const url = "https://api.openai.com/v1/chat/completions";

  const payload = {
    model: model,
    temperature: 0.2,
    max_tokens: 2500,
    messages: [
      { role: "system", content: systemPrompt },
      { role: "user", content: userPrompt }
    ]
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: { Authorization: "Bearer " + apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const resp = UrlFetchApp.fetch(url, options);
  const text = resp.getContentText();

  let data;
  try {
    data = JSON.parse(text);
  } catch (e) {
    throw new Error("Failed to parse OpenAI response: " + text);
  }

  if (data.error) {
    throw new Error("OpenAI error: " + (data.error.message || JSON.stringify(data.error)));
  }

  const answer =
    data.choices &&
    data.choices[0] &&
    data.choices[0].message &&
    data.choices[0].message.content;

  return (answer || "").trim();
}


/**
 * Delete prompt.
 */
function BF_deleteBackfillPrompt(columnId) {
  var ss = SpreadsheetApp.getActive();
  var name = (typeof AIA !== "undefined" && AIA.BACKFILL_SHEET) || "Backfill";
  var sheet = ss.getSheetByName(name);
  if (!sheet) return "Backfill sheet not found.";

  var cfgRow = BF_findBackfillConfigRow_(sheet, columnId);
  if (cfgRow < 2) return 'Column_ID "' + columnId + '" not found.';

  sheet.getRange(cfgRow, 2).clearContent();
  return 'Prompt deleted for "' + columnId + '".';
}


/*************************************************
 * 3) Menu helper — ask for row range + dispatch
 **************************************************/

function BF_runBackfillForMenu_(columnId) {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();

  const rangeInfo = BF_promptForRowRange_(ui);
  if (!rangeInfo) return;

  const startRow = rangeInfo.startRow;
  const endRow = rangeInfo.endRow;

  try {
    const result = BF_runBackfillForColumnId_(columnId, startRow, endRow);

    const backfillSheetName = (typeof AIA !== "undefined" && AIA.BACKFILL_SHEET) || "Backfill";
    const backfillSheet = ss.getSheetByName(backfillSheetName);
    if (backfillSheet) {
      const cfgRow = BF_findBackfillConfigRow_(backfillSheet, columnId);
      if (cfgRow > 0) {
        const logMsg =
          'Last run for "' + columnId + '": rows ' + startRow + "-" + endRow +
          " (" + result.rowsProcessed + " rows) at " + new Date().toLocaleString();
        backfillSheet.getRange(cfgRow, 3).setValue(logMsg);
      }
    }

    SpreadsheetApp.getUi().alert(
      'Backfill complete for "' + columnId + '".\n' +
      "MMCrawl rows: " + startRow + "-" + endRow + "\n" +
      "Rows processed: " + result.rowsProcessed
    );
  } catch (e) {
    Logger.log("Backfill error for " + columnId + ": " + e);
    SpreadsheetApp.getUi().alert('Backfill failed for "' + columnId + '":\n' + e);
  }
}

/**
 * Prompt the user for a row range "From-To" and return {startRow, endRow}
 * Data rows start at 2 (row 1 is header).
 */
function BF_promptForRowRange_(ui) {
  const ss = SpreadsheetApp.getActive();
  const mmSheetName = (typeof AIA !== "undefined" && AIA.MMCRAWL_SHEET) || "MMCrawl";
  const mmSheet = ss.getSheetByName(mmSheetName);

  if (!mmSheet) {
    ui.alert('Data sheet "' + mmSheetName + '" not found.');
    return null;
  }

  const lastDataRow = mmSheet.getLastRow();
  const exampleEnd = Math.max(lastDataRow, 10);

  const resp = ui.prompt(
    "Backfill row range",
    "Enter MMCrawl row range in the form From-To.\n" +
      "Example: 2-" + exampleEnd + "\n\n" +
      "Header row is 1, so first data row is 2.",
    ui.ButtonSet.OK_CANCEL
  );

  if (resp.getSelectedButton() !== ui.Button.OK) return null;

  const text = resp.getResponseText().trim();
  const match = text.match(/^(\d+)\s*-\s*(\d+)$/);
  if (!match) {
    ui.alert('Invalid range "' + text + '". Use format like 2-' + exampleEnd + ".");
    return null;
  }

  let startRow = parseInt(match[1], 10);
  let endRow = parseInt(match[2], 10);

  if (endRow < startRow) {
    const tmp = startRow;
    startRow = endRow;
    endRow = tmp;
  }
  if (startRow < 2) startRow = 2;
  if (endRow < 2) endRow = 2;

  return { startRow, endRow };
}

/*************************************************
 * 4) Core GPT backfill logic (all manual columns)
 *    + "by hand" protection + Equipment/CNC rules
 **************************************************/

function BF_runBackfillForColumnId_(columnId, startRow, endRow) {
  const ss = SpreadsheetApp.getActive();
  const backfillSheetName = (typeof AIA !== "undefined" && AIA.BACKFILL_SHEET) || "Backfill";
  const mmSheetName = (typeof AIA !== "undefined" && AIA.MMCRAWL_SHEET) || "MMCrawl";

  const backfillSheet = ss.getSheetByName(backfillSheetName);
  if (!backfillSheet) throw new Error('Config sheet "' + backfillSheetName + '" not found.');

  const mmSheet = ss.getSheetByName(mmSheetName);
  if (!mmSheet) throw new Error('Data sheet "' + mmSheetName + '" not found.');

  const cfgRow = BF_findBackfillConfigRow_(backfillSheet, columnId);
  if (cfgRow < 2) {
    throw new Error('Column_ID "' + columnId + '" not found in Backfill sheet.');
  }
  const promptTemplate = backfillSheet.getRange(cfgRow, 2).getValue().toString();
  if (!promptTemplate) {
    throw new Error('GPT_Prompt is blank in Backfill for Column_ID "' + columnId + '".');
  }

  const lastDataRow = mmSheet.getLastRow();
  if (startRow > lastDataRow) return { rowsProcessed: 0 };
  if (endRow > lastDataRow) endRow = lastDataRow;
  if (endRow < startRow) return { rowsProcessed: 0 };

  const lastCol = mmSheet.getLastColumn();
  const headerRow = mmSheet.getRange(1, 1, 1, lastCol).getValues()[0];

  // Find target column in MMCrawl
  let targetColIndex = -1;
  for (let c = 0; c < headerRow.length; c++) {
    const headerName = (headerRow[c] || "").toString().trim();
    if (headerName === columnId) {
      targetColIndex = c + 1; // 1-based
      break;
    }
  }
  if (targetColIndex === -1) {
    throw new Error(
      'Column header "' + columnId + '" not found in sheet "' + mmSheetName + '". ' +
      "Make sure it matches Backfill.Column_ID exactly."
    );
  }

  const numRows = endRow - startRow + 1;
  if (numRows <= 0) return { rowsProcessed: 0 };

  const rowsValues = mmSheet.getRange(startRow, 1, numRows, lastCol).getValues();
  const resultValues = [];

  // Find URL column once for better toasts
  const urlColIndex = BF_findHeaderIndex_(headerRow, [
    "Public Website Homepage URL",
    "Company Website URL",
    "Company Website",
    "Website"
  ]);

  // Find Equipment column index once (used for CNC columns)
  const equipmentColIndex = BF_findHeaderIndexExact_(headerRow, "Equipment");

  for (let r = 0; r < numRows; r++) {
    const sheetRowNumber = startRow + r;
    const rowData = rowsValues[r];

    const existing = rowData[targetColIndex - 1];
    const existingStr = (existing === null || existing === undefined) ? "" : existing.toString();

    // SPECIAL CASE: Years of operation — if already filled, keep and skip GPT
    if (columnId === "Years of operation" && existingStr !== "") {
      resultValues.push([existingStr]);
      continue;
    }

      // === SPECIAL CASE: 2nd Address (client rule) ===
  if (columnId === "2nd Address") {
    const trimmed = existingStr.trim();
    const lower   = trimmed.toLowerCase();

    // Treat these as "no usable 2nd address":
    //  - empty
    //  - "NI" / "ni" / "Refer to Site" variants
    //  - literal "" (two quote characters)
    const isEmptyLike =
      !trimmed ||
      lower === "ni" ||
      lower === '"ni"' ||
      lower === "refer to site" ||
      lower === '"refer to site"' ||
      trimmed === '""';

    if (isEmptyLike) {
      // Force a true blank cell
      resultValues.push([""]);
    } else {
      // Real 2nd address already present → keep it
      resultValues.push([existingStr]);
    }
    continue; // skip GPT for this column
  }

    // SPECIAL CASE: Equipment — only re-fill when empty / NI / refer to site
    if (columnId === "Equipment") {
      const lower = existingStr.trim().toLowerCase();
      if (
        lower &&
        lower !== "ni" &&
        lower !== '"ni"' &&
        lower !== "refer to site" &&
        lower !== '"refer to site"'
      ) {
        // keep existing rich list; do not overwrite
        resultValues.push([existingStr]);
        continue;
      }
    }

    // SPECIAL CASE: CNC 3-axis / CNC 5-axis — logic based on Equipment cell
    if (columnId === "CNC 3-axis" || columnId === "CNC 5-axis") {
      let eqVal = "";
      if (equipmentColIndex !== -1) {
        eqVal = (rowData[equipmentColIndex] || "").toString().trim();
      }
      const eqLower = eqVal.toLowerCase();

      if (!eqVal) {
        // no equipment info at all
        resultValues.push(["NI"]);
        continue;
      }
      if (eqLower === "refer to site" || eqLower === '"refer to site"') {
        resultValues.push(["refer to site"]);
        continue;
      }
      if (eqLower === "ni" || eqLower === '"ni"') {
        resultValues.push(["NI"]);
        continue;
      }
      // else: we DO send the row to GPT to interpret the normalized Equipment line
      // and decide Yes/NI/refer to site.
    }

    // Toast with URL
    let urlForToast = "";
    if (urlColIndex !== -1) {
      urlForToast = (rowData[urlColIndex] || "").toString().trim();
    }
    SpreadsheetApp.getActive().toast(
      'Backfill "' + columnId + '" – MMCrawl row ' + sheetRowNumber + " of " + endRow +
      (urlForToast ? ("\n" + urlForToast) : ""),
      "Backfill progress",
      4
    );

    const rowText = BF_formatRowForPrompt_(headerRow, rowData, sheetRowNumber);

    let systemPrompt = promptTemplate;
    if (systemPrompt.indexOf("<<<ROW_DATA_HERE>>>") !== -1) {
      systemPrompt = systemPrompt.replace("<<<ROW_DATA_HERE>>>", rowText);
    } else {
      systemPrompt += "\n\nMMCrawl row:\n" + rowText;
    }

    // Call OpenAI to get new cell value
    let cellValue = BF_callOpenAI_Backfill_(systemPrompt, columnId);

    // Re-order multi-part answers for specific columns:
    // Best = (site), then (calc ...), then (public: ...)
    if (columnId === "Number of employees") {
      cellValue = BF_normalizeEmployeesOrder_(cellValue);
    } else if (columnId === "Estimated Revenues") {
      cellValue = BF_normalizeRevenueOrder_(cellValue);
    }

    // IMPORTANT: protect any client "by hand" notes in existing cell
    cellValue = BF_mergeByHand_(existingStr, cellValue);

    resultValues.push([cellValue]);
  }

  // Write results back
  mmSheet.getRange(startRow, targetColIndex, numRows, 1).setValues(resultValues);
  SpreadsheetApp.getActive().toast(
    'Backfill "' + columnId + '" finished.',
    "Backfill progress",
    3
  );

  return { rowsProcessed: numRows };
}

/**
 * Find config row in Backfill sheet for a given Column_ID.
 */
function BF_findBackfillConfigRow_(backfillSheet, columnId) {
  const lastRow = backfillSheet.getLastRow();
  if (lastRow < 2) return -1;

  const values = backfillSheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    const v = (values[i][0] || "").toString().trim();
    if (v === columnId) return i + 2;
  }
  return -1;
}

/*************************************************
 * 5) Prompt formatting + OpenAI call
 **************************************************/

function BF_formatRowForPrompt_(headers, rowValues, rowNumber) {
  const lines = [];
  if (rowNumber) lines.push("Sheet row: " + rowNumber);

  for (let i = 0; i < headers.length; i++) {
    const headerName = (headers[i] || "").toString().trim();
    if (!headerName) continue;
    const val = rowValues[i];
    const valueStr = (val === "" || val === null || val === undefined) ? "" : val.toString();
    lines.push(headerName + ": " + valueStr);
  }
  return lines.join("\n");
}

function BF_callOpenAI_Backfill_(systemPrompt, columnId) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
  if (!apiKey) {
    throw new Error(
      "OPENAI_API_KEY not set in Script Properties. " +
      "Set it under: Extensions → Apps Script → Project Settings → Script properties."
    );
  }

  const model = (typeof AIA !== "undefined" && AIA.MODEL) || "gpt-4o";
  const url = "https://api.openai.com/v1/chat/completions";

  const payload = {
    model: model,
    temperature: 0.15,
    max_tokens: 80,
    messages: [
      { role: "system", content: systemPrompt },
      {
        role: "user",
        content:
          'Return ONLY the final value that should be written into the "' +
          columnId +
          '" cell for this MMCrawl row. Do not add explanations or extra text.'
      }
    ]
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: { Authorization: "Bearer " + apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const resp = UrlFetchApp.fetch(url, options);
  const text = resp.getContentText();
  let data;
  try {
    data = JSON.parse(text);
  } catch (e) {
    throw new Error("Failed to parse OpenAI response: " + text);
  }

  if (data.error) {
    throw new Error("OpenAI error: " + (data.error.message || JSON.stringify(data.error)));
  }

  const answer =
    data.choices &&
    data.choices[0] &&
    data.choices[0].message &&
    data.choices[0].message.content;

  return (answer || "").trim();
}

/*************************************************
 * 6) Small utilities
 **************************************************/

function BF_findHeaderIndex_(headers, candidates) {
  const lowerCandidates = candidates.map(function (c) { return c.toLowerCase(); });
  for (let i = 0; i < headers.length; i++) {
    const h = (headers[i] || "").toString().trim().toLowerCase();
    if (!h) continue;
    if (lowerCandidates.indexOf(h) !== -1) return i;
  }
  return -1;
}

function BF_findHeaderIndexExact_(headers, label) {
  for (let i = 0; i < headers.length; i++) {
    const h = (headers[i] || "").toString().trim();
    if (h === label) return i;
  }
  return -1;
}

function BF_normalizeUrl_(url) {
  if (!url) return "";
  let s = url.toString().trim().toLowerCase();
  s = s.replace(/^https?:\/\//, "");
  s = s.replace(/\/+$/, "");
  return s;
}

/** ===== Replace your existing BF_simplifyNewsValue_() with this version ===== */
function BF_simplifyNewsValue_(rawValue, pubYear, label) {
  let v = (rawValue || "").toString().trim();
  let year = pubYear || "";

  // Boolean columns must stay clean: ONLY "Yes" or "NI" (no "(News: YEAR)")
  if (BF_isBooleanNewsColumn_(label)) {
    return BF_normalizeYesNI_(v); // "Yes" or "NI" or original if neither
  }

  if (!year) {
    const ym = v.match(/\b(19|20)\d{2}\b/);
    if (ym) year = ym[0];
  }

  // Remove any existing "(News ...)" tags
  v = v.replace(/\(News[^)]*\)/gi, "").trim();

  // Remove trailing ", YEAR"
  if (year) {
    const reYearComma = new RegExp("[,\\s]*" + year + "\\s*$");
    v = v.replace(reYearComma, "").trim();
  }

  v = v.replace(/[;,.\s]+$/, "").trim();
  if (!v) v = (rawValue || "").toString().trim();

  if (year) return v + " (News: " + year + ")";
  return v + " (News)";
}

/********************************************************************
 * PATCH: Backfill from News — preserve existing hyperlinks + add ALL
 *        new hyperlinks (multi-link support)
 *
 * What this fixes:
 * - Your current BF_applyNewsValueToCell_() overwrites the whole cell’s
 *   RichText, but only re-links the NEW segment → older segments lose links.
 * - This patch preserves existing RichText link runs, then adds links for:
 *    (A) any "(News: https://...)" tokens (links to that URL)
 *    (B) any "(News: 2024)" / "(News)" tokens in the NEW segment (links to storyUrl)
 *
 * HOW TO APPLY:
 * 1) Replace your existing BF_applyNewsValueToCell_() with the version below.
 * 2) Add the NEW helper functions below anywhere in Backfill_Agent.gs
 ********************************************************************/


/** ===== REPLACE your existing BF_applyNewsValueToCell_() with this version ===== */
function BF_applyNewsValueToCell_(sheet, row, col, cleanedValue, storyUrl, label) {
  const cell = sheet.getRange(row, col);

  const existingRich = cell.getRichTextValue();
  const currentText = existingRich ? String(existingRich.getText() || "") : String(cell.getDisplayValue() || "");
  const currentRaw = (currentText || "").toString().trim();

  const currNorm = BF_normalizeYesNI_(currentRaw);
  const incomingNorm = BF_normalizeYesNI_(cleanedValue);

  // BOOLEAN COLUMN RULES (your screenshots)
  if (BF_isBooleanNewsColumn_(label)) {
    // If already "Yes" -> keep, no changes
    if (currNorm === "Yes") return false;

    // If incoming is "Yes" and current is blank/NI/refer -> set ONLY "Yes" with hyperlink
    const isEmptyLike =
      !currentRaw ||
      currNorm === "NI" ||
      /^"?refer to site"?$/i.test(currentRaw);

    if (incomingNorm === "Yes" && isEmptyLike) {
      if (!storyUrl) {
        cell.setValue("Yes");
      } else {
        const rt = SpreadsheetApp.newRichTextValue()
          .setText("Yes")
          .setLinkUrl(0, 3, storyUrl)
          .build();
        cell.setRichTextValue(rt);
      }
      return true;
    }

    // If incoming is NI or unknown -> do nothing (do not overwrite with NI)
    return false;
  }

  // NON-BOOLEAN: append behavior
  const isReferToSite = /^"?refer to site"?$/i.test(currentRaw);
  const isEmpty = !currentRaw || isReferToSite;

  let newValue;
  if (isEmpty) newValue = cleanedValue;
  else if (currentRaw.indexOf(cleanedValue) !== -1) newValue = currentRaw;
  else newValue = currentRaw + " ; " + cleanedValue;

  const textChanged = newValue !== currentRaw;

  // If no story URL and no embedded "(News: https://...)" tokens, just set the value
  const hasEmbeddedNewsUrls = BF_textHasEmbeddedNewsUrl_(newValue);
  if (!storyUrl && !hasEmbeddedNewsUrls) {
    if (textChanged) cell.setValue(newValue);
    return textChanged;
  }

  // Build RichText while preserving existing link runs
  const builder = SpreadsheetApp.newRichTextValue().setText(newValue);

  // 1) Preserve previous hyperlinks (only safe if old text is prefix of new text)
  if (existingRich && currentText && newValue.indexOf(currentText) === 0) {
    BF_copyRichTextLinks_(builder, existingRich);
  }

  // 2) Apply links for NEW segment
  //    NEW segment = (isEmpty ? whole cell : trailing cleanedValue part)
  let segStart = 0;
  if (!isEmpty) segStart = newValue.length - String(cleanedValue || "").length;
  if (segStart < 0) segStart = 0;

  // 2A) Link embedded "(News: https://...)" tokens anywhere (these are precise)
  BF_applyEmbeddedNewsUrlLinks_(builder, newValue);

  // 2B) Link "(News: YEAR)" or "(News)" tokens in the NEW segment to storyUrl (if provided)
  if (storyUrl) {
    BF_applyNewsTokenLinksToStory_(builder, newValue, segStart, newValue.length, storyUrl);
  }

  cell.setRichTextValue(builder.build());
  return textChanged;
}


/** =========================
 * NEW helper functions
 * Add these anywhere in Backfill_Agent.gs
 * ========================= */

/** Returns true if text contains "(News: http...)" token(s). */
function BF_textHasEmbeddedNewsUrl_(text) {
  const t = String(text || "");
  return /\(News:\s*https?:\/\/[^\s\)"]+\)/i.test(t);
}

/**
 * Copy all existing link runs from a RichTextValue into a RichTextValueBuilder.
 * Assumes the builder text begins with the old text (prefix mapping).
 */
function BF_copyRichTextLinks_(builder, oldRich) {
  try {
    const runs = oldRich.getRuns();
    if (!runs || !runs.length) return;

    runs.forEach(function (run) {
      const url = run.getLinkUrl();
      if (!url) return;

      const start = run.getStartIndex();
      const end = run.getEndIndex();
      if (start == null || end == null) return;

      builder.setLinkUrl(start, end, url);
    });
  } catch (e) {
    // If getRuns() is unavailable in this runtime, we fail gracefully (no preservation).
  }
}

/**
 * Apply hyperlinks for embedded "(News: https://...)" tokens.
 * Links the entire token "(News: ...)" to the URL inside.
 */
function BF_applyEmbeddedNewsUrlLinks_(builder, text) {
  const t = String(text || "");
  const re = /\(News:\s*(https?:\/\/[^\s\)"]+)\)/gi;
  let m;
  while ((m = re.exec(t)) !== null) {
    const url = m[1];
    const start = m.index;
    const end = m.index + m[0].length;
    builder.setLinkUrl(start, end, url);
  }
}

/**
 * Apply hyperlinks to "(News: YYYY)" and "(News)" tokens within [start,end)
 * using the provided storyUrl (used for your "(News: 2024)" suffix segments).
 */
function BF_applyNewsTokenLinksToStory_(builder, text, start, end, storyUrl) {
  const t = String(text || "");
  const s = Math.max(0, parseInt(start || 0, 10));
  const e = Math.min(t.length, parseInt(end || t.length, 10));
  if (!storyUrl || s >= e) return;

  const seg = t.slice(s, e);

  // Match "(News: 2024)" OR "(News)"
  const re = /\(News(?::\s*(?:19|20)\d{2})?\)/gi;
  let m;
  while ((m = re.exec(seg)) !== null) {
    const tokenStart = s + m.index;
    const tokenEnd = tokenStart + m[0].length;
    builder.setLinkUrl(tokenStart, tokenEnd, storyUrl);
  }

  // Backward compatibility: also link the word "News" itself if someone uses "... News ..."
  // only inside the new segment.
  const reWord = /\bNews\b/g;
  while ((m = reWord.exec(seg)) !== null) {
    const ws = s + m.index;
    const we = ws + 4;
    builder.setLinkUrl(ws, we, storyUrl);
  }
}


/**
 * Preserve any cell content that contains the phrase "by hand".
 */
function BF_mergeByHand_(existingStr, newStr) {
  const existing = (existingStr || "").toString();
  let updated = (newStr || "").toString();

  if (!existing) {
    // Nothing to preserve
    return updated;
  }

  const hasByHand = existing.toLowerCase().indexOf("by hand") !== -1;
  if (!hasByHand) {
    // Normal case: prefer new value, but if GPT returns empty, keep existing
    return updated || existing;
  }

  // Existing value includes "by hand" — must NOT be removed
  if (!updated) {
    // GPT returned nothing → keep original
    return existing;
  }

  if (updated.toLowerCase().indexOf("by hand") !== -1) {
    // GPT already preserved the note
    return updated;
  }

  // Combine original (with "by hand") plus new content
  if (updated === existing) {
    return updated;
  }
  return existing + " ; " + updated;
}

/**
 * Normalize order for "Number of employees" cell.
 * Priority:
 *   0: segments tagged with "(site"
 *   1: segments tagged with "(calc"
 *   2: segments tagged with "(public:"
 */
function BF_normalizeEmployeesOrder_(value) {
  if (value == null) return value;
  var trimmed = String(value).trim();
  if (!trimmed || trimmed === "NI") return trimmed;

  var parts = trimmed.split(/\s*;\s*/).filter(function (p) { return p; });
  if (parts.length <= 1) return trimmed;

  function rankEmployeePart(p) {
    var s = p.toLowerCase();
    if (s.indexOf("(site") !== -1)   return 0; // best
    if (s.indexOf("(calc") !== -1)   return 1; // middle
    if (s.indexOf("(public:") !== -1) return 2; // least
    return 1; // unknown → treat like calc
  }

  parts.sort(function (a, b) {
    return rankEmployeePart(a) - rankEmployeePart(b);
  });

  return parts.join("; ");
}

/**
 * Normalize order for "Estimated Revenues" cell.
 * Priority:
 *   0: "(news" or "(site"
 *   1: "(calc"
 *   2: "(public:"
 */
function BF_normalizeRevenueOrder_(value) {
  if (value == null) return value;
  var trimmed = String(value).trim();
  if (!trimmed || trimmed === "NI") return trimmed;

  var parts = trimmed.split(/\s*;\s*/).filter(function (p) { return p; });
  if (parts.length <= 1) return trimmed;

  function rankRevenuePart(p) {
    var s = p.toLowerCase();
    if (s.indexOf("(news") !== -1)  return 0; // best factual
    if (s.indexOf("(site") !== -1)  return 0; // treat site same as news
    if (s.indexOf("(calc") !== -1)  return 1;
    if (s.indexOf("(public:") !== -1) return 2;
    return 1;
  }

  parts.sort(function (a, b) {
    return rankRevenuePart(a) - rankRevenuePart(b);
  });

  return parts.join("; ");
}


/**
 * Sidebar action: run the standard backfill engine for a column + row range.
 * Returns { rowsProcessed: number }
 */
function BF_quickBackfill(columnId, startRow, endRow) {
  startRow = parseInt(startRow, 10);
  endRow = parseInt(endRow, 10);

  if (!columnId) throw new Error("Missing columnId.");
  if (!startRow || !endRow) throw new Error("Invalid row range.");
  if (startRow < 2 || endRow < 2) throw new Error("Row range must be >= 2.");

  if (endRow < startRow) {
    const t = startRow; startRow = endRow; endRow = t;
  }

  // Reuse your existing engine:
  const result = BF_runBackfillForColumnId_(columnId, startRow, endRow);

  return {
    message: 'Backfill "' + columnId + '" complete. Rows: ' + startRow + "-" + endRow +
             " (" + (result && result.rowsProcessed ? result.rowsProcessed : (endRow - startRow + 1)) + " rows)"
  };
}

function BF_getMMCrawlDataRange() {
  const ss = SpreadsheetApp.getActive();
  const mmSheetName = (typeof AIA !== "undefined" && AIA.MMCRAWL_SHEET) || "MMCrawl";
  const mmSheet = ss.getSheetByName(mmSheetName);
  if (!mmSheet) throw new Error('Data sheet "' + mmSheetName + '" not found.');

  const lastRow = mmSheet.getLastRow();
  // Data starts at row 2; if sheet only has header, endRow becomes 2
  const startRow = 2;
  const endRow = Math.max(2, lastRow);

  return { startRow: startRow, endRow: endRow };
}

/** ===== Add these NEW helper functions anywhere in Backfill_Agent.gs ===== */
function BF_isBooleanNewsColumn_(label) {
  const s = (label || "").toString().trim().toLowerCase();
  return (
    s === "medical" ||
    s === "family business" ||
    s === "cnc 3-axis" ||
    s === "cnc 5-axis"
  );
}

function BF_normalizeYesNI_(value) {
  const s = (value || "").toString().trim().toLowerCase();

  // Accept common variants produced by summaries/specific values
  if (s === "yes") return "Yes";
  if (s.indexOf("yes") === 0) return "Yes"; // "Yes (News: 2020)" etc.
  if (s === "ni") return "NI";
  if (s.indexOf("ni") === 0) return "NI";

  return (value || "").toString().trim();
}
