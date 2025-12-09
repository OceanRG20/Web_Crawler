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
    .addItem("Number of employees", "BF_runBackfill_NumberOfEmployees")
    .addItem("Estimated Revenues", "BF_runBackfill_EstimatedRevenues")
    .addItem("Square footage (facility)", "BF_runBackfill_SquareFootage")
    .addItem("Years of operation", "BF_runBackfill_YearsOfOperation")
    .addItem("Equipment", "BF_runBackfill_Equipment");

  // --- Backfill main menu ---
  ui.createMenu("Backfill")
    .addSubMenu(manualSubMenu)
    .addSeparator()
    .addItem("▶ Backfill from News", "BF_runBackfill_FromNews")
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

/**
 * Equipment backfill (special: uses MMCrawl website URL + site bundle).
 * - Only touches rows where Equipment is blank, "NI", or "refer to site".
 * - For other rows, keeps the existing Equipment value.
 */
function BF_runBackfill_Equipment() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();

  const rangeInfo = BF_promptForRowRange_(ui);
  if (!rangeInfo) return;

  const startRow = rangeInfo.startRow;
  const endRow = rangeInfo.endRow;

  try {
    const result = BF_runBackfillForColumnId_WithSite_("Equipment", startRow, endRow);

    const backfillSheetName =
      (typeof AIA !== "undefined" && AIA.BACKFILL_SHEET) || "Backfill";
    const backfillSheet = ss.getSheetByName(backfillSheetName);
    if (backfillSheet) {
      const cfgRow = BF_findBackfillConfigRow_(backfillSheet, "Equipment");
      if (cfgRow > 0) {
        const logMsg =
          'Last run for "Equipment": rows ' + startRow + "-" + endRow +
          " (" + result.rowsProcessed + " rows) at " + new Date().toLocaleString();
        backfillSheet.getRange(cfgRow, 3).setValue(logMsg);
      }
    }

    ui.alert(
      'Backfill complete for "Equipment".\n' +
      "MMCrawl rows: " + startRow + "-" + endRow + "\n" +
      "Rows processed (GPT calls): " + result.rowsProcessed
    );
  } catch (e) {
    Logger.log("Backfill error for Equipment: " + e);
    ui.alert('Backfill failed for "Equipment":\n' + e);
  }
}

/*************************************************
 * 2) Backfill from News (non-GPT, merges "Specific Value")
 **************************************************/

function BF_runBackfill_FromNews() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();

  const mmSheetName =
    (typeof AIA !== "undefined" && AIA.MMCRAWL_SHEET) || "MMCrawl";
  const newsSheetName = "News Raw";

  const mmSheet = ss.getSheetByName(mmSheetName);
  const newsSheet = ss.getSheetByName(newsSheetName);

  if (!mmSheet) {
    ui.alert('Data sheet "' + mmSheetName + '" not found.');
    return;
  }
  if (!newsSheet) {
    ui.alert('News sheet "' + newsSheetName + '" not found.');
    return;
  }

  const rangeInfo = BF_promptForRowRange_(ui);
  if (!rangeInfo) return;

  let startRow = rangeInfo.startRow;
  let endRow = rangeInfo.endRow;

  const mmLastRow = mmSheet.getLastRow();
  if (startRow > mmLastRow) {
    ui.alert("Start row is beyond MMCrawl data.");
    return;
  }
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
    ui.alert(
      'No website column found in MMCrawl. ' +
      'Expected header like "Company Website URL" or "Public Website Homepage URL".'
    );
    return;
  }

  // Read MMCrawl rows
  const mmNumRows = mmLastRow - 1;
  const mmData = mmSheet.getRange(2, 1, mmNumRows, mmLastCol).getValues();

  const targetRowIndexMin = startRow - 2;
  const targetRowIndexMax = endRow - 2;

  // News Raw setup
  const newsLastRow = newsSheet.getLastRow();
  const newsLastCol = newsSheet.getLastColumn();
  if (newsLastRow < 2) {
    ui.alert("No data rows found in News Raw.");
    return;
  }

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

  if (newsUrlCol === -1) {
    ui.alert('Column "Company Website URL" (or equivalent) not found in News Raw.');
    return;
  }
  if (newsSpecificCol === -1) {
    ui.alert('Column "Specific Value" not found in News Raw.');
    return;
  }

  const newsNumRows = newsLastRow - 1;
  const newsData = newsSheet.getRange(2, 1, newsNumRows, newsLastCol).getValues();

  // Build map: normalized company URL -> [{ specVal, pubYear, storyUrl }]
  const newsMap = {};

  for (let i = 0; i < newsNumRows; i++) {
    const row = newsData[i];
    const url = (row[newsUrlCol] || "").toString().trim();
    const specVal = (row[newsSpecificCol] || "").toString().trim();
    if (!url || !specVal) continue;

    let pubYear = "";
    if (newsPubDateCol !== -1) {
      const pubStr = (row[newsPubDateCol] || "").toString();
      const ym = pubStr.match(/\b(19|20)\d{2}\b/);
      if (ym) pubYear = ym[0];
    }

    const storyUrl =
      newsStoryUrlCol !== -1
        ? (row[newsStoryUrlCol] || "").toString().trim()
        : "";

    const norm = BF_normalizeUrl_(url);
    if (!norm) continue;

    if (!newsMap[norm]) newsMap[norm] = [];
    newsMap[norm].push({ specVal: specVal, pubYear: pubYear, storyUrl: storyUrl });
  }

  // Label synonyms: Specific Value label -> MMCrawl header
  const LABEL_SYNONYM = {
    "Family Ownership": "Family business",
    "Family ownership": "Family business"
  };

  let appliedCount = 0;
  const ssActive = SpreadsheetApp.getActive();

  // Process MMCrawl rows in selected range
  for (let idx = targetRowIndexMin; idx <= targetRowIndexMax; idx++) {
    if (idx < 0 || idx >= mmNumRows) continue;

    const sheetRowNumber = idx + 2;
    const row = mmData[idx];

    const url = (row[mmUrlCol] || "").toString().trim();
    if (!url) continue;

    const norm = BF_normalizeUrl_(url);
    if (!norm) continue;

    const newsEntries = newsMap[norm];
    if (!newsEntries || newsEntries.length === 0) continue; // no news for this company

    ssActive.toast(
      "Backfill from News – MMCrawl row " + sheetRowNumber + " of " + endRow,
      "Backfill from News",
      3
    );

    for (let s = 0; s < newsEntries.length; s++) {
      const specEntry = newsEntries[s];
      const specVal = specEntry.specVal;
      const pubYear = specEntry.pubYear;
      const storyUrl = specEntry.storyUrl;

      // Specific Value format is like:  Label ; "Value"
      const regex = /([^;]+?)\s*;\s*"([^"]+)"/g;
      let match;
      while ((match = regex.exec(specVal)) !== null) {
        let label = match[1].trim();
        const rawValue = match[2].trim();
        if (!label || !rawValue) continue;

        if (LABEL_SYNONYM[label]) {
          label = LABEL_SYNONYM[label];
        }

        const colIndex = BF_findHeaderIndexExact_(mmHeaders, label);
        if (colIndex === -1) continue; // no matching column

        const cleanedValue = BF_simplifyNewsValue_(rawValue, pubYear);

        const changed = BF_applyNewsValueToCell_(
          mmSheet,
          sheetRowNumber,
          colIndex + 1,
          cleanedValue,
          storyUrl
        );
        if (changed) appliedCount++;
      }
    }
  }

  ssActive.toast("Backfill from News finished.", "Backfill from News", 3);
  ui.alert(
    "Backfill from News complete.\n" +
      "MMCrawl rows processed: " +
      startRow +
      "-" +
      endRow +
      "\n" +
      "Values applied to MMCrawl cells: " +
      appliedCount
  );
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

    const backfillSheetName =
      (typeof AIA !== "undefined" && AIA.BACKFILL_SHEET) || "Backfill";
    const backfillSheet = ss.getSheetByName(backfillSheetName);
    if (backfillSheet) {
      const cfgRow = BF_findBackfillConfigRow_(backfillSheet, columnId);
      if (cfgRow > 0) {
        const logMsg =
          'Last run for "' +
          columnId +
          '": rows ' +
          startRow +
          "-" +
          endRow +
          " (" +
          result.rowsProcessed +
          " rows) at " +
          new Date().toLocaleString();
        backfillSheet.getRange(cfgRow, 3).setValue(logMsg);
      }
    }

    ui.alert(
      'Backfill complete for "' +
        columnId +
        '".\n' +
        "MMCrawl rows: " +
        startRow +
        "-" +
        endRow +
        "\n" +
        "Rows processed: " +
        result.rowsProcessed
    );
  } catch (e) {
    Logger.log("Backfill error for " + columnId + ": " + e);
    ui.alert('Backfill failed for "' + columnId + '":\n' + e);
  }
}

/**
 * Prompt the user for a row range "From-To" and return {startRow, endRow}
 * Data rows start at 2 (row 1 is header).
 */
function BF_promptForRowRange_(ui) {
  const ss = SpreadsheetApp.getActive();
  const mmSheetName =
    (typeof AIA !== "undefined" && AIA.MMCRAWL_SHEET) || "MMCrawl";
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
      "Example: 2-" +
      exampleEnd +
      "\n\n" +
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
 * 4) Core GPT backfill logic (generic columns)
 **************************************************/

function BF_runBackfillForColumnId_WithSite_(columnId, startRow, endRow) {
  const ss = SpreadsheetApp.getActive();
  const backfillSheetName =
    (typeof AIA !== "undefined" && AIA.BACKFILL_SHEET) || "Backfill";
  const mmSheetName =
    (typeof AIA !== "undefined" && AIA.MMCRAWL_SHEET) || "MMCrawl";

  const backfillSheet = ss.getSheetByName(backfillSheetName);
  if (!backfillSheet)
    throw new Error('Config sheet "' + backfillSheetName + '" not found.');

  const mmSheet = ss.getSheetByName(mmSheetName);
  if (!mmSheet) throw new Error('Data sheet "' + mmSheetName + '" not found.');

  const cfgRow = BF_findBackfillConfigRow_(backfillSheet, columnId);
  if (cfgRow < 2) {
    throw new Error(
      'Column_ID "' + columnId + '" not found in Backfill sheet.'
    );
  }
  const promptTemplate = backfillSheet
    .getRange(cfgRow, 2)
    .getValue()
    .toString();
  if (!promptTemplate) {
    throw new Error(
      'GPT_Prompt is blank in Backfill for Column_ID "' + columnId + '".'
    );
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
      'Column header "' +
        columnId +
        '" not found in sheet "' +
        mmSheetName +
        '". ' +
        "Make sure it matches Backfill.Column_ID exactly."
    );
  }

  // Website column index (for site bundle)
  const urlColIndex = BF_findHeaderIndex_(headerRow, [
    "Company Website URL",
    "Public Website Homepage URL",
    "Company Website",
    "Website"
  ]);
  if (urlColIndex === -1) {
    throw new Error(
      'Website column not found in "' +
        mmSheetName +
        '". Expected header like "Company Website URL" or "Public Website Homepage URL".'
    );
  }

  const numRows = endRow - startRow + 1;
  if (numRows <= 0) return { rowsProcessed: 0 };

  const rowsValues = mmSheet.getRange(startRow, 1, numRows, lastCol).getValues();
  const resultValues = [];
  let processed = 0;

  for (let r = 0; r < numRows; r++) {
    const sheetRowNumber = startRow + r;
    const rowData = rowsValues[r];

    const existing = rowData[targetColIndex - 1];
    const existingStr =
      existing === null || existing === undefined ? "" : existing.toString();
    const existingNorm = existingStr.trim().toLowerCase();

    const hasRealValue =
      existingNorm !== "" &&
      existingNorm !== "ni" &&
      existingNorm !== "refer to site";

    // If Equipment already has a real value, keep it and skip GPT.
    if (hasRealValue) {
      resultValues.push([existingStr]);
      continue;
    }

    // Website URL for this row
    const siteUrl = (rowData[urlColIndex] || "").toString().trim();

    // Show URL in the toast notification
    const urlMsg = siteUrl ? siteUrl : "(no URL)";
    const shortUrl =
      urlMsg.length > 80 ? urlMsg.substring(0, 77) + "..." : urlMsg;

    SpreadsheetApp.getActive().toast(
      'Equipment Backfill\nRow ' +
        sheetRowNumber +
        " of " +
        endRow +
        "\nURL: " +
        shortUrl,
      "Backfill progress",
      6
    );


    const rowText = BF_formatRowForPrompt_(headerRow, rowData, sheetRowNumber);

    // Fetch site bundle text (re-use Raw_Data helper).
    let siteText = "";
    if (siteUrl && typeof AIA_fetchSiteBundleText_ === "function") {
      try {
        const bundle = AIA_fetchSiteBundleText_(siteUrl, 15000);
        siteText = (bundle && bundle.text) || "";
      } catch (err) {
        Logger.log("Equipment site fetch error for " + siteUrl + ": " + err);
      }
    }

    let systemPrompt = promptTemplate;
    if (systemPrompt.indexOf("<<<ROW_DATA_HERE>>>") !== -1) {
      systemPrompt = systemPrompt.replace("<<<ROW_DATA_HERE>>>", rowText);
    } else {
      systemPrompt += "\n\nMMCrawl row:\n" + rowText;
    }

    if (siteText) {
      systemPrompt +=
        "\n\n### SITE_TEXT (plain text extracted from " +
        siteUrl +
        ")\n" +
        siteText +
        "\n\nUse ONLY this SITE_TEXT plus the rules above to decide the final Equipment line.";
    } else {
      systemPrompt +=
        "\n\n(No additional SITE_TEXT was available for this row; if you cannot confirm any equipment, follow the NI / Refer to Site rules.)";
    }

    let cellValue = BF_callOpenAI_Backfill_(systemPrompt, columnId);
    cellValue = BF_mergeByHand_(existingStr, cellValue);

    resultValues.push([cellValue]);
    processed++;
  }

  mmSheet
    .getRange(startRow, targetColIndex, numRows, 1)
    .setValues(resultValues);
  SpreadsheetApp.getActive().toast(
    'Backfill "' + columnId + '" finished.',
    "Backfill progress",
    3
  );

  return { rowsProcessed: processed };
}


/*************************************************
 * 4b) Core GPT backfill logic WITH SITE TEXT (Equipment)
 **************************************************/

function BF_runBackfillForColumnId_WithSite_(columnId, startRow, endRow) {
  const ss = SpreadsheetApp.getActive();
  const backfillSheetName =
    (typeof AIA !== "undefined" && AIA.BACKFILL_SHEET) || "Backfill";
  const mmSheetName =
    (typeof AIA !== "undefined" && AIA.MMCRAWL_SHEET) || "MMCrawl";

  const backfillSheet = ss.getSheetByName(backfillSheetName);
  if (!backfillSheet)
    throw new Error('Config sheet "' + backfillSheetName + '" not found.');

  const mmSheet = ss.getSheetByName(mmSheetName);
  if (!mmSheet) throw new Error('Data sheet "' + mmSheetName + '" not found.');

  const cfgRow = BF_findBackfillConfigRow_(backfillSheet, columnId);
  if (cfgRow < 2) {
    throw new Error(
      'Column_ID "' + columnId + '" not found in Backfill sheet.'
    );
  }
  const promptTemplate = backfillSheet
    .getRange(cfgRow, 2)
    .getValue()
    .toString();
  if (!promptTemplate) {
    throw new Error(
      'GPT_Prompt is blank in Backfill for Column_ID "' + columnId + '".'
    );
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
      'Column header "' +
        columnId +
        '" not found in sheet "' +
        mmSheetName +
        '". ' +
        "Make sure it matches Backfill.Column_ID exactly."
    );
  }

  // Website column index (for site bundle)
  const urlColIndex = BF_findHeaderIndex_(headerRow, [
    "Company Website URL",
    "Public Website Homepage URL",
    "Company Website",
    "Website"
  ]);
  if (urlColIndex === -1) {
    throw new Error(
      'Website column not found in "' +
        mmSheetName +
        '". Expected header like "Company Website URL" or "Public Website Homepage URL".'
    );
  }

  const numRows = endRow - startRow + 1;
  if (numRows <= 0) return { rowsProcessed: 0 };

  const rowsValues = mmSheet.getRange(startRow, 1, numRows, lastCol).getValues();
  const resultValues = [];
  let processed = 0;

  for (let r = 0; r < numRows; r++) {
    const sheetRowNumber = startRow + r;
    const rowData = rowsValues[r];

    const existing = rowData[targetColIndex - 1];
    const existingStr =
      existing === null || existing === undefined ? "" : existing.toString();
    const existingNorm = existingStr.trim().toLowerCase();

    const hasRealValue =
      existingNorm !== "" &&
      existingNorm !== "ni" &&
      existingNorm !== "refer to site";

    // If Equipment already has a real value, keep it and skip GPT.
    if (hasRealValue) {
      resultValues.push([existingStr]);
      continue;
    }

    SpreadsheetApp.getActive().toast(
      'Backfill "' +
        columnId +
        '" – MMCrawl row ' +
        sheetRowNumber +
        " of " +
        endRow,
      "Backfill progress",
      3
    );

    const rowText = BF_formatRowForPrompt_(headerRow, rowData, sheetRowNumber);

    // Website URL for this row
    const siteUrl = (rowData[urlColIndex] || "").toString().trim();

    // Fetch site bundle text (re-use Raw_Data helper).
    let siteText = "";
    if (siteUrl && typeof AIA_fetchSiteBundleText_ === "function") {
      try {
        const bundle = AIA_fetchSiteBundleText_(siteUrl, 15000);
        siteText = (bundle && bundle.text) || "";
      } catch (err) {
        Logger.log("Equipment site fetch error for " + siteUrl + ": " + err);
      }
    }

    let systemPrompt = promptTemplate;
    if (systemPrompt.indexOf("<<<ROW_DATA_HERE>>>") !== -1) {
      systemPrompt = systemPrompt.replace("<<<ROW_DATA_HERE>>>", rowText);
    } else {
      systemPrompt += "\n\nMMCrawl row:\n" + rowText;
    }

    // Inject SITE_TEXT block so the model actually sees the website content.
    if (siteText) {
      systemPrompt +=
        "\n\n### SITE_TEXT (plain text extracted from " +
        siteUrl +
        ")\n" +
        siteText +
        "\n\nUse ONLY this SITE_TEXT plus the rules above to decide the final Equipment line.";
    } else {
      systemPrompt +=
        "\n\n(No additional SITE_TEXT was available for this row; if you cannot confirm any equipment, follow the NI / Refer to Site rules.)";
    }

    // Call OpenAI
    let cellValue = BF_callOpenAI_Backfill_(systemPrompt, columnId);
    cellValue = BF_mergeByHand_(existingStr, cellValue);

    resultValues.push([cellValue]);
    processed++;
  }

  // Write results back
  mmSheet
    .getRange(startRow, targetColIndex, numRows, 1)
    .setValues(resultValues);
  SpreadsheetApp.getActive().toast(
    'Backfill "' + columnId + '" finished.',
    "Backfill progress",
    3
  );

  return { rowsProcessed: processed };
}

/*************************************************
 * 5) Backfill config helpers
 **************************************************/

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
 * 6) Prompt formatting + OpenAI call
 **************************************************/

function BF_formatRowForPrompt_(headers, rowValues, rowNumber) {
  const lines = [];
  if (rowNumber) lines.push("Sheet row: " + rowNumber);

  for (let i = 0; i < headers.length; i++) {
    const headerName = (headers[i] || "").toString().trim();
    if (!headerName) continue;
    const val = rowValues[i];
    const valueStr =
      val === "" || val === null || val === undefined ? "" : val.toString();
    lines.push(headerName + ": " + valueStr);
  }
  return lines.join("\n");
}

function BF_callOpenAI_Backfill_(systemPrompt, columnId) {
  const apiKey =
    PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
  if (!apiKey) {
    throw new Error(
      "OPENAI_API_KEY not set in Script Properties. " +
        "Set it under: Extensions → Apps Script → Project Settings → Script properties."
    );
  }

  const model =
    (typeof AIA !== "undefined" && AIA.MODEL) || "gpt-4o";
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
    throw new Error(
      "OpenAI error: " + (data.error.message || JSON.stringify(data.error))
    );
  }

  const answer =
    data.choices &&
    data.choices[0] &&
    data.choices[0].message &&
    data.choices[0].message.content;

  return (answer || "").trim();
}

/*************************************************
 * 7) Small utilities
 **************************************************/

function BF_findHeaderIndex_(headers, candidates) {
  const lowerCandidates = candidates.map(function (c) {
    return c.toLowerCase();
  });
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

/**
 * Simplify raw Specific Value into "Value (News: YEAR)" or "Value (News)".
 */
function BF_simplifyNewsValue_(rawValue, pubYear) {
  let v = (rawValue || "").toString().trim();
  let year = pubYear || "";

  if (!year) {
    const ym = v.match(/\b(19|20)\d{2}\b/);
    if (ym) year = ym[0];
  }

  // Remove any existing "(News ...)" tags
  v = v.replace(/\(News[^)]*\)/gi, "").trim();

  // Remove trailing ", YEAR" if it exists
  if (year) {
    const reYearComma = new RegExp("[,\\s]*" + year + "\\s*$");
    v = v.replace(reYearComma, "").trim();
  }

  // Clean trailing punctuation
  v = v.replace(/[;,.\s]+$/, "").trim();

  if (!v) v = (rawValue || "").toString().trim();

  if (year) {
    return v + " (News: " + year + ")";
  } else {
    return v + " (News)";
  }
}

/**
 * Apply one cleaned news value into a specific MMCrawl cell,
 * and hyperlink the word "News" (inside this new segment) to storyUrl.
 *
 * Rules:
 *  - If existing is blank or "refer to site": replace with cleanedValue.
 *  - Else: append "; cleanedValue" unless already present.
 * Returns true if text content changed.
 */
function BF_applyNewsValueToCell_(sheet, row, col, cleanedValue, storyUrl) {
  const cell = sheet.getRange(row, col);
  const currentRaw = (cell.getValue() || "").toString().trim();
  const isReferToSite = /^refer to site$/i.test(currentRaw);
  const isEmpty = !currentRaw || isReferToSite;

  let newValue;
  if (isEmpty) {
    newValue = cleanedValue;
  } else if (currentRaw.indexOf(cleanedValue) !== -1) {
    newValue = currentRaw; // already there
  } else {
    newValue = currentRaw + " ; " + cleanedValue;
  }

  const textChanged = newValue !== currentRaw;

  // If no story URL, just write plain text
  if (!storyUrl) {
    if (textChanged) cell.setValue(newValue);
    return textChanged;
  }

  // Build rich text where only new segment's "News" word is hyperlinked
  const builder = SpreadsheetApp.newRichTextValue().setText(newValue);

  const segStart = newValue.length - cleanedValue.length;
  const segEnd = newValue.length;

  let searchPos = segStart;
  while (true) {
    const idx = newValue.indexOf("News", searchPos);
    if (idx === -1 || idx >= segEnd) break;
    builder.setLinkUrl(idx, idx + 4, storyUrl);
    searchPos = idx + 4;
  }

  cell.setRichTextValue(builder.build());
  return textChanged;
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
