/*************************************************
 * Backfill_Agent.gs — Column-specific Backfill via OpenAI
 *
 * CONFIG SHEET: "Backfill"
 *   Row 1 headers:
 *     A: Column_ID      (must match header text in MMCrawl row 1)
 *     B: GPT_Prompt     (template; may contain <<<ROW_DATA_HERE>>>)
 *     C: Result         (optional log for last run)
 *
 * DATA SHEET: "MMCrawl"  (or AIA.MMCRAWL_SHEET if defined)
 *   Row 1: column headers
 *   Row 2+: company rows
 *
 * Script Properties required:
 *   OPENAI_API_KEY  – your OpenAI key
 *
 * Optional global config object:
 *   var AIA = {
 *     MMCRAWL_SHEET: "MMCrawl",
 *     BACKFILL_SHEET: "Backfill",
 *     MODEL: "gpt-4o",
 *   };
 **************************************************/

/** ===== Menu hook (called from main onOpen) ===== */
function onOpen_Backfill(ui) {
  ui = ui || SpreadsheetApp.getUi();

  const menu = ui.createMenu("Backfill")
    // one submenu item per backfill column
    .addItem("Number of employees", "BF_runBackfill_NumberOfEmployees")
    .addItem("Estimated Revenues", "BF_runBackfill_EstimatedRevenues"); // NEW

  menu.addToUi();
}

/*************************************************
 * 1) PUBLIC ENTRY FUNCTIONS (one per column)
 **************************************************/

/**
 * Backfill for the "Number of employees" column.
 * Uses the GPT prompt in Backfill sheet where Column_ID = "Number of employees".
 */
function BF_runBackfill_NumberOfEmployees() {
  BF_runBackfillForMenu_("Number of employees");
}

/**
 * Backfill for the "Estimated Revenues" column.
 * Uses the GPT prompt in Backfill sheet where Column_ID = "Estimated Revenues".
 */
function BF_runBackfill_EstimatedRevenues() {
  BF_runBackfillForMenu_("Estimated Revenues");
}

/*************************************************
 * 2) MENU HELPER — ASK ROW RANGE, THEN RUN
 **************************************************/

/**
 * Common handler for a menu click for a specific Column_ID.
 */
function BF_runBackfillForMenu_(columnId) {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();

  // Ask for MMCrawl row range (simple From-To prompt)
  const rangeInfo = BF_promptForRowRange_(ui);
  if (!rangeInfo) {
    return; // user cancelled or invalid
  }

  const startRow = rangeInfo.startRow;
  const endRow = rangeInfo.endRow;

  try {
    const result = BF_runBackfillForColumnId_(columnId, startRow, endRow);

    // Log to Backfill sheet (optional)
    const backfillSheetName = (typeof AIA !== "undefined" && AIA.BACKFILL_SHEET) || "Backfill";
    const backfillSheet = ss.getSheetByName(backfillSheetName);
    if (backfillSheet) {
      const cfgRow = BF_findBackfillConfigRow_(backfillSheet, columnId);
      if (cfgRow > 0) {
        const logMsg =
          "Last run for \"" + columnId + "\": rows " + startRow + "-" + endRow +
          " (" + result.rowsProcessed + " rows) at " + new Date().toLocaleString();
        backfillSheet.getRange(cfgRow, 3).setValue(logMsg); // column C
      }
    }

    ui.alert(
      "Backfill complete for \"" + columnId + "\".\n" +
      "MMCrawl rows: " + startRow + "-" + endRow + "\n" +
      "Rows processed: " + result.rowsProcessed
    );
  } catch (e) {
    Logger.log("Backfill error for " + columnId + ": " + e);
    ui.alert("Backfill failed for \"" + columnId + "\":\n" + e);
  }
}

/**
 * Simple & clean row-range selector.
 *
 * Shows ONE prompt:
 *    "Enter MMCrawl row range (example: 2-50)"
 *
 * Returns:
 *   { startRow, endRow }
 * or null if cancelled/invalid.
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

  if (resp.getSelectedButton() !== ui.Button.OK) {
    return null; // cancelled
  }

  const text = resp.getResponseText().trim();
  const match = text.match(/^(\d+)\s*-\s*(\d+)$/);
  if (!match) {
    ui.alert('Invalid range "' + text + '". Use format like 2-' + exampleEnd + ".");
    return null;
  }

  let startRow = parseInt(match[1], 10);
  let endRow = parseInt(match[2], 10);

  // Normalize
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
 * 3) CORE BACKFILL LOGIC (REUSABLE FOR ANY COLUMN)
 **************************************************/

/**
 * Backfill a single MMCrawl column by Column_ID over a range of rows.
 * Uses the GPT_Prompt from Backfill sheet for that Column_ID.
 *
 * Shows toast progress: "Backfill 'col' – MMCrawl row X of Y".
 *
 * Returns { rowsProcessed: N }.
 */
function BF_runBackfillForColumnId_(columnId, startRow, endRow) {
  const ss = SpreadsheetApp.getActive();
  const backfillSheetName = (typeof AIA !== "undefined" && AIA.BACKFILL_SHEET) || "Backfill";
  const mmSheetName = (typeof AIA !== "undefined" && AIA.MMCRAWL_SHEET) || "MMCrawl";

  const backfillSheet = ss.getSheetByName(backfillSheetName);
  if (!backfillSheet) {
    throw new Error('Config sheet "' + backfillSheetName + '" not found.');
  }

  const mmSheet = ss.getSheetByName(mmSheetName);
  if (!mmSheet) {
    throw new Error('Data sheet "' + mmSheetName + '" not found.');
  }

  // Find config row and prompt template for this Column_ID
  const cfgRow = BF_findBackfillConfigRow_(backfillSheet, columnId);
  if (cfgRow < 2) {
    throw new Error('Column_ID "' + columnId + '" not found in Backfill sheet.');
  }
  const promptTemplate = backfillSheet.getRange(cfgRow, 2).getValue().toString();
  if (!promptTemplate) {
    throw new Error('GPT_Prompt is blank in Backfill for Column_ID "' + columnId + '".');
  }

  // Clamp row range within MMCrawl actual data
  const lastDataRow = mmSheet.getLastRow();
  if (startRow > lastDataRow) {
    return { rowsProcessed: 0 };
  }
  if (endRow > lastDataRow) {
    endRow = lastDataRow;
  }
  if (endRow < startRow) {
    return { rowsProcessed: 0 };
  }

  const lastCol = mmSheet.getLastColumn();
  const headerRow = mmSheet.getRange(1, 1, 1, lastCol).getValues()[0];

  // Find the target column index in MMCrawl by header text
  let targetColIndex = -1;
  for (let c = 0; c < headerRow.length; c++) {
    const headerName = (headerRow[c] || "").toString().trim();
    if (headerName === columnId) {
      targetColIndex = c + 1;
      break;
    }
  }
  if (targetColIndex === -1) {
    throw new Error(
      'Column header "' + columnId + '" not found in sheet "' + mmSheetName + '". ' +
      'Make sure it matches Backfill.Column_ID exactly.'
    );
  }

  const numRows = endRow - startRow + 1;
  if (numRows <= 0) {
    return { rowsProcessed: 0 };
  }

  // Read all rows in one batch
  const rowsValues = mmSheet.getRange(startRow, 1, numRows, lastCol).getValues();
  const resultValues = [];

  for (let r = 0; r < numRows; r++) {
    const sheetRowNumber = startRow + r;
    const rowData = rowsValues[r];

    // Toast progress so you can see it working row-by-row
    ss.toast(
      'Backfill "' + columnId + '" – MMCrawl row ' + sheetRowNumber + " of " + endRow,
      "Backfill progress",
      3
    );

    // Convert MMCrawl row into text: "Header: value"
    const rowText = BF_formatRowForPrompt_(headerRow, rowData, sheetRowNumber);

    // Merge template + rowText
    let systemPrompt = promptTemplate;
    if (systemPrompt.indexOf("<<<ROW_DATA_HERE>>>") !== -1) {
      systemPrompt = systemPrompt.replace("<<<ROW_DATA_HERE>>>", rowText);
    } else {
      systemPrompt += "\n\nMMCrawl row:\n" + rowText;
    }

    // Call GPT for this row
    const cellValue = BF_callOpenAI_Backfill_(systemPrompt, columnId);

    resultValues.push([cellValue]);
  }

  // Write results into MMCrawl target column
  mmSheet.getRange(startRow, targetColIndex, numRows, 1).setValues(resultValues);

  // Clear final toast quickly
  ss.toast("Backfill \"" + columnId + "\" finished.", "Backfill progress", 3);

  return { rowsProcessed: numRows };
}

/**
 * Find the row index in Backfill sheet where Column_ID matches.
 * Returns row number (>=2) or -1 if not found.
 */
function BF_findBackfillConfigRow_(backfillSheet, columnId) {
  const lastRow = backfillSheet.getLastRow();
  if (lastRow < 2) return -1;

  const values = backfillSheet.getRange(2, 1, lastRow - 1, 1).getValues(); // A2:A
  for (let i = 0; i < values.length; i++) {
    const v = (values[i][0] || "").toString().trim();
    if (v === columnId) {
      return i + 2; // row index
    }
  }
  return -1;
}

/*************************************************
 * 4) PROMPT FORMATTING + OPENAI CALL
 **************************************************/

/**
 * Format one MMCrawl row as text for the prompt:
 *   Sheet row: 2
 *   Company Name: Promark Tool and Manufacturing
 *   State: ON
 *   ...
 */
function BF_formatRowForPrompt_(headers, rowValues, rowNumber) {
  const lines = [];
  if (rowNumber) {
    lines.push("Sheet row: " + rowNumber);
  }

  for (let i = 0; i < headers.length; i++) {
    const headerName = (headers[i] || "").toString().trim();
    if (!headerName) continue;

    const val = rowValues[i];
    const valueStr = (val === "" || val === null) ? "" : val.toString();
    lines.push(headerName + ": " + valueStr);
  }

  return lines.join("\n");
}

/**
 * Low-level OpenAI call for backfill.
 * - systemPrompt: full column-specific instructions + row data
 * - columnId: used only in the user message
 *
 * Returns: trimmed string for the cell.
 */
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
    temperature: 0.15, // small flexibility for ranges, still factual
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
