/*************************************************
 * AI_Agent.gs — Google Sheets × OpenAI bridge
 * Menu on open:
 *   ▶ Raw Text Data Generation  (auto-finds Prompt ID = Raw_Data)
 *   🔑 Set OpenAI API Key
 * Results → column C; Timestamp → column D
 **************************************************/

/** ===== Safe UI getter (prevents editor-run crashes) ===== */
function AIA_safeUi_() {
  try {
    return SpreadsheetApp.getUi();
  } catch (e) {
    return null;
  }
}

/** ===== MASTER onOpen — runs every time the spreadsheet loads ===== */
function onOpen(e) {
  const ui = AIA_safeUi_(); // may be null if no UI context
  try {
    if (typeof onOpen_Code === "function") onOpen_Code(ui);
  } catch (err) {
    console.log("onOpen_Code:", err);
  }
  try {
    if (typeof onOpen_Agent === "function") onOpen_Agent(ui);
  } catch (err) {
    console.log("onOpen_Agent:", err);
  }
}

/* ========= NAMESPACE ========= */
var AIA = {
  // AI Integration sheet (NOTE THE SPACE)
  SHEET_NAME: "AI Integration",
  RESULT_COL: 3, // C
  WHEN_COL: 4, // D

  // Candidate sheet
  CANDIDATE_SHEET: "Candidate",
  COL_NO: 1, // A
  COL_URL: 2, // B
  COL_SOURCE: 3, // C

  // OpenAI
  MODEL: "gpt-4o-mini",
  TEMP: 0.2,
};

/* ========= MENU HOOK ========= */
function onOpen_Agent(ui) {
  ui = ui || AIA_safeUi_();
  if (!ui) return;

  ui.createMenu("AI Integration")
    .addItem("▶ Raw Text Data Generation", "AI_runRawTextData") // no selection needed
    .addSeparator()
    .addItem("🔑 Set OpenAI API Key", "AI_setApiKey")
    .addSeparator()
    .addItem("✔ Authorize (first-time only)", "AI_authorize") // helper to trigger scopes
    .addToUi();
}

/* ========= PUBLIC ACTIONS ========= */

/** One-time helper so new users can trigger OAuth scopes easily */
function AI_authorize() {
  // Harmless external request; prompts for "Connect to an external service"
  UrlFetchApp.fetch("https://example.com", { muteHttpExceptions: true });
}

/** Save/Update the OpenAI API key in Script Properties */
function AI_setApiKey() {
  const ui = AIA_safeUi_();
  if (!ui) return;
  const res = ui.prompt(
    "OpenAI API Key",
    'Paste your key (starts with "sk-"). It will be stored in Script Properties.',
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) return;

  const key = (res.getResponseText() || "").trim();
  if (!/^sk-/.test(key)) {
    ui.alert(
      'That does not look like an OpenAI key (should start with "sk-").'
    );
    return;
  }
  PropertiesService.getScriptProperties().setProperty("OPENAI_API_KEY", key);
  ui.alert("Saved. You can now run prompts from the AI menu.");
}

/** Generate Raw Text Data for the row whose Prompt ID = Raw_Data (no selection required) */
function AI_runRawTextData() {
  const ss = SpreadsheetApp.getActive();
  const ui = AIA_safeUi_();

  const sheet = ss.getSheetByName(AIA.SHEET_NAME);
  if (!sheet) {
    if (ui) ui.alert(`Sheet "${AIA.SHEET_NAME}" not found.`);
    return;
  }

  const row = AIA_findPromptRow_("Raw_Data");
  if (!row) {
    if (ui) {
      ss.toast('No row with Prompt ID "Raw_Data" found.', "AI Integration", 5);
      ui.alert(
        "Not found",
        'Could not find any row in column A with Prompt ID "Raw_Data".',
        ui.ButtonSet.OK
      );
    }
    return;
  }

  const template = String(sheet.getRange(row, 2).getValue() || "").trim();
  if (!template) {
    if (ui) ui.alert('No template found in column B for Prompt ID "Raw_Data".');
    return;
  }

  if (AIA_hasNoCandidates_()) {
    AIA_notifyNoCandidates_(sheet, row);
    return;
  }

  const finalPrompt = AIA_buildRawDataPrompt_(template);

  try {
    if (ui) ss.toast("Calling GPT…", "AI Integration", 5);
    const answer = AIA_callOpenAI_(finalPrompt);
    sheet.getRange(row, AIA.RESULT_COL).setValue(answer);
    sheet.getRange(row, AIA.WHEN_COL).setValue(new Date());
    if (ui) ss.toast(`Done: ${AIA_truncate_(answer, 90)}`, "AI Integration", 7);
  } catch (err) {
    const msg = String(err);
    sheet.getRange(row, AIA.RESULT_COL).setValue("ERROR: " + msg);
    sheet.getRange(row, AIA.WHEN_COL).setValue(new Date());
    if (ui) {
      ss.toast(`Error: ${AIA_truncate_(msg, 90)}`, "AI Integration", 8);
      ui.alert("OpenAI Error", msg, ui.ButtonSet.OK);
    }
  }
}

/* ========= ROW LOCATOR ========= */
function AIA_findPromptRow_(id) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(AIA.SHEET_NAME);
  if (!sh) return 0;
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const colA = sh
    .getRange(2, 1, last - 1, 1)
    .getDisplayValues()
    .map((r) =>
      String(r[0] || "")
        .trim()
        .toLowerCase()
    );
  const target = String(id || "")
    .trim()
    .toLowerCase();
  const idx = colA.findIndex((v) => v === target);
  return idx >= 0 ? 2 + idx : 0;
}

/* ========= PROMPT BUILDERS ========= */
function AIA_buildRawDataPrompt_(template) {
  const list = AIA_getCandidateRows_();
  const formatted = AIA_formatCandidateList_(list);
  const inputBlock =
    "\n\n### Input List\n" +
    "Below is the list of company websites to process. For each, generate **Raw Text Data** in the requested output format. If a URL is unreachable, note it explicitly and continue.\n\n" +
    formatted +
    "\n";
  return template + inputBlock;
}

function AIA_getCandidateRows_() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(AIA.CANDIDATE_SHEET);
  if (!sheet) throw new Error(`Sheet "${AIA.CANDIDATE_SHEET}" not found.`);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, 3).getDisplayValues();
  const rows = [];
  for (let i = 0; i < values.length; i++) {
    const [no, url, source] = values[i].map((s) => String(s || "").trim());
    if (!url) continue;
    rows.push({ no: no || String(i + 1), url, source: source || "" });
  }
  return rows;
}

function AIA_formatCandidateList_(rows) {
  if (!rows.length) return "(No candidates found in Candidate sheet)";
  const lines = rows.map(
    (r) => `${r.no}) ${r.url}${r.source ? "  |  Source: " + r.source : ""}`
  );
  return lines.join("\n");
}

/* ========= NO-CANDIDATE HANDLING ========= */
function AIA_hasNoCandidates_() {
  return AIA_getCandidateRows_().length === 0;
}
function AIA_notifyNoCandidates_(sheet, row) {
  const msg = "No Candidate Company list";
  const ss = SpreadsheetApp.getActive();
  sheet.getRange(row, AIA.RESULT_COL).setValue(msg);
  sheet.getRange(row, AIA.WHEN_COL).setValue(new Date());
  const ui = AIA_safeUi_();
  if (ui) {
    ss.toast(msg, "AI Integration", 5);
    ui.alert(msg);
  }
}

/* ========= OPENAI CALL ========= */
function AIA_callOpenAI_(userPrompt) {
  const key =
    PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
  if (!key)
    throw new Error(
      "Missing OpenAI API key. Use “AI Integration → Set OpenAI API Key”."
    );

  const url = "https://api.openai.com/v1/chat/completions";
  const payload = {
    model: AIA.MODEL,
    temperature: AIA.TEMP,
    messages: [
      {
        role: "system",
        content:
          "You are a professional research assistant. Return plain text only unless asked for JSON.",
      },
      { role: "user", content: userPrompt },
    ],
  };

  const resp = UrlFetchApp.fetch(url, {
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
  const answer = data?.choices?.[0]?.message?.content;
  if (!answer) throw new Error("No content returned from OpenAI.");
  return String(answer).trim();
}

/* ========= HELPERS ========= */
function AIA_truncate_(s, n) {
  s = String(s || "");
  return s.length <= n ? s : s.slice(0, n - 1) + "…";
}

/* ========= OPTIONAL: installable onOpen trigger (backup) ========= */
function AI_installOpenTrigger() {
  ScriptApp.getProjectTriggers()
    .filter((t) => t.getHandlerFunction() === "AI_masterOnOpen")
    .forEach((t) => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("AI_masterOnOpen")
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onOpen()
    .create();
}
function AI_masterOnOpen() {
  const ui = AIA_safeUi_();
  try {
    if (typeof onOpen_Code === "function") onOpen_Code(ui);
  } catch (err) {
    console.log("onOpen_Code:", err);
  }
  try {
    if (typeof onOpen_Agent === "function") onOpen_Agent(ui);
  } catch (err) {
    console.log("onOpen_Agent:", err);
  }
}
