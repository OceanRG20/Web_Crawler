/*************************************************
 * AI_Agent.gs — Google Sheets × OpenAI bridge
 *
 * Menu on open:
 *   AI Integration
 *     ▶ Raw Text Data Generation      (Prompt ID = Raw_Data, via OpenAI)
 *     ▶ Add new candidates to MMCrawl (Prompt ID = Add_MMCrawl, via OpenAI)
 *     ▶ Reject Company                (Prompt ID = Reject_Company, via OpenAI)
 *     ————
 *     🔑 Set OpenAI API Key
 *     ✔ Authorize (first-time only)
 *
 * Tabs used:
 *   - AI Integration (prompts + results)
 *   - Candidate     (No | URL | Source) — used by Raw_Data
 *   - MMCrawl       (main DB)
 **************************************************/

/** ===== Safe UI getter ===== */
function AIA_safeUi_() {
  try {
    return SpreadsheetApp.getUi();
  } catch (_) {
    return null;
  }
}

/** ===== MASTER onOpen (ordered menus) ===== */
function onOpen(e) {
  const ui = AIA_safeUi_();

  // 1) Record View (if you have a Code / Record View module)
  try {
    if (typeof onOpen_Code === "function") onOpen_Code(ui);
  } catch (err) {
    console.log("onOpen_Code:", err);
  }

  // 2) Auto Filter
  try {
    AIA_addAutoFilterMenu_(ui);
  } catch (err) {
    console.log("AIA_addAutoFilterMenu_:", err);
  }

  // 3) AI Integration (MMCrawl menu)
  try {
    if (typeof onOpen_Agent === "function") onOpen_Agent(ui);
  } catch (err) {
    console.log("onOpen_Agent:", err);
  }

  // 4) ✅ News menu (from News_Agent.gs)
  try {
    if (typeof onOpen_News === "function") onOpen_News(ui);
  } catch (err) {
    console.log("onOpen_News:", err);
  }

  // 4) ✅ Backfill menu (from Backfill_Agent.gs)
  try {
    if (typeof onOpen_Backfill === "function") onOpen_Backfill(ui);
  } catch (err) {
    console.log("onOpen_Backfill:", err);
  }
}

function onInstall(e) {
  onOpen(e);
}

function onInstall(e) {
  onOpen(e);
}

/* ========= CONFIG ========= */
var AIA = {
  SHEET_NAME: "AI Integration", // tab with prompts
  RESULT_COL: 3, // C
  WHEN_COL: 4,   // D
  PREVIEW_COL: 5, // E  (Prompt Preview)

  CANDIDATE_SHEET: "Candidate", // used by Raw_Data
  COL_NO: 1,
  COL_URL: 2,
  COL_SOURCE: 3,

  MMCRAWL_SHEET: "MMCrawl", // target DB for Add_MMCrawl

  // OpenAI (GPT) for Raw_Data, Add_MMCrawl, Reject_Company
  MODEL: "gpt-4o",
  TEMP: 0.2,
  MAX_TOKENS: 5000,

  // Prompt preview (off by default)
  DEBUG_SHOW_PROMPT: false,
  PREVIEW_MAX_CHARS: 18000,
  PREVIEW_ONLY_FIRST_IN_LOOP: true,
};

/* ========= AUTO FILTER MENU INJECTION ========= */
/**
 * Adds the Auto Filter menu. Prefers AutoFilter.addMenu(ui) if present in Auto_Filter.gs,
 * otherwise builds a minimal fallback menu wired to the global wrapper functions.
 */
function AIA_addAutoFilterMenu_(ui) {
  ui = ui || AIA_safeUi_();
  if (!ui) return;

  // Preferred: use the module hook
  try {
    if (typeof AutoFilter !== "undefined" && AutoFilter.addMenu) {
      AutoFilter.addMenu(ui);
      return;
    }
  } catch (err) {
    console.log("AutoFilter.addMenu error:", err);
  }

  // Fallback: create a minimal menu that calls global wrappers in Auto_Filter.gs
  try {
    ui.createMenu("Auto Filter")
      .addItem("Open Filter Dialog…", "AutoFilter_openFilterDialog")
      .addItem("Run Last Filter", "AutoFilter_runLastFilter")
      .addSeparator()
      .addItem("Clear Filter Flags", "AutoFilter_clearFilterFlags")
      .addItem("Show All Rows (remove filter)", "AutoFilter_removeSheetFilter")
      .addToUi();
  } catch (err2) {
    console.log("Auto Filter fallback menu error:", err2);
  }
}

/* ========= MENU HOOK ========= */
function onOpen_Agent(ui) {
  ui = ui || AIA_safeUi_();
  if (!ui) return;

  const menu = ui.createMenu("MMCrawl")
    .addItem("▶ Raw Text Data Generation", "AI_runRawTextData")
    .addItem("▶ Add new candidates to MMCrawl", "AI_runAddMMCrawl")
    .addItem("▶ Reject Company", "AI_runRejectCompany")
    .addSeparator()
    .addItem("🔑 Set OpenAI API Key", "AI_setApiKey")
    .addSeparator()
    .addItem("✔ Authorize (first-time only)", "AI_authorize");

  menu.addToUi();
}

/* ========= AUTHORIZATION ========= */
function AI_authorize() {
  // Dummy call to trigger OAuth scopes
  UrlFetchApp.fetch("https://example.com", { muteHttpExceptions: true });
}

/* ========= API KEY (OpenAI only) ========= */
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
    ui.alert('That does not look like an OpenAI key (should start with "sk-").');
    return;
  }
  PropertiesService.getScriptProperties().setProperty("OPENAI_API_KEY", key);
  ui.alert("Saved. You can now run prompts from the AI menu (OpenAI-based items).");
}

/* ========= RAW DATA GENERATION — per-candidate processing (OpenAI) ========= */
function AI_runRawTextData() {
  const ss = SpreadsheetApp.getActive();
  const ui = AIA_safeUi_();
  const sheet = ss.getSheetByName(AIA.SHEET_NAME);
  if (!sheet) {
    if (ui) ui.alert('Sheet "' + AIA.SHEET_NAME + '" not found.');
    return;
  }

  const row = AIA_findPromptRow_("Raw_Data");
  if (!row) {
    if (ui) ui.alert('Prompt ID "Raw_Data" not found.');
    return;
  }

  const template = String(sheet.getRange(row, 2).getValue() || "").trim();
  if (!template) {
    if (ui) ui.alert('No template found in column B for "Raw_Data".');
    return;
  }

  const candidates = AIA_getCandidateRows_();
  if (!candidates.length) {
    AIA_notifyNoCandidates_(sheet, row);
    return;
  }

  let outputs = [];
  for (let i = 0; i < candidates.length; i++) {
    const obj = candidates[i];
    const no = obj.no;
    const url = obj.url;
    const source = obj.source;
    if (ui) {
      ss.toast(
        "Raw_Data " + (i + 1) + "/" + candidates.length + ": " + url,
        "AI Integration",
        4
      );
    }

    const bundle = AIA_fetchSiteBundleText_(url, 20000);
    const hints = AIA_extractHintsFromBundle_(bundle.htmls || [], url);

    const prompt = AIA_buildSingleRawPrompt_(
      template,
      no,
      url,
      source,
      bundle.text,
      hints
    );

    let ans = "";
    try {
      ans = AIA_callOpenAI_(prompt);
      ans = AIA_fixIdentityNA_(ans, hints);
      ans = AIA_normalizePhoneLine_(ans);

      if (!/\nSource:\s*\n- /i.test(ans)) {
        ans = ans.replace(/\s*$/, "\n\nSource:\n- " + (source || "Not Sure"));
      }
    } catch (err) {
      ans =
        "Row " + no + "\n" + url + "\n\n" + AIA_extractDomain_(url) + "\n" +
        (hints.company || "N/A") + "\n" +
        (hints.street || "N/A") + "\n" +
        (hints.cityStateZip || "N/A") + "\n" +
        "Phone: " + (hints.phoneFmt || "N/A") + "\n\n" +
        "Company: " + (hints.company || "N/A") + "\n\n" +
        "About:\n- ERROR: " + String(err) + "\n\n" +
        "Services:\n- N/A\n\n" +
        "Industries:\n- N/A\n\n" +
        "Employees:\n- N/A\n\n" +
        "Revenue:\n- N/A\n\n" +
        "Square Footage:\n- N/A\n\n" +
        "Ownership:\n- N/A\n\n" +
        "Equipment:\n- null\n\n" +
        "Source:\n- " + (source || "Not Sure");
    }

    outputs.push(String(ans || "").trim());
    Utilities.sleep(250);
  }

  const finalText = outputs.join("\n\n");
  sheet.getRange(row, AIA.RESULT_COL).setValue(finalText);
  sheet.getRange(row, AIA.WHEN_COL).setValue(new Date());
  if (ui) {
    ss.toast(
      "Raw_Data complete: " + candidates.length + " URLs processed.",
      "AI Integration",
      6
    );
  }
}

/* Build a single-company Raw_Data prompt (FORCES Source echo) */
function AIA_buildSingleRawPrompt_(
  template,
  no,
  url,
  source,
  excerptText,
  hints
) {
  const domain = AIA_extractDomain_(url);

  const guidance =
    "\n\n### Important\n" +
    "- Output plain text ONLY in the exact Output Format. No extra headings, no commentary, no JSON.\n" +
    "- The 5-line identity header (Domain, Official Name, Street, City/State ZIP, Phone) MUST NOT contain 'N/A'. Use HINTS below if present and consistent.\n" +
    "- Core identity must come from the SOURCE_EXCERPT or HINTS (derived from the site bundle). If still unknown after both, leave the line blank (do not write 'N/A').\n" +
    "- You MAY supplement Employees, Revenue, Ownership/Leadership/Corporate parent, and Industries with clearly matched public sources (same entity).\n" +
    "  When you do, annotate like: 'N/A (site); ~300 (public: LinkedIn 2024)' or 'N/A (site); ~$17.9M (public: Datanyze est.)'.\n" +
    "- Never mix similarly named companies from other states/domains.\n" +
    "- **Your record MUST end with a 'Source:' section containing exactly one bullet with the input source name provided below.**\n";

  const hintsBlock =
    "### HINTS (parsed from site bundle)\n" +
    "Company Name (guess): " + (hints.company || "") + "\n" +
    "Street Address (guess): " + (hints.street || "") + "\n" +
    "City/State/ZIP (guess): " + (hints.cityStateZip || "") + "\n" +
    "Phone (guess): " + (hints.phoneFmt || hints.phoneRaw || "") + "\n";

  const block =
    "\nRow " + no + "\n" + url + "\n\n" + domain + "\n" +
    "<<SOURCE_EXCERPT_BEGIN>>\n" +
    (excerptText || "") +
    "\n<<SOURCE_EXCERPT_END>>\n" +
    hintsBlock +
    "Input_Source_Name: " + (source || "Not Sure") + "\n";

  return template + guidance + block;
}

/* ========= ADD MMCrawl — per-record processing (OpenAI) ========= */
function AI_runAddMMCrawl() {
  const ss = SpreadsheetApp.getActive();
  const ui = AIA_safeUi_();
  const integ = ss.getSheetByName(AIA.SHEET_NAME);
  if (!integ) {
    if (ui) ui.alert('Sheet "' + AIA.SHEET_NAME + '" not found.');
    return;
  }

  const addRow = AIA_findPromptRow_("Add_MMCrawl");
  if (!addRow) {
    if (ui) ui.alert('Prompt ID "Add_MMCrawl" not found.');
    return;
  }

  const rawRow = AIA_findPromptRow_("Raw_Data");
  if (!rawRow) {
    AIA_setAndNotifyEmpty_(integ, addRow, "Raw_Data row not found.");
    return;
  }

  const template = String(integ.getRange(addRow, 2).getValue() || "").trim();
  if (!template) {
    if (ui) ui.alert('No template found in column B for "Add_MMCrawl".');
    return;
  }

  const rawInput = String(
    integ.getRange(rawRow, AIA.RESULT_COL).getDisplayValue() || ""
  ).trim();
  if (!rawInput || /^no candidate company list$/i.test(rawInput)) {
    AIA_setAndNotifyEmpty_(
      integ,
      addRow,
      'No usable Raw Text Data (empty or "No Candidate Company list").'
    );
    return;
  }

  const candMap = AIA_buildCandidateSourceMap_();
  const chunks = AIA_splitRawDataRows_(rawInput);
  if (!chunks.length) {
    if (ui) ui.alert('Could not detect any "Row X" blocks in Raw Text Data.');
    return;
  }

  let appendedTotal = 0;
  integ.getRange(addRow, AIA.RESULT_COL).clearContent();

  for (let i = 0; i < chunks.length; i++) {
    const piece = chunks[i].trim();
    if (!piece) continue;

    if (ui) {
      ss.toast(
        "Add_MMCrawl " + (i + 1) + "/" + chunks.length,
        "AI Integration",
        4
      );
    }

    const urlInPiece = AIA_extractFirstUrl_(piece);
    const lookedSource = AIA_lookupSourceByUrl_(candMap, urlInPiece) || "";

    const prompt =
      template +
      "\n\n### Input: Raw Text Data (single record)\n" +
      piece +
      "\n\nReturn **JSON array only** (one object) matching the exact MMCrawl schema; no markdown;";

    try {
      const ans = AIA_callOpenAI_(prompt);
      integ.getRange(addRow, AIA.RESULT_COL).setValue(ans);
      integ.getRange(addRow, AIA.WHEN_COL).setValue(new Date());

      const items = AIA_extractJsonArray_(ans);
      if (items.length) {
        items.forEach(function (obj) {
          const hasSrc =
            (obj.Source != null && String(obj.Source).trim() !== "") ||
            (obj.source != null && String(obj.source).trim() !== "");
          if (!hasSrc && lookedSource) obj["Source"] = lookedSource;

          if (!obj["Domain from URL"] && urlInPiece) {
            obj["Domain from URL"] = AIA_extractDomain_(urlInPiece);
          }
          if (!obj["Public Website Homepage URL"] && urlInPiece) {
            obj["Public Website Homepage URL"] = urlInPiece;
          }
        });
        appendedTotal += AIA_appendToMMCrawl_(items);
      }
    } catch (err) {
      integ.getRange(addRow, AIA.RESULT_COL).setValue("ERROR: " + String(err));
      integ.getRange(addRow, AIA.WHEN_COL).setValue(new Date());
    }
    Utilities.sleep(250);
  }

  if (ui) {
    ss.toast(
      "Add_MMCrawl complete. " + appendedTotal + " row(s) added.",
      "AI Integration",
      6
    );
  }
  try {
    if (typeof mmcrawlRemoveDuplicateUrls === "function") {
      mmcrawlRemoveDuplicateUrls(false);
    }
  } catch (_) {}
}

/* Split Raw_Data text into per-row blocks starting at "Row <number>" lines */
function AIA_splitRawDataRows_(rawText) {
  const lines = String(rawText || "").split(/\r?\n/);
  const blocks = [];
  let cur = [];
  function pushCur() {
    if (cur.length) {
      blocks.push(cur.join("\n").trim());
      cur = [];
    }
  }
  for (let i = 0; i < lines.length; i++) {
    const L = lines[i];
    if (/^\s*Row\s+\d+\s*$/i.test(L.trim())) {
      pushCur();
      cur.push(L);
    } else {
      cur.push(L);
    }
  }
  pushCur();
  return blocks.filter(function (b) {
    return /\bRow\s+\d+\b/i.test(b);
  });
}

/* ========= Reject_Company runner ========= */
function AI_runRejectCompany() {
  const ss  = SpreadsheetApp.getActive();
  const ui  = AIA_safeUi_();
  const integ = ss.getSheetByName(AIA.SHEET_NAME);
  if (!integ) {
    if (ui) ui.alert(`Sheet "${AIA.SHEET_NAME}" not found.`);
    return;
  }

  const row = AIA_findPromptRow_("Reject_Company");
  if (!row) {
    if (ui) ui.alert('Prompt ID "Reject_Company" not found in column A.');
    return;
  }

  const template = String(integ.getRange(row, 2).getValue() || "").trim();
  if (!template) {
    if (ui) ui.alert('No template found in column B for "Reject_Company".');
    return;
  }

  const mm = ss.getSheetByName(AIA.MMCRAWL_SHEET);
  if (!mm) {
    if (ui) ui.alert(`Tab "${AIA.MMCRAWL_SHEET}" not found.`);
    return;
  }

  const lastRow = mm.getLastRow();
  const lastCol = mm.getLastColumn();
  if (lastRow <= 1) {
    if (ui) ui.alert("MMCrawl has no data rows.");
    return;
  }

  // Build JSON array of MMCrawl rows
  const headers = mm.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const data    = mm.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const rowsJson = data.map((vals) => {
    const obj = {};
    headers.forEach((h, i) => (obj[h] = vals[i]));
    return obj;
  });

  const prompt =
    template +
    "\n\n### Input: MMCrawl rows (JSON array)\n" +
    JSON.stringify(rowsJson, null, 2) +
    "\n\nReturn ONLY a JSON array of objects as described (no markdown, no comments).";

  try {
    if (ui) ss.toast("Running Reject_Company…", "AI Integration", 5);

    const ans = AIA_callOpenAI_Reject_(prompt);
    integ.getRange(row, AIA.RESULT_COL).setValue(ans);
    integ.getRange(row, AIA.WHEN_COL).setValue(new Date());

    const items = AIA_extractJsonArray_(ans) || [];

    // Filter to only real rejects (Reason starts with "Reject —")
    const rejects = items.filter((it) => {
      const reason = (
        it["Reason"] ||
        it.reason ||
        ""
      ).toString().trim().toLowerCase();
      return reason.startsWith("reject");
    });

    // 1) Append rejects to Rejected Companies
    if (rejects.length) {
      AIA_appendToRejectedCompanies_(rejects);
    }

    // 2) Remove those rejected companies from MMCrawl (match by URL)
    const removedCount = rejects.length
      ? AIA_removeRejectedFromMMCrawl_(rejects)
      : 0;

    if (ui) {
      ss.toast(
        `Reject_Company complete. ${rejects.length} rejected; ${removedCount} row(s) removed from MMCrawl.`,
        "AI Integration",
        7
      );
    }
  } catch (err) {
    const msg = String(err);
    integ.getRange(row, AIA.RESULT_COL).setValue("ERROR: " + msg);
    integ.getRange(row, AIA.WHEN_COL).setValue(new Date());
    if (ui) ui.alert("Reject_Company error", msg, ui.ButtonSet.OK);
  }
}


/* ========= OpenAI call for Reject_Company (separate system prompt) ========= */
function AIA_callOpenAI_Reject_(userPrompt) {
  const key =
    PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
  if (!key) {
    throw new Error(
      'Missing OpenAI API key. Use "AI Integration → Set OpenAI API Key".'
    );
  }

  const payload = {
    model: AIA.MODEL,
    temperature: 0.1,
    max_tokens: 4000,
    messages: [
      {
        role: "system",
        content:
          "You are an M&A analyst working for a private equity client acquiring North American moldmaking and tool & die companies. " +
          "You receive MMCrawl rows as JSON objects from a Google Sheet. Each object is one company record. " +
          "Using the written rules in the user prompt (reject criteria), decide which companies should be moved to the Rejected Companies sheet. " +
          "You must be conservative: only reject when the rules clearly apply. Borderline or ambiguous cases should NOT be rejected. " +
          'Your output must be ONLY a JSON array with objects containing the exact keys: "Company Name", "Company URL", "Reason", "Source".'
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
  if (code < 200 || code >= 300) {
    throw new Error("HTTP " + code + ": " + text);
  }
  const data = JSON.parse(text);
  const answer =
    data &&
    data.choices &&
    data.choices[0] &&
    data.choices[0].message &&
    data.choices[0].message.content;
  if (!answer) {
    throw new Error("No content returned from OpenAI (Reject_Company).");
  }
  return String(answer).trim();
}

/* ========= Move MMCrawl rows to "Rejected Companies" ========= */
/*
Expected JSON from Reject_Company:
[
  {
    "Company Name": "De Boer Tool",
    "Company URL": "https://deboertool.com/",
    "Reason": "Reject — Not moldmaking/tool & die",
    "Source": "MMCrawl fields"
  },
  ...
]
*/
function AIA_appendToRejectedCompanies_(items) {
  if (!items || !items.length) return;

  const ss = SpreadsheetApp.getActive();

  // Use the shared constant if available, else literal tab name
  const rejSh =
    ss.getSheetByName(typeof REJECTED_SHEET !== "undefined" ? REJECTED_SHEET : "Rejected Companies");
  if (!rejSh) throw new Error('Missing tab: "Rejected Companies"');

  const lastCol = rejSh.getLastColumn();
  const headers = rejSh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);

  function colIdx(name) {
    const canon = String(name).toLowerCase();
    return headers.findIndex((h) => String(h).toLowerCase() === canon);
  }

  // Your Rejected Companies headers: A=Company Name, B=URL, C=Reason, D=Source
  const cCompany = colIdx("company name");
  const cUrl     = colIdx("url");
  const cReason  = colIdx("reason");
  const cSource  = colIdx("source");

  const rows = items.map((it) => {
    const row = new Array(lastCol).fill("");

    const name   = (it["Company Name"] || it.company || "").toString().trim();
    const url    = (it["Company URL"] || it.url || "").toString().trim();
    const reason = (it["Reason"] || it.reason || "").toString().trim();
    const source = (it["Source"] || it.source || "").toString().trim();

    if (cCompany >= 0) row[cCompany] = name;
    if (cUrl     >= 0) row[cUrl]     = url;
    if (cReason  >= 0) row[cReason]  = reason;
    if (cSource  >= 0) row[cSource]  = source;

    return row;
  });

  if (!rows.length) return;

  const startRow = rejSh.getLastRow() + 1;
  rejSh.getRange(startRow, 1, rows.length, lastCol).setValues(rows);
}

/* ========= Remove rejected companies from MMCrawl (by URL match) ========= */
/*
  rejectItems: array of objects from Reject_Company with at least:
  {
    "Company Name": "...",
    "Company URL": "https://example.com/",
    "Reason": "Reject — ...",
    ...
  }
*/
function AIA_removeRejectedFromMMCrawl_(rejectItems) {
  if (!rejectItems || !rejectItems.length) return 0;

  const ss = SpreadsheetApp.getActive();
  const mm = ss.getSheetByName(AIA.MMCRAWL_SHEET);
  if (!mm) throw new Error(`Tab "${AIA.MMCRAWL_SHEET}" not found.`);

  const lastRow = mm.getLastRow();
  const lastCol = mm.getLastColumn();
  if (lastRow <= 1) return 0;

  const headers = mm.getRange(1, 1, 1, lastCol).getValues()[0].map(String);

  // Find the URL column in MMCrawl ("Public Website Homepage URL")
  const urlIdx =
    typeof getHeaderIndexSmart_ === "function"
      ? getHeaderIndexSmart_(headers, "Public Website Homepage URL")
      : headers.findIndex(
          (h) =>
            String(h).toLowerCase() ===
            "public website homepage url".toLowerCase()
        );

  if (urlIdx < 0) {
    // No URL column; nothing we can safely delete
    return 0;
  }

  const data = mm.getRange(2, 1, lastRow - 1, lastCol).getValues();

  // Build normalized keys for each MMCrawl row based on its URL
  const rowKeys = data.map((row) => {
    const raw = row[urlIdx] || "";
    return AIA_normalizeCandidateUrl_(raw);
  });

  const rowsToDelete = new Set();

  // For each rejected company, normalize its URL and match against MMCrawl rows
  rejectItems.forEach((it) => {
    const url = (
      it["Company URL"] ||
      it["Public Website Homepage URL"] ||
      it.url ||
      ""
    )
      .toString()
      .trim();
    if (!url) return;

    const key = AIA_normalizeCandidateUrl_(url);
    if (!key) return;

    rowKeys.forEach((rk, i) => {
      if (!rk) return;
      if (rk === key) {
        // Row index in sheet = 2 (data starts) + i
        rowsToDelete.add(2 + i);
      }
    });
  });

  const toDelete = Array.from(rowsToDelete).sort((a, b) => b - a);
  toDelete.forEach((r) => mm.deleteRow(r));

  return toDelete.length;
}

/* ========= Helpers ========= */
function AIA_truncate_(s, n) {
  s = String(s || "");
  return s.length <= n ? s : s.slice(0, n - 1) + "…";
}

function AIA_extractDomain_(url) {
  try {
    const u = new URL(/^(https?:)?\/\//i.test(url) ? url : "http://" + url);
    return u.hostname.replace(/^www\./i, "");
  } catch (e) {
    return String(url || "").replace(/^https?:\/\//i, "").split("/")[0];
  }
}

/* ========= Site bundle fetch + hints extraction ========= */
function AIA_fetchSiteBundleText_(baseUrl, maxChars) {
  const root = String(baseUrl || "").replace(/\/+$/, "");
  if (!root) return { text: "", htmls: [] };

  const paths = [
    "",
    "/contact",
    "/contact/",
    "/contact-us",
    "/contactus.html/",
    "/contact.html",
    "/contact-us.html",
    "/contact_us.html",
    "/contact.php",
    "/who-to-contact/",
    "/contact-us/",
    "/about",
    "/about/",
    "/about-us",
    "/about-us/",
    "/about.html",
    "/about-us.html",
    "/about.php",
    "/locations",
    "/locations/",
    "/location",
    "/location/",
    "/location.html",
  ];

  const urls = [];
  paths.forEach(function (p) {
    const u = p ? root + p : root;
    if (urls.indexOf(u) === -1) urls.push(u);
  });

  const htmls = [];
  const texts = [];
  urls.forEach(function (u) {
    try {
      const resp = UrlFetchApp.fetch(u, {
        muteHttpExceptions: true,
        followRedirects: true,
        validateHttpsCertificates: true,
        headers: {
          "User-Agent": "Mozilla/5.0 (compatible; GoogleAppsScript/1.0)",
        },
      });
      const html = resp.getContentText();
      htmls.push(html);
      const text = html
        .replace(/<script[\s\S]*?<\/script>/gi, "")
        .replace(/<style[\s\S]*?<\/style>/gi, "")
        .replace(/<[^>]+>/g, " ")
        .replace(/\s+/g, " ")
        .trim();
      texts.push(text);
    } catch (e) {
      // ignore per-URL errors
    }
    Utilities.sleep(120);
  });
  const joined = texts.join("\n").slice(0, maxChars || 20000);
  return { text: joined, htmls };
}

/* ===== Helper: choose best phone candidate from HTML ===== */
function AIA_pickBestPhoneFromHtml_(htmlText) {
  const txt = String(htmlText || "");
  if (!txt) return "";

  const re =
    /(?:\+?1[\s\-\.\)]*)?\(?\d{3}\)?[\s\-\.]?\d{3}[\s\-\.]?\d{4}(?:\s*(?:ext|x)\s*\d{1,5})?/gi;

  let best = "";
  let bestScore = -1;
  let m;
  while ((m = re.exec(txt)) !== null) {
    const raw = m[0];
    const idx = m.index;
    const windowStart = Math.max(0, idx - 120);
    const windowEnd = Math.min(txt.length, idx + 120);
    const ctx = txt.slice(windowStart, windowEnd).toLowerCase();

    let score = 0;
    if (ctx.includes("phone") || ctx.includes("tel") || ctx.includes("call")) {
      score += 3;
    }
    if (ctx.match(/\b(ext\.?|extension)\b/)) score += 1;
    if (ctx.match(/\b[A-Z]{2}\s*\d{5}(?:-\d{4})?\b/i)) score += 2;
    if (
      ctx.match(
        /\b(st(?:reet)?|rd\.?|road|dr\.?|drive|hwy\.?|highway|parkway|pkwy\.?)\b/i
      )
    ) {
      score += 1;
    }

    if (score > bestScore) {
      bestScore = score;
      best = raw;
    }
  }
  return best;
}

function AIA_extractHintsFromBundle_(htmls, url) {
  const out = {
    company: "",
    street: "",
    cityStateZip: "",
    phoneRaw: "",
    phoneFmt: "",
  };
  const all = (htmls || []).join("\n");

  try {
    const blocks =
      all.match(
        /<script[^>]*type=["']application\/ld\+json["'][^>]*>([\s\S]*?)<\/script>/gi
      ) || [];
    for (let i = 0; i < blocks.length; i++) {
      const m = blocks[i].match(/<script[^>]*>([\s\S]*?)<\/script>/i);
      if (!m) continue;
      const raw = m[1];
      try {
        const obj = JSON.parse(raw);
        const cand = Array.isArray(obj) ? obj : [obj];
        for (let j = 0; j < cand.length; j++) {
          const o = cand[j];
          const n = o.name || o.legalName || "";
          const tel = (o.telephone || o.phone || "").toString();
          const adr = o.address || {};
          const street = adr.streetAddress || "";
          const city = adr.addressLocality || "";
          const state = adr.addressRegion || "";
          const zip = adr.postalCode || "";
          if (!out.company && n) out.company = String(n).trim();
          if (!out.phoneRaw && tel) out.phoneRaw = String(tel).trim();
          if (!out.street && street) out.street = String(street).trim();
          const csz = [city, state, zip]
            .filter(Boolean)
            .join(", ")
            .replace(", ,", ",");
          if (!out.cityStateZip && csz) out.cityStateZip = csz;
        }
      } catch (_) {}
    }
  } catch (_) {}

  try {
    if (!out.company) {
      const t = all.match(/<title[^>]*>([\s\S]*?)<\/title>/i);
      if (t) {
        out.company = String(t[1] || "")
          .replace(/\s*\|.*$/, "")
          .replace(/[-–—].*$/, "")
          .trim();
      }
    }
  } catch (_) {}

  if (!out.phoneRaw) {
    const best = AIA_pickBestPhoneFromHtml_(all);
    if (best) out.phoneRaw = best;
  }
  out.phoneFmt = AIA_formatNANP_(out.phoneRaw);

  if (!out.street || !out.cityStateZip) {
    const streetMatch = all.match(
      /\b\d{1,6}\s+[A-Za-z0-9\.\- ]+\s(?:Road|Rd\.?|Street|St\.?|Drive|Dr\.?|Avenue|Ave\.?|Boulevard|Blvd\.?|Lane|Ln\.?|Court|Ct\.?|Parkway|Pkwy\.?|Circle|Cir\.?)\b[^\n<]{0,80}/i
    );
    if (streetMatch && !out.street) out.street = streetMatch[0].trim();
    const usMatch = all.match(
      /\b([A-Za-z][A-Za-z\.\- ]+),\s*([A-Z]{2})\s*(\d{5}(?:-\d{4})?)\b/
    );
    const caMatch = all.match(
      /\b([A-Za-z][A-Za-z\.\- ]+),\s*(AB|BC|MB|NB|NL|NS|NT|NU|ON|PE|QC|SK|YT)\s*([A-Z]\d[A-Z]\s?\d[A-Z]\d)\b/
    );
    if (!out.cityStateZip && (usMatch || caMatch)) {
      const mm = usMatch || caMatch;
      out.cityStateZip = (mm[1].trim() + ", " + mm[2] + " " + mm[3])
        .replace(/\s+/g, " ")
        .trim();
    }
  }

  if (!out.company) out.company = AIA_guessNameFromUrl_(url);
  return out;
}

/* ========= Post-processing: identity & phone normalization ========= */
function AIA_formatNANP_(raw) {
  if (!raw) return "";
  const digits = String(raw).replace(/[^\d]/g, "");
  if (digits.length === 11 && digits.startsWith("1")) {
    return (
      "(" +
      digits.slice(1, 4) +
      ") " +
      digits.slice(4, 7) +
      "-" +
      digits.slice(7, 11)
    );
  }
  if (digits.length === 10) {
    return (
      "(" +
      digits.slice(0, 3) +
      ") " +
      digits.slice(3, 6) +
      "-" +
      digits.slice(6, 10)
    );
  }
  return "";
}

function AIA_normalizePhoneLine_(recordText) {
  const lines = String(recordText || "").split(/\r?\n/);
  for (let i = 0; i < lines.length; i++) {
    if (/^\s*Phone\s*:/.test(lines[i])) {
      const val = lines[i].replace(/^\s*Phone\s*:\s*/i, "").trim();
      const fmt = AIA_formatNANP_(val);
      lines[i] = ("Phone: " + (fmt || (val ? val : ""))).trim();
      break;
    }
  }
  return lines.join("\n");
}

function AIA_fixIdentityNA_(recordText, hints) {
  const txt = String(recordText || "");
  const parts = txt.split(/^\s*Company\s*:/im);
  if (parts.length < 2) return txt;
  const before = parts[0];
  const afterStart = txt.slice(before.length);

  const lines = before.split(/\r?\n/);

  function replaceIfNA(idx, val) {
    if (!lines[idx]) return;
    const s = lines[idx].trim();
    if (!s || /^N\/A$/i.test(s)) lines[idx] = val || "";
  }

  let domainIdx = -1;
  for (let i = 0; i < lines.length; i++) {
    const s = lines[i].trim();
    if (
      s &&
      s.includes(".") &&
      !/\s/.test(s) &&
      !/^http/i.test(s) &&
      !/^row/i.test(s)
    ) {
      domainIdx = i;
      break;
    }
  }

  if (domainIdx >= 0) {
    const nameIdx = domainIdx + 1;
    const streetIdx = domainIdx + 2;
    const cityIdx = domainIdx + 3;
    const phoneIdx = domainIdx + 4;

    replaceIfNA(nameIdx, hints.company || "");
    replaceIfNA(streetIdx, hints.street || "");
    replaceIfNA(cityIdx, hints.cityStateZip || "");
    if (/^\s*Phone\s*:/.test(lines[phoneIdx] || "")) {
      const fmt = hints.phoneFmt || "";
      if (fmt) lines[phoneIdx] = "Phone: " + fmt;
      else lines[phoneIdx] = lines[phoneIdx].replace(/N\/A/i, "").trim();
    }
  }
  return lines.join("\n") + afterStart;
}

/* ========= Cross-file helper ========= */
function getHeaderIndexSmart_(headers, name) {
  const canon = String(name || "")
    .toLowerCase()
    .normalize("NFKC")
    .replace(/[\u2010-\u2015]/g, "-")
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
  for (let i = 0; i < headers.length; i++) {
    const h = String(headers[i] || "")
      .toLowerCase()
      .normalize("NFKC")
      .replace(/[\u2010-\u2015]/g, "-")
      .replace(/[^a-z0-9]+/g, " ")
      .trim();
    if (h === canon) return i;
  }
  const keys = canon.split(" ").filter(Boolean);
  for (let i = 0; i < headers.length; i++) {
    const h = String(headers[i] || "")
      .toLowerCase()
      .normalize("NFKC")
      .replace(/[\u2010-\u2015]/g, "-")
      .replace(/[^a-z0-9]+/g, " ")
      .trim();
    if (keys.every(function (k) { return h.includes(k); })) return i;
  }
  return -1;
}

/* ========= Missing helper referenced above ========= */
function AIA_setAndNotifyEmpty_(sheet, row, msg) {
  sheet.getRange(row, AIA.RESULT_COL).setValue("[]");
  sheet.getRange(row, AIA.WHEN_COL).setValue(new Date());
  const ui = AIA_safeUi_();
  if (ui) ui.alert(msg || "No input.");
}

function AIA_notifyNoCandidates_(sheet, row) {
  AIA_setAndNotifyEmpty_(
    sheet,
    row,
    "No candidate URLs found in Candidate sheet."
  );
}

/* ========= NEW: Candidate Source backfill helpers (for Add_MMCrawl only) ========= */
function AIA_buildCandidateSourceMap_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(AIA.CANDIDATE_SHEET);
  const map = {};
  if (!sh) return map;
  const last = sh.getLastRow();
  if (last < 2) return map;
  const vals = sh.getRange(2, 1, last - 1, 3).getDisplayValues();
  for (let i = 0; i < vals.length; i++) {
    const url = String(vals[i][1] || "").trim();
    const src = String(vals[i][2] || "").trim();
    if (!url) continue;
    const key = AIA_normalizeCandidateUrl_(url);
    if (key) map[key] = { source: src, url: url };
  }
  return map;
}

function AIA_normalizeCandidateUrl_(u) {
  try {
    const url = new URL(/^(https?:)?\/\//i.test(u) ? u : "http://" + u);
    return (
      url.hostname.replace(/^www\./i, "") +
      (url.pathname === "/" ? "" : url.pathname.replace(/\/+$/, ""))
    );
  } catch (_) {
    return String(u)
      .replace(/^https?:\/\//i, "")
      .replace(/^www\./i, "")
      .replace(/\/+$/, "");
  }
}

function AIA_extractFirstUrl_(text) {
  const m = String(text || "").match(/https?:\/\/\S+/i);
  return m ? m[0].trim() : "";
}

function AIA_lookupSourceByUrl_(candMap, url) {
  if (!url) return "";
  const key = AIA_normalizeCandidateUrl_(url);
  return (candMap[key] && candMap[key].source) || "";
}

/* ========= Candidate sheet readers ========= */
function AIA_getCandidateRows_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(AIA.CANDIDATE_SHEET);
  if (!sh) throw new Error('Sheet "' + AIA.CANDIDATE_SHEET + '" not found.');
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const vals = sh.getRange(2, 1, lastRow - 1, 3).getDisplayValues();
  const out = [];
  for (let i = 0; i < vals.length; i++) {
    const row = vals[i].map(function (s) {
      return String(s || "").trim();
    });
    const no = row[0];
    const url = row[1];
    const source = row[2];
    if (!url) continue;
    out.push({ no: no || String(i + 1), url: url, source: source || "" });
  }
  return out;
}

/* ========= OpenAI call (Raw_Data, Add_MMCrawl) ========= */
function AIA_callOpenAI_(userPrompt) {
  const key =
    PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
  if (!key) {
    throw new Error(
      'Missing OpenAI API key. Use "AI Integration → Set OpenAI API Key".'
    );
  }

  const resp = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", {
    method: "post",
    contentType: "application/json",
    muteHttpExceptions: true,
    headers: { Authorization: "Bearer " + key },
    payload: JSON.stringify({
      model: AIA.MODEL,
      temperature: AIA.TEMP,
      max_tokens: AIA.MAX_TOKENS,
      messages: [
        {
          role: "system",
          content:
            "You extract company profiles using the provided SOURCE_EXCERPT (site-first grounding) and HINTS parsed from the site. " +
            'Identity header (Domain, Official Name, Street, City/State ZIP, Phone) must not contain "N/A"; prefer HINTS when available. ' +
            "You may supplement Employees, Square footage (facility), and Estimated Revenues ONLY if clearly matched public sources refer to the SAME entity (same domain + city/state). " +
            'When you use public data, annotate provenance in parentheses (e.g., "~80 (public: LinkedIn 2024)", "~$17.9M (public: ZoomInfo 2024)", "50,000 (public: BF 2024)"). ' +
            "For Square footage (facility), Number of employees, and Estimated Revenues, always provide numeric/currency values and keep distinct datapoints separated by semicolons. " +
            "The visible tokens before parentheses for these three fields must be numeric/currency only (with optional ~, +, or ranges). " +
            'Use parentheses for sources and explanations (e.g., "(site)", "(public: LinkedIn 2024)", "(calc from employees)"). ' +
            "Never average conflicting estimates into a single number; preserve distinct datapoints. " +
            "For Equipment, aggressively extract and normalize any CNC/EDM/gun drill/laser welder/CMM/tryout press information; only output 'Equipment: null' if no machinery at all is described on the site. " +
            "Never mix similarly named companies from other states/domains. Return plain text only in the exact Output Format. " +
            'Never output bare numeric-only values for Square footage (facility), Number of employees, or Estimated Revenues; always include at least one parenthetical source or explanation (e.g., "(site)", "(public: ZoomInfo 2024)", "(estimate)").'
        },
        { role: "user", content: userPrompt },
      ],
    }),
  });

  const code = resp.getResponseCode();
  const text = resp.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error("HTTP " + code + ": " + text);
  }
  const data = JSON.parse(text);
  const answer =
    data &&
    data.choices &&
    data.choices[0] &&
    data.choices[0].message &&
    data.choices[0].message.content;
  if (!answer) {
    throw new Error("No content returned from OpenAI.");
  }
  return String(answer).trim();
}

/* ========= Prompt row locator ========= */
function AIA_findPromptRow_(id) {
  const sh = SpreadsheetApp.getActive().getSheetByName(AIA.SHEET_NAME);
  if (!sh) return 0;
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const colA = sh
    .getRange(2, 1, last - 1, 1)
    .getDisplayValues()
    .map(function (r) {
      return String(r[0] || "").trim().toLowerCase();
    });
  const target = String(id || "").trim().toLowerCase();
  const idx = colA.findIndex(function (v) {
    return v === target;
  });
  return idx >= 0 ? 2 + idx : 0;
}

/* ========= JSON helpers ========= */
function AIA_extractJsonArray_(text) {
  if (!text) return [];
  let t = String(text).trim();
  const fence =
    t.match(/```json([\s\S]*?)```/i) || t.match(/```([\s\S]*?)```/);
  if (fence) t = fence[1].trim();
  let obj = null;
  try {
    obj = JSON.parse(t);
  } catch (_) {
    const m = t.match(/(\{[\s\S]*\}|\[[\s\S]*\])/);
    if (m) {
      try {
        obj = JSON.parse(m[1]);
      } catch (_) {}
    }
  }
  if (!obj) return [];
  if (Array.isArray(obj)) {
    return obj.filter(function (v) {
      return v && typeof v === "object";
    });
  }
  if (typeof obj === "object") return [obj];
  return [];
}

function AIA_extractJsonObject_(text) {
  const arr = AIA_extractJsonArray_(text);
  if (arr.length === 1) return arr[0];
  if (arr.length > 1) return arr[0];
  try {
    return JSON.parse(String(text || ""));
  } catch (_) {
    return {};
  }
}

function AIA_jsonString_(obj) {
  try {
    return JSON.stringify(obj, null, 2);
  } catch (_) {
    return "[]";
  }
}

/* ========= Guess Name From URL (used by Raw_Data hints) ========= */
function AIA_guessNameFromUrl_(url) {
  try {
    const host = new URL(
      /^(https?:)?\/\//.test(url) ? url : "http://" + url
    ).hostname.replace(/^www\./, "");
    const base = host.split(".")[0];
    return base
      ? base
          .replace(/[-_]/g, " ")
          .replace(/\s+/g, " ")
          .trim()
          .replace(/\b\w/g, function (c) {
            return c.toUpperCase();
          })
      : host;
  } catch (_) {
    const host = String(url || "")
      .replace(/^https?:\/\//, "")
      .replace(/^www\./, "")
      .split("/")[0];
    const base = host.split(".")[0];
    return base
      ? base
          .replace(/[-_]/g, " ")
          .replace(/\s+/g, " ")
          .trim()
          .replace(/\b\w/g, function (c) {
            return c.toUpperCase();
          })
      : host;
  }
}

/* ========= Append to MMCrawl ========= */
function AIA_appendToMMCrawl_(items) {
  if (!items || !items.length) return 0;
  const sh = SpreadsheetApp.getActive().getSheetByName(AIA.MMCRAWL_SHEET);
  if (!sh) throw new Error("Missing tab: " + AIA.MMCRAWL_SHEET);
  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const matrix = items.map(function (it) {
    return AIA_mapItemToRow_(it, headers);
  });
  const startRow = sh.getLastRow() + 1;
  sh.getRange(startRow, 1, matrix.length, lastCol).setValues(matrix);
  try {
    if (typeof mmcrawlRemoveDuplicateUrls === "function") {
      mmcrawlRemoveDuplicateUrls(false);
    }
  } catch (_) {}
  return matrix.length;
}

function AIA_mapItemToRow_(item, headers) {
  const row = new Array(headers.length).fill("");
  const norm = function (v) {
    return v == null ? "" : String(v).trim();
  };

  const j = {};
  Object.keys(item || {}).forEach(function (k) {
    j[k.toLowerCase()] = item[k];
  });

  function v() {
    for (let i = 0; i < arguments.length; i++) {
      const key = String(arguments[i]).toLowerCase();
      if (key in j) {
        const val = norm(j[key]);
        if (val) return val;
      }
    }
    return "";
  }

  function hIdx(name) {
    const canon = String(name).toLowerCase();
    return headers.findIndex(function (h) {
      return String(h).toLowerCase() === canon;
    });
  }

  function put(name, value) {
    if (!value) return;
    const i = hIdx(name);
    if (i >= 0) row[i] = value;
  }

  const Hmap = typeof H === "object" && H ? H : {};
  const CN = Hmap.company || "Company Name";
  const TS = Hmap.status || "Target Status";
  const WEB = Hmap.website || "Public Website Homepage URL";
  const DOM = Hmap.domain || "Domain from URL";
  const SRC = Hmap.source || "Source";
  const STR = Hmap.street || "Street Address";
  const CTY = Hmap.city || "City";
  const STA = Hmap.state || "State";
  const ZIP = Hmap.zip || "Zipcode";
  const PHN = Hmap.phone || "Phone";
  const IND = Hmap.industries || "Industries served";
  const PRD = Hmap.products || "Products and services offered";

  put(CN, v("Company Name", "company", "name", "company_name"));
  put(TS, v("Target Status", "target status", "status"));
  put(
    "Status (proposed)",
    v("Status (proposed)", "status (proposed)", "status_proposed")
  );
  put(WEB, v("Public Website Homepage URL", "website", "url", "homepage"));
  put(DOM, v("Domain from URL", "domain", "host"));
  put(SRC, v("Source", "source"));
  put(STR, v("Street Address", "street address", "street", "address"));
  put(CTY, v("City", "city", "town"));
  put(STA, v("State", "state", "province", "region code"));
  put(ZIP, v("Zipcode", "zipcode", "zip", "postal code", "postcode"));
  put(PHN, v("Phone", "telephone", "phone number"));
  put(IND, v("Industries served", "industries served", "industries"));
  put(
    PRD,
    v(
      "Products and services offered",
      "products and services offered",
      "products",
      "services"
    )
  );

  put(
    "Square footage (facility)",
    v("Square footage (facility)", "facility size", "square footage", "sqft")
  );
  put(
    "Number of employees",
    v("Number of employees", "# employees", "employees", "headcount")
  );
  put(
    "Estimated Revenues",
    v("Estimated Revenues", "estimated revenue", "revenue", "revenues")
  );
  put(
    "Years of operation",
    v("Years of operation", "years", "years in business")
  );
  put("Ownership", v("Ownership", "owner", "ownership / owner"));
  put("Equipment", v("Equipment"));
  put("CNC 3-axis", v("CNC 3-axis", "cnc 3 axis", "3-axis", "3 axis"));
  put("CNC 5-axis", v("CNC 5-axis", "cnc 5 axis", "5-axis", "5 axis"));
  put(
    "Spares/ Repairs",
    v("Spares/ Repairs", "spares/repairs", "repairs", "spares")
  );
  put("Family business", v("Family business", "family"));
  put("2nd Address", v("2nd Address", "second address", "address 2"));
  put("Region", v("Region"));
  put("Medical", v("Medical"));
  put(
    "Notes (Approach/ Contacts/ Info)",
    v(
      "Notes (Approach/ Contacts/ Info)",
      "notes",
      "notes (approach/contacts/info)"
    )
  );

  return row;
}
