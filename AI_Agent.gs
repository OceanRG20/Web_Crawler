/*************************************************
 * AI_Agent.gs — Google Sheets × OpenAI bridge
 * Menu on open:
 *   ▶ Raw Text Data Generation  (Prompt ID = Raw_Data)
 *   ▶ Add new candidates to MMCrawl  (Prompt ID = Add_MMCrawl)
 *   ▶ News Search                 (Prompt ID = News_Search)
 *   🔑 Set OpenAI API Key
 *   ✔ Authorize (first-time only)
 * Results → column C; Timestamp → column D
 **************************************************/

/** ===== Safe UI getter ===== */
function AIA_safeUi_() {
  try {
    return SpreadsheetApp.getUi();
  } catch (_) {
    return null;
  }
}

/** ===== MASTER onOpen ===== */
function onOpen(e) {
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

/* ========= CONFIG ========= */
var AIA = {
  SHEET_NAME: "AI Integration", // tab with prompts
  RESULT_COL: 3, // C
  WHEN_COL: 4, // D

  CANDIDATE_SHEET: "Candidate", // used by Raw_Data & News_Search
  COL_NO: 1,
  COL_URL: 2,
  COL_SOURCE: 3,

  MMCRAWL_SHEET: "MMCrawl", // target DB for Add_MMCrawl
  NEWSRAW_SHEET: "News Raw", // target DB for News_Search

  MODEL: "gpt-4o-mini",
  TEMP: 0.2,
};

/* ========= MENU HOOK ========= */
function onOpen_Agent(ui) {
  ui = ui || AIA_safeUi_();
  if (!ui) return;

  ui.createMenu("AI Integration")
    .addItem("▶ Raw Text Data Generation", "AI_runRawTextData")
    .addItem("▶ Add new candidates to MMCrawl", "AI_runAddMMCrawl")
    .addItem("▶ News Search", "AI_runNewsSearch")
    .addSeparator()
    .addItem("🔑 Set OpenAI API Key", "AI_setApiKey")
    .addSeparator()
    .addItem("✔ Authorize (first-time only)", "AI_authorize")
    .addToUi();
}

/* ========= AUTHORIZATION ========= */
function AI_authorize() {
  // harmless call to request UrlFetch scope for new users
  UrlFetchApp.fetch("https://example.com", { muteHttpExceptions: true });
}

/* ========= API KEY ========= */
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

/* ========= RAW DATA GENERATION (uses Candidate list) ========= */
function AI_runRawTextData() {
  AI_runPromptById_("Raw_Data", (template, row, sheet) => {
    if (AIA_hasNoCandidates_()) {
      AIA_notifyNoCandidates_(sheet, row);
      return null;
    }
    return AIA_buildRawDataPrompt_(template);
  });
}

/* ========= ADD MMCrawl ========= */
function AI_runAddMMCrawl() {
  const ss = SpreadsheetApp.getActive();
  const ui = AIA_safeUi_();
  const integ = ss.getSheetByName(AIA.SHEET_NAME);
  if (!integ) {
    if (ui) ui.alert(`Sheet "${AIA.SHEET_NAME}" not found.`);
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

  const prompt =
    template +
    "\n\n### Input: Raw Text Data\n" +
    "Use the following **Raw Text Data** to construct MMCrawl CSV rows exactly as requested:\n\n" +
    rawInput +
    "\n";

  try {
    if (ui) ss.toast("Running Add_MMCrawl…", "AI Integration", 5);
    const answer = AIA_callOpenAI_(prompt);
    integ.getRange(addRow, AIA.RESULT_COL).setValue(answer);
    integ.getRange(addRow, AIA.WHEN_COL).setValue(new Date());

    const items = AIA_extractJsonArray_(answer);
    if (!items.length) {
      if (ui)
        ss.toast(
          "Add_MMCrawl: could not parse JSON. Check the result cell.",
          "AI Integration",
          6
        );
      return;
    }
    const appended = AIA_appendToMMCrawl_(items);
    if (ui)
      ss.toast(`Added ${appended} row(s) to MMCrawl`, "AI Integration", 6);
  } catch (err) {
    const msg = String(err);
    integ.getRange(addRow, AIA.RESULT_COL).setValue("ERROR: " + msg);
    integ.getRange(addRow, AIA.WHEN_COL).setValue(new Date());
    if (ui) {
      ss.toast(`Error: ${AIA_truncate_(msg, 90)}`, "AI Integration", 8);
      ui.alert("OpenAI Error", msg, ui.ButtonSet.OK);
    }
  }
}

/* ========= NEWS SEARCH =========
 * - Iterates Candidate URLs one by one.
 * - Calls GPT per URL with your News_Search template + "Return JSON only."
 * - Parses and collects JSON; if none, pushes a minimal object with is_estimated:true.
 * - De-duplicates by canonical news_url and (company_name + headline) OR the spaced keys.
 * - Writes a single combined JSON array to the News_Search row’s Result cell.
 * - Also appends results to the "News Raw" sheet (header-mapped).
 */
function AI_runNewsSearch() {
  const ss = SpreadsheetApp.getActive();
  const ui = AIA_safeUi_();
  const sheet = ss.getSheetByName(AIA.SHEET_NAME);
  if (!sheet) {
    if (ui) ui.alert(`Sheet "${AIA.SHEET_NAME}" not found.`);
    return;
  }

  const row = AIA_findPromptRow_("News_Search");
  if (!row) {
    if (ui) ui.alert('Prompt ID "News_Search" not found.');
    return;
  }

  const template = String(sheet.getRange(row, 2).getValue() || "").trim();
  if (!template) {
    if (ui) ui.alert('No template found in column B for "News_Search".');
    return;
  }

  const urls = AIA_getCandidateUrls_();
  if (!urls.length) {
    AIA_notifyNoCandidates_(sheet, row);
    return;
  }

  const collected = [];
  for (let i = 0; i < urls.length; i++) {
    const url = urls[i];
    if (ui)
      ss.toast(
        `News Search ${i + 1}/${urls.length}: ${url}`,
        "AI Integration",
        4
      );

    const prompt =
      template +
      "\n\n### Input Company Website:\n" +
      url +
      "\n\nReturn **JSON array only** with objects that follow the exact schema; no markdown; no commentary.";

    try {
      const ans = AIA_callOpenAI_(prompt);
      const parsed = AIA_extractJsonArray_(ans);
      if (parsed.length) {
        parsed.forEach((o) => collected.push(o));
      } else {
        collected.push({
          "Company Name": AIA_guessNameFromUrl_(url),
          "Company Website URL": url,
          "News Story URL": "",
          Headline: "No news",
          "Publication Date": "",
          "Publisher or Source": "",
          "GPT Summary":
            "No news found after applying filters (local/regional/trade priority; junk excluded).",
          is_estimated: true,
        });
      }
    } catch (err) {
      collected.push({
        "Company Name": AIA_guessNameFromUrl_(url),
        "Company Website URL": url,
        "News Story URL": "",
        Headline: "Error fetching news",
        "Publication Date": "",
        "Publisher or Source": "",
        "GPT Summary": String(err),
        is_estimated: true,
      });
    }
    Utilities.sleep(300); // gentle pacing
  }

  // De-duplicate combined results (supports both key styles)
  const deduped = AIA_dedupeArticles_(collected);

  // Save JSON back to AI Integration
  const jsonOutput = JSON.stringify(deduped, null, 2);
  sheet.getRange(row, AIA.RESULT_COL).setValue(jsonOutput);
  sheet.getRange(row, AIA.WHEN_COL).setValue(new Date());

  // Append to News Raw
  const appended = AIA_appendToNewsRaw_(deduped);
  if (ui)
    ss.toast(
      `News Search complete. ${deduped.length} articles total; ${appended} row(s) added to News Raw.`,
      "AI Integration",
      7
    );

  try {
    if (typeof newsRawRemoveDuplicateStories === "function")
      newsRawRemoveDuplicateStories(false);
  } catch (_) {}
}

/* ========= GENERIC EXECUTION HANDLER ========= */
function AI_runPromptById_(promptId, builderFn) {
  const ss = SpreadsheetApp.getActive();
  const ui = AIA_safeUi_();
  const sheet = ss.getSheetByName(AIA.SHEET_NAME);
  if (!sheet) {
    if (ui) ui.alert(`Sheet "${AIA.SHEET_NAME}" not found.`);
    return;
  }

  const row = AIA_findPromptRow_(promptId);
  if (!row) {
    if (ui) ui.alert(`Prompt ID "${promptId}" not found in column A.`);
    return;
  }

  const template = String(sheet.getRange(row, 2).getValue() || "").trim();
  if (!template) {
    if (ui) ui.alert(`No template found in column B for "${promptId}".`);
    return;
  }

  const finalPrompt = builderFn(template, row, sheet);
  if (!finalPrompt) return;

  try {
    if (ui) ss.toast(`Running ${promptId}…`, "AI Integration", 5);
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

/* ========= PROMPT ROW LOCATOR ========= */
function AIA_findPromptRow_(id) {
  const sh = SpreadsheetApp.getActive().getSheetByName(AIA.SHEET_NAME);
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
    "Below is the list of company websites to process. For each, generate **Raw Text Data** in the requested output format. " +
    "If a URL is unreachable, note it explicitly and continue.\n\n" +
    formatted +
    "\n";
  return template + inputBlock;
}

/* ========= Candidate sheet readers ========= */
function AIA_getCandidateRows_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(AIA.CANDIDATE_SHEET);
  if (!sh) throw new Error(`Sheet "${AIA.CANDIDATE_SHEET}" not found.`);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const values = sh.getRange(2, 1, lastRow - 1, 3).getDisplayValues();
  const rows = [];
  for (let i = 0; i < values.length; i++) {
    const [no, url, source] = values[i].map((s) => String(s || "").trim());
    if (!url) continue;
    rows.push({ no: no || String(i + 1), url, source: source || "" });
  }
  return rows;
}
function AIA_getCandidateUrls_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(AIA.CANDIDATE_SHEET);
  if (!sh) return [];
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  return sh
    .getRange(2, AIA.COL_URL, lastRow - 1, 1)
    .getDisplayValues()
    .map((r) => String(r[0] || "").trim())
    .filter(Boolean);
}
function AIA_formatCandidateList_(rows) {
  if (!rows.length) return "(No candidates found in Candidate sheet)";
  return rows
    .map(
      (r) => `${r.no}) ${r.url}${r.source ? "  |  Source: " + r.source : ""}`
    )
    .join("\n");
}

/* ========= No-candidate handling ========= */
function AIA_hasNoCandidates_() {
  return AIA_getCandidateRows_().length === 0;
}
function AIA_notifyNoCandidates_(sheet, row) {
  const msg = "No Candidate Company list";
  sheet.getRange(row, AIA.RESULT_COL).setValue(msg);
  sheet.getRange(row, AIA.WHEN_COL).setValue(new Date());
  const ui = AIA_safeUi_();
  if (ui) ui.alert(msg);
}
function AIA_setAndNotifyEmpty_(sheet, row, reason) {
  const ss = SpreadsheetApp.getActive();
  const ui = AIA_safeUi_();
  sheet.getRange(row, AIA.RESULT_COL).clearContent();
  sheet.getRange(row, AIA.WHEN_COL).setValue(new Date());
  const msg = reason || "No input available.";
  if (ui) {
    ss.toast(msg, "AI Integration", 5);
    ui.alert(msg);
  }
}

/* ========= OpenAI call ========= */
function AIA_callOpenAI_(userPrompt) {
  const key =
    PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
  if (!key)
    throw new Error(
      "Missing OpenAI API key. Use “AI Integration → Set OpenAI API Key”."
    );

  const resp = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", {
    method: "post",
    contentType: "application/json",
    muteHttpExceptions: true,
    headers: { Authorization: "Bearer " + key },
    payload: JSON.stringify({
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
    }),
  });

  const code = resp.getResponseCode();
  const text = resp.getContentText();
  if (code < 200 || code >= 300) throw new Error(`HTTP ${code}: ${text}`);
  const data = JSON.parse(text);
  const answer = data?.choices?.[0]?.message?.content;
  if (!answer) throw new Error("No content returned from OpenAI.");
  return String(answer).trim();
}

/* ========= JSON helpers ========= */
function AIA_extractJsonArray_(text) {
  if (!text) return [];
  let t = String(text).trim();
  const fence = t.match(/```json([\s\S]*?)```/i) || t.match(/```([\s\S]*?)```/);
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
  if (Array.isArray(obj)) return obj.filter((v) => v && typeof v === "object");
  if (typeof obj === "object") return [obj];
  return [];
}
function AIA_jsonString_(obj) {
  try {
    return JSON.stringify(obj, null, 2);
  } catch (_) {
    return "[]";
  }
}

/* ========= News utils: canonicalize & dedupe (supports both schemas) ========= */
function AIA_canonicalUrl_(u) {
  if (!u) return "";
  try {
    const addProto = /^(https?:)?\/\//i.test(u) ? u : "http://" + u;
    const url = new URL(addProto);
    [
      "utm_source",
      "utm_medium",
      "utm_campaign",
      "utm_term",
      "utm_content",
      "utm_id",
      "gclid",
      "fbclid",
      "mc_cid",
      "mc_eid",
      "igshid",
    ].forEach((k) => url.searchParams.delete(k));
    url.hash = "";
    let host = url.hostname.toLowerCase().replace(/^www\./, "");
    let path = url.pathname.replace(/\/+$/, "") || "/";
    const qs = url.search ? "?" + url.searchParams.toString() : "";
    return host + path + qs;
  } catch (_) {
    return String(u)
      .replace(/^https?:\/\//i, "")
      .replace(/^www\./i, "")
      .replace(/\/+$/, "");
  }
}
function AIA_readUrlFromArticle_(a) {
  return a.news_url || a["News Story URL"] || a.url || "";
}
function AIA_readCompanyFromArticle_(a) {
  return a.company_name || a["Company Name"] || a.company || "";
}
function AIA_readHeadlineFromArticle_(a) {
  return a.headline || a["Headline"] || "";
}
function AIA_dedupeArticles_(arr) {
  const byUrl = new Set();
  const byKey = new Set();
  const out = [];
  for (let i = 0; i < (arr || []).length; i++) {
    const a = arr[i] || {};
    const rawUrl = AIA_readUrlFromArticle_(a);
    const urlKey = rawUrl ? AIA_canonicalUrl_(rawUrl) : "";
    const key =
      String(AIA_readCompanyFromArticle_(a)).toLowerCase().trim() +
      "|" +
      String(AIA_readHeadlineFromArticle_(a)).toLowerCase().trim();
    if (urlKey && byUrl.has(urlKey)) continue;
    if (byKey.has(key)) continue;
    if (urlKey) byUrl.add(urlKey);
    byKey.add(key);
    out.push(a);
  }
  return out;
}
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
          .replace(/\b\w/g, (c) => c.toUpperCase())
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
          .replace(/\b\w/g, (c) => c.toUpperCase())
      : host;
  }
}

/* ========= Append to MMCrawl ========= */
function AIA_appendToMMCrawl_(items) {
  if (!items || !items.length) return 0;
  const sh = SpreadsheetApp.getActive().getSheetByName(AIA.MMCRAWL_SHEET);
  if (!sh) throw new Error(`Missing tab: ${AIA.MMCRAWL_SHEET}`);
  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const matrix = items.map((it) => AIA_mapItemToRow_(it, headers));
  const startRow = sh.getLastRow() + 1;
  sh.getRange(startRow, 1, matrix.length, lastCol).setValues(matrix);
  try {
    if (typeof mmcrawlRemoveDuplicateUrls === "function")
      mmcrawlRemoveDuplicateUrls(false);
  } catch (_) {}
  return matrix.length;
}
function AIA_mapItemToRow_(item, headers) {
  const row = new Array(headers.length).fill("");
  const norm = (v) => (v == null ? "" : String(v).trim());

  const j = {};
  Object.keys(item || {}).forEach((k) => (j[k.toLowerCase()] = item[k]));

  function v() {
    // alias picker
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
    if (typeof getHeaderIndexSmart_ === "function")
      return getHeaderIndexSmart_(headers, name);
    const canon = String(name).toLowerCase();
    return headers.findIndex((h) => String(h).toLowerCase() === canon);
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

/* ========= Append to News Raw (supports both key styles) ========= */
function AIA_appendToNewsRaw_(articles) {
  if (!articles || !articles.length) return 0;
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(AIA.NEWSRAW_SHEET);
  if (!sh) throw new Error(`Missing tab: ${AIA.NEWSRAW_SHEET}`);

  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);

  function colIdx(name) {
    if (typeof getHeaderIndexSmart_ === "function")
      return getHeaderIndexSmart_(headers, name);
    const canon = String(name).toLowerCase();
    return headers.findIndex((h) => String(h).toLowerCase() === canon);
  }

  const cCompany = colIdx("Company Name");
  const cWebsite = colIdx("Company Website URL");
  const cUrl = colIdx("News Story URL");
  const cHeadline = colIdx("Headline");
  const cDate = colIdx("Publication Date");
  const cPublisher = colIdx("Publisher or Source");
  const cSummary = colIdx("GPT Summary");

  function asDate(v) {
    if (!v) return "";
    const s = String(v).trim();
    const d = new Date(s);
    return isNaN(d.getTime()) ? s : d;
  }

  function get(a, spacedKey, snakeKey) {
    return a[spacedKey] != null
      ? a[spacedKey]
      : a[snakeKey] != null
      ? a[snakeKey]
      : "";
  }

  const matrix = articles.map((a) => {
    const row = new Array(lastCol).fill("");
    const company = get(a, "Company Name", "company_name");
    const website = get(a, "Company Website URL", "company_website");
    const url = get(a, "News Story URL", "news_url");
    const headline = get(a, "Headline", "headline");
    const pubDate = get(a, "Publication Date", "publication_date");
    const publisher = get(a, "Publisher or Source", "publisher");
    const summary = get(a, "GPT Summary", "summary");

    if (cCompany >= 0) row[cCompany] = company || "";
    if (cWebsite >= 0) row[cWebsite] = website || "";
    if (cUrl >= 0) row[cUrl] = url || "";
    if (cHeadline >= 0) row[cHeadline] = headline || "";
    if (cDate >= 0) row[cDate] = asDate(pubDate || "");
    if (cPublisher >= 0) row[cPublisher] = publisher || "";
    if (cSummary >= 0) row[cSummary] = summary || "";
    return row;
  });

  if (!matrix.length) return 0;

  const startRow = sh.getLastRow() + 1;
  sh.getRange(startRow, 1, matrix.length, lastCol).setValues(matrix);
  return matrix.length;
}

/* ========= Helpers ========= */
function AIA_truncate_(s, n) {
  s = String(s || "");
  return s.length <= n ? s : s.slice(0, n - 1) + "…";
}
