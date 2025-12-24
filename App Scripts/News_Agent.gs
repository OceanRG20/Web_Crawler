/*************************************************
 * News_Agent.gs — News menu + GPT news crawler + URL repair
 *
 * Sheets used:
 *   - "AI Integration"   (prompts + result log)
 *   - "News Source"      (input/output: Company URL, Name, News URL, Source)
 *   - "News Raw"         (output: normalized news rows)
 *
 * Prompt IDs in AI Integration:
 *   - "News_Search"  (news crawling / summarization)
 *   - "News_Url"     (URL validation + repair; stores FINAL rebuilt rows list)
 **************************************************/

/** ===== Menu hook (called from AI_Agent.onOpen) ===== */
function onOpen_News(ui) {
  ui = ui || SpreadsheetApp.getUi();
  ui.createMenu("News")
    .addItem("▶ News Searching", "NS_runNewsSearch")
    .addItem("✅ Check News Source URLs", "NS_checkNewsSourceUrls")
    .addToUi();
}

/** ===== Special Values defaults (expanded) ===== */
const NS_DEFAULT_SPECIAL_VALUES = {
  "Square footage (facility)": "",
  "Number of employees": "",
  "Estimated Revenues": "",
  "Years of operation": "",
  "Ownership": "",
  "Equipment": "",
  "Spares/ Repairs": "",
  "Family business": "",
  "2nd Address": "",
  "Region": "",
  "Medical": ""
};

/** ===== Main runner (kept as your current logic) ===== */
/** ===== Main runner (UPDATED: do NOT skip any News Source rows) ===== */
function NS_runNewsSearch() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();

  const aiSheet = ss.getSheetByName("AI Integration");
  if (!aiSheet) {
    ui.alert('Sheet "AI Integration" not found.');
    return;
  }

  const promptRow = NS_findPromptRow_("News_Search");
  if (!promptRow) {
    ui.alert('Prompt ID "News_Search" not found in AI Integration!A:A.');
    return;
  }

  const basePrompt = String(aiSheet.getRange(promptRow, 2).getValue() || "").trim();
  if (!basePrompt) {
    ui.alert('No prompt content in "AI Integration" for Prompt ID "News_Search".');
    return;
  }

  const newsRows = NS_getRowsFromNewsSource_();
  if (!newsRows.length) {
    ui.alert('No data rows found in sheet "News Source".');
    aiSheet.getRange(promptRow, 3).setValue("[]");
    aiSheet.getRange(promptRow, 4).setValue(new Date());
    return;
  }

  ui.alert(
    "News Search",
    "Processing " + newsRows.length + " news URL row(s) from News Source…",
    ui.ButtonSet.OK
  );

  const allArticles = [];
  let idx = 0;

  // Fallback object to guarantee at least 1 output per input row
  function makeFallbackNoNewsObject_(row, reason) {
    const now = new Date();
    const isoDate = now.toISOString().slice(0, 10);
    const timeStr = now.toTimeString().slice(0, 8);

    return {
      "Company Name": row.companyName || "",
      "Company Website URL": row.companyUrl || "",
      "News Story URL": row.newsUrl || "",
      "Headline": "No news found",
      "Publication Date": "",
      "Publisher or Source": "",
      "GPT Summary":
        "No news found after applying filters on " +
        isoDate + " " + timeStr +
        (reason ? (". Reason: " + String(reason)) : "."),
      "Confidence Score": "",
      "Source": row.source || "",
      "Special Values": JSON.parse(JSON.stringify(NS_DEFAULT_SPECIAL_VALUES)),
      "MMCrawl Updates": ""
    };
  }

  newsRows.forEach((row) => {
    idx++;

    ss.toast(
      "News Search " + idx + "/" + newsRows.length + " — " + (row.newsUrl || row.companyUrl),
      "News",
      5
    );

    // IMPORTANT CHANGE:
    // - Do NOT pre-fetch / soft-404 repair / 404/429 logic here.
    // - Always call News_Search for every row as provided in News Source.

    try {
      const fullPrompt = NS_buildPromptForRow_(basePrompt, row);
      const rawAnswer = NS_callOpenAIForNews_(fullPrompt);
      const articles = NS_extractJsonArray_(rawAnswer);

      // Guarantee at least 1 output row per News Source input row
      const finalArticles = (articles && articles.length)
        ? articles
        : [makeFallbackNoNewsObject_(row, "Model returned empty/invalid JSON array for this input.")];

      const enriched = finalArticles.map((obj) => {
        let copy = Object.assign({}, obj);

        if (!copy["Source"]) copy["Source"] = row.source || "";
        if (!copy["Company Name"]) copy["Company Name"] = row.companyName || "";
        if (!copy["Company Website URL"]) copy["Company Website URL"] = row.companyUrl || "";
        if (!copy["News Story URL"]) copy["News Story URL"] = row.newsUrl || "";

        if (!copy.hasOwnProperty("Special Values")) {
          copy["Special Values"] = JSON.parse(JSON.stringify(NS_DEFAULT_SPECIAL_VALUES));
        } else if (!copy["Special Values"] || typeof copy["Special Values"] !== "object") {
          copy["Special Values"] = JSON.parse(JSON.stringify(NS_DEFAULT_SPECIAL_VALUES));
        } else {
          Object.keys(NS_DEFAULT_SPECIAL_VALUES).forEach((k) => {
            if (!copy["Special Values"].hasOwnProperty(k)) copy["Special Values"][k] = "";
          });
        }

        if (!copy.hasOwnProperty("MMCrawl Updates")) copy["MMCrawl Updates"] = "";

        copy = NS_enforceSpecialValuesFromSummary_(copy);
        return copy;
      });

      allArticles.push.apply(allArticles, enriched);

    } catch (err) {
      // Guarantee at least 1 output for hard failures too
      allArticles.push(makeFallbackNoNewsObject_(row, "Error while running News_Search: " + String(err)));
    }

    Utilities.sleep(250);
  });

  const jsonOut = JSON.stringify(allArticles, null, 2);
  aiSheet.getRange(promptRow, 3).setValue(jsonOut);
  aiSheet.getRange(promptRow, 4).setValue(new Date());

  NS_writeResultsToNewsRaw_(allArticles);

  ss.toast(
    "News Search complete — " + allArticles.length + " article object(s) created.",
    "News",
    8
  );
}


/** ===== Prompt row locator in AI Integration ===== */
function NS_findPromptRow_(id) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName("AI Integration");
  if (!sh) return 0;
  const last = sh.getLastRow();
  if (last < 2) return 0;

  const vals = sh
    .getRange(2, 1, last - 1, 1)
    .getDisplayValues()
    .map((r) => String(r[0] || "").trim().toLowerCase());

  const needle = String(id || "").trim().toLowerCase();
  const idx = vals.findIndex((v) => v === needle);
  return idx >= 0 ? 2 + idx : 0;
}

/** ===== Read rows from "News Source" ===== */
function NS_getRowsFromNewsSource_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName("News Source");
  if (!sh) throw new Error('Sheet "News Source" not found.');

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const vals = sh.getRange(2, 1, lastRow - 1, 4).getDisplayValues();
  const out = [];

  vals.forEach((r, i) => {
    const companyUrl = String(r[0] || "").trim();
    const companyName = String(r[1] || "").trim();
    const newsUrl = String(r[2] || "").trim();
    const source = String(r[3] || "").trim();

    if (!companyUrl && !newsUrl) return;

    out.push({
      rowIndex: i + 2,
      companyUrl: companyUrl,
      companyName: companyName,
      newsUrl: newsUrl,
      source: source,
    });
  });

  return out;
}

/** ===== Build per-row prompt ===== */
function NS_buildPromptForRow_(basePrompt, row) {
  const now = new Date();
  const isoDate = now.toISOString().slice(0, 10);

  const companyName = row.companyName || "";
  const companyUrl = row.companyUrl || "";
  const newsUrl = row.newsUrl || "";
  const source = row.source || "";

  const summaryTemplate =
    "\n\n### Additional formatting requirements\n" +
    "- GPT Summary must follow the prompt’s PRODUCTION RULE: ALWAYS include OVERVIEW; include other sections ONLY when the article contains concrete info for that section.\n" +
    "- Do NOT output filler like “No direct quotes appear…”, “Not discussed…”, or any “No content…” statements.\n" +
    "- Keep the GPT Summary JSON-safe (no unescaped line breaks other than \\n) and non-promotional.\n";

  const noNewsNote =
    "\nIf you truly find no acceptable articles after filtering (or if the given News Story URL is invalid/unreachable), return a single JSON object in the exact 'NO NEWS CASE' format from the instructions, " +
    "including empty 'Special Values' and 'MMCrawl Updates', and set the GPT Summary to include the phrase: \"No news found after applying filters on " +
    isoDate +
    " <HH:MM:SS>\" where you substitute the current time.\n";

  const scenarioBlock =
    "\n\n### Input record\n" +
    "- Company Name (from sheet): " + companyName + "\n" +
    "- Company Website URL (from sheet): " + companyUrl + "\n" +
    "- News Story URL (from sheet; may be blank): " + newsUrl + "\n" +
    "- Source note: " + (source || "N/A") + "\n" +
    "\nRemember:\n" +
    "- If a News Story URL is provided, you MUST skip all searching and only analyze that URL, returning exactly one JSON object.\n" +
    "- If no News Story URL is provided, follow the full search protocol.\n" +
    "- Always populate 'Special Values' and 'MMCrawl Updates' as specified in the main prompt.\n" +
    "- Special Values are SUMMARY-LOCKED: only output a Special Value when that exact value also appears in GPT Summary.\n";

  return basePrompt + summaryTemplate + noNewsNote + scenarioBlock;
}

/** ===== OpenAI call for News_Search ===== */
function NS_callOpenAIForNews_(userPrompt) {
  const key = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
  if (!key) throw new Error('Missing OpenAI API key. Use "AI Integration → Set OpenAI API Key".');

  const model = (typeof AIA !== "undefined" && AIA && AIA.MODEL) ? AIA.MODEL : "gpt-4o";

  const payload = {
    model: model,
    temperature: 0.2,
    max_tokens: 6000,
    messages: [
      {
        role: "system",
        content:
          "You are an MBA-trained analyst with 5+ years researching U.S. precision metal and plastics manufacturers. " +
          "Follow the user instructions exactly. Return ONLY a strict JSON array (no markdown)."
      },
      { role: "user", content: userPrompt }
    ]
  };

  const resp = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", {
    method: "post",
    contentType: "application/json",
    muteHttpExceptions: true,
    headers: { Authorization: "Bearer " + key },
    payload: JSON.stringify(payload)
  });

  const code = resp.getResponseCode();
  const text = resp.getContentText();
  if (code < 200 || code >= 300) throw new Error("OpenAI HTTP " + code + ": " + text);

  const data = JSON.parse(text);
  const answer = data && data.choices && data.choices[0] && data.choices[0].message && data.choices[0].message.content;
  if (!answer) throw new Error("No content returned from OpenAI (News_Search).");
  return String(answer).trim();
}

/** ===== JSON helpers ===== */
function NS_extractJsonArray_(text) {
  if (!text) return [];
  let t = String(text).trim();

  const fence = t.match(/```json([\s\S]*?)```/i) || t.match(/```([\s\S]*?)```/i);
  if (fence) t = fence[1].trim();

  let obj = null;
  try {
    obj = JSON.parse(t);
  } catch (err) {
    const m = t.match(/(\{[\s\S]*\}|\[[\s\S]*\])/);
    if (m) {
      try { obj = JSON.parse(m[1]); } catch (_) {}
    }
  }

  if (!obj) return [];
  if (Array.isArray(obj)) return obj.filter((v) => v && typeof v === "object");
  if (typeof obj === "object") return [obj];
  return [];
}

/** ===== Enforce Special Values must appear in GPT Summary (summary-locked) ===== */
function NS_enforceSpecialValuesFromSummary_(obj) {
  const summary = String(obj["GPT Summary"] || "");
  const sv = (obj["Special Values"] && typeof obj["Special Values"] === "object")
    ? obj["Special Values"]
    : {};

  Object.keys(NS_DEFAULT_SPECIAL_VALUES).forEach((k) => {
    if (!sv.hasOwnProperty(k)) sv[k] = "";
  });

  if (!summary.trim()) {
    Object.keys(sv).forEach((k) => { sv[k] = ""; });
    obj["Special Values"] = sv;
    return obj;
  }

  const sumLower = summary.toLowerCase();

  Object.keys(sv).forEach((k) => {
    const v = (sv[k] === null || sv[k] === undefined) ? "" : String(sv[k]).trim();
    if (!v) { sv[k] = ""; return; }
    if (!sumLower.includes(v.toLowerCase())) sv[k] = "";
  });

  obj["Special Values"] = sv;
  return obj;
}

/** ===== Append JSON objects to "News Raw" ===== */
function NS_writeResultsToNewsRaw_(jsonArr) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName("News Raw");
  if (!sheet) throw new Error("Sheet 'News Raw' not found.");

  if (!jsonArr || !jsonArr.length) return;

  const lastRow = sheet.getLastRow();
  let row = (lastRow < 1 ? 2 : lastRow + 1);

  jsonArr.forEach((obj) => {
    let finalValue = "";
    const mv = obj["MMCrawl Updates"];

    if (mv === null || mv === undefined) finalValue = "";
    else if (typeof mv === "string") finalValue = mv;
    else {
      try { finalValue = JSON.stringify(mv); }
      catch (e) { finalValue = String(mv); }
    }

    const values = [
      obj["Company Name"] || "",
      obj["Company Website URL"] || "",
      obj["News Story URL"] || "",
      obj["Headline"] || "",
      obj["Publication Date"] || "",
      obj["Publisher or Source"] || "",
      obj["GPT Summary"] || "",
      obj["Source"] || "",
      finalValue
    ];

    sheet.getRange(row, 1, 1, values.length).setValues([values]);
    row++;
  });
}

/** ===== URL helpers ===== */

/** Actionable URL check: ONLY 404 and 429. */
function NS_checkUrlActionable404_429_(url) {
  const headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9"
  };

  try {
    const resp = UrlFetchApp.fetch(url, {
      method: "get",
      followRedirects: true,
      muteHttpExceptions: true,
      validateHttpsCertificates: true,
      headers
    });

    const code = resp.getResponseCode();
    if (code === 404 || code === 429) return { actionable: true, code: code, reason: "HTTP " + code };
    return { actionable: false, code: code, reason: "" };
  } catch (e) {
    return { actionable: false, code: 0, reason: String(e && e.message ? e.message : e) };
  }
}

/** Fetch HTML (best-effort). Returns { ok, code, html }. */
function NS_fetchHtml_(url) {
  try {
    const resp = UrlFetchApp.fetch(url, {
      method: "get",
      followRedirects: true,
      muteHttpExceptions: true,
      validateHttpsCertificates: true,
      headers: {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "en-US,en;q=0.9"
      }
    });
    const code = resp.getResponseCode();
    const html = String(resp.getContentText() || "");
    return { ok: true, code: code, html: html };
  } catch (e) {
    return { ok: false, code: 0, html: "" };
  }
}
/**
 * Soft-404 detector (PRE-VERSION style: simple, reliable string checks).
 * Returns true when the HTML content indicates "not found" even if HTTP 200.
 */
function NS_isSoft404_(url, html) {
  if (!html) return false;
  const h = String(html).toLowerCase();

  const needles = [
    "sorry, we couldn't find the page",
    "couldn't find the page you're looking for",
    "could not find the page",
    "page not found",
    "the page you requested cannot be found",
    "we can't find the page",
    "sorry, we can’t find that page.",
    "sorry, we can't find that page",
    "sorry, we can’t find that page",
    "we can’t find that page",
    "we can't find that page"
  ];

  let hits = 0;
  needles.forEach((n) => { if (h.indexOf(n) >= 0) hits++; });

  if (h.indexOf("couldn't find the page") >= 0) return true;
  if (h.indexOf("page not found") >= 0) return true;
  if (hits >= 2) return true;

  return false;
}

/**
 * Extract candidate URLs from a soft-404 page (e.g., "Closest matches").
 * - Supports absolute and relative URLs.
 * - Keeps only article-like URLs.
 * - De-dupes and preserves page order.
 */
function NS_extractCandidateUrlsFromHtml_(html, baseUrl) {
  const out = [];
  if (!html) return out;

  const base = String(baseUrl || "").trim();

  function toAbsolute_(href) {
    let u = String(href || "").trim();
    if (!u) return "";
    u = u.replace(/&amp;/g, "&");

    if (/^https?:\/\//i.test(u)) return u;
    if (u.indexOf("//") === 0) return "https:" + u;

    // relative -> origin + path
    const m = base.match(/^(https?:\/\/[^\/]+)(\/.*)?$/i);
    const origin = m ? m[1] : "";
    if (!origin) return "";
    if (u.charAt(0) !== "/") u = "/" + u;
    return origin + u;
  }

  // Pull hrefs
  const re = /href\s*=\s*["']([^"']+)["']/gi;
  let m;
  while ((m = re.exec(html)) !== null) {
    const abs = toAbsolute_(m[1]);
    if (!abs) continue;

    const ul = abs.toLowerCase();

    // Article-like patterns (tune as needed)
    const isArticleLike =
      ul.indexOf("/articles/") >= 0 ||
      ul.indexOf("/article/") >= 0 ||
      ul.indexOf("/news/") >= 0;

    if (!isArticleLike) continue;

    // Exclusions (avoid directories / taxonomies)
    if (ul.indexOf("/suppliers/") >= 0) continue;
    if (ul.indexOf("/topics/") >= 0) continue;
    if (ul.indexOf("/topic/") >= 0) continue;
    if (ul.indexOf("/tag/") >= 0) continue;
    if (ul.indexOf("/category/") >= 0) continue;
    if (ul.indexOf("/search") >= 0) continue;

    out.push(abs);
  }

  // Dedupe, preserve order
  const uniq = [];
  const seen = {};
  out.forEach((u) => {
    const key = String(u).toLowerCase();
    if (seen[key]) return;
    seen[key] = true;
    uniq.push(u);
  });

  return uniq;
}

/**
 * Rebuild News Source sheet from output rows (A:E).
 * Expected columns:
 * A: Company URL
 * B: Company Name
 * C: News URL
 * D: Source
 * E: Archive_Flag
 */
function NS_rebuildNewsSourceSheetFromRows_(rows) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName("News Source");
  if (!sh) throw new Error('Sheet "News Source" not found.');

  const out = [];
  (rows || []).forEach((r) => {
    if (!Array.isArray(r)) return;
    const a = String(r[0] || "").trim();
    const b = String(r[1] || "").trim();
    const c = String(r[2] || "").trim();
    const d = String(r[3] || "").trim();
    const e = String(r[4] || "").trim();

    if (!a && !b && !c && !d && !e) return;

    // Only allow "Yes" or blank
    const flag = (e === "Yes") ? "Yes" : "";
    out.push([a, b, c, d, flag]);
  });

  // Clear old data area (row 2 down)
  const last = sh.getLastRow();
  if (last >= 2) {
    sh.getRange(2, 1, last - 1, sh.getMaxColumns()).clearContent().clearNote();
  }

  if (out.length) {
    sh.getRange(2, 1, out.length, 5).setValues(out);
  }
}

/**
 * ✅ CHECK NEWS SOURCE URLS (LOCAL REBUILD + SAVE RESULT)
 *
 * Rules implemented exactly as you described:
 * - HTTP 200 with real content -> keep
 * - HTTP 404 / 429 -> remove
 * - HTTP 200 soft-404:
 *      - if closest matches exist -> replace with matches, set Archive_Flag="Yes"
 *      - if no matches -> remove
 * - blocked / fetch failed -> keep as-is
 *
 * Also saves the FINAL rebuilt rows list into AI Integration -> Result for Prompt ID "News_Url",
 * then rebuilds News Source from that JSON.
 */
function NS_checkNewsSourceUrls() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();

  const aiSheet = ss.getSheetByName("AI Integration");
  if (!aiSheet) {
    ui.alert('Sheet "AI Integration" not found.');
    return;
  }

  const promptRow = NS_findPromptRow_("News_Url");
  if (!promptRow) {
    ui.alert('Prompt ID "News_Url" not found in AI Integration!A:A.');
    return;
  }

  const sh = ss.getSheetByName("News Source");
  if (!sh) {
    ui.alert('Sheet "News Source" not found.');
    return;
  }

  const lastRow = sh.getLastRow();
  if (lastRow < 2) {
    ui.alert('No data rows found in sheet "News Source".');
    return;
  }

  // News Source columns:
  // A: Company URL, B: Company Name, C: News URL, D: Source, E: Archive_Flag
  const vals = sh.getRange(2, 1, lastRow - 1, 5).getDisplayValues();

  const inputRows = vals
    .map(r => [
      String(r[0] || "").trim(),
      String(r[1] || "").trim(),
      String(r[2] || "").trim(),
      String(r[3] || "").trim(),
      String(r[4] || "").trim()
    ])
    .filter(r => r.some(x => String(x || "").trim() !== ""));

  if (!inputRows.length) {
    ui.alert("News Source is empty (no rows to check).");
    return;
  }

  // Output + counters
  let checked_rows = 0;
  let kept = 0;
  let removed_404_429 = 0;
  let replaced_soft404 = 0;
  let removed_soft404_no_matches = 0;
  let added_matches = 0;

  const outRows = [];
  const seenNewsUrls = {}; // dedupe by news URL lower

  // Fetch headers
  const headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9"
  };

  function addRow_(companyUrl, companyName, newsUrl, source, flag) {
    const nu = String(newsUrl || "").trim();
    if (!nu) return;
    const key = nu.toLowerCase();
    if (seenNewsUrls[key]) return;
    seenNewsUrls[key] = true;
    outRows.push([
      String(companyUrl || "").trim(),
      String(companyName || "").trim(),
      nu,
      String(source || "").trim(),
      (flag === "Yes") ? "Yes" : ""
    ]);
  }

  for (let i = 0; i < inputRows.length; i++) {
    const row = inputRows[i];
    const companyUrl = row[0];
    const companyName = row[1];
    const newsUrl = row[2];
    const source = row[3];
    // const existingFlag = row[4]; // we will recompute

    if (!newsUrl) {
      // If no News URL, keep row as-is (rare, but safe)
      addRow_(companyUrl, companyName, newsUrl, source, "");
      continue;
    }

    checked_rows++;
    ss.toast("Checking " + checked_rows + "/" + inputRows.length + " — " + newsUrl, "News", 5);

    let resp, code, html;
    try {
      resp = UrlFetchApp.fetch(newsUrl, {
        method: "get",
        followRedirects: true,
        muteHttpExceptions: true,
        validateHttpsCertificates: true,
        headers
      });
      code = resp.getResponseCode();
      html = String(resp.getContentText() || "");
    } catch (e) {
      // BLOCKED/UNKNOWN -> keep original
      kept++;
      addRow_(companyUrl, companyName, newsUrl, source, "");
      continue;
    }

    // HARD REMOVE
    if (code === 404 || code === 429) {
      removed_404_429++;
      continue;
    }

    // If not a normal HTML success, treat as unknown -> keep
    // (e.g., 403, 406, 500, etc.)
    if (!(code >= 200 && code < 400)) {
      kept++;
      addRow_(companyUrl, companyName, newsUrl, source, "");
      continue;
    }

    // SOFT-404 detection (HTTP 200 but "can't find" content)
    if (NS_isSoft404_(newsUrl, html)) {
      const candidates = NS_extractCandidateUrlsFromHtml_(html, newsUrl).slice(0, 6);

      if (candidates.length) {
        replaced_soft404++;
        candidates.forEach((u) => {
          add_matches:
          addRow_(companyUrl, companyName, u, source, "Yes");
        });
        added_matches += candidates.length;
      } else {
        // Soft-404 but no matches -> remove
        removed_soft404_no_matches++;
      }
      continue;
    }

    // Otherwise keep (assume valid article)
    kept++;
    addRow_(companyUrl, companyName, newsUrl, source, "");
  }

  const resultObj = {
    checked_rows: checked_rows,
    kept: kept,
    removed_404_429: removed_404_429,
    replaced_soft404: replaced_soft404,
    removed_soft404_no_matches: removed_soft404_no_matches,
    added_matches: added_matches,
    rows: outRows,
    notes: "Local rebuild: kept valid pages, removed 404/429, replaced soft-404 pages with closest-match URLs (Archive_Flag=Yes), removed soft-404 with no matches, kept blocked/unknown unchanged."
  };

  // Save JSON into AI Integration Result/Date for News_Url
  aiSheet.getRange(promptRow, 3).setValue(JSON.stringify(resultObj, null, 2));
  aiSheet.getRange(promptRow, 4).setValue(new Date());

  // Rebuild News Source from result rows
  NS_rebuildNewsSourceSheetFromRows_(outRows);

  ui.alert(
    "Check News Source URLs complete",
    "Checked: " + checked_rows +
      "\nKept: " + kept +
      "\nRemoved (404/429): " + removed_404_429 +
      "\nReplaced (soft-404): " + replaced_soft404 +
      "\nRemoved (soft-404, no matches): " + removed_soft404_no_matches +
      "\nAdded matches: " + added_matches +
      "\n\nNews Source was rebuilt and AI Integration → Result (News_Url) was updated.",
    ui.ButtonSet.OK
  );

  ss.toast("News Source rebuilt. AI Integration (News_Url) Result updated.", "News", 8);
}


/** ===== OpenAI call for News_Url (URL validation + rebuild) ===== */
function NS_callOpenAIForUrlRepair_(userPrompt) {
  const key = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
  if (!key) throw new Error('Missing OpenAI API key. Use "AI Integration → Set OpenAI API Key".');

  const model = (typeof AIA !== "undefined" && AIA && AIA.MODEL) ? AIA.MODEL : "gpt-4o";

  const payload = {
    model: model,
    temperature: 0.0,
    max_tokens: 6000,
    messages: [
      {
        role: "system",
        content:
          "You are a News Source URL validation + repair agent. " +
          "Follow the provided rules exactly. Return ONLY strict JSON (no markdown)."
      },
      { role: "user", content: userPrompt }
    ]
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
  if (code < 200 || code >= 300) throw new Error("OpenAI HTTP " + code + ": " + text);

  const data = JSON.parse(text);
  const answer = data && data.choices && data.choices[0] && data.choices[0].message && data.choices[0].message.content;
  if (!answer) throw new Error("No content returned from OpenAI (News_Url).");

  return String(answer).trim();
}


