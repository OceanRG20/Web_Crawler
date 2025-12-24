/*************************************************
 * News_Agent.gs — News menu + GPT news crawler + GPT URL repair
 *
 * Sheets used:
 *   - "AI Integration"   (prompts + result log)
 *   - "News Source"      (input/output: Company URL, Name, News URL, Source)
 *   - "News Raw"         (output: normalized news rows)
 *
 * Prompt IDs in AI Integration:
 *   - "News_Search"  (news crawling / summarization)
 *   - "News_Url"     (URL validation + repair; returns FINAL rebuilt rows list)
 **************************************************/

/** ===== Menu hook (called from AI_Agent.onOpen) ===== */
function onOpen_News(ui) {
  ui = ui || SpreadsheetApp.getUi();
  ui.createMenu("News")
    .addItem("▶ News Searching", "NS_runNewsSearch")
    .addItem("✅ Check News Source URLs", "NS_checkNewsSourceUrls") // GPT-driven rebuild
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

/** ===== Main runner ===== */
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
    "Processing " + newsRows.length + " candidate news URLs from News Source…",
    ui.ButtonSet.OK
  );

  const allArticles = [];
  let idx = 0;

  newsRows.forEach((row) => {
    idx++;
    ss.toast(
      "News Search " + idx + "/" + newsRows.length + " — " + (row.newsUrl || row.companyUrl),
      "News",
      5
    );

    /**
     * POLICY (keep as you had):
     * - NEVER SKIP a row.
     * - 404/429: try Wayback; if none, keep URL in sheet but run SEARCH MODE for this run.
     * - Soft-404 (HTML says not found): remove URL from sheet, append candidate URLs, run SEARCH MODE for this run.
     * - All other statuses (including 403/406 bot blocks): do NOT remove; proceed.
     */
    if (row.newsUrl) {
      // 1) Soft-404 check (content-based)
      const fx = NS_fetchHtml_(row.newsUrl);
      if (fx.ok && NS_isSoft404_(row.newsUrl, fx.html)) {
        const candidates = NS_extractCandidateUrlsFromHtml_(fx.html).slice(0, 6); // cap
        NS_applySoft404RepairToSheet_(
          row.rowIndex,
          row.companyUrl,
          row.companyName,
          row.newsUrl,
          candidates,
          row.source // preserve original source
        );

        // Run search mode for THIS run
        row.newsUrl = "";
      } else {
        // 2) Only actionable statuses 404/429
        const chk = NS_checkUrlActionable404_429_(row.newsUrl);
        if (chk.actionable) {
          const wb = NS_tryWayback_(row.newsUrl);

          if (wb.ok && wb.url) {
            const old = row.newsUrl;
            row.newsUrl = wb.url;

            // write back to News Source column C
            try {
              const ns = ss.getSheetByName("News Source");
              if (ns && row.rowIndex) {
                ns.getRange(row.rowIndex, 3).setValue(wb.url);
                ns.getRange(row.rowIndex, 3).setNote("Auto-repaired via Wayback from: " + old);
              }
            } catch (_) {}

          } else {
            // No Wayback: keep sheet URL but run search mode THIS run
            const original = row.newsUrl;
            row.newsUrl = "";

            try {
              const ns = ss.getSheetByName("News Source");
              if (ns && row.rowIndex) {
                ns.getRange(row.rowIndex, 3).setNote(
                  "404/429 detected; no Wayback snapshot. Kept URL in sheet, ran SEARCH MODE for this run.\n" +
                  "Original: " + original + "\nChecked: " + new Date().toISOString()
                );
              }
            } catch (_) {}
          }
        }
      }
    }

    try {
      const fullPrompt = NS_buildPromptForRow_(basePrompt, row);
      const rawAnswer = NS_callOpenAIForNews_(fullPrompt);
      const articles = NS_extractJsonArray_(rawAnswer);

      const enriched = articles.map((obj) => {
        let copy = Object.assign({}, obj);

        if (!copy["Source"]) copy["Source"] = row.source || "";
        if (!copy["Company Name"]) copy["Company Name"] = row.companyName || "";
        if (!copy["Company Website URL"]) copy["Company Website URL"] = row.companyUrl || "";

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
      const now = new Date();
      const isoDate = now.toISOString().slice(0, 10);
      const timeStr = now.toTimeString().slice(0, 8);

      allArticles.push({
        "Company Name": row.companyName || "",
        "Company Website URL": row.companyUrl || "",
        "News Story URL": row.newsUrl || "",
        "Headline": "Error fetching news",
        "Publication Date": "",
        "Publisher or Source": "",
        "GPT Summary":
          "Error while running News_Search for this URL at " +
          isoDate + " " + timeStr +
          ": " + String(err),
        "Confidence Score": "",
        "Source": row.source || "",
        "Special Values": JSON.parse(JSON.stringify(NS_DEFAULT_SPECIAL_VALUES)),
        "MMCrawl Updates": ""
      });
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
/**
 * Expected headers in News Source:
 *   A: Company URL
 *   B: Company Name
 *   C: News URL
 *   D: Source
 */
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
  if (!key) {
    throw new Error('Missing OpenAI API key. Use "AI Integration → Set OpenAI API Key".');
  }

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
  if (code < 200 || code >= 300) throw new Error("OpenAI HTTP " + code + ": " + text);

  const data = JSON.parse(text);
  const answer = data?.choices?.[0]?.message?.content;
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

function NS_extractJsonObject_(text) {
  if (!text) return null;
  let t = String(text).trim();

  const fence = t.match(/```json([\s\S]*?)```/i) || t.match(/```([\s\S]*?)```/i);
  if (fence) t = fence[1].trim();

  // prefer object
  let obj = null;
  try {
    obj = JSON.parse(t);
  } catch (err) {
    const m = t.match(/(\{[\s\S]*\})/);
    if (m) {
      try { obj = JSON.parse(m[1]); } catch (_) {}
    }
  }
  return (obj && typeof obj === "object" && !Array.isArray(obj)) ? obj : null;
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

/** ===== URL + Wayback + Soft-404 helpers ===== */

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

/** Soft-404 detector: content indicates "not found" even if HTTP 200. */
function NS_isSoft404_(url, html) {
  if (!html) return false;
  const h = html.toLowerCase();

  const needles = [
    "sorry, we couldn't find the page",
    "sorry, we couldn’t find the page",
    "couldn't find the page you're looking for",
    "could not find the page",
    "page not found",
    "the page you requested cannot be found",
    "we can't find the page",
    "we can’t find the page",
    "sorry, we couldn't find that page",
    "sorry, we can’t find that page"
  ];

  let hits = 0;
  needles.forEach((n) => { if (h.indexOf(n) >= 0) hits++; });

  // conservative threshold
  return hits >= 1;
}

/**
 * Extract candidate URLs from a soft-404 page (e.g., "Closest matches").
 * Returns a de-duped list of absolute URLs.
 */
function NS_extractCandidateUrlsFromHtml_(html) {
  const out = [];
  if (!html) return out;

  const re = /href\s*=\s*["']([^"']+)["']/gi;
  let m;
  while ((m = re.exec(html)) !== null) {
    let u = String(m[1] || "").trim();
    if (!u) continue;

    u = u.replace(/&amp;/g, "&");
    if (!/^https?:\/\//i.test(u)) continue;

    // Keep only likely article links (expand safely if needed)
    if (
      u.indexOf("ptonline.com/articles/") >= 0 ||
      u.indexOf("plasticstoday.com/") >= 0 ||
      u.indexOf("moldmakingtechnology.com/") >= 0
    ) {
      out.push(u);
    }
  }

  const uniq = [];
  const seen = {};
  out.forEach((u) => {
    const key = u.toLowerCase();
    if (seen[key]) return;
    seen[key] = true;
    uniq.push(u);
  });

  return uniq;
}

/**
 * Soft-404 repair in sheet:
 * - Clear the bad News URL in place
 * - Append candidate URLs as new rows with SAME company + SAME source
 */
function NS_applySoft404RepairToSheet_(rowIndex, companyUrl, companyName, badUrl, candidates, originalSource) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName("News Source");
  if (!sh) return;

  // Clear bad URL
  sh.getRange(rowIndex, 3).clearContent();
  sh.getRange(rowIndex, 3).setNote(
    "Soft-404 detected; URL removed.\nOriginal: " + badUrl +
    "\nChecked: " + new Date().toISOString()
  );

  // Keep Source as-is (do not overwrite with a label)
  // If you want to mark, prefer Note not value change.

  if (!candidates || !candidates.length) return;

  // Dedupe against existing News URLs
  const last = sh.getLastRow();
  const existing = (last >= 2)
    ? sh.getRange(2, 3, last - 1, 1).getDisplayValues().flat().map(s => String(s || "").trim().toLowerCase())
    : [];

  const toAdd = [];
  candidates.slice(0, 6).forEach((u) => {
    const key = String(u).trim().toLowerCase();
    if (!key) return;
    if (existing.indexOf(key) >= 0) return;
    toAdd.push([companyUrl || "", companyName || "", u, originalSource || ""]);
  });

  if (!toAdd.length) return;
  sh.getRange(sh.getLastRow() + 1, 1, toAdd.length, 4).setValues(toAdd);
}

/** Wayback repair helper (used by News_Search flow) */
function NS_tryWayback_(url) {
  try {
    const api = "https://archive.org/wayback/available?url=" + encodeURIComponent(url);
    const resp = UrlFetchApp.fetch(api, { method: "get", muteHttpExceptions: true });
    const code = resp.getResponseCode();
    if (code < 200 || code >= 300) return { ok: false, url: "" };

    const json = JSON.parse(resp.getContentText() || "{}");
    const closest = json && json.archived_snapshots && json.archived_snapshots.closest;
    if (closest && closest.available && closest.url) {
      return { ok: true, url: closest.url };
    }
    return { ok: false, url: "" };
  } catch (e) {
    return { ok: false, url: "" };
  }
}

/** =========================================================
 *  ✅ CHECK NEWS SOURCE URLS (GPT-DRIVEN REBUILD)
 *
 *  Flow:
 *   1) Read all rows from News Source (A:D).
 *   2) Load prompt "News_Url" from AI Integration.
 *   3) Call OpenAI with: prompt + JSON input rows.
 *   4) Save returned JSON into AI Integration Result/Date for News_Url.
 *   5) Rebuild News Source sheet rows from returned "rows".
 * ========================================================= */
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

  const basePrompt = String(aiSheet.getRange(promptRow, 2).getValue() || "").trim();
  if (!basePrompt) {
    ui.alert('No prompt content in "AI Integration" for Prompt ID "News_Url".');
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

  // Read ALL existing rows
  const vals = sh.getRange(2, 1, lastRow - 1, 4).getDisplayValues();
  const inputRows = vals
    .map(r => [String(r[0] || "").trim(), String(r[1] || "").trim(), String(r[2] || "").trim(), String(r[3] || "").trim()])
    .filter(r => r.some(x => String(x || "").trim() !== "")); // remove fully empty

  if (!inputRows.length) {
    ui.alert("News Source is empty (no rows to check).");
    return;
  }

  ss.toast("Sending " + inputRows.length + " News Source row(s) to News_Url…", "News", 8);

  // Build user prompt: base + explicit JSON payload
  const userPrompt =
    basePrompt +
    "\n\n====================\nINPUT_ROWS_JSON\n====================\n" +
    JSON.stringify({ rows: inputRows }, null, 2) +
    "\n";

  // Call OpenAI
  let raw = "";
  try {
    raw = NS_callOpenAIForUrlRepair_(userPrompt);
  } catch (e) {
    ui.alert("News_Url call failed:\n\n" + String(e));
    return;
  }

  // Parse JSON object
  const obj = NS_extractJsonObject_(raw);
  if (!obj) {
    ui.alert("News_Url returned invalid JSON. See AI Integration Result for raw output.");
    aiSheet.getRange(promptRow, 3).setValue(String(raw || ""));
    aiSheet.getRange(promptRow, 4).setValue(new Date());
    return;
  }

  // Save result JSON to AI Integration
  aiSheet.getRange(promptRow, 3).setValue(JSON.stringify(obj, null, 2));
  aiSheet.getRange(promptRow, 4).setValue(new Date());

  // Validate rows output
  const outRows = Array.isArray(obj.rows) ? obj.rows : null;
  if (!outRows) {
    ui.alert('News_Url JSON missing "rows" array. No sheet update performed.');
    return;
  }

  // Rebuild News Source from output rows
  NS_rebuildNewsSourceSheetFromRows_(outRows);

  // Summary (best-effort; fields depend on your prompt)
  const checked = (typeof obj.checked_rows === "number") ? obj.checked_rows : inputRows.length;
  const kept = (typeof obj.kept === "number") ? obj.kept : "";
  const removed = (typeof obj.removed_404_429 === "number") ? obj.removed_404_429 : "";
  const replaced = (typeof obj.replaced_soft404 === "number") ? obj.replaced_soft404 : "";
  const added = (typeof obj.added_matches === "number") ? obj.added_matches : "";

  ui.alert(
    "Check News Source URLs complete",
    "Checked: " + checked +
      (kept !== "" ? ("\nKept: " + kept) : "") +
      (removed !== "" ? ("\nRemoved (404/429): " + removed) : "") +
      (replaced !== "" ? ("\nReplaced (soft404): " + replaced) : "") +
      (added !== "" ? ("\nAdded matches: " + added) : "") +
      "\n\nNews Source was rebuilt from AI Integration → Result (News_Url).",
    ui.ButtonSet.OK
  );

  ss.toast("News Source rebuilt from News_Url result.", "News", 8);
}

/** OpenAI call for News_Url (URL validation + rebuild) */
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
          "You are a URL validation and repair agent. Follow the provided rules exactly. " +
          "Return ONLY strict JSON (no markdown)."
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
  const answer = data?.choices?.[0]?.message?.content;
  if (!answer) throw new Error("No content returned from OpenAI (News_Url).");
  return String(answer).trim();
}

/** Rebuild News Source sheet from output rows (A:D). */
function NS_rebuildNewsSourceSheetFromRows_(rows) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName("News Source");
  if (!sh) throw new Error('Sheet "News Source" not found.');

  // Normalize rows to 4 columns strings
  const out = [];
  (rows || []).forEach((r) => {
    if (!Array.isArray(r)) return;
    const a = String(r[0] || "").trim();
    const b = String(r[1] || "").trim();
    const c = String(r[2] || "").trim();
    const d = String(r[3] || "").trim();
    if (!a && !b && !c && !d) return;
    out.push([a, b, c, d]);
  });

  // Clear old data area (row 2 down)
  const last = sh.getLastRow();
  if (last >= 2) {
    sh.getRange(2, 1, last - 1, sh.getMaxColumns()).clearContent().clearNote();
  }

  // Write new rows
  if (out.length) {
    sh.getRange(2, 1, out.length, 4).setValues(out);
  }
}
