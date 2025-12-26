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
 *
 * NOTE (per your request):
 *   - Archive_Flag is REMOVED completely (no column E, no flag logic).
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

/** =========================================================
 *  ▶ NEWS SEARCH
 *  - No skipping: every News Source row produces output (at least 1 object).
 *  - If URL is soft-404 w/ “Closest matches”, it expands and processes those matches.
 *  - If URL is 404/429/blocked/unknown, it writes a fallback “No news found” object.
 * ========================================================= */
function NS_runNewsSearch() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();

  const aiSheet = ss.getSheetByName("AI Integration");
  if (!aiSheet) { ui.alert('Sheet "AI Integration" not found.'); return; }

  const promptRow = NS_findPromptRow_("News_Search");
  if (!promptRow) { ui.alert('Prompt ID "News_Search" not found in AI Integration!A:A.'); return; }

  const basePrompt = String(aiSheet.getRange(promptRow, 2).getValue() || "").trim();
  if (!basePrompt) { ui.alert('No prompt content in "AI Integration" for Prompt ID "News_Search".'); return; }

  const newsRows = NS_getRowsFromNewsSource_(); // A:D
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

  function makeFallback_(row, reason) {
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
        "No news found after applying filters on " + isoDate + " " + timeStr +
        (reason ? (". Reason: " + String(reason)) : "."),
      "Confidence Score": "",
      "Source": row.source || "",
      "Special Values": JSON.parse(JSON.stringify(NS_DEFAULT_SPECIAL_VALUES)),
      "MMCrawl Updates": ""
    };
  }

  function enrich_(obj, row) {
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
  }

  newsRows.forEach((row) => {
    idx++;

    ss.toast(
      "News Search " + idx + "/" + newsRows.length + " — " + (row.newsUrl || row.companyUrl),
      "News",
      5
    );

    try {
      const url = String(row.newsUrl || "").trim();

      // If there is a URL, pre-check for soft-404 w/ closest matches.
      if (url) {
        const fx = NS_fetchHtml_(url);

        // 404/429: do NOT skip; output fallback for this row
        if (fx.ok && (fx.code === 404 || fx.code === 429)) {
          allArticles.push(enrich_(makeFallback_(row, "URL returned HTTP " + fx.code + "."), row));
          return;
        }

        // Soft-404: expand closest matches (if any)
        if (fx.ok && fx.code >= 200 && fx.code < 400 && NS_isSoft404_(url, fx.html)) {
          const candidates = NS_extractCandidateUrlsFromHtml_(fx.html, url).slice(0, 6);

          if (candidates.length) {
            let produced = 0;

            candidates.forEach((u) => {
              const tempRow = Object.assign({}, row, { newsUrl: u });
              try {
                const fullPrompt = NS_buildPromptForRow_(basePrompt, tempRow);
                const rawAnswer = NS_callOpenAIForNews_(fullPrompt);
                const articles = NS_extractJsonArray_(rawAnswer);

                if (articles && articles.length) {
                  articles.forEach((a) => {
                    allArticles.push(enrich_(a, tempRow));
                    produced++;
                  });
                }
              } catch (e2) {
                // continue; fallback if produced=0 overall
              }
              Utilities.sleep(150);
            });

            if (produced === 0) {
              allArticles.push(enrich_(makeFallback_(row, "Soft-404 with closest matches, but model returned no usable JSON."), row));
            }
            return;
          }

          // Soft-404 but no matches
          allArticles.push(enrich_(makeFallback_(row, "Soft-404 detected; no closest-match URLs found."), row));
          return;
        }

        // Blocked/unknown/fetch failed
        if (!fx.ok) {
          allArticles.push(enrich_(makeFallback_(row, "Fetch failed/blocked/unknown."), row));
          return;
        }
      }

      // Normal path: call News_Search once for the row as-is.
      const fullPrompt = NS_buildPromptForRow_(basePrompt, row);
      const rawAnswer = NS_callOpenAIForNews_(fullPrompt);
      const articles = NS_extractJsonArray_(rawAnswer);

      if (articles && articles.length) {
        articles.forEach((a) => allArticles.push(enrich_(a, row)));
      } else {
        allArticles.push(enrich_(makeFallback_(row, "Model returned empty/invalid JSON array for this input."), row));
      }

    } catch (err) {
      allArticles.push(enrich_(makeFallback_(row, "Error while running News_Search: " + String(err)), row));
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

/** ===== Read rows from "News Source" (A:D) =====
 * Columns:
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
      companyUrl,
      companyName,
      newsUrl,
      source
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

/** =========================================================
 *  ✅ CHECK NEWS SOURCE URLS (LOCAL REBUILD)
 *  - Removes hard 404/429 rows.
 *  - Removes soft-404 rows; if closest matches exist, replaces with matches.
 *  - Keeps blocked/unknown rows unchanged.
 *  - Rebuilds "News Source" with ONLY A:D (no Archive_Flag).
 *
 *  IMPORTANT FIX:
 *  - Soft-404 detection is strengthened (handles curly apostrophes / HTML entities),
 *    so “Sorry, we couldn’t find…” pages are reliably removed/replaced.
 * ========================================================= */
function NS_checkNewsSourceUrls() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();

  const sh = ss.getSheetByName("News Source");
  if (!sh) { ui.alert('Sheet "News Source" not found.'); return; }

  const lastRow = sh.getLastRow();
  if (lastRow < 2) { ui.alert('No data rows found in sheet "News Source".'); return; }

  // Read A:D (Company URL, Company Name, News URL, Source)
  const vals = sh.getRange(2, 1, lastRow - 1, 4).getDisplayValues();

  const inputRows = vals
    .map(r => [
      String(r[0] || "").trim(),
      String(r[1] || "").trim(),
      String(r[2] || "").trim(),
      String(r[3] || "").trim()
    ])
    .filter(r => r.some(x => String(x || "").trim() !== ""));

  if (!inputRows.length) {
    ui.alert("News Source is empty (no rows to check).");
    return;
  }

  let checked_rows = 0;
  let kept = 0;
  let removed_404_429 = 0;
  let replaced_soft404 = 0;
  let removed_soft404_no_matches = 0;
  let added_matches = 0;

  const outRows = [];
  const seenNewsUrls = {}; // dedupe by news URL lower

  const headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9"
  };

  function addRow_(companyUrl, companyName, newsUrl, source) {
    const nu = String(newsUrl || "").trim();
    if (!nu) return;
    const key = nu.toLowerCase();
    if (seenNewsUrls[key]) return;
    seenNewsUrls[key] = true;
    outRows.push([
      String(companyUrl || "").trim(),
      String(companyName || "").trim(),
      nu,
      String(source || "").trim()
    ]);
  }

  for (let i = 0; i < inputRows.length; i++) {
    const row = inputRows[i];
    const companyUrl = row[0];
    const companyName = row[1];
    const newsUrl = row[2];
    const source = row[3];

    if (!newsUrl) continue;

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
      // BLOCKED/UNKNOWN -> keep original row unchanged
      kept++;
      addRow_(companyUrl, companyName, newsUrl, source);
      continue;
    }

    // HARD REMOVE
    if (code === 404 || code === 429) {
      removed_404_429++;
      continue;
    }

    // Non-2xx/3xx -> keep as unknown
    if (!(code >= 200 && code < 400)) {
      kept++;
      addRow_(companyUrl, companyName, newsUrl, source);
      continue;
    }

    // SOFT-404 -> remove original; replace with closest matches if available
    if (NS_isSoft404_(newsUrl, html)) {
      const candidates = NS_extractCandidateUrlsFromHtml_(html, newsUrl).slice(0, 6);

      if (candidates.length) {
        replaced_soft404++;
        candidates.forEach((u) => addRow_(companyUrl, companyName, u, source));
        added_matches += candidates.length;
      } else {
        removed_soft404_no_matches++;
      }
      continue;
    }

    // Otherwise keep
    kept++;
    addRow_(companyUrl, companyName, newsUrl, source);
  }

  // Save result JSON into AI Integration (optional; keeps your existing logging pattern)
  const aiSheet = ss.getSheetByName("AI Integration");
  const promptRow = NS_findPromptRow_("News_Url");
  if (aiSheet && promptRow) {
    const resultObj = {
      checked_rows: checked_rows,
      kept: kept,
      removed_404_429: removed_404_429,
      replaced_soft404: replaced_soft404,
      removed_soft404_no_matches: removed_soft404_no_matches,
      added_matches: added_matches,
      rows: outRows,
      notes: "Local rebuild: kept valid/unknown pages, removed 404/429, replaced soft-404 pages with closest-match URLs, removed soft-404 with no matches."
    };
    aiSheet.getRange(promptRow, 3).setValue(JSON.stringify(resultObj, null, 2));
    aiSheet.getRange(promptRow, 4).setValue(new Date());
  }

  // Rebuild News Source from output rows (A:D)
  NS_rebuildNewsSourceSheetFromRows_(outRows);

  ui.alert(
    "Check News Source URLs complete",
    "Checked: " + checked_rows +
      "\nKept: " + kept +
      "\nRemoved (404/429): " + removed_404_429 +
      "\nReplaced (soft-404): " + replaced_soft404 +
      "\nRemoved (soft-404, no matches): " + removed_soft404_no_matches +
      "\nAdded matches: " + added_matches +
      "\n\nNews Source was rebuilt (A:D).",
    ui.ButtonSet.OK
  );

  ss.toast("News Source rebuilt (A:D).", "News", 8);
}

/** ===== URL helpers ===== */

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
 * Soft-404 detector (ROBUST):
 * - Handles curly apostrophes and common HTML entities.
 * - Designed to catch PT / MMT / Gardner “Sorry, we couldn’t find…” pages reliably.
 */
function NS_isSoft404_(url, html) {
  if (!html) return false;

  let h = String(html).toLowerCase();

  // Normalize apostrophes (HTML entities + unicode)
  h = h
    .replace(/&#39;|&apos;|&#x27;/g, "'")
    .replace(/&#8217;|&#x2019;|&rsquo;|&lsquo;/g, "'")
    .replace(/\u2019/g, "'");

  // Normalize whitespace
  h = h.replace(/\s+/g, " ");

  const needles = [
    "sorry, we couldn't find the page you're looking for",
    "sorry, we couldn't find the page you are looking for",
    "sorry, we couldn't find the page",
    "sorry, we couldn't find that page",
    "we couldn't find the page you're looking for",
    "we couldn't find that page",
    "we can't find that page",
    "we can't find the page",
    "page not found",
    "the page you requested cannot be found",
    "the page you're looking for may no longer exist",
    "the page you’re looking for may no longer exist",
    "sorry, we can't find that page",
    "sorry, we can't find the page"
  ];

  for (let i = 0; i < needles.length; i++) {
    if (h.indexOf(needles[i]) >= 0) return true;
  }

  // Extra catch: many not-found pages include “closest matches”
  if (h.indexOf("closest matches") >= 0) {
    if (h.indexOf("sorry") >= 0 || h.indexOf("couldn't find") >= 0 || h.indexOf("can't find") >= 0) return true;
  }

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

    const m = base.match(/^(https?:\/\/[^\/]+)(\/.*)?$/i);
    const origin = m ? m[1] : "";
    if (!origin) return "";
    if (u.charAt(0) !== "/") u = "/" + u;
    return origin + u;
  }

  const re = /href\s*=\s*["']([^"']+)["']/gi;
  let m;
  while ((m = re.exec(html)) !== null) {
    const abs = toAbsolute_(m[1]);
    if (!abs) continue;

    const ul = abs.toLowerCase();

    // Article-like patterns
    const isArticleLike =
      ul.indexOf("/articles/") >= 0 ||
      ul.indexOf("/article/") >= 0 ||
      ul.indexOf("/news/") >= 0 ||
      ul.indexOf("/viewpoint/") >= 0;

    if (!isArticleLike) continue;

    // Exclusions
    if (ul.indexOf("/suppliers/") >= 0) continue;
    if (ul.indexOf("/supplier/") >= 0) continue;
    if (ul.indexOf("/topics/") >= 0) continue;
    if (ul.indexOf("/topic/") >= 0) continue;
    if (ul.indexOf("/tag/") >= 0) continue;
    if (ul.indexOf("/category/") >= 0) continue;
    if (ul.indexOf("/search") >= 0) continue;

    out.push(abs);
  }

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
 * Rebuild News Source sheet from output rows (A:D).
 * Expected columns:
 * A: Company URL
 * B: Company Name
 * C: News URL
 * D: Source
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

    if (!a && !b && !c && !d) return;
    out.push([a, b, c, d]);
  });

  // Clear old data area (row 2 down)
  const last = sh.getLastRow();
  if (last >= 2) {
    sh.getRange(2, 1, last - 1, sh.getMaxColumns()).clearContent().clearNote();
  }

  // Write new rows (A:D)
  if (out.length) {
    sh.getRange(2, 1, out.length, 4).setValues(out);
  }
}

/** ===== OpenAI call for News_Url (kept for compatibility; not required by LOCAL checker) ===== */
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
    payload: JSON.stringify(payload)
  });

  const code = resp.getResponseCode();
  const text = resp.getContentText();
  if (code < 200 || code >= 300) throw new Error("OpenAI HTTP " + code + ": " + text);

  const data = JSON.parse(text);
  const answer = data && data.choices && data.choices[0] && data.choices[0].message && data.choices[0].message.content;
  if (!answer) throw new Error("No content returned from OpenAI (News_Url).");

  return String(answer).trim();
}
