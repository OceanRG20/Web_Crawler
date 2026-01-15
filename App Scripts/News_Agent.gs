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

/** ===== Special Values defaults ===== */
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

    try {
      const fullPrompt = NS_buildPromptForRow_(basePrompt, row);

      const rawAnswer = NS_callOpenAIForNews_(fullPrompt);
      const articles = NS_extractJsonArray_(rawAnswer);

      const finalArticles = (articles && articles.length)
        ? articles
        : [makeFallbackNoNewsObject_(row, "Model returned empty/invalid JSON array for this input.")];

      const enriched = finalArticles.map((obj) => {
        let copy = Object.assign({}, obj);

        // Ensure key fields are present
        if (!copy["Source"]) copy["Source"] = row.source || "";
        if (!copy["Company Name"]) copy["Company Name"] = row.companyName || "";
        if (!copy["Company Website URL"]) copy["Company Website URL"] = row.companyUrl || "";
        if (!copy["News Story URL"]) copy["News Story URL"] = row.newsUrl || "";

        // Ensure Special Values is an object with expected keys
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

        // Enforce your pipeline rules
        copy = NS_enforceSpecialValuesFromSummary_(copy);
        copy = NS_adjustMmcrawlUpdateYears_(copy);

        return copy;
      });

      allArticles.push.apply(allArticles, enriched);

    } catch (err) {
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

/** ===== Build per-row prompt (UPDATED: inject fetched article text in direct-URL mode) ===== */
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

  // NEW: Fetch article text for direct-URL mode, inject it so the model can extract facts.
  let articleBlock = "";
  if (newsUrl) {
    const articleText = NS_fetchArticleText_(newsUrl, 12000);
    if (articleText) {
      articleBlock =
        "\n\n### Article text (fetched by system)\n" +
        "You MUST base all extraction ONLY on the text below.\n" +
        "If a data point is not explicitly stated in this text, do NOT output it.\n\n" +
        articleText +
        "\n";
    } else {
      articleBlock =
        "\n\n### Article text (fetched by system)\n" +
        "The system could not fetch readable text from the URL (blocked/paywall/format).\n" +
        "In this case, do NOT guess numbers. Only output what you can confirm.\n";
    }
  }

  return basePrompt + summaryTemplate + noNewsNote + scenarioBlock + articleBlock;
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

/**
 * Enforce year logic on MMCrawl Updates lines:
 * - Trailing year must be earlier of:
 *     publication year vs. year(s) mentioned in the bullet text
 * - If no mentioned year in bullet text, default trailing year to publication year (if available)
 *
 * Expected MMCrawl Updates line format:
 *   Field ; "Bullet text (News: <URL>)", 2021
 */
function NS_adjustMmcrawlUpdateYears_(obj) {
  const updatesRaw = obj["MMCrawl Updates"];
  if (!updatesRaw) return obj;

  const pub = String(obj["Publication Date"] || "").trim();
  const pubYearMatch = pub.match(/\b(19|20)\d{2}\b/);
  const pubYear = pubYearMatch ? parseInt(pubYearMatch[0], 10) : null;

  let lines = [];
  if (Array.isArray(updatesRaw)) {
    lines = updatesRaw.map(x => String(x || "")).filter(Boolean);
  } else {
    lines = String(updatesRaw || "").split("\n").map(s => s.trim()).filter(Boolean);
  }
  if (!lines.length) return obj;

  const fixed = lines.map((line) => {
    const original = String(line || "").trim();
    if (!original) return "";

    const m = original.match(/^([\s\S]*?;\s*")([\s\S]*?)("\s*,\s*)(\d{4})\s*$/);
    if (!m) return original;

    const prefix = m[1];
    const bulletText = m[2];
    const mid = m[3];
    const trailingYear = parseInt(m[4], 10);

    const yearsInText = [];
    const re = /\b(19|20)\d{2}\b/g;
    let ym;
    while ((ym = re.exec(bulletText)) !== null) {
      yearsInText.push(parseInt(ym[0], 10));
    }

    let effectiveYear = trailingYear;

    if (pubYear) {
      if (yearsInText.length) {
        const minMentioned = Math.min.apply(null, yearsInText);
        effectiveYear = Math.min(pubYear, minMentioned);
      } else {
        effectiveYear = pubYear;
      }
    } else {
      if (yearsInText.length) {
        effectiveYear = Math.min.apply(null, yearsInText);
      }
    }

    return prefix + bulletText + mid + String(effectiveYear);
  }).filter(Boolean);

  obj["MMCrawl Updates"] = fixed.join("\n");
  return obj;
}

/** ===== HTML fetch + extract article text (NEW) ===== */

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

/** Basic HTML -> readable text (best-effort). */
function NS_htmlToText_(html) {
  if (!html) return "";
  let t = String(html);

  t = t.replace(/<script[\s\S]*?<\/script>/gi, " ");
  t = t.replace(/<style[\s\S]*?<\/style>/gi, " ");
  t = t.replace(/<br\s*\/?>/gi, "\n");
  t = t.replace(/<\/p\s*>/gi, "\n");
  t = t.replace(/<[^>]+>/g, " ");

  t = t.replace(/&nbsp;/g, " ");
  t = t.replace(/&amp;/g, "&");
  t = t.replace(/&quot;/g, "\"");
  t = t.replace(/&#39;/g, "'");
  t = t.replace(/&lt;/g, "<");
  t = t.replace(/&gt;/g, ">");

  t = t.replace(/[ \t\r\f\v]+/g, " ");
  t = t.replace(/\n\s+/g, "\n");
  t = t.replace(/\n{3,}/g, "\n\n");

  return t.trim();
}

/** Fetch and extract article text (size-limited). */
function NS_fetchArticleText_(url, maxChars) {
  const u = String(url || "").trim();
  if (!u) return "";

  const r = NS_fetchHtml_(u);
  if (!r || !r.ok || !r.html) return "";

  const txt = NS_htmlToText_(r.html);
  if (!txt) return "";

  const limit = (maxChars && maxChars > 0) ? maxChars : 12000;
  return txt.length > limit ? txt.slice(0, limit) : txt;
}

/** ===== Append JSON objects to "News Raw" ===== */
function NS_writeResultsToNewsRaw_(jsonArr) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName("News Raw");
  if (!sheet) throw new Error("Sheet 'News Raw' not found.");

  if (!jsonArr || !jsonArr.length) return;

  const lastRow = sheet.getLastRow();
  let row = (lastRow < 1 ? 2 : lastRow + 1);

  function normalizeSpecificValue_(mv) {
    if (mv === null || mv === undefined) return "";

    if (typeof mv === "string") {
      const s = mv.trim();
      if (!s) return "";
      if ((s.startsWith("[") && s.endsWith("]")) || (s.startsWith("{") && s.endsWith("}"))) {
        try {
          const parsed = JSON.parse(s);
          return normalizeSpecificValue_(parsed);
        } catch (_) {
          return s;
        }
      }
      return s;
    }

    if (Array.isArray(mv)) {
      return mv
        .map((x) => {
          if (x === null || x === undefined) return "";
          if (typeof x === "string") return x.trim();
          try { return JSON.stringify(x); } catch (_) { return String(x); }
        })
        .filter(Boolean)
        .join("\n");
    }

    if (typeof mv === "object") {
      const lines = [];
      Object.keys(mv).forEach((k) => {
        const key = String(k || "").trim();
        if (!key) return;

        let v = mv[k];
        if (v === null || v === undefined) return;

        if (Array.isArray(v)) {
          v = v.map((z) => (z === null || z === undefined) ? "" : String(z).trim()).filter(Boolean).join(", ");
        } else if (typeof v === "object") {
          try { v = JSON.stringify(v); } catch (_) { v = String(v); }
        } else {
          v = String(v).trim();
        }

        if (!v) return;
        lines.push(key + ' ; "' + v.replace(/"/g, '\\"') + '"');
      });

      return lines.join("\n");
    }

    return String(mv);
  }

  jsonArr.forEach((obj) => {
    const specificValueText = normalizeSpecificValue_(obj["MMCrawl Updates"]);

    const values = [
      obj["Company Name"] || "",
      obj["Company Website URL"] || "",
      obj["News Story URL"] || "",
      obj["Headline"] || "",
      obj["Publication Date"] || "",
      obj["Publisher or Source"] || "",
      obj["GPT Summary"] || "",
      obj["Source"] || "",
      specificValueText
    ];

    sheet.getRange(row, 1, 1, values.length).setValues([values]);
    row++;
  });
}

/** ===== Soft-404 detector (simple string checks) ===== */
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

    const isArticleLike =
      ul.indexOf("/articles/") >= 0 ||
      ul.indexOf("/article/") >= 0 ||
      ul.indexOf("/news/") >= 0;

    if (!isArticleLike) continue;

    if (ul.indexOf("/suppliers/") >= 0) continue;
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

    const flag = (e === "Yes") ? "Yes" : "";
    out.push([a, b, c, d, flag]);
  });

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
 * Rules:
 * - HTTP 200 with real content -> keep
 * - HTTP 404 / 429 -> remove
 * - HTTP 200 soft-404:
 *      - if closest matches exist -> replace with matches, set Archive_Flag="Yes"
 *      - if no matches -> remove
 * - blocked / fetch failed -> keep as-is
 *
 * Saves final JSON into AI Integration -> Result for Prompt ID "News_Url",
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

  let checked_rows = 0;
  let kept = 0;
  let removed_404_429 = 0;
  let replaced_soft404 = 0;
  let removed_soft404_no_matches = 0;
  let added_matches = 0;

  const outRows = [];
  const seenNewsUrls = {};

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

    if (!newsUrl) {
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
      kept++;
      addRow_(companyUrl, companyName, newsUrl, source, "");
      continue;
    }

    if (code === 404 || code === 429) {
      removed_404_429++;
      continue;
    }

    if (!(code >= 200 && code < 400)) {
      kept++;
      addRow_(companyUrl, companyName, newsUrl, source, "");
      continue;
    }

    if (NS_isSoft404_(newsUrl, html)) {
      const candidates = NS_extractCandidateUrlsFromHtml_(html, newsUrl).slice(0, 6);

      if (candidates.length) {
        replaced_soft404++;
        candidates.forEach((u) => {
          addRow_(companyUrl, companyName, u, source, "Yes");
        });
        added_matches += candidates.length;
      } else {
        removed_soft404_no_matches++;
      }
      continue;
    }

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

  aiSheet.getRange(promptRow, 3).setValue(JSON.stringify(resultObj, null, 2));
  aiSheet.getRange(promptRow, 4).setValue(new Date());

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
