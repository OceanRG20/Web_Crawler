/*************************************************
 * News_Agent.gs — News menu + GPT news crawler
 *
 * Sheets used:
 *   - "AI Integration"   (prompts + result log)
 *   - "News Source"      (input: Company URL, Name, News URL, Source)
 *   - "News Raw"         (output: normalized news rows)
 *
 * Prompt:
 *   Uses Prompt ID = "News_Search" from AI Integration!A:A
 **************************************************/

/** ===== Menu hook (called from AI_Agent.onOpen) ===== */
function onOpen_News(ui) {
  ui = ui || SpreadsheetApp.getUi();
  ui.createMenu("News")
    .addItem("▶ News Searching", "NS_runNewsSearch")
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

  const basePrompt =
    String(aiSheet.getRange(promptRow, 2).getValue() || "").trim();
  if (!basePrompt) {
    ui.alert('No prompt content in "AI Integration" for Prompt ID "News_Search".');
    return;
  }

  const newsRows = NS_getRowsFromNewsSource_();
  if (!newsRows.length) {
    ui.alert('No data rows found in sheet "News Source".');
    // still log an empty JSON array
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
    const msg = "News Search " + idx + "/" + newsRows.length;
    ss.toast(msg + " — " + (row.newsUrl || row.companyUrl), "News", 5);

    try {
      // If DIRECT-URL mode, try to auto-repair broken URLs via Wayback BEFORE calling OpenAI
      if (row.newsUrl) {
        const chk = NS_checkUrlReachable_(row.newsUrl);
        if (!chk.ok) {
          const wb = NS_tryWayback_(row.newsUrl);
          if (wb.ok) {
            const old = row.newsUrl;
            row.newsUrl = wb.url;

            // write back to News Source column C
            try {
              const ns = ss.getSheetByName("News Source");
              if (ns && row.rowIndex) {
                ns.getRange(row.rowIndex, 3).setValue(wb.url); // Col C = News URL
                ns.getRange(row.rowIndex, 3).setNote("Auto-repaired via Wayback from: " + old);
              }
            } catch (_) {}
          }
        }
      }

      const fullPrompt = NS_buildPromptForRow_(basePrompt, row);
      const rawAnswer = NS_callOpenAIForNews_(fullPrompt);
      const articles = NS_extractJsonArray_(rawAnswer);

      // Attach fallback Source / Company fields if missing in objects
      const enriched = articles.map((obj) => {
        let copy = Object.assign({}, obj);

        if (!copy["Source"]) {
          copy["Source"] = row.source || "";
        }
        if (!copy["Company Name"]) {
          copy["Company Name"] = row.companyName || "";
        }
        if (!copy["Company Website URL"]) {
          copy["Company Website URL"] = row.companyUrl || "";
        }

        // Ensure new keys exist, even if the model omitted them
        if (!copy.hasOwnProperty("Special Values")) {
          copy["Special Values"] = JSON.parse(JSON.stringify(NS_DEFAULT_SPECIAL_VALUES));
        } else if (!copy["Special Values"] || typeof copy["Special Values"] !== "object") {
          copy["Special Values"] = JSON.parse(JSON.stringify(NS_DEFAULT_SPECIAL_VALUES));
        } else {
          // ensure all keys exist
          Object.keys(NS_DEFAULT_SPECIAL_VALUES).forEach((k) => {
            if (!copy["Special Values"].hasOwnProperty(k)) copy["Special Values"][k] = "";
          });
        }

        if (!copy.hasOwnProperty("MMCrawl Updates")) {
          copy["MMCrawl Updates"] = "";
        }

        // Enforce: Special Values must be present in GPT Summary (summary-locked)
        copy = NS_enforceSpecialValuesFromSummary_(copy);

        return copy;
      });

      allArticles.push.apply(allArticles, enriched);
    } catch (err) {
      // On error, create a simple "error" record for this row
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

  // Write raw JSON back to AI Integration (Result + Date)
  const jsonOut = JSON.stringify(allArticles, null, 2);
  aiSheet.getRange(promptRow, 3).setValue(jsonOut);
  aiSheet.getRange(promptRow, 4).setValue(new Date());

  // Append to News Raw sheet
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
 *   C: News URL  (optional; if present, crawl only this URL)
 *   D: Source    (e.g., GPT-5.1, Internal, etc.)
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

    if (!companyUrl && !newsUrl) return; // skip empty rows

    out.push({
      rowIndex: i + 2, // actual sheet row number
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

  // Production-friendly: do NOT force "No content..." filler
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
    "- If no News Story URL is provided, follow the full search protocol, including allowed PR sources (e.g., PRWeb, Reuters), but keep only one canonical PR version when duplicates exist.\n" +
    "- When deduplicating, prefer the article with the richest relevant quotes and, when similar in scope, prefer the earliest publication date.\n" +
    "- Always populate 'Special Values' and 'MMCrawl Updates' as specified in the main prompt.\n" +
    "- Special Values are SUMMARY-LOCKED: only output a Special Value when that exact value also appears in GPT Summary.\n";

  return basePrompt + summaryTemplate + noNewsNote + scenarioBlock;
}

/** ===== OpenAI call for News_Search ===== */
function NS_callOpenAIForNews_(userPrompt) {
  const key =
    PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
  if (!key) {
    throw new Error(
      'Missing OpenAI API key. Use "AI Integration → Set OpenAI API Key".'
    );
  }

  // Try to reuse AIA.MODEL if defined; fallback to gpt-4o
  const model =
    (typeof AIA !== "undefined" && AIA && AIA.MODEL) ? AIA.MODEL : "gpt-4o";

  const payload = {
    model: model,
    temperature: 0.2,
    max_tokens: 6000,
    messages: [
      {
        role: "system",
        content:
          "You are an MBA-trained analyst with 5+ years researching U.S. precision metal and plastics manufacturers (mold building; tool & die; injection molding). " +
          "Your job is to perform a thorough news search OR, when a specific News Story URL is provided, to analyze only that URL. " +
          "You MUST follow all instructions in the user prompt, including the PRODUCTION GPT Summary rules, population of 'Special Values' (summary-locked), and extraction of 'MMCrawl Updates'. " +
          "Return ONLY a strict JSON array of article objects with the exact required keys (no markdown, no commentary)."
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
    throw new Error("OpenAI HTTP " + code + ": " + text);
  }

  const data = JSON.parse(text);
  const answer =
    data &&
    data.choices &&
    data.choices[0] &&
    data.choices[0].message &&
    data.choices[0].message.content;
  if (!answer) {
    throw new Error("No content returned from OpenAI (News_Search).");
  }
  return String(answer).trim();
}

/** ===== JSON helpers (local copy) ===== */
function NS_extractJsonArray_(text) {
  if (!text) return [];
  let t = String(text).trim();

  const fence =
    t.match(/```json([\s\S]*?)```/i) ||
    t.match(/```([\s\S]*?)```/i);
  if (fence) t = fence[1].trim();

  let obj = null;
  try {
    obj = JSON.parse(t);
  } catch (err) {
    const m = t.match(/(\{[\s\S]*\}|\[[\s\S]*\])/);
    if (m) {
      try {
        obj = JSON.parse(m[1]);
      } catch (_) {}
    }
  }

  if (!obj) return [];
  if (Array.isArray(obj)) {
    return obj.filter((v) => v && typeof v === "object");
  }
  if (typeof obj === "object") return [obj];
  return [];
}

/** ===== Enforce Special Values must appear in GPT Summary (summary-locked) ===== */
function NS_enforceSpecialValuesFromSummary_(obj) {
  const summary = String(obj["GPT Summary"] || "");
  const sv = (obj["Special Values"] && typeof obj["Special Values"] === "object")
    ? obj["Special Values"]
    : {};

  // Ensure all keys exist
  Object.keys(NS_DEFAULT_SPECIAL_VALUES).forEach((k) => {
    if (!sv.hasOwnProperty(k)) sv[k] = "";
  });

  // If no summary, blank all values
  if (!summary.trim()) {
    Object.keys(sv).forEach((k) => { sv[k] = ""; });
    obj["Special Values"] = sv;
    return obj;
  }

  const sumLower = summary.toLowerCase();

  // Strict: literal value must appear in GPT Summary
  Object.keys(sv).forEach((k) => {
    const v = (sv[k] === null || sv[k] === undefined) ? "" : String(sv[k]).trim();
    if (!v) {
      sv[k] = "";
      return;
    }
    if (!sumLower.includes(v.toLowerCase())) {
      sv[k] = "";
    }
  });

  obj["Special Values"] = sv;
  return obj;
}

/** ===== Append JSON objects to "News Raw" =====
 *
 * News Raw expected headers:
 *   A: Company Name
 *   B: Company Website URL
 *   C: News Story URL
 *   D: Headline
 *   E: Publication Date
 *   F: Publisher or Source
 *   G: GPT Summary
 *   H: Source
 *   I: Special Values (JSON)
 *   J: MMCrawl Updates (string / JSON)
 */
function NS_writeResultsToNewsRaw_(jsonArr) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName("News Raw");
  if (!sheet) throw new Error("Sheet 'News Raw' not found.");

  if (!jsonArr || !jsonArr.length) return;

  const lastRow = sheet.getLastRow();
  let row = (lastRow < 1 ? 2 : lastRow + 1);

  jsonArr.forEach((obj) => {
    // What you want in column I ("Specific Value")
    let finalValue = "";
    const mv = obj["MMCrawl Updates"];

    if (mv === null || mv === undefined) {
      finalValue = "";
    } else if (typeof mv === "string") {
      finalValue = mv;
    } else {
      // object/array/number/bool -> stringify so it never becomes [object Object]
      try {
        finalValue = JSON.stringify(mv);
      } catch (e) {
        finalValue = String(mv);
      }
    }

    const values = [
      obj["Company Name"] || "",          // A
      obj["Company Website URL"] || "",   // B
      obj["News Story URL"] || "",        // C
      obj["Headline"] || "",              // D
      obj["Publication Date"] || "",      // E
      obj["Publisher or Source"] || "",   // F
      obj["GPT Summary"] || "",           // G
      obj["Source"] || "",                // H
      finalValue                           // I  ✅ Specific Value
      // J intentionally unused/blank
    ];

    sheet.getRange(row, 1, 1, values.length).setValues([values]);
    row++;
  });
}



/** ===== URL reachability + Wayback repair helpers ===== */
function NS_checkUrlReachable_(url) {
  try {
    const resp = UrlFetchApp.fetch(url, {
      method: "get",
      followRedirects: true,
      muteHttpExceptions: true,
      validateHttpsCertificates: true
    });
    const code = resp.getResponseCode();
    const ok = (code >= 200 && code < 400);
    return { ok: ok, code: code, error: "" };
  } catch (e) {
    return { ok: false, code: 0, error: String(e && e.message ? e.message : e) };
  }
}

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
