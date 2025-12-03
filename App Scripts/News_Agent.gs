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
      const fullPrompt = NS_buildPromptForRow_(basePrompt, row);
      const rawAnswer = NS_callOpenAIForNews_(fullPrompt);
      const articles = NS_extractJsonArray_(rawAnswer);

      // Attach fallback Source / Company fields if missing in objects
      const enriched = articles.map((obj) => {
        const copy = Object.assign({}, obj);

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
          copy["Special Values"] = {
            "Square footage (facility)": "",
            "Number of employees": "",
            "Estimated Revenues": "",
            "Family business": "",
            "Medical": ""
          };
        }
        if (!copy.hasOwnProperty("MMCrawl Updates")) {
          copy["MMCrawl Updates"] = "";
        }

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
        "Special Values": {
          "Square footage (facility)": "",
          "Number of employees": "",
          "Estimated Revenues": "",
          "Family business": "",
          "Medical": ""
        },
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

  vals.forEach((r) => {
    const companyUrl = String(r[0] || "").trim();
    const companyName = String(r[1] || "").trim();
    const newsUrl = String(r[2] || "").trim();
    const source = String(r[3] || "").trim();

    if (!companyUrl && !newsUrl) return; // skip empty rows

    out.push({
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

  // Extra guidance (still compatible with updated News_Search prompt)
  const summaryTemplate =
    "\n\n### Additional formatting requirements\n" +
    "- For each accepted article, structure GPT Summary exactly as required in the prompt (sections 1–9 with OVERVIEW, LEADERSHIP & OWNERSHIP, FACILITY/EQUIPMENT, etc.).\n" +
    "- Always explicitly state when information is not mentioned in the article (e.g., “No direct quotes appear in this article.”).\n" +
    "- Keep the GPT Summary JSON-safe (no unescaped line breaks other than \\n) and non-promotional.\n";

  const noNewsNote =
    "\nIf you truly find no acceptable articles after filtering (or if the given News Story URL is invalid), return a single JSON object in the exact 'NO NEWS CASE' format from the instructions, " +
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
    "- Always populate 'Special Values' and 'MMCrawl Updates' as specified in the main prompt.\n";

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
          "You MUST follow all instructions in the user prompt, including construction of GPT Summary sections, population of 'Special Values', and extraction of 'MMCrawl Updates'. " +
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
    const specialValues = obj["Special Values"] || {
      "Square footage (facility)": "",
      "Number of employees": "",
      "Estimated Revenues": "",
      "Family business": "",
      "Medical": ""
    };

    const mmcrawlUpdates = obj["MMCrawl Updates"] !== undefined
      ? obj["MMCrawl Updates"]
      : "";

    const values = [
      obj["Company Name"] || "",
      obj["Company Website URL"] || "",
      obj["News Story URL"] || "",
      obj["Headline"] || "",
      obj["Publication Date"] || "",     // Column E
      obj["Publisher or Source"] || "",  // Column F
      obj["GPT Summary"] || "",          // Column G
      obj["Source"] || "",               // Column H
      (typeof mmcrawlUpdates === "string"
        ? mmcrawlUpdates
        : JSON.stringify(mmcrawlUpdates)) // Column J
    ];

    sheet.getRange(row, 1, 1, values.length).setValues([values]);
    row++;
  });
}
