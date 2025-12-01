/*************************************************
 * News_Agent.gs — Google Sheets × OpenAI (News only)
 *
 * Menu:
 *   News
 *     ▶ News Search        (Prompt ID = News_Search, via OpenAI)
 *
 * Sheets:
 *   - AI Integration   (prompts + results)
 *   - News Source      (Company URL | Company Name | News URL | Source)
 **************************************************/

// Config for the News agent
var NS = {
  INTEG_SHEET: "AI Integration",   // same as AIA.SHEET_NAME
  NEWS_SHEET: "News Source",

  // News Source columns (1-based)
  COL_COMPANY_URL: 1,   // A: Company URL
  COL_COMPANY_NAME: 2,  // B: Company Name
  COL_NEWS_URL: 3,      // C: News URL (optional)
  COL_SOURCE: 4,        // D: Source (optional)

  // Where to write in AI Integration (matches AIA.RESULT_COL / WHEN_COL)
  RESULT_COL: 3,  // C
  WHEN_COL: 4,    // D
  PREVIEW_COL: 5, // E (optional prompt preview)

  // OpenAI model for News_Search
  MODEL: "gpt-4o",
  TEMP: 0.25,
  MAX_TOKENS: 5000,

  DEBUG_SHOW_PROMPT: false,
  PREVIEW_MAX_CHARS: 18000
};

/*************************************************
 * MENU: called from global onOpen() in AI_Agent.gs
 **************************************************/
function onOpen_News(ui) {
  ui = ui || (typeof AIA_safeUi_ === "function" ? AIA_safeUi_() : SpreadsheetApp.getUi());
  if (!ui) return;

  ui.createMenu("News")
    .addItem("▶ News Search", "NS_runNewsSearch")
    .addToUi();
}

/*************************************************
 * MAIN RUNNER — News Search
 **************************************************/
function NS_runNewsSearch() {
  const ss = SpreadsheetApp.getActive();
  const ui = (typeof AIA_safeUi_ === "function" ? AIA_safeUi_() : SpreadsheetApp.getUi());

  // 1) Get News_Search prompt from AI Integration
  const integ = ss.getSheetByName(NS.INTEG_SHEET);
  if (!integ) {
    if (ui) ui.alert('Sheet "' + NS.INTEG_SHEET + '" not found.');
    return;
  }

  const newsRow = (typeof AIA_findPromptRow_ === "function")
    ? AIA_findPromptRow_("News_Search")
    : NS_findPromptRow_Fallback_("News_Search", integ);

  if (!newsRow) {
    if (ui) ui.alert('Prompt ID "News_Search" not found in AI Integration.');
    return;
  }

  const template = String(integ.getRange(newsRow, 2).getValue() || "").trim();
  if (!template) {
    if (ui) ui.alert('No template found in column B for "News_Search".');
    return;
  }

  // 2) Read candidates from News Source (skip header row)
  const newsSheet = ss.getSheetByName(NS.NEWS_SHEET);
  if (!newsSheet) {
    if (ui) ui.alert('Sheet "' + NS.NEWS_SHEET + '" not found.');
    return;
  }

  const lastRow = newsSheet.getLastRow();
  if (lastRow <= 1) {
    if (ui) ui.alert('"' + NS.NEWS_SHEET + '" has no data rows.');
    return;
  }

  // Data rows only: row 2..lastRow
  const data = newsSheet
    .getRange(2, 1, lastRow - 1, 4)
    .getDisplayValues();

  // 3) Loop rows, build per-row input, call OpenAI, collect all articles
  const allArticles = [];
  const total = data.length;

  if (ui) {
    ss.toast("News Search: " + total + " candidate row(s)…", "News", 4);
  }

  for (let i = 0; i < total; i++) {
    const rowVals = data[i].map(v => String(v || "").trim());
    const companyUrl = rowVals[NS.COL_COMPANY_URL - 1] || "";
    const companyName = rowVals[NS.COL_COMPANY_NAME - 1] || "";
    const newsUrl     = rowVals[NS.COL_NEWS_URL - 1] || "";
    const source      = rowVals[NS.COL_SOURCE - 1] || "";

    // Skip completely empty rows
    if (!companyUrl && !companyName && !newsUrl) continue;

    const candidateIndex = i + 1; // 1..N for user display (NOT sheet row)
    if (ui) {
      const label = (companyName || companyUrl || "Candidate " + candidateIndex);
      ss.toast(
        "News Search " + candidateIndex + "/" + total + " — " + label,
        "News",
        4
      );
    }

    // Build the JSON input object for this row
    const inputObj = {
      "Company Name": companyName,
      "Company Website URL": companyUrl,
      "News Story URL": newsUrl,
      "Source": source
    };

    // Use existing helper from AI_Agent if available, else JSON.stringify
    const jsonInput = (typeof AIA_jsonString_ === "function")
      ? AIA_jsonString_(inputObj)
      : JSON.stringify(inputObj, null, 2);

    let prompt =
      template +
      "\n\n### Input Row (News Source)\n" +
      jsonInput +
      "\n\n" +
      "Remember:\n" +
      '- If "News Story URL" is non-empty, analyze ONLY that URL and return a SINGLE JSON object in a one-element array.\n' +
      '- If "News Story URL" is empty, run the full search protocol and return zero or more objects in a JSON array.\n' +
      "- Output MUST be a JSON array only — no markdown, no commentary.";

    // Optional preview
    if (NS.DEBUG_SHOW_PROMPT && i === 0) {
      NS_writePreview_(integ, newsRow, prompt);
      if (ui) {
        const res = ui.alert(
          "News_Search Preview",
          "Full prompt written to column E. Proceed?",
          ui.ButtonSet.OK_CANCEL
        );
        if (res !== ui.Button.OK) return;
      }
    }

    try {
      const ans = NS_callOpenAI_News_(prompt);

      // Parse via existing helper if available
      const articles = (typeof AIA_extractJsonArray_ === "function")
        ? AIA_extractJsonArray_(ans)
        : NS_extractJsonArray_Fallback_(ans);

      if (articles && articles.length) {
        articles.forEach(a => allArticles.push(a));
      }
    } catch (err) {
      // On error, push a structured "No news / Error" object for this row
      allArticles.push({
        "Company Name": companyName || "",
        "Company Website URL": companyUrl || "",
        "News Story URL": newsUrl || "",
        "Headline": "Error in News_Search",
        "Publication Date": "",
        "Publisher or Source": "",
        "GPT Summary": String(err),
        "Confidence Score": "",
        "Source": source || ""
      });
    }

    Utilities.sleep(250);
  }

  // 4) Write combined JSON array to AI Integration result cell
  const resultJson = (typeof AIA_jsonString_ === "function")
    ? AIA_jsonString_(allArticles)
    : JSON.stringify(allArticles, null, 2);

  integ.getRange(newsRow, NS.RESULT_COL).setValue(resultJson);
  integ.getRange(newsRow, NS.WHEN_COL).setValue(new Date());

  if (ui) {
    ss.toast(
      "News_Search complete — " + allArticles.length + " article object(s) returned.",
      "News",
      6
    );
  }
}

/*************************************************
 * OpenAI call for News_Search
 **************************************************/
function NS_callOpenAI_News_(userPrompt) {
  const key = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
  if (!key) {
    throw new Error('Missing OpenAI API key. Use "AI Integration → Set OpenAI API Key".');
  }

  const payload = {
    model: NS.MODEL,
    temperature: NS.TEMP,
    max_tokens: NS.MAX_TOKENS,
    messages: [
      {
        role: "system",
        content:
          "You are an MBA-trained analyst specializing in news research for U.S. precision manufacturers. " +
          "Follow the detailed instructions in the user prompt. " +
          "Your ONLY output must be a syntactically valid JSON array of objects that match the requested schema. " +
          "Do not add any explanations, headings, or markdown."
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

/*************************************************
 * Small helpers
 **************************************************/

// Fallback prompt-row finder if AIA_findPromptRow_ is not available
function NS_findPromptRow_Fallback_(id, sheet) {
  const last = sheet.getLastRow();
  if (last < 2) return 0;
  const colA = sheet.getRange(2, 1, last - 1, 1).getDisplayValues();
  const target = String(id || "").trim().toLowerCase();
  for (let i = 0; i < colA.length; i++) {
    if (String(colA[i][0] || "").trim().toLowerCase() === target) {
      return 2 + i;
    }
  }
  return 0;
}

// Simple preview helper (optional)
function NS_writePreview_(sheet, row, text) {
  try {
    const headerCell = sheet.getRange(1, NS.PREVIEW_COL);
    if (!String(headerCell.getValue() || "").trim()) {
      headerCell.setValue("Prompt Preview");
    }
    const preview = String(text || "");
    const toWrite =
      preview.length > NS.PREVIEW_MAX_CHARS
        ? preview.slice(0, NS.PREVIEW_MAX_CHARS) + "\n...[truncated]"
        : preview;
    sheet.getRange(row, NS.PREVIEW_COL).setValue(toWrite);
  } catch (e) {
    // ignore
  }
}

// Fallback JSON-array extractor (if AIA_extractJsonArray_ is unavailable)
function NS_extractJsonArray_Fallback_(text) {
  if (!text) return [];
  let t = String(text).trim();
  const fence = t.match(/```json([\s\S]*?)```/i) || t.match(/```([\s\S]*?)```/);
  if (fence) t = fence[1].trim();
  let obj = null;
  try {
    obj = JSON.parse(t);
  } catch (_) {
    const m = t.match(/(\[[\s\S]*\]|\{[\s\S]*\})/);
    if (m) {
      try {
        obj = JSON.parse(m[1]);
      } catch (_) {}
    }
  }
  if (!obj) return [];
  if (Array.isArray(obj)) {
    return obj.filter(v => v && typeof v === "object");
  }
  if (typeof obj === "object") return [obj];
  return [];
}
