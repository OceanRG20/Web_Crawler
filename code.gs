/**** MoldMakerCSV1 — Clean Codebase (MMCrawl + Record View + News Raw/List + News Record) ****/
/**** Assumes News List is a formula view of News Raw. No script-driven filters. **************/

/* ===== CONFIG ===== */
const SHEET_NAME = "MMCrawl"; // master company table
const VIEW_SHEET = "Record View"; // single record UI
const NEWS_RAW = "News Raw"; // full news dataset (you maintain only this)
const NEWS_LIST = "News List"; // formula-driven view of News Raw (by current company)
const NEWS_RECORD = "News Record"; // single news UI
const REJECTED_SHEET = "Rejected Companies";
const LOG_SHEET = "Logs";
const HEADER_ROW = 1;

/* ===== Column header names (MMCrawl) ===== */
const H = {
  company: "Company Name",
  website: "Public Website Homepage URL",
  domain: "Domain from URL",
  street: "Street Address",
  city: "City",
  state: "State",
  zip: "Zipcode",
  phone: "Phone",
  source: "Source",
  status: "Target Status",
  ownership: "Ownership",

  industries: "Industries served",
  products: "Products and services offered",

  sqft: "Square footage (facility)",
  employees: "Number of employees",
  revenue: "Estimated Revenues",
  years: "Years of operation",
  equipment: "Equipment",
  cnc3: "CNC 3-axis",
  cnc5: "CNC 5-axis",
  spares: "Spares/Repairs",
  family: "Family business",
  addr2: "2nd Address",

  notes: "Notes (Approach/Contacts/Info)", // ← column AF in your sheet
};

/* ===== Label variants (Record View col A) ===== */
const L = {
  company: ["Company Name"],
  website: ["Public Website Homepage URL", "Website"],
  domain: ["Domain", "Domain from URL"],
  source: ["Source"],
  status: ["Target Status", "Status"],
  ownership: ["Ownership", "Owner", "Ownership / Owner"],

  addr: ["Address & Phone", "Address"],

  industries: [
    "Industries served",
    "Industries Served",
    "Industries – served",
    "Industries",
  ],
  products: [
    "Products and services offered",
    "Products & services offered",
    "Products and Services Offered",
    "Products & Services",
  ],

  sqft: [
    "Square footage (facility)",
    "Square Footage",
    "Sq Ft",
    "Square footage",
  ],
  employees: ["Number of employees", "# Employees", "Employees"],
  revenue: [
    "Estimated Revenues",
    "Estimated Revenue",
    "Revenue (est)",
    "Estimated revenues",
  ],
  years: ["Years of operation", "Years of operaation", "Years"],

  equipment: ["Equipment"],
  cnc3: ["CNC 3-axis", "CNC 3 axis", "3-axis", "3 Axis"],
  cnc5: ["CNC 5-axis", "CNC 5 axis", "5-axis", "5 Axis"],
  spares: ["Spares/Repairs", "Spares / Repairs", "Repairs"],
  family: ["Family business", "Family"],
  addr2: ["2nd Address", "Second Address"],

  notes: ["Notes (Approach/Contacts/Info)", "Notes"],
  rownum: ["Row #", "Row Number"],
};

/* ===== News summary header autodetect aliases ===== */
const NEWS_SUMMARY_ALIASES = [
  "summary",
  "abstract",
  "description",
  "notes",
  "story",
  "blurb",
];

/* ===== Cached state ===== */
const DP = () => PropertiesService.getDocumentProperties();
const KEY_LAST_COMPANY = "LAST_COMPANY";
const KEY_LAST_COMPANY_ROW = "LAST_COMPANY_ROW";
const KEY_LAST_NEWSROW = "LAST_NEWS_ROW";

/* ===================== LOGGING (human-readable) ===================== */
function ensureLogSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(LOG_SHEET);
  if (!sh) {
    sh = ss.insertSheet(LOG_SHEET);
    sh.getRange(1, 1).setValue("Timestamp");
    sh.getRange(1, 2).setValue("Action");
    sh.getRange(1, 3).setValue("Details");
    sh.setFrozenRows(1);
    sh.setColumnWidths(1, 3, 320);
  }
  return sh;
}
function logPretty_(action, text) {
  try {
    ensureLogSheet_().appendRow([new Date(), action, text]);
  } catch (_) {}
}

/* Helpers for friendly text */
function headerAt_(headers, colIdx1) {
  return headers && headers[colIdx1 - 1]
    ? String(headers[colIdx1 - 1])
    : "C" + colIdx1;
}
function sRange_(r1, r2) {
  return r1 === r2 ? String(r1) : `${r1}-${r2}`;
}
function sList_(arr, max = 5) {
  const a = (arr || []).filter(Boolean);
  return a.length <= max ? a.join("; ") : a.slice(0, max).join("; ") + "; …";
}

/* ===================== SIZE TRACKING (adds/deletes) ===================== */
const SIZE_KEYS = {
  [SHEET_NAME]: "SIZE_MMCRAWL",
  [NEWS_RAW]: "SIZE_NEWSRAW",
  [REJECTED_SHEET]: "SIZE_REJECTED",
};
function initSizeTracking_() {
  [SHEET_NAME, NEWS_RAW, REJECTED_SHEET].forEach(updateSizeProperty_);
}
function updateSizeProperty_(sheetName) {
  const sh = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (sh) DP().setProperty(SIZE_KEYS[sheetName], String(sh.getLastRow()));
}
function detectAndLogAddsDeletes_(sheetName) {
  const sh = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sh) return;
  const key = SIZE_KEYS[sheetName];
  const prev = Number(DP().getProperty(key) || "0");
  const cur = sh.getLastRow();
  if (prev === 0) {
    DP().setProperty(key, String(cur));
    return;
  }
  if (cur === prev) return;

  const delta = cur - prev;
  const lastCol = sh.getLastColumn();
  const headers =
    cur > HEADER_ROW
      ? sh.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0]
      : [];

  if (delta > 0) {
    const startRow = prev + 1,
      endRow = cur,
      addedCount = delta;
    let samples = [];
    try {
      if (sheetName === SHEET_NAME) {
        const cIdx = getHeaderIndexSmart_(headers, H.company);
        const idCol = cIdx >= 0 ? cIdx + 1 : 1;
        samples = sh
          .getRange(startRow, idCol, addedCount, 1)
          .getValues()
          .flat()
          .map((v) => String(v || ""))
          .filter(Boolean);
      } else if (sheetName === NEWS_RAW) {
        const uIdx = getHeaderIndexSmart_(headers, "News Story URL");
        const hIdx = getHeaderIndexSmart_(headers, "Headline");
        const vals = sh.getRange(startRow, 1, addedCount, lastCol).getValues();
        samples = vals
          .map(
            (r) =>
              (uIdx >= 0 ? String(r[uIdx] || "") : "") ||
              (hIdx >= 0 ? String(r[hIdx] || "") : "")
          )
          .filter(Boolean);
      } else if (sheetName === REJECTED_SHEET) {
        samples = sh
          .getRange(startRow, 1, addedCount, 1)
          .getValues()
          .flat()
          .map((v) => String(v || ""))
          .filter(Boolean);
      }
    } catch (_) {}
    logPretty_(
      "added",
      `${sheetName} rows ${sRange_(
        startRow,
        endRow
      )} added (${addedCount} rows).` +
        (samples.length ? ` Sample: ${sList_(samples)}` : "")
    );
  } else {
    const deletedCount = Math.abs(delta);
    if (cur < prev && sh.getLastRow() === cur) {
      logPretty_(
        "deleted",
        `${sheetName} rows ${sRange_(
          cur + 1,
          prev
        )} deleted (${deletedCount} rows).`
      );
    } else {
      logPretty_("deleted", `${sheetName} deleted ${deletedCount} rows.`);
    }
  }
  DP().setProperty(key, String(cur));
}

/* ====== Record View menu + init — called by AI_Agent.gs → onOpen ====== */
function onOpen_Code(ui) {
  ui = ui || SpreadsheetApp.getUi();

  // 1) Always add the Record View menu
  try {
    ui.createMenu("Record View")
      .addItem("Search by Company Name", "rvSearchPrompt")
      .addToUi();
  } catch (err) {
    Logger.log("Record View menu: " + err);
  }

  // 2) Log "open" and init size tracking
  try {
    const user = Session.getActiveUser().getEmail() || "unknown";
    logPretty_("open", `Open: ${user}`);
    initSizeTracking_();
  } catch (err) {
    Logger.log("init/log on open: " + err);
  }

  // 3) Auto-dedupe on open (and sync size so deletes don’t double-log)
  try {
    const d1 = mmcrawlRemoveDuplicateUrls(true);
    if (d1 && d1.count)
      logPretty_("delete_duplication", prettyDupMsg_("MMCrawl", d1));
    updateSizeProperty_(SHEET_NAME);
  } catch (err) {
    Logger.log("MMCrawl dedupe: " + err);
  }

  try {
    const d2 = newsRawRemoveDuplicateStories(true);
    if (d2 && d2.count)
      logPretty_("delete_duplication", prettyDupMsg_("News Raw", d2));
    updateSizeProperty_(NEWS_RAW);
  } catch (err) {
    Logger.log("News Raw dedupe: " + err);
  }
}

/* ====== Spreadsheet-level change tracking ====== */
function onChange(e) {
  if (!e || !e.changeType) return;
  const sheetName =
    e.source && e.source.getActiveSheet()
      ? e.source.getActiveSheet().getName()
      : "(unknown)";
  logPretty_(String(e.changeType), `${sheetName}`);
  [SHEET_NAME, NEWS_RAW, REJECTED_SHEET].forEach(detectAndLogAddsDeletes_);
}

/* ====== Core UX wiring ====== */
function onSelectionChange(e) {
  const sh = e && e.range && e.range.getSheet();
  if (!sh) return;
  const name = sh.getName();

  if (name === SHEET_NAME) {
    const row = e.range.getRow();
    if (row > HEADER_ROW) {
      const comp = getCompanyNameAtRow_(row);
      if (!comp) return;
      setLastCompany_(comp);
      setLastCompanyRow_(row);
      updateRecordViewFromRow_(row);
      const nrow = findFirstNewsListRow_();
      if (nrow) {
        setLastNewsRow_(nrow);
        const nrec = getNewsRowFromNewsListRow_(nrow);
        if (nrec) renderNewsToView_(nrec);
      } else {
        clearLastNewsRow_();
        clearNewsRecordValues_();
      }
    }
    return;
  }

  if (name === NEWS_LIST) {
    const row = e.range.getRow();
    if (row > HEADER_ROW) {
      setLastNewsRow_(row);
      const nrec = getNewsRowFromNewsListRow_(row);
      if (nrec) renderNewsToView_(nrec);
    }
    return;
  }

  if (name === NEWS_RECORD) {
    let row = getLastNewsRow_();
    if (!row || row <= HEADER_ROW) row = findFirstNewsListRow_();
    if (row) {
      const nrec = getNewsRowFromNewsListRow_(row);
      if (nrec) renderNewsToView_(nrec);
    } else {
      clearNewsRecordValues_();
    }
    return;
  }
}

/* ====== Record View: Labels ====== */
function rvBuildLabels() {
  const sh = getViewSheet_();
  const order = [
    "company",
    "website",
    "domain",
    "source",
    "status",
    "ownership",
    "addr",
    "industries",
    "products",
    "sqft",
    "employees",
    "revenue",
    "years",
    "equipment",
    "cnc3",
    "cnc5",
    "spares",
    "family",
    "addr2",
    "notes",
    "rownum",
  ];
  for (let i = 0; i < order.length; i++) {
    const key = order[i];
    const row = i + 1;
    sh.getRange(row, 1).setValue((L[key] && L[key][0]) || key);
    sh.getRange(row, 2).clearContent();
  }
  sh.getRange(1, 1, order.length, 1).setFontWeight("bold");
  const labelMap = buildLabelRowMap_(sh);
  const notesRow = labelMap["notes"];
  if (notesRow) {
    sh.getRange(notesRow, 2).setWrap(true);
    sh.setRowHeight(notesRow, 60);
  }
  sh.setColumnWidth(1, 220);
  sh.setColumnWidth(2, 520);
  SpreadsheetApp.getUi().alert("Record View labels refreshed.");
}

/* ====== Record View: Show selected row ====== */
function rvShowActiveRow() {
  const row = getLastCompanyRow_();
  if (row > HEADER_ROW) {
    updateRecordViewFromRow_(row);
    return;
  }
  const rec = getActiveCompanyRecord_();
  if (!rec)
    return SpreadsheetApp.getUi().alert(
      'Put the cursor on a data row in "' + SHEET_NAME + '".'
    );
  renderToView_(rec);
}

/* ====== Record View: Prompt search ====== */
function rvSearchPrompt() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt(
    "Search by Company Name",
    "Enter any part of a name:",
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  quickSearch_(resp.getResponseText());
}

/* ====== onEdit: logs + dedupe + RV notes ====== */
function onEdit(e) {
  const rng = e && e.range;
  if (!rng) return;
  const sh = rng.getSheet();
  const name = sh.getName();

  if (name === SHEET_NAME || name === NEWS_RAW || name === REJECTED_SHEET) {
    const rows = rng.getNumRows(),
      cols = rng.getNumColumns();
    const lastCol = sh.getLastColumn();
    const headers = sh.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0];
    if (rows === 1 && cols === 1) {
      const r = rng.getRow(),
        c = rng.getColumn();
      const header = headerAt_(headers, c);
      const oldV = "oldValue" in e ? e.oldValue : "";
      const newV = "value" in e ? e.value : String(rng.getValue() ?? "");
      logPretty_(
        "update",
        `${name} R${r} "${header}": "${oldV || ""}" → "${newV || ""}"`
      );
    } else {
      const r1 = rng.getRow(),
        r2 = r1 + rows - 1;
      logPretty_(
        "update",
        `${name} rows ${sRange_(r1, r2)} (${rows} rows, ${cols} cols)`
      );
    }
    detectAndLogAddsDeletes_(name);
  }

  if (name === SHEET_NAME) {
    try {
      const d = mmcrawlRemoveDuplicateUrls(true);
      if (d && d.count) {
        logPretty_("delete_duplication", prettyDupMsg_("MMCrawl", d));
        updateSizeProperty_(SHEET_NAME);
      }
    } catch (_) {}
    return;
  }
  if (name === NEWS_RAW) {
    try {
      const d = newsRawRemoveDuplicateStories(true);
      if (d && d.count) {
        logPretty_("delete_duplication", prettyDupMsg_("News Raw", d));
        updateSizeProperty_(NEWS_RAW);
      }
    } catch (_) {}
    return;
  }

  if (name === VIEW_SHEET) {
    const searchCell = getSearchInput_(); // D1
    if (
      rng.getRow() === searchCell.getRow() &&
      rng.getColumn() === searchCell.getColumn()
    ) {
      quickSearch_(e.value);
      return;
    }
    const labelMap = buildLabelRowMap_(sh);
    const notesRow = labelMap["notes"];
    if (notesRow && rng.getColumn() === 2 && rng.getRow() === notesRow) {
      const rownumRow = labelMap["rownum"];
      if (!rownumRow) return;
      const rowNum = Number(sh.getRange(rownumRow, 2).getValue());
      if (!rowNum || rowNum <= HEADER_ROW) return;
      const notesText = String(rng.getValue() || "");
      saveNotesToData_(rowNum, notesText);
      logPretty_(
        "notes_update",
        `MMCrawl notes saved (row ${rowNum}, ${notesText.length} chars)`
      );
      SpreadsheetApp.getActive().toast("Notes saved", "Record View", 2);
    }
  }
}

/* ====== Quick search ====== */
function getSearchInput_() {
  return getViewSheet_().getRange(1, 4);
} // D1
function quickSearch_(raw) {
  const inputCell = getSearchInput_();
  const q = (raw || "").toString().trim().toLowerCase();
  if (!q) return;
  const { headers, data } = getTable_(SHEET_NAME);
  const cIdx = getHeaderIndexSmart_(headers, H.company);
  if (cIdx === -1) {
    SpreadsheetApp.getUi().alert('Could not find column "' + H.company + '".');
    return;
  }
  for (let i = 0; i < data.length; i++) {
    const name = (data[i][cIdx] || "").toString().trim().toLowerCase();
    if (name.indexOf(q) !== -1) {
      const rowNum = HEADER_ROW + 1 + i;
      selectDataRow_(rowNum);
      setLastCompanyRow_(rowNum);
      const comp = String(data[i][cIdx] || "");
      setLastCompany_(comp);
      updateRecordViewFromRow_(rowNum);
      Utilities.sleep(50);
      const nrow = findFirstNewsListRow_();
      if (nrow) {
        setLastNewsRow_(nrow);
        const nrec = getNewsRowFromNewsListRow_(nrow);
        if (nrec) renderNewsToView_(nrec);
      } else {
        clearLastNewsRow_();
        clearNewsRecordValues_();
      }
      inputCell.setValue("");
      SpreadsheetApp.getActive().toast("Found: " + comp, "Search", 2);
      return;
    }
  }
  SpreadsheetApp.getActive().toast('No match for "' + raw + '"', "Search", 3);
}

/* ====== News Record: Labels ====== */
function newsBuildLabels() {
  const rec = getNewsRecordSheet_();
  const { headers } = getTable_(NEWS_RAW);
  if (!headers.length)
    return SpreadsheetApp.getUi().alert(
      'No headers found in "' + NEWS_RAW + '".'
    );
  const map = detectNewsMapping_(headers);
  const labels = map.order
    .filter((h) => h !== map.summaryHeader)
    .concat([map.summaryHeader]);
  rec.clear();
  for (let i = 0; i < labels.length; i++) {
    rec.getRange(i + 1, 1).setValue(String(labels[i]));
    rec.getRange(i + 1, 2).setValue("");
  }
  rec.getRange(1, 1, labels.length, 1).setFontWeight("bold");
  rec
    .getRange(1, 2, labels.length, 1)
    .setWrap(true)
    .setVerticalAlignment("top");
  rec.setColumnWidth(1, 220);
  rec.setColumnWidth(2, 700);
  const sumRow = labels.indexOf(map.summaryHeader) + 1;
  if (sumRow > 0) {
    rec.setRowHeight(sumRow, 120);
  }
  SpreadsheetApp.getUi().alert(
    'News Record labels refreshed from "' + NEWS_RAW + '".'
  );
}

/* ====== Renderers ====== */
function updateRecordViewFromRow_(rowNum) {
  if (!rowNum || rowNum <= HEADER_ROW) return;
  const rec = getCompanyRecordAtRow_(rowNum);
  if (rec) renderToView_(rec);
}
function renderToView_(rec) {
  const view = getViewSheet_();
  const m = rec._map || {};
  const labelMap = buildLabelRowMap_(view);
  const { headers } = getTable_(SHEET_NAME);
  const headerMap = {};
  Object.keys(H).forEach(
    (k) => (headerMap[k] = getHeaderIndexSmart_(headers, H[k]))
  );
  const rowsToClear = Object.values(labelMap);
  if (rowsToClear.length) {
    const minR = Math.min.apply(null, rowsToClear);
    const maxR = Math.max.apply(null, rowsToClear);
    view.getRange(minR, 2, maxR - minR + 1, 1).clearContent();
  }
  function put(key, val) {
    const r = labelMap[key];
    if (r) view.getRange(r, 2).setValue(val || "");
  }
  put("company", valByKey_(m, headerMap, headers, "company"));
  put("website", valByKey_(m, headerMap, headers, "website"));
  put("domain", valByKey_(m, headerMap, headers, "domain"));
  put("source", valByKey_(m, headerMap, headers, "source"));
  put("status", valByKey_(m, headerMap, headers, "status"));
  put("ownership", valByKey_(m, headerMap, headers, "ownership"));
  const line1 = valByKey_(m, headerMap, headers, "street");
  const line2 = [
    valByKey_(m, headerMap, headers, "city"),
    valByKey_(m, headerMap, headers, "state"),
    valByKey_(m, headerMap, headers, "zip"),
  ]
    .filter(Boolean)
    .join(", ")
    .replace(", ,", ",");
  const line3 = valByKey_(m, headerMap, headers, "phone");
  put("addr", [line1, line2, line3].filter(Boolean).join("\n"));
  put("industries", valByKey_(m, headerMap, headers, "industries"));
  put("products", valByKey_(m, headerMap, headers, "products"));
  put("sqft", valByKey_(m, headerMap, headers, "sqft"));
  put("employees", valByKey_(m, headerMap, headers, "employees"));
  put("revenue", valByKey_(m, headerMap, headers, "revenue"));
  put("years", valByKey_(m, headerMap, headers, "years"));
  put("equipment", valByKey_(m, headerMap, headers, "equipment"));
  put("cnc3", valByKey_(m, headerMap, headers, "cnc3"));
  put("cnc5", valByKey_(m, headerMap, headers, "cnc5"));
  put("spares", valByKey_(m, headerMap, headers, "spares"));
  put("family", valByKey_(m, headerMap, headers, "family"));
  put("addr2", valByKey_(m, headerMap, headers, "addr2"));
  put("notes", valByKey_(m, headerMap, headers, "notes"));
  put("rownum", String(rec._row));
  const nrow = labelMap["notes"];
  if (nrow) {
    view.getRange(nrow, 2).setWrap(true);
    view.setRowHeight(nrow, 60);
  }
}
function renderNewsToView_(newsRec) {
  const recSh = getNewsRecordSheet_();
  const headers = newsRec.headers;
  const values = newsRec.values;
  const map = detectNewsMapping_(headers);
  const labels = map.order
    .filter((h) => h !== map.summaryHeader)
    .concat([map.summaryHeader]);
  const maxRow = Math.max(recSh.getLastRow(), labels.length);
  if (maxRow > 0) recSh.getRange(1, 2, maxRow, 1).clearContent();
  const rowMap = {};
  headers.forEach((h, i) => (rowMap[String(h)] = values[i]));
  labels.forEach((label, idx) => {
    const r = idx + 1;
    const val = safeStr_(rowMap[label]);
    recSh.getRange(r, 2).setValue(val);
  });
  const sumRow = labels.indexOf(map.summaryHeader) + 1;
  if (sumRow > 0) {
    recSh.getRange(sumRow, 2).setWrap(true);
    recSh.setRowHeight(sumRow, 120);
  }
}

/* ====== News mapping + row access ====== */
function detectNewsMapping_(headers) {
  const norm = headers.map(normHeader_);
  let summaryIdx = -1;
  for (let i = 0; i < headers.length; i++) {
    const h = norm[i];
    if (NEWS_SUMMARY_ALIASES.some((a) => h === a || h.includes(a))) {
      summaryIdx = i;
      break;
    }
  }
  if (summaryIdx < 0) summaryIdx = headers.length - 1;
  const summaryHeader = String(headers[summaryIdx]);
  const prefs = [
    "Company Name",
    "Date",
    "Headline",
    "Title",
    "Source",
    "Outlet",
    "Publisher",
    "URL",
    "Link",
    "Author",
    "Type",
    "Category",
  ];
  const presentPref = [],
    used = new Set();
  prefs.forEach((p) => {
    const idx = headers.findIndex((h) => normHeader_(h) === normHeader_(p));
    if (idx >= 0 && idx !== summaryIdx && !used.has(idx)) {
      presentPref.push(String(headers[idx]));
      used.add(idx);
    }
  });
  headers.forEach((h, i) => {
    if (i === summaryIdx) return;
    if (!used.has(i)) {
      presentPref.push(String(h));
      used.add(i);
    }
  });
  return { order: presentPref.concat([summaryHeader]), summaryHeader };
}

/* First visible news row in News List (formula view). */
function findFirstNewsListRow_() {
  const sh = getNewsListSheet_();
  const last = sh.getLastRow();
  if (last <= HEADER_ROW) return 0;
  const vals = sh.getRange(HEADER_ROW + 1, 1, last - HEADER_ROW, 1).getValues();
  for (let i = 0; i < vals.length; i++)
    if (safeStr_(vals[i][0])) return HEADER_ROW + 1 + i;
  return 0;
}
function getNewsRowFromNewsListRow_(rowNum) {
  const sh = getNewsListSheet_();
  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0];
  const vals = sh.getRange(rowNum, 1, 1, lastCol).getValues()[0];
  return { row: rowNum, headers, values: vals };
}

/* ====== Low-level helpers ====== */
function getViewSheet_() {
  return mustGetSheet_(VIEW_SHEET);
}
function getDataSheet_() {
  return mustGetSheet_(SHEET_NAME);
}
function getNewsListSheet_() {
  return mustGetSheet_(NEWS_LIST);
}
function getNewsRecordSheet_() {
  return mustGetSheet_(NEWS_RECORD);
}
function mustGetSheet_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error("Missing tab: " + name);
  return sh;
}
function getTable_(sheetName) {
  const sh = SpreadsheetApp.getActive().getSheetByName(sheetName);
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow <= HEADER_ROW) return { headers: [], data: [] };
  const headers = sh.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0];
  const data = sh
    .getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, lastCol)
    .getValues();
  return { headers, data };
}
function getActiveCompanyRecord_() {
  const sh = getDataSheet_();
  const r = sh.getActiveCell() ? sh.getActiveCell().getRow() : 0;
  if (r <= HEADER_ROW) return null;
  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0];
  const vals = sh.getRange(r, 1, 1, lastCol).getValues()[0];
  return { _row: r, _map: arrayRowToMap_(headers, vals) };
}
function getCompanyRecordAtRow_(row) {
  const sh = getDataSheet_();
  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0];
  const vals = sh.getRange(row, 1, 1, lastCol).getValues()[0];
  return { _row: row, _map: arrayRowToMap_(headers, vals) };
}
function getCompanyNameAtRow_(row) {
  const sh = getDataSheet_();
  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0];
  const vals = sh.getRange(row, 1, 1, lastCol).getValues()[0];
  const cIdx = getHeaderIndexSmart_(headers, H.company);
  return cIdx >= 0 ? safeStr_(vals[cIdx]) : "";
}
function arrayRowToMap_(headers, vals) {
  const m = {};
  headers.forEach((h, i) => (m[String(h)] = vals[i]));
  return m;
}
function selectDataRow_(rowNum) {
  const sh = getDataSheet_();
  sh.activate();
  sh.setCurrentCell(sh.getRange(rowNum, 1));
  SpreadsheetApp.flush();
}

/* Label + header helpers */
function buildLabelRowMap_(viewSh) {
  const last = Math.max(1, viewSh.getLastRow());
  const colA = viewSh
    .getRange(1, 1, last, 1)
    .getValues()
    .map((r) => String(r[0] || "").trim());
  const normA = colA.map(normHeader_);
  const map = {};
  Object.keys(L).forEach((key) => {
    const targets = L[key].map(normHeader_);
    let found = 0;
    for (let i = 0; i < normA.length; i++) {
      if (!normA[i]) continue;
      if (targets.includes(normA[i])) {
        found = i + 1;
        break;
      }
    }
    if (found) map[key] = found;
  });
  return map;
}
function valByKey_(m, headerMap, headers, key) {
  const idx = headerMap[key];
  if (idx == null || idx < 0) return "";
  const actualHeader = headers[idx];
  return safeStr_(m[actualHeader]);
}

/* Matching / strings */
function getHeaderIndexSmart_(headers, name) {
  const canon = normHeader_(name);
  for (let i = 0; i < headers.length; i++)
    if (normHeader_(headers[i]) === canon) return i;
  const keys = canon.split(" ").filter(Boolean);
  for (let i = 0; i < headers.length; i++) {
    const h = normHeader_(headers[i]);
    if (keys.every((k) => h.includes(k))) return i;
  }
  return -1;
}
function normHeader_(v) {
  return safeStr_(v)
    .toLowerCase()
    .normalize("NFKC")
    .replace(/[\u2010-\u2015]/g, "-")
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}
function safeStr_(v) {
  return v == null ? "" : String(v);
}

/* Data write */
function saveNotesToData_(rowNum, notesText) {
  const dataSh = getDataSheet_();
  const lastCol = dataSh.getLastColumn();
  const headers = dataSh.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0];
  const nIdx = getHeaderIndexSmart_(headers, H.notes);
  if (nIdx === -1) {
    SpreadsheetApp.getUi().alert("Notes column not found: " + H.notes);
    return;
  }
  dataSh.getRange(Number(rowNum), nIdx + 1).setValue(notesText);
}

/* Cached state */
function setLastCompany_(name) {
  DP().setProperty(KEY_LAST_COMPANY, String(name));
}
function getLastCompany_() {
  return DP().getProperty(KEY_LAST_COMPANY) || "";
}
function setLastCompanyRow_(row) {
  DP().setProperty(KEY_LAST_COMPANY_ROW, String(row));
}
function getLastCompanyRow_() {
  return Number(DP().getProperty(KEY_LAST_COMPANY_ROW) || "0");
}
function setLastNewsRow_(row) {
  DP().setProperty(KEY_LAST_NEWSROW, String(row));
}
function getLastNewsRow_() {
  return Number(DP().getProperty(KEY_LAST_NEWSROW) || "0");
}
function clearLastNewsRow_() {
  DP().deleteProperty(KEY_LAST_NEWSROW);
}

/* Utility */
function clearNewsRecordValues_() {
  const recSh = getNewsRecordSheet_();
  const last = recSh.getLastRow();
  if (last > 0) recSh.getRange(1, 2, last, 1).clearContent();
}
function centerDataRowViewport_(row) {
  const sh = getDataSheet_();
  const winHalf = 20;
  const startRow = Math.max(HEADER_ROW + 1, row - winHalf);
  const endRow = Math.max(row + winHalf, row + 1);
  const height = Math.min(sh.getMaxRows(), endRow) - startRow + 1;
  sh.activate();
  sh.setActiveRange(sh.getRange(startRow, 1, height, 1));
  SpreadsheetApp.flush();
  Utilities.sleep(30);
  sh.setCurrentCell(sh.getRange(row, 1));
  SpreadsheetApp.flush();
}

/* ===================== DEDUPERS ===================== */
function prettyDupMsg_(sheetName, d) {
  if (!d || !d.count) return `${sheetName} duplicates removed (0)`;
  const parts = (d.pairs || [])
    .slice(0, 5)
    .map((p) => `r${p.row}→keep r${p.keepRow} ${p.label}=${p.key}`);
  return (
    `${sheetName} duplicates removed (${d.count}): ${parts.join("; ")}` +
    (d.count > 5 ? "; …" : "")
  );
}

/* MMCrawl duplicate remover (by domain/homepage) */
function mmcrawlRemoveDuplicateUrls(returnDiagnostics) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!sh) return returnDiagnostics ? { count: 0 } : undefined;
  const lastRow = sh.getLastRow(),
    lastCol = sh.getLastColumn();
  if (lastRow <= HEADER_ROW)
    return returnDiagnostics ? { count: 0 } : undefined;
  const headers = sh
    .getRange(HEADER_ROW, 1, 1, lastCol)
    .getValues()[0]
    .map(String);
  const data = sh
    .getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, lastCol)
    .getValues();

  function normHeader(v) {
    return String(v || "")
      .toLowerCase()
      .normalize("NFKC")
      .replace(/[\u2010-\u2015]/g, "-")
      .replace(/[^a-z0-9]+/g, " ")
      .trim();
  }
  function hIdx(name) {
    const target = normHeader(name);
    let idx = headers.findIndex((h) => normHeader(h) === target);
    if (idx >= 0) return idx;
    const parts = target.split(" ").filter(Boolean);
    idx = headers.findIndex((h) => {
      const hh = normHeader(h);
      return parts.every((p) => hh.includes(p));
    });
    return idx;
  }

  const idxDomain = hIdx("Domain from URL");
  const idxUrl = hIdx("Public Website Homepage URL");
  if (idxDomain < 0 && idxUrl < 0)
    return returnDiagnostics ? { count: 0 } : undefined;

  function normalizeCompanyKey(raw) {
    let s = String(raw || "")
      .trim()
      .toLowerCase();
    if (!s) return "";
    try {
      const urlStr = /^(https?:)?\/\//.test(s) ? s : "http://" + s;
      const u = new URL(urlStr);
      s = u.hostname || s;
    } catch (_) {}
    s = s
      .replace(/^https?:\/\//, "")
      .replace(/^www\./, "")
      .replace(/\/+$/, "")
      .split("/")[0];
    return s;
  }

  const seen = new Map(),
    rowsToDelete = [],
    pairs = [];
  for (let i = 0; i < data.length; i++) {
    const rowVals = data[i];
    const keyRaw =
      (idxDomain >= 0 ? rowVals[idxDomain] : "") ||
      (idxUrl >= 0 ? rowVals[idxUrl] : "");
    const key = normalizeCompanyKey(keyRaw);
    if (!key) continue;
    const rowNum = HEADER_ROW + 1 + i;
    if (seen.has(key)) {
      rowsToDelete.push(rowNum);
      pairs.push({ row: rowNum, keepRow: seen.get(key), key, label: "host" });
    } else {
      seen.set(key, rowNum);
    }
  }
  if (!rowsToDelete.length) return returnDiagnostics ? { count: 0 } : undefined;
  rowsToDelete.sort((a, b) => b - a).forEach((r) => sh.deleteRow(r));
  return returnDiagnostics
    ? { count: rowsToDelete.length, rows: rowsToDelete.slice(), pairs }
    : undefined;
}

/* News Raw duplicate remover (by "News Story URL") */
function newsRawRemoveDuplicateStories(returnDiagnostics) {
  const sh = SpreadsheetApp.getActive().getSheetByName(NEWS_RAW);
  if (!sh) return returnDiagnostics ? { count: 0 } : undefined;
  const lastRow = sh.getLastRow(),
    lastCol = sh.getLastColumn();
  if (lastRow <= HEADER_ROW)
    return returnDiagnostics ? { count: 0 } : undefined;
  const headers = sh
    .getRange(HEADER_ROW, 1, 1, lastCol)
    .getValues()[0]
    .map(String);
  const data = sh
    .getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, lastCol)
    .getValues();

  function normHeader(v) {
    return String(v || "")
      .toLowerCase()
      .normalize("NFKC")
      .replace(/[\u2010-\u2015]/g, "-")
      .replace(/[^a-z0-9]+/g, " ")
      .trim();
  }
  function colIdx(name) {
    const target = normHeader(name);
    let idx = headers.findIndex((h) => normHeader(h) === target);
    if (idx >= 0) return idx;
    const parts = target.split(" ").filter(Boolean);
    idx = headers.findIndex((h) => {
      const hh = normHeader(h);
      return parts.every((p) => hh.includes(p));
    });
    return idx;
  }

  const urlIdx = colIdx("News Story URL");
  if (urlIdx < 0) return returnDiagnostics ? { count: 0 } : undefined;

  function normalizeNewsUrl(raw) {
    let s = String(raw || "").trim();
    if (!s) return "";
    const lc = s.toLowerCase();
    if (lc === "no news" || lc === "none" || lc === "n/a") return "";
    try {
      const urlStr = /^(https?:)?\/\//i.test(s) ? s : "http://" + s;
      const u = new URL(urlStr);
      const toDelete = [
        "gclid",
        "fbclid",
        "mc_cid",
        "mc_eid",
        "igshid",
        "utm_source",
        "utm_medium",
        "utm_campaign",
        "utm_term",
        "utm_content",
        "utm_id",
      ];
      toDelete.forEach((k) => u.searchParams.delete(k));
      u.hash = "";
      let host = u.hostname.toLowerCase().replace(/^www\./, "");
      let path = u.pathname.replace(/\/+$/, "");
      if (path === "") path = "/";
      return host + path + (u.search ? "?" + u.searchParams.toString() : "");
    } catch (_) {
      return s
        .replace(/^https?:\/\//i, "")
        .replace(/^www\./i, "")
        .replace(/\/+$/, "");
    }
  }

  const seen = new Map(),
    rowsToDelete = [],
    pairs = [];
  for (let i = 0; i < data.length; i++) {
    const raw = data[i][urlIdx];
    const key = normalizeNewsUrl(raw);
    if (!key) continue;
    const rowNum = HEADER_ROW + 1 + i;
    if (seen.has(key)) {
      rowsToDelete.push(rowNum);
      pairs.push({ row: rowNum, keepRow: seen.get(key), key, label: "url" });
    } else {
      seen.set(key, rowNum);
    }
  }
  if (!rowsToDelete.length) return returnDiagnostics ? { count: 0 } : undefined;
  rowsToDelete.sort((a, b) => b - a).forEach((r) => sh.deleteRow(r));
  return returnDiagnostics
    ? { count: rowsToDelete.length, rows: rowsToDelete.slice(), pairs }
    : undefined;
}
