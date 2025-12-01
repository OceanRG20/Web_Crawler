/******************************************************
 * AutoFilterCount.gs
 * Auto-popup (toast) with count of visible rows on MMCrawl
 * whenever the filter is changed.
 *
 * HOW IT WORKS
 * - onOpen() ensures an installable onChange trigger exists.
 * - FT_onChange() fires on any sheet change; if the active
 *   sheet is "MMCrawl" and a filter is present, it computes
 *   visible rows and shows a toast.
 ******************************************************/

var FT_CFG = {
  SHEET_NAME: "MMCrawl", // watch this sheet
  NONEMPTY_COL: 1, // use column A (Company Name) to decide if a row is a data row
};

function onOpen() {
  // Ensure the onChange trigger exists (idempotent)
  try {
    const id = SpreadsheetApp.getActive().getId();
    const exists = ScriptApp.getProjectTriggers().some(
      (t) => t.getHandlerFunction && t.getHandlerFunction() === "FT_onChange"
    );
    if (!exists) {
      ScriptApp.newTrigger("FT_onChange")
        .forSpreadsheet(id)
        .onChange()
        .create();
    }
  } catch (_) {}
}

/**
 * Installable onChange trigger handler.
 * Shows a toast with "Visible rows: X / Y (hidden: Z)".
 * If none visible: "No rows match current filter."
 */
function FT_onChange(e) {
  try {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getActiveSheet();
    if (!sh || sh.getName() !== FT_CFG.SHEET_NAME) return;

    // Only react when a filter exists (basic filter or a filter view)
    const hasBasic = !!sh.getFilter();
    const hasView = sh.getFilterViews && sh.getFilterViews().length > 0;
    if (!hasBasic && !hasView) return;

    const stats = FT_computeVisibleStats_(sh, FT_CFG.NONEMPTY_COL);
    const msg =
      stats.visible === 0
        ? "No rows match the current filter."
        : `Visible rows: ${stats.visible} / ${stats.total} (hidden: ${stats.hidden})`;
    ss.toast(msg, "Filtered Count", 6);
  } catch (err) {
    try {
      SpreadsheetApp.getActive().toast(String(err), "Filtered Count", 6);
    } catch (_) {}
  }
}

/* ===== Helper: count visible (non-empty) rows ===== */
function FT_computeVisibleStats_(sh, nonEmptyColIndex) {
  const HEADER_ROW = 1;
  const lastRow = sh.getLastRow();
  if (lastRow <= HEADER_ROW) return { total: 0, visible: 0, hidden: 0 };

  const col = Math.max(1, nonEmptyColIndex || 1);
  const values = sh
    .getRange(HEADER_ROW + 1, col, lastRow - HEADER_ROW, 1)
    .getDisplayValues();

  let total = 0,
    visible = 0,
    hidden = 0;
  for (let r = HEADER_ROW + 1; r <= lastRow; r++) {
    const hasData =
      String(values[r - HEADER_ROW - 1][0] || "").trim().length > 0;
    if (!hasData) continue; // ignore blank rows
    total++;
    const isHidden =
      (sh.isRowHiddenByFilter && sh.isRowHiddenByFilter(r)) || false;
    if (isHidden) hidden++;
    else visible++;
  }
  return { total, visible, hidden };
}
