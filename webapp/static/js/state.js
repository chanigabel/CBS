/**
 * Frontend state shared by the page scripts.
 */

// ---------------------------------------------------------------------------
// Application State
// ---------------------------------------------------------------------------

// Sessions keyed by session ID.
const sessions = new Map();

const state = {
    sessionId: null,
    currentSheet: null,
    sheetData: null,
    selectedRows: new Set(),
    // columnFilters: Map<colName, Set<string>> — active value filters per column
    columnFilters: new Map(),
};

Object.assign(window, { sessions, state });
