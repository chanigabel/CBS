/**
 * Excel standardization Web App - Frontend JavaScript
 * Vanilla JS, no external dependencies, fully offline-capable.
 */

// ---------------------------------------------------------------------------
// Application State
// ---------------------------------------------------------------------------

// sessions: Map<sessionId, SessionMeta>
// SessionMeta: { sessionId, filename, sheetNames, lastSheet, isNormalized,
//               hasEdits, sheetStats }
// Each session keeps its own last-viewed sheet and standardization flag so
// switching between files restores the correct view.
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
