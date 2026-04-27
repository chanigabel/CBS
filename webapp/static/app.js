/**
 * Excel Normalization Web App - Frontend JavaScript
 * Vanilla JS, no external dependencies, fully offline-capable.
 */

// ---------------------------------------------------------------------------
// Application State
// ---------------------------------------------------------------------------

// sessions: Map<sessionId, SessionMeta>
// SessionMeta: { sessionId, filename, sheetNames, lastSheet, isNormalized,
//               hasEdits, sheetStats }
// Each session keeps its own last-viewed sheet and normalization flag so
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

// ---------------------------------------------------------------------------
// Error Handling
// ---------------------------------------------------------------------------

function showError(message) {
    const banner = document.getElementById('error-banner');
    document.getElementById('error-message').textContent = message;
    banner.classList.remove('hidden');
}

function dismissError() {
    document.getElementById('error-banner').classList.add('hidden');
}

// ---------------------------------------------------------------------------
// API Helpers
// ---------------------------------------------------------------------------

async function apiCall(method, url, body = null) {
    const options = { method, headers: {} };
    if (body && !(body instanceof FormData)) {
        options.headers['Content-Type'] = 'application/json';
        options.body = JSON.stringify(body);
    } else if (body instanceof FormData) {
        options.body = body;
    }

    const response = await fetch(url, options);

    if (!response.ok) {
        let detail = `HTTP ${response.status}`;
        try { const err = await response.json(); detail = err.detail || detail; } catch (_) {}
        throw new Error(detail);
    }

    const ct = response.headers.get('content-type') || '';
    if (ct.includes('application/zip') || ct.includes('application/vnd.openxmlformats')) {
        return response;
    }
    return response.json();
}

// ---------------------------------------------------------------------------
// Upload Flow — U-10: XHR with progress feedback
// ---------------------------------------------------------------------------

/**
 * U-10: Upload a single file using XMLHttpRequest so we can report progress.
 * Returns a Promise that resolves with the parsed JSON response.
 */
function uploadWithProgress(file, onProgress) {
    return new Promise((resolve, reject) => {
        const xhr = new XMLHttpRequest();
        const fd = new FormData();
        fd.append('file', file);

        xhr.upload.addEventListener('progress', e => {
            if (e.lengthComputable) {
                onProgress(Math.round((e.loaded / e.total) * 100));
            }
        });

        xhr.addEventListener('load', () => {
            if (xhr.status >= 200 && xhr.status < 300) {
                try {
                    resolve(JSON.parse(xhr.responseText));
                } catch (_) {
                    reject(new Error('Invalid server response'));
                }
            } else {
                let detail = `HTTP ${xhr.status}`;
                try { detail = JSON.parse(xhr.responseText).detail || detail; } catch (_) {}
                reject(new Error(detail));
            }
        });

        xhr.addEventListener('error', () => reject(new Error('Network error during upload')));
        xhr.addEventListener('abort', () => reject(new Error('Upload cancelled')));

        xhr.open('POST', '/api/upload');
        xhr.send(fd);
    });
}

async function handleUpload(event) {
    event.preventDefault();
    dismissError();

    const fileInput = document.getElementById('file-input');
    const files = Array.from(fileInput.files);
    if (!files.length) return;

    const uploadBtn = document.getElementById('upload-btn');
    const statusDiv = document.getElementById('upload-status');

    uploadBtn.disabled = true;

    let successCount = 0;
    const errors = [];

    for (const file of files) {
        try {
            // U-10: Show per-file upload progress
            statusDiv.textContent = `Uploading ${file.name}: 0%`;
            const data = await uploadWithProgress(file, pct => {
                statusDiv.textContent = `Uploading ${file.name}: ${pct}%`;
            });
            sessions.set(data.session_id, {
                sessionId: data.session_id,
                filename: file.name,
                sheetNames: data.sheet_names,
                lastSheet: data.sheet_names[0] || null,
                isNormalized: false,
                hasEdits: false,
                sheetStats: {},
            });
            successCount++;
        } catch (err) {
            errors.push(`${file.name}: ${err.message}`);
        }
    }

    if (errors.length) showError(`Some uploads failed:\n${errors.join('\n')}`);

    if (successCount === 0) {
        statusDiv.textContent = '';
        uploadBtn.disabled = false;
        return;
    }

    statusDiv.textContent = `${successCount} file(s) uploaded.`;
    renderSessionSwitcher();

    const lastSession = [...sessions.values()].at(-1);
    await activateSession(lastSession.sessionId);

    uploadBtn.disabled = false;
}

// ---------------------------------------------------------------------------
// Session Switcher — U-05: sheet stats badges on file tabs
// ---------------------------------------------------------------------------

function renderSessionSwitcher() {
    const switcher = document.getElementById('session-switcher');
    const tabs = document.getElementById('session-tabs');

    if (sessions.size === 0) {
        switcher.classList.add('hidden');
        return;
    }

    tabs.innerHTML = '';
    sessions.forEach(({ sessionId, filename, isNormalized, sheetStats }) => {
        const btn = document.createElement('button');
        btn.className = 'sheet-tab file-tab';
        btn.dataset.sessionId = sessionId;

        // U-05: Show a warning badge if any sheet has < 100% success rate
        const hasWarning = sheetStats && Object.values(sheetStats).some(s => s.success_rate < 1.0);
        let label = filename;
        if (isNormalized) label += hasWarning ? ' ⚠' : ' ✓';
        btn.textContent = label;
        btn.title = filename;

        if (sessionId === state.sessionId) btn.classList.add('active');
        btn.onclick = () => activateSession(sessionId);
        tabs.appendChild(btn);
    });

    // Bulk-export controls (commented out in original, preserved)
    let bulkBar = document.getElementById('bulk-export-bar');
    if (!bulkBar) {
        bulkBar = document.createElement('div');
        bulkBar.id = 'bulk-export-bar';
        bulkBar.className = 'bulk-export-bar';
        switcher.appendChild(bulkBar);
    }
    bulkBar.innerHTML = '';

    switcher.classList.remove('hidden');
}

function _highlightActiveSession() {
    document.querySelectorAll('#session-tabs .file-tab').forEach(btn => {
        btn.classList.toggle('active', btn.dataset.sessionId === state.sessionId);
    });
}

// ---------------------------------------------------------------------------
// Session Activation — preserves per-session state
// ---------------------------------------------------------------------------

async function activateSession(sessionId) {
    const session = sessions.get(sessionId);
    if (!session) return;

    // Save current sheet back to the outgoing session before switching
    if (state.sessionId && state.sessionId !== sessionId) {
        const outgoing = sessions.get(state.sessionId);
        if (outgoing) outgoing.lastSheet = state.currentSheet;
    }

    state.sessionId = sessionId;
    state.currentSheet = null;
    state.sheetData = null;
    state.selectedRows.clear();

    _highlightActiveSession();

    renderSheetSelector(session.sheetNames, session.sheetStats || {});
    document.getElementById('action-bar').classList.remove('hidden');

    document.getElementById('grid-section').classList.add('hidden');
    document.getElementById('grid-container').innerHTML = '';

    // Load institution metadata for this session
    await loadInstitution();

    // Restore the last sheet this session had open, or default to first sheet
    const sheetToLoad = session.lastSheet || session.sheetNames[0];
    if (sheetToLoad) {
        await loadSheet(sheetToLoad);
    }
}

// ---------------------------------------------------------------------------
// Sheet Selector — U-05: show per-sheet success rate badge
// ---------------------------------------------------------------------------

function renderSheetSelector(sheetNames, sheetStats) {
    const selector = document.getElementById('sheet-selector');
    const tabs = document.getElementById('sheet-tabs');

    tabs.innerHTML = '';
    sheetNames.forEach(name => {
        const btn = document.createElement('button');
        btn.className = 'sheet-tab';
        btn.setAttribute('role', 'tab');
        btn.onclick = () => loadSheet(name);

        // U-05: Annotate sheet tab with success rate when available
        const stat = (sheetStats || {})[name];
        if (stat && stat.success_rate < 1.0) {
            btn.textContent = `${name} ⚠ ${Math.round(stat.success_rate * 100)}%`;
            btn.title = `${stat.rows} rows — ${Math.round(stat.success_rate * 100)}% normalized successfully`;
            btn.classList.add('sheet-tab-warning');
        } else {
            btn.textContent = name;
        }

        tabs.appendChild(btn);
    });

    selector.classList.remove('hidden');
}

function setActiveSheetTab(sheetName) {
    document.querySelectorAll('#sheet-tabs .sheet-tab').forEach(btn => {
        // Match by the base name (strip any appended stats badge)
        btn.classList.toggle('active', btn.textContent.startsWith(sheetName));
    });
}

// ---------------------------------------------------------------------------
// Sheet Data Loading
// ---------------------------------------------------------------------------

async function loadSheet(sheetName) {
    if (!state.sessionId) return;
    dismissError();

    state.currentSheet = sheetName;
    state.selectedRows.clear();
    state.columnFilters.clear();
    setActiveSheetTab(sheetName);

    // Persist last-viewed sheet on the session record
    const session = sessions.get(state.sessionId);
    if (session) session.lastSheet = sheetName;

    const gridSection = document.getElementById('grid-section');
    const gridTitle = document.getElementById('grid-title');
    const gridContainer = document.getElementById('grid-container');

    gridTitle.textContent = sheetName;
    gridContainer.innerHTML = '<div style="padding:20px;text-align:center">Loading... <span class="loading"></span></div>';
    gridSection.classList.remove('hidden');

    try {
        const data = await apiCall('GET', `/api/workbook/${state.sessionId}/sheet/${encodeURIComponent(sheetName)}`);
        state.sheetData = data;
        renderGrid(data, getFilteredRows(data.rows));
    } catch (err) {
        showError(`Failed to load sheet '${sheetName}': ${err.message}`);
        gridContainer.innerHTML = '';
    }
}

// ---------------------------------------------------------------------------
// Column Filtering — multi-select, value-driven, multi-column AND logic
// ---------------------------------------------------------------------------

/**
 * Return the subset of rows that pass ALL active column filters.
 * If no filters are active, returns the full rows array.
 */
function getFilteredRows(rows) {
    if (state.columnFilters.size === 0) return rows;
    return rows.filter(row => {
        for (const [col, values] of state.columnFilters) {
            if (values.size === 0) continue;
            const cell = row[col];
            const cellStr = (cell !== null && cell !== undefined) ? String(cell).trim() : '';
            if (!values.has(cellStr)) return false;
        }
        return true;
    });
}

/**
 * Get sorted distinct string values for a column across all (unfiltered) rows.
 */
function getDistinctValues(col) {
    if (!state.sheetData) return [];
    const seen = new Set();
    state.sheetData.rows.forEach(row => {
        const v = row[col];
        seen.add((v !== null && v !== undefined) ? String(v).trim() : '');
    });
    return [...seen].sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
}

/**
 * Open (or close) the filter dropdown for a column.
 * Closes any other open dropdown first.
 */
function openFilterDropdown(col, btnEl) {
    // Close any existing open dropdown
    const existing = document.querySelector('.col-filter-dropdown');
    if (existing) {
        const wasForSameCol = existing.dataset.col === col;
        existing.remove();
        if (wasForSameCol) return;   // toggle: clicking again closes it
    }

    const values = getDistinctValues(col);
    const activeSet = state.columnFilters.get(col) || new Set();

    const dropdown = document.createElement('div');
    dropdown.className = 'col-filter-dropdown';
    dropdown.dataset.col = col;

    // Search box (only shown when there are many values)
    if (values.length > 8) {
        const search = document.createElement('input');
        search.type = 'text';
        search.placeholder = 'חיפוש...';
        search.className = 'col-filter-search';
        search.addEventListener('input', () => {
            const q = search.value.toLowerCase();
            dropdown.querySelectorAll('.col-filter-item').forEach(item => {
                item.style.display = item.dataset.val.toLowerCase().includes(q) ? '' : 'none';
            });
        });
        dropdown.appendChild(search);
    }

    // Value list
    const list = document.createElement('div');
    list.className = 'col-filter-list';

    values.forEach(val => {
        const label = document.createElement('label');
        label.className = 'col-filter-item';
        label.dataset.val = val;

        const cb = document.createElement('input');
        cb.type = 'checkbox';
        cb.checked = activeSet.has(val);
        cb.addEventListener('change', () => {
            let set = state.columnFilters.get(col);
            if (!set) { set = new Set(); state.columnFilters.set(col, set); }
            if (cb.checked) set.add(val); else set.delete(val);
            if (set.size === 0) state.columnFilters.delete(col);
            applyFilters();
            // Update the filter button active state without closing the dropdown
            updateFilterButtonState(col);
        });

        const text = document.createElement('span');
        text.textContent = val === '' ? '(ריק)' : val;
        if (val === '') text.style.fontStyle = 'italic';

        label.appendChild(cb);
        label.appendChild(text);
        list.appendChild(label);
    });
    dropdown.appendChild(list);

    // Footer: clear this column's filter
    const footer = document.createElement('div');
    footer.className = 'col-filter-footer';
    const clearBtn = document.createElement('button');
    clearBtn.textContent = 'נקה סינון';
    clearBtn.className = 'col-filter-clear-btn';
    clearBtn.addEventListener('click', () => {
        state.columnFilters.delete(col);
        applyFilters();
        updateFilterButtonState(col);
        dropdown.querySelectorAll('input[type=checkbox]').forEach(cb => cb.checked = false);
    });
    footer.appendChild(clearBtn);
    dropdown.appendChild(footer);

    // Position below the button
    document.body.appendChild(dropdown);
    const rect = btnEl.getBoundingClientRect();
    dropdown.style.top = (rect.bottom + window.scrollY + 2) + 'px';
    // RTL: align right edge of dropdown to right edge of button
    const dropW = dropdown.offsetWidth;
    dropdown.style.left = (rect.right + window.scrollX - dropW) + 'px';

    // Close on outside click
    function onOutsideClick(e) {
        if (!dropdown.contains(e.target) && e.target !== btnEl) {
            dropdown.remove();
            document.removeEventListener('mousedown', onOutsideClick, true);
        }
    }
    document.addEventListener('mousedown', onOutsideClick, true);
}

function updateFilterButtonState(col) {
    const btn = document.querySelector(`.col-filter-btn[data-col="${CSS.escape(col)}"]`);
    if (!btn) return;
    const active = state.columnFilters.has(col) && state.columnFilters.get(col).size > 0;
    btn.classList.toggle('col-filter-active', active);
    btn.title = active ? 'סינון פעיל — לחץ לעריכה' : 'סנן לפי עמודה זו';
}

/**
 * Re-render only the tbody rows based on current filters.
 * Avoids a full table rebuild for performance.
 */
function applyFilters() {
    if (!state.sheetData) return;
    const filtered = getFilteredRows(state.sheetData.rows);

    // Update stats
    const statsDiv = document.getElementById('grid-stats');
    const total = state.sheetData.rows.length;
    const shown = filtered.length;
    const cols = state.sheetData.field_names.length;
    if (state.columnFilters.size > 0) {
        statsDiv.textContent = `מציג ${shown} מתוך ${total} שורות × ${cols} עמודות`;
    } else {
        statsDiv.textContent = `${total} rows × ${cols} columns`;
    }

    // Re-render the grid with filtered rows
    renderGrid(state.sheetData, filtered);

    // Update "clear all" button visibility
    const clearAllBtn = document.getElementById('clear-all-filters-btn');
    if (clearAllBtn) clearAllBtn.classList.toggle('hidden', state.columnFilters.size === 0);
}

function clearAllFilters() {
    state.columnFilters.clear();
    // Close any open dropdown
    const existing = document.querySelector('.col-filter-dropdown');
    if (existing) existing.remove();
    applyFilters();
}

// ---------------------------------------------------------------------------
// Full-screen grid overlay
// ---------------------------------------------------------------------------

function openGridOverlay() {
    const overlay          = document.getElementById('grid-overlay');
    const overlayContainer = document.getElementById('grid-overlay-container');
    const overlayTitle     = document.getElementById('grid-overlay-title');
    const overlayStats     = document.getElementById('grid-overlay-stats');
    if (!overlay || !overlayContainer || !state.sheetData) return;

    // Copy current title and stats into the overlay bar
    overlayTitle.textContent = document.getElementById('grid-title')?.textContent || '';
    const total = state.sheetData.rows.length;
    const filtered = getFilteredRows(state.sheetData.rows);
    const shown = filtered.length;
    overlayStats.textContent = state.columnFilters.size > 0
        ? `מציג ${shown} מתוך ${total} שורות`
        : `${total} שורות`;

    // Render the full interactive grid into the overlay container
    renderGrid(state.sheetData, filtered, overlayContainer);

    overlay.classList.remove('hidden');
    document.body.classList.add('grid-overlay-open');
    document.addEventListener('keydown', _overlayEscHandler);
}

function closeGridOverlay() {
    const overlay = document.getElementById('grid-overlay');
    if (!overlay) return;
    overlay.classList.add('hidden');
    document.body.classList.remove('grid-overlay-open');
    document.removeEventListener('keydown', _overlayEscHandler);

    // Sync any changes made inside the overlay back to the normal grid
    if (state.sheetData) {
        renderGrid(state.sheetData, getFilteredRows(state.sheetData.rows));
    }
}

function _overlayEscHandler(e) {
    if (e.key === 'Escape') closeGridOverlay();
}

function renderGrid(sheetData, rows, targetContainer) {
    const container = targetContainer || document.getElementById('grid-container');
    const statsDiv  = targetContainer ? null : document.getElementById('grid-stats');

    const displayRows = rows !== undefined ? rows : sheetData.rows;

    if (!sheetData.rows || sheetData.rows.length === 0) {
        container.innerHTML = '<p style="padding:20px;color:#718096">No data rows found in this sheet.</p>';
        if (statsDiv) statsDiv.textContent = '';
        if (!targetContainer) updateDeleteButton();
        return;
    }

    const displayColumns = sheetData.field_names;

    function colClass(col) {
        if (col.endsWith('_corrected')) return 'corrected';
        if (col.endsWith('_status'))   return 'status';
        return 'original';
    }

    const table = document.createElement('table');
    table.className = 'data-grid';

    // Header
    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');

    const thCheck = document.createElement('th');
    thCheck.className = 'col-select';
    const selectAll = document.createElement('input');
    selectAll.type = 'checkbox';
    selectAll.title = 'Select all rows';
    selectAll.addEventListener('change', () => toggleSelectAll(selectAll.checked, displayRows));
    thCheck.appendChild(selectAll);
    headerRow.appendChild(thCheck);

    const thDel = document.createElement('th');
    thDel.className = 'col-del';
    headerRow.appendChild(thDel);

    displayColumns.forEach(col => {
        const th = document.createElement('th');
        const cls = colClass(col);
        if (cls === 'corrected') th.className = 'corrected-header';
        else if (cls === 'status') th.className = 'status-header';

        // Column label + filter button wrapper
        const headerWrap = document.createElement('div');
        headerWrap.className = 'col-header-wrap';

        const label = document.createElement('span');
        label.textContent = col;
        headerWrap.appendChild(label);

        const filterBtn = document.createElement('button');
        filterBtn.className = 'col-filter-btn';
        filterBtn.dataset.col = col;
        const isActive = state.columnFilters.has(col) && state.columnFilters.get(col).size > 0;
        if (isActive) filterBtn.classList.add('col-filter-active');
        filterBtn.title = isActive ? 'סינון פעיל — לחץ לעריכה' : 'סנן לפי עמודה זו';
        filterBtn.textContent = '▾';
        filterBtn.addEventListener('click', e => {
            e.stopPropagation();
            openFilterDropdown(col, filterBtn);
        });
        headerWrap.appendChild(filterBtn);

        th.appendChild(headerWrap);
        headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);
    table.appendChild(thead);

    // Body
    const tbody = document.createElement('tbody');
    displayRows.forEach((row) => {
        const tr = document.createElement('tr');
        tr.dataset.rowUid = row._row_uid;

        const tdCheck = document.createElement('td');
        tdCheck.className = 'col-select';
        const cb = document.createElement('input');
        cb.type = 'checkbox';
        cb.checked = state.selectedRows.has(row._row_uid);
        cb.addEventListener('change', () => {
            if (cb.checked) { state.selectedRows.add(row._row_uid); tr.classList.add('row-selected'); }
            else            { state.selectedRows.delete(row._row_uid); tr.classList.remove('row-selected'); }
            updateDeleteButton();
            const visibleRows = document.querySelectorAll('.data-grid tbody tr').length;
            selectAll.checked = state.selectedRows.size === visibleRows;
            selectAll.indeterminate = state.selectedRows.size > 0 && state.selectedRows.size < visibleRows;
        });
        if (cb.checked) tr.classList.add('row-selected');
        tdCheck.appendChild(cb);
        tr.appendChild(tdCheck);

        // U-07: Count changed corrected fields for this row and show a badge
        const changedFields = displayColumns.filter(col => {
            if (!col.endsWith('_corrected')) return false;
            const origVal = row[col.replace(/_corrected$/, '')];
            const corrVal = row[col];
            if (corrVal === null || corrVal === undefined) return false;
            const origStr = (origVal !== null && origVal !== undefined) ? String(origVal).trim() : '';
            const corrStr = String(corrVal).trim();
            return corrStr !== '' && corrStr !== origStr;
        });

        const tdDel = document.createElement('td');
        tdDel.className = 'col-del';
        const delBtn = document.createElement('button');
        delBtn.className = 'btn-row-delete';
        delBtn.textContent = '✕';
        delBtn.title = 'Delete this row';
        delBtn.addEventListener('click', () => deleteSingleRow(row._row_uid));
        tdDel.appendChild(delBtn);

        // U-07: Add change-count badge when there are corrections
        if (changedFields.length > 0) {
            tr.classList.add('row-has-changes');
            const badge = document.createElement('span');
            badge.className = 'change-badge';
            badge.textContent = changedFields.length;
            const fieldNames = changedFields.map(f => f.replace(/_corrected$/, '')).join(', ');
            badge.title = `${changedFields.length} field(s) corrected: ${fieldNames}`;
            tdDel.appendChild(badge);
        }

        tr.appendChild(tdDel);

        displayColumns.forEach(col => {
            const td = document.createElement('td');
            const value = row[col];
            td.textContent = value !== null && value !== undefined ? String(value) : '';

            const cls = colClass(col);
            if (cls === 'corrected') {
                // U-03: Normalize comparison to avoid false "changed" highlights on
                // type mismatches (e.g. gender original="ז" str vs corrected=1 int).
                const origVal = row[col.replace(/_corrected$/, '')];
                const origStr = (origVal !== null && origVal !== undefined) ? String(origVal).trim() : '';
                const corrStr = (value !== null && value !== undefined) ? String(value).trim() : '';
                td.className = (corrStr !== '' && corrStr !== origStr)
                    ? 'corrected-changed' : 'corrected-cell';
            } else if (cls === 'status') {
                // U-02: Visually distinguish error status cells from empty ones
                const statusText = String(value || '').trim();
                td.className = statusText !== '' ? 'status-cell status-error' : 'status-cell status-ok';
            }

            if (cls !== 'status') td.addEventListener('click', () => makeEditable(td, row._row_uid, col));
            tr.appendChild(td);
        });

        tbody.appendChild(tr);
    });

    table.appendChild(tbody);
    container.innerHTML = '';
    container.appendChild(table);

    const total = sheetData.rows.length;
    const shown = displayRows.length;
    if (statsDiv) {
        if (state.columnFilters.size > 0) {
            statsDiv.textContent = `מציג ${shown} מתוך ${total} שורות × ${displayColumns.length} עמודות`;
        } else {
            statsDiv.textContent = `${total} rows × ${displayColumns.length} columns`;
        }
    }
    if (!targetContainer) updateDeleteButton();
}

// ---------------------------------------------------------------------------
// Row Selection
// ---------------------------------------------------------------------------

function toggleSelectAll(checked, displayRows) {
    state.selectedRows.clear();
    document.querySelectorAll('.data-grid tbody tr').forEach((tr) => {
        const cb = tr.querySelector('input[type=checkbox]');
        if (!cb) return;
        cb.checked = checked;
        const rowUid = tr.dataset.rowUid;
        if (checked) { state.selectedRows.add(rowUid); tr.classList.add('row-selected'); }
        else         { tr.classList.remove('row-selected'); }
    });
    updateDeleteButton();
}

function updateDeleteButton() {
    const btn = document.getElementById('delete-rows-btn');
    if (!btn) return;
    const n = state.selectedRows.size;
    btn.disabled = n === 0;
    btn.textContent = n > 0 ? `🗑 Delete ${n} row${n > 1 ? 's' : ''}` : '🗑 Delete rows';
}

// ---------------------------------------------------------------------------
// Row Deletion — U-06: confirm before bulk delete
// ---------------------------------------------------------------------------

async function deleteSingleRow(rowUid) { await _deleteRows([rowUid]); }

async function deleteSelectedRows() {
    if (state.selectedRows.size === 0) return;
    const n = state.selectedRows.size;
    // U-06: Ask for confirmation when deleting more than one row
    if (n > 1) {
        const confirmed = confirm(
            `Delete ${n} rows?\n\nThis cannot be undone without re-running normalization.`
        );
        if (!confirmed) return;
    }
    await _deleteRows([...state.selectedRows]);
}

async function _deleteRows(rowUids) {
    if (!state.sessionId || !rowUids.length) return;
    dismissError();

    try {
        const result = await apiCall(
            'DELETE',
            `/api/workbook/${state.sessionId}/sheet/${encodeURIComponent(state.currentSheet)}/rows`,
            { row_uids: rowUids }
        );

        const uidSet = new Set(rowUids);
        state.sheetData.rows = state.sheetData.rows.filter(r => !uidSet.has(r._row_uid));

        // Remove deleted UIDs from selectedRows
        uidSet.forEach(uid => state.selectedRows.delete(uid));

        const filtered = getFilteredRows(state.sheetData.rows);
        renderGrid(state.sheetData, filtered);
        document.getElementById('grid-stats').textContent =
            `Deleted ${result.deleted_count} row(s). ${result.remaining_rows} rows remaining.`;
    } catch (err) {
        showError(`Delete failed: ${err.message}`);
        updateDeleteButton();
    }
}

// ---------------------------------------------------------------------------
// Inline Cell Editing
// ---------------------------------------------------------------------------

function makeEditable(td, rowUid, fieldName) {
    if (td.querySelector('input[type=text]')) return;

    const currentValue = td.textContent;
    td.className = (td.className || '') + ' editing';
    td.textContent = '';

    const input = document.createElement('input');
    input.type = 'text';
    input.value = currentValue;
    td.appendChild(input);
    input.focus();
    input.select();

    async function commitEdit() {
        const newValue = input.value;
        if (newValue === currentValue) {
            td.textContent = currentValue;
            td.className = td.className.replace(' editing', '');
            return;
        }
        try {
            await apiCall(
                'PATCH',
                `/api/workbook/${state.sessionId}/sheet/${encodeURIComponent(state.currentSheet)}/cell`,
                { row_uid: rowUid, field_name: fieldName, new_value: newValue }
            );
            // Find the row in sheetData.rows by matching _row_uid
            const editedRow = state.sheetData?.rows.find(r => r._row_uid === rowUid);
            if (editedRow) editedRow[fieldName] = newValue;
            td.textContent = newValue;
            td.className = td.className.replace(' editing', '');
            if (fieldName.endsWith('_corrected')) {
                // U-03: Use string comparison to avoid false highlights
                const origVal = editedRow ? editedRow[fieldName.replace(/_corrected$/, '')] : null;
                const origStr = (origVal !== null && origVal !== undefined) ? String(origVal).trim() : '';
                td.className = (newValue.trim() !== '' && newValue.trim() !== origStr)
                    ? 'corrected-changed' : 'corrected-cell';
            }
            // Mark session as having unsaved edits (U-01)
            const session = sessions.get(state.sessionId);
            if (session) session.hasEdits = true;
        } catch (err) {
            showError(`Edit failed: ${err.message}`);
            td.textContent = currentValue;
            td.className = td.className.replace(' editing', '');
        }
    }

    input.addEventListener('blur', commitEdit);
    input.addEventListener('keydown', e => {
        if (e.key === 'Enter') { e.preventDefault(); input.blur(); }
        else if (e.key === 'Escape') {
            td.textContent = currentValue;
            td.className = td.className.replace(' editing', '');
        }
    });
}

// ---------------------------------------------------------------------------
// Normalization — U-01, U-05
// ---------------------------------------------------------------------------

async function runNormalization() {
    if (!state.sessionId) return;
    dismissError();

    // U-01: Warn the user if they have unsaved manual edits that will be
    // overwritten by re-normalization.
    const session = sessions.get(state.sessionId);
    if (session && session.hasEdits) {
        const confirmed = confirm(
            'Running normalization will discard your manual cell edits.\n\n' +
            'Continue?'
        );
        if (!confirmed) return;
    }

    const btn = document.getElementById('normalize-btn');
    btn.disabled = true;
    btn.innerHTML = '⏳ Normalizing... <span class="loading"></span>';

    try {
        // Normalize the entire workbook (all sheets) — this is the default and
        // correct behavior.  The backend supports ?sheet=<name> for single-sheet
        // normalization but the UI button must always normalize all sheets so
        // that identifier_status, MosadID, and all corrected fields are
        // consistent across every sheet in the workbook.
        const result = await apiCall('POST',
            `/api/workbook/${state.sessionId}/normalize`);

        // U-05: Cache per-sheet stats on the session so tabs can show badges
        if (session) {
            session.isNormalized = true;
            session.hasEdits = false;  // edits were replayed by the server (F-01)
            if (!session.sheetStats) session.sheetStats = {};
            result.per_sheet_stats.forEach(s => {
                session.sheetStats[s.sheet_name] = s;
            });
            renderSessionSwitcher();
            _highlightActiveSession();
            // Re-render sheet tabs with updated stats badges
            renderSheetSelector(session.sheetNames, session.sheetStats);
            if (state.currentSheet) setActiveSheetTab(state.currentSheet);
        }

        // Reload the current sheet so the grid shows the corrected values
        if (state.currentSheet) await loadSheet(state.currentSheet);

        const stats = result.per_sheet_stats
            .map(s => `${s.sheet_name}: ${s.rows} rows (${(s.success_rate * 100).toFixed(1)}% success)`)
            .join(' | ');
        document.getElementById('grid-stats').textContent =
            `Normalization complete (${result.sheets_processed} sheet${result.sheets_processed !== 1 ? 's' : ''}) — ${stats}`;
    } catch (err) {
        showError(`Normalization failed: ${err.message}`);
    } finally {
        btn.disabled = false;
        btn.innerHTML = '▶ Run Normalization';
    }
}

// ---------------------------------------------------------------------------
// Single-file Export
// ---------------------------------------------------------------------------

async function exportWorkbook() {
    if (!state.sessionId) return;
    dismissError();

    const btn = document.getElementById('export-btn');
    btn.disabled = true;
    btn.innerHTML = '⏳ Exporting... <span class="loading"></span>';

    try {
        await _downloadFile(`/api/workbook/${state.sessionId}/export`, 'POST', 'normalized.xlsx');
    } catch (err) {
        showError(`Export failed: ${err.message}`);
    } finally {
        btn.disabled = false;
        btn.innerHTML = '⬇ Export / Download';
    }
}

// ---------------------------------------------------------------------------
// Bulk Export (ZIP)
// ---------------------------------------------------------------------------

async function exportBulk(sessionIds) {
    if (!sessionIds.length) return;
    dismissError();

    const btn = document.getElementById('export-all-btn') || document.querySelector('.bulk-export-bar .btn');
    if (btn) { btn.disabled = true; btn.innerHTML = '⏳ Exporting... <span class="loading"></span>'; }

    try {
        const response = await fetch('/api/export/bulk', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ session_ids: sessionIds }),
        });

        if (!response.ok) {
            let detail = `HTTP ${response.status}`;
            try { const err = await response.json(); detail = err.detail || detail; } catch (_) {}
            throw new Error(detail);
        }

        const blob = await response.blob();
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'export_bulk.zip';
        document.body.appendChild(a); a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    } catch (err) {
        showError(`Bulk export failed: ${err.message}`);
    } finally {
        if (btn) { btn.disabled = false; btn.innerHTML = '⬇ Export all as ZIP'; }
    }
}

// Export only the sessions whose file-tab checkboxes are checked
async function exportSelected() {
    const checked = [...document.querySelectorAll('.file-tab-check:checked')]
        .map(cb => cb.dataset.sessionId);
    if (!checked.length) {
        showError('Select at least one file to export.');
        return;
    }
    await exportBulk(checked);
}

// ---------------------------------------------------------------------------
// Shared download helper
// ---------------------------------------------------------------------------

async function _downloadFile(url, method, defaultFilename) {
    const response = await fetch(url, { method });
    if (!response.ok) {
        let detail = `HTTP ${response.status}`;
        try { const err = await response.json(); detail = err.detail || detail; } catch (_) {}
        throw new Error(detail);
    }
    const blob = await response.blob();
    const cd = response.headers.get('content-disposition') || '';
    const match = cd.match(/filename="?([^"]+)"?/);
    const filename = match ? match[1] : defaultFilename;
    const objUrl = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = objUrl; a.download = filename;
    document.body.appendChild(a); a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(objUrl);
}

// ---------------------------------------------------------------------------
// Initialize — U-08: keyboard shortcuts
// ---------------------------------------------------------------------------

document.addEventListener('DOMContentLoaded', () => {
    const form = document.getElementById('upload-form');
    if (form) form.addEventListener('submit', handleUpload);

    const fileInput = document.getElementById('file-input');
    if (fileInput) {
        fileInput.addEventListener('change', () => {
            const label = document.querySelector('.file-label');
            if (!label) return;
            const files = Array.from(fileInput.files);
            label.textContent = files.length === 1
                ? files[0].name
                : `${files.length} files selected`;
        });
    }

    // U-08: Keyboard shortcuts for power users.
    // Ctrl+Enter (or Cmd+Enter on Mac) = Run Normalization
    // Ctrl+S (or Cmd+S on Mac) = Export / Download
    document.addEventListener('keydown', e => {
        const mod = e.ctrlKey || e.metaKey;
        if (!mod) return;

        if (e.key === 'Enter') {
            e.preventDefault();
            if (state.sessionId) runNormalization();
        } else if (e.key === 's') {
            e.preventDefault();
            if (state.sessionId) exportWorkbook();
        }
    });

    // Institution bar: save on blur for all inputs
    const instId    = document.getElementById('inst-id');
    const instName  = document.getElementById('inst-name');
    const instType1 = document.getElementById('inst-type-1');
    const instType2 = document.getElementById('inst-type-2');
    const instType3 = document.getElementById('inst-type-3');

    function saveInstitution() {
        if (!state.sessionId) return;
        const types = [instType1, instType2, instType3]
            .map(el => el ? el.value.trim() : '')
            .filter(v => v !== '');
        apiCall('PATCH', `/api/workbook/${state.sessionId}/institution`, {
            mosad_id:    instId   ? instId.value   : undefined,
            mosad_name:  instName ? instName.value : undefined,
            mosad_types: types,
        }).catch(err => showError(`Failed to save institution: ${err.message}`));
    }

    if (instId)    instId.addEventListener('blur', saveInstitution);
    if (instName)  instName.addEventListener('blur', saveInstitution);

    // Rebuild the apply-dropdown whenever a type input changes, then save.
    [instType1, instType2, instType3].forEach(el => {
        if (!el) return;
        el.addEventListener('input', updateMosadTypeDropdown);
        el.addEventListener('blur', () => { saveInstitution(); updateMosadTypeDropdown(); });
    });
});

// ---------------------------------------------------------------------------
// Institution bar
// ---------------------------------------------------------------------------

/**
 * Rebuild the apply-dropdown from the actual user-entered type values.
 * Only non-empty inputs appear as selectable options.
 * Preserves the currently selected value when possible.
 */
function updateMosadTypeDropdown() {
    const sel = document.getElementById('inst-type-apply-select');
    if (!sel) return;

    const inputIds = ['inst-type-1', 'inst-type-2', 'inst-type-3'];
    const entries = inputIds
        .map(id => { const el = document.getElementById(id); return el ? el.value.trim() : ''; })
        .filter(v => v !== '');

    // Remember the currently selected value so we can restore it.
    const prevValue = sel.value || '';

    sel.innerHTML = '';

    if (entries.length === 0) {
        const opt = document.createElement('option');
        opt.value = '';
        opt.textContent = 'הזן ערך סוג מוסד';
        opt.disabled = true;
        opt.selected = true;
        sel.appendChild(opt);
        return;
    }

    let restored = false;
    entries.forEach(value => {
        const opt = document.createElement('option');
        opt.value = value;
        opt.textContent = value;
        if (value === prevValue) { opt.selected = true; restored = true; }
        sel.appendChild(opt);
    });
    if (!restored) sel.options[0].selected = true;
}

async function loadInstitution() {
    if (!state.sessionId) return;
    try {
        const data = await apiCall('GET', `/api/workbook/${state.sessionId}/institution`);
        const instId    = document.getElementById('inst-id');
        const instName  = document.getElementById('inst-name');
        const instType1 = document.getElementById('inst-type-1');
        const instType2 = document.getElementById('inst-type-2');
        const instType3 = document.getElementById('inst-type-3');
        if (instId)    instId.value    = data.mosad_id   || '';
        if (instName)  instName.value  = data.mosad_name || '';
        const types = data.mosad_types || [];
        if (instType1) instType1.value = types[0] || '';
        if (instType2) instType2.value = types[1] || '';
        if (instType3) instType3.value = types[2] || '';
        // Rebuild dropdown with the loaded real values.
        updateMosadTypeDropdown();
    } catch (_) { /* non-critical */ }
}

async function applyMosadType() {
    if (!state.sessionId) return;

    // Read the actual selected user-entered value directly from the dropdown.
    const applySelect = document.getElementById('inst-type-apply-select');
    const mosadType   = applySelect ? applySelect.value.trim() : '';

    if (!mosadType) {
        showError('הזן ערך סוג מוסד לפני ההחלה.');
        return;
    }

    try {
        const result = await apiCall('POST',
            `/api/workbook/${state.sessionId}/mosad-type/apply`,
            { mosad_type: mosadType });
        document.getElementById('grid-stats').textContent =
            `סוג מוסד "${mosadType}" הוחל על ${result.updated_rows} שורות`;
        if (state.currentSheet) await loadSheet(state.currentSheet);
    } catch (err) {
        showError(`Failed to apply MosadType: ${err.message}`);
    }
}
