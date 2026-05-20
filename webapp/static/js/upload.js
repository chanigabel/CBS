// ---------------------------------------------------------------------------
// Upload flow
// ---------------------------------------------------------------------------

/**
 * Upload a single file using XMLHttpRequest so progress can be reported.
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
                    reject(new Error('תגובת השרת אינה תקינה'));
                }
            } else {
                let detail = `HTTP ${xhr.status}`;
                try {
                    const err = JSON.parse(xhr.responseText);
                    detail = typeof formatApiErrorDetail === 'function'
                        ? formatApiErrorDetail(err.detail || detail)
                        : (err.detail || detail);
                } catch (_) {}
                reject(new Error(detail));
            }
        });

        xhr.addEventListener('error', () => reject(new Error('שגיאת רשת במהלך העלאת הקובץ')));
        xhr.addEventListener('abort', () => reject(new Error('העלאת הקובץ בוטלה')));

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
            // Show per-file upload progress.
            statusDiv.textContent = `העלה ${file.name}: 0%`;
            const data = await uploadWithProgress(file, pct => {
                statusDiv.textContent = `העלה ${file.name}: ${pct}%`;
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

    if (errors.length) showError(`העלאות נכשלו:\n${errors.join('\n')}`);

    if (successCount === 0) {
        statusDiv.textContent = '';
        uploadBtn.disabled = false;
        return;
    }

    statusDiv.textContent = `הועלו ${successCount} קובץ(ים).`;
    renderSessionSwitcher();

    const lastSession = [...sessions.values()].at(-1);
    await activateSession(lastSession.sessionId);

    uploadBtn.disabled = false;
}

// ---------------------------------------------------------------------------
// Session switcher
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

        // Show a warning badge if any sheet has < 100% success rate.
        const hasWarning = sheetStats && Object.values(sheetStats).some(s => s.success_rate < 1.0);
        let label = filename;
        if (isNormalized) label += hasWarning ? ' ⚠' : ' ✓';
        btn.textContent = label;
        btn.title = filename;

        if (sessionId === state.sessionId) btn.classList.add('active');
        btn.onclick = () => activateSession(sessionId);
        tabs.appendChild(btn);
    });

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
// Session activation
// ---------------------------------------------------------------------------

async function activateSession(sessionId) {
    const session = sessions.get(sessionId);
    if (!session) return;

    // Save the outgoing sheet before switching sessions.
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
    document.getElementById('processing-report-section')?.classList.add('hidden');

    // Load institution metadata for this session.
    await loadInstitution();

    // Restore the last sheet this session had open, or default to the first sheet.
    const sheetToLoad = session.lastSheet || session.sheetNames[0];
    if (sheetToLoad) {
        await loadSheet(sheetToLoad);
    }
}

// ---------------------------------------------------------------------------
// Sheet selector
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

        // Annotate sheet tabs with success-rate badges when available.
        const stat = (sheetStats || {})[name];
        if (stat && stat.success_rate < 1.0) {
            btn.textContent = `${name} ⚠ ${Math.round(stat.success_rate * 100)}%`;
            btn.title = `${stat.rows} שורות — ${Math.round(stat.success_rate * 100)}% תוקננו בהצלחה`;
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
        // Match by the base name, ignoring any appended stats badge.
        btn.classList.toggle('active', btn.textContent.startsWith(sheetName));
    });
}

// ---------------------------------------------------------------------------
// Sheet Data Loading
// ---------------------------------------------------------------------------

async function loadSheet(sheetName) {
    if (!state.sessionId) return;
    dismissError();

    // Clear row selections only when navigating to a different sheet.
    // Reloading the same sheet (e.g. after applying institution type) must
    // preserve selections so the user can apply a second type to other rows.
    if (sheetName !== state.currentSheet) {
        state.selectedRows.clear();
    }

    state.currentSheet = sheetName;
    state.columnFilters.clear();
    setActiveSheetTab(sheetName);

    // Persist last-viewed sheet on the session record
    const session = sessions.get(state.sessionId);
    if (session) session.lastSheet = sheetName;

    const gridSection = document.getElementById('grid-section');
    const gridTitle = document.getElementById('grid-title');
    const gridContainer = document.getElementById('grid-container');

    gridTitle.textContent = sheetName;
    gridContainer.innerHTML = '<div style="padding:20px;text-align:center">טוען... <span class="loading"></span></div>';
    gridSection.classList.remove('hidden');

    try {
        const data = await apiCall('GET', `/api/workbook/${state.sessionId}/sheet/${encodeURIComponent(sheetName)}`);
        state.sheetData = data;
        renderGrid(data, getFilteredRows(data.rows));
        if (typeof refreshColumnMappingControls === 'function') {
            await refreshColumnMappingControls(data);
        }
    } catch (err) {
        if (isSessionNotFoundError(err)) {
            clearStaleSessionState('הסשן הסתיים. יש להעלות קובץ מחדש.');
            return;
        }
        showError(`טעינת הגיליון '${sheetName}' נכשלה: ${err.message}`);
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

/**
 * Populate the sheet selector dropdown with the current session's sheet names.
 * Called when a session is activated or sheets change.
 */
function updateInstSheetSelector() {
    const sel = document.getElementById('inst-sheet-select');
    if (!sel || !state.sessionId) return;
    const session = sessions.get(state.sessionId);
    if (!session) return;

    const prev = sel.value;
    sel.innerHTML = '';
    (session.sheetNames || []).forEach(name => {
        const opt = document.createElement('option');
        opt.value = name;
        opt.textContent = name;
        if (name === prev) opt.selected = true;
        sel.appendChild(opt);
    });
    // Default to current sheet if available
    if (!prev && state.currentSheet) sel.value = state.currentSheet;
}

/**
 * Show/hide the sheet selector and selected-rows hint based on selected scope.
 */
function onScopeChange() {
    const scope = (document.getElementById('inst-scope-select') || {}).value || 'workbook';
    const sheetSel = document.getElementById('inst-sheet-select');
    const hint = document.getElementById('inst-selected-rows-hint');

    if (sheetSel) sheetSel.classList.toggle('hidden', scope === 'workbook');
    if (hint)     hint.classList.toggle('hidden', scope !== 'selected_rows');

    if (scope !== 'workbook') updateInstSheetSelector();
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
        // Populate sheet selector in case scope is already set to sheet/rows.
        updateInstSheetSelector();
        onScopeChange();
    } catch (_) { /* non-critical */ }
}

/** Legacy workbook-wide apply (kept for backward compatibility). */
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
        showError(`החלת סוג המוסד נכשלה: ${err.message}`);
    }
}

/**
 * Validate that a value is numeric-only and at least 3 digits.
 * Returns an error message string, or null if valid.
 */
function validateNumericMin3(value, label) {
    if (!value) return null; // empty is allowed (not required here)
    if (!/^\d+$/.test(value)) return `${label} חייב להכיל ספרות בלבד`;
    if (value.length < 3)     return `${label} חייב להכיל לפחות 3 ספרות`;
    return null;
}

/**
 * Scoped apply: workbook / sheet / selected_rows.
 *
 * For "selected_rows" scope the function reads state.selectedRows (the same
 * set used by row deletion) and sends the selected _row_uid values to the
 * backend.  The user selects rows in the grid first, then clicks "החל סוג".
 *
 * The function can be called multiple times with different row selections and
 * different sug_mosad values to build up to 3 groups per sheet.  Each call
 * replaces the previous selected_rows config for that sheet.
 */
async function applyMosadTypeScoped() {
    if (!state.sessionId) return;
    dismissError();

    const scope       = (document.getElementById('inst-scope-select') || {}).value || 'workbook';
    const applySelect = document.getElementById('inst-type-apply-select');
    const sugMosad    = applySelect ? applySelect.value.trim() : '';
    const mosadId     = (document.getElementById('inst-id') || {}).value?.trim() || '';
    const sheetName   = (document.getElementById('inst-sheet-select') || {}).value || state.currentSheet || '';

    // Validate sug_mosad for workbook/sheet scopes
    if (scope !== 'selected_rows') {
        if (!sugMosad) { showError('הזן ערך סוג מוסד לפני ההחלה.'); return; }
        const sugErr = validateNumericMin3(sugMosad, 'סוג מוסד');
        if (sugErr) { showError(sugErr); return; }
    }

    let body = { scope, mosad_id: mosadId || undefined };

    if (scope === 'workbook') {
        body.sug_mosad = sugMosad;

    } else if (scope === 'sheet') {
        if (!sheetName) { showError('בחר גיליון להחלה.'); return; }
        body.sug_mosad  = sugMosad;
        body.sheet_name = sheetName;

    } else { // selected_rows
        // For selected_rows, always use the sheet currently displayed in the
        // grid — that is where the user selected the rows.  The inst-sheet-select
        // dropdown is only relevant for the "sheet" scope.
        const selectedRowsSheet = state.currentSheet || '';
        if (!selectedRowsSheet) { showError('טען גיליון לפני ההחלה.'); return; }

        // Validate sug_mosad for selected_rows too
        if (!sugMosad) { showError('הזן ערך סוג מוסד לפני ההחלה.'); return; }
        const sugErr = validateNumericMin3(sugMosad, 'סוג מוסד');
        if (sugErr) { showError(sugErr); return; }

        // Use the currently selected rows (same mechanism as row deletion)
        const rowUids = [...state.selectedRows];
        if (rowUids.length === 0) {
            showError('סמן שורות בטבלה לפני ההחלה.');
            return;
        }

        body.sheet_name    = selectedRowsSheet;
        body.selected_rows = [{ sug_mosad: sugMosad, row_uids: rowUids }];
    }

    try {
        const result = await apiCall('POST',
            `/api/workbook/${state.sessionId}/mosad-type/apply-scoped`, body);
        const scopeLabel = scope === 'workbook'       ? 'כל הגיליונות'
                         : scope === 'sheet'          ? `גיליון "${result.sheet_name}"`
                         :                              `${result.updated_rows} שורות נבחרות`;

        // After a successful selected_rows apply, clear the selection so the
        // user starts fresh for the next group.  Do this BEFORE reloading the
        // sheet so the grid renders with no checkboxes ticked.
        // For workbook/sheet scopes there is nothing to clear.
        if (scope === 'selected_rows') {
            state.selectedRows.clear();
            // Uncheck any visible checkboxes immediately (the grid reload below
            // will also render them unchecked, but this avoids a visual flash).
            document.querySelectorAll('.data-grid tbody input[type=checkbox]')
                .forEach(cb => { cb.checked = false; });
            document.querySelectorAll('.data-grid tbody tr.row-selected')
                .forEach(tr => tr.classList.remove('row-selected'));
            updateDeleteButton();
        }

        document.getElementById('grid-stats').textContent =
            `סוג מוסד "${sugMosad}" הוחל על ${result.updated_rows} שורות (${scopeLabel})`;
        if (state.currentSheet) await loadSheet(state.currentSheet);
    } catch (err) {
        showError(`שגיאה בהחלת סוג מוסד: ${err.message}`);
    }
}

Object.assign(window, { uploadWithProgress, handleUpload, renderSessionSwitcher, _highlightActiveSession, activateSession, renderSheetSelector, setActiveSheetTab, loadSheet, updateMosadTypeDropdown, updateInstSheetSelector, onScopeChange, loadInstitution, applyMosadType, validateNumericMin3, applyMosadTypeScoped });
