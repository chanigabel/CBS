/**
 * Excel standardization Web App - Frontend JavaScript wiring.
 * Global behavior lives in static/js/*.js modules loaded before this file.
 */

const STANDARDIZATION_STATE_KEY = 'excelStandardizationState:v1';

function initApp() {
    const actionBar = document.getElementById('action-bar');
    const reportSection = document.getElementById('processing-report-section');
    if (actionBar && reportSection && actionBar.previousElementSibling !== reportSection) {
        actionBar.parentNode.insertBefore(reportSection, actionBar);
    }

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
                : `נבחרו ${files.length} קבצים`;
        });
    }

    initKeyboardShortcuts();

    // Institution metadata form.
    const instId    = document.getElementById('inst-id');
    const instName  = document.getElementById('inst-name');
    const instType1 = document.getElementById('inst-type-1');
    const instType2 = document.getElementById('inst-type-2');
    const instType3 = document.getElementById('inst-type-3');

    function saveInstitution() {
        if (!state.sessionId) return;
        const rawId = instId ? instId.value.trim() : '';
        const types = [instType1, instType2, instType3]
            .map(el => el ? el.value.trim() : '')
            .filter(v => v !== '');
        for (const t of types) {
            const tErr = validateNumericMin3(t, 'סוג מוסד');
            if (tErr) { showError(tErr); return; }
        }
        apiCall('PATCH', `/api/workbook/${state.sessionId}/institution`, {
            mosad_id: rawId || undefined,
            mosad_name: instName ? instName.value.trim() || undefined : undefined,
            mosad_types: types,
        }).catch(err => showError(`שמירת פרטי המוסד נכשלה: ${err.message}`));
    }

    [instId, instName, instType1, instType2, instType3].forEach(el => {
        if (el) el.addEventListener('blur', saveInstitution);
    });

    const applySelect = document.getElementById('inst-type-apply-select');
    if (applySelect) {
        applySelect.addEventListener('change', () => {
            const v = applySelect.value;
            if (instType1 && !instType1.value) instType1.value = v;
            else if (instType2 && !instType2.value) instType2.value = v;
            else if (instType3 && !instType3.value) instType3.value = v;
        });
    }

    // Initialize dependent UI controls.
    updateMosadTypeDropdown();
    updateInstSheetSelector();
    onScopeChange();
    if (typeof updateUndoButton === 'function') updateUndoButton();
    if (typeof updateDeleteButton === 'function') updateDeleteButton();

    restoreStandardizationUiState().catch(err => {
        console.warn('Failed to restore standardization UI state', err);
    });

    window.addEventListener('pagehide', saveStandardizationUiState);
}

function keyboardMessage(message) {
    const stats = document.getElementById('grid-stats');
    if (stats) stats.textContent = message;
}

function isKeyboardEditableTarget(target) {
    if (!target || target === document || target === window) return false;
    if (target.isContentEditable) return true;
    if (target.closest && target.closest('[contenteditable="true"]')) return true;
    if (target.matches && target.matches('textarea, select, [role="textbox"]')) return true;
    if (target.matches && target.matches('input')) {
        const type = (target.getAttribute('type') || 'text').toLowerCase();
        return !['button', 'submit', 'reset', 'checkbox', 'radio'].includes(type);
    }
    return false;
}

function isGridShortcutContext(event) {
    const target = event?.target || document.activeElement;
    if (target?.closest && target.closest('#grid-section, #grid-overlay')) return true;
    return Boolean(state.gridShortcutActive && state.sheetData);
}

function isOverlayOpen() {
    const overlay = document.getElementById('grid-overlay');
    return Boolean(overlay && !overlay.classList.contains('hidden'));
}

function hasVisibleGridRows() {
    return Boolean(state.sheetData && getFilteredRows(state.sheetData.rows).length > 0);
}

function isButtonEnabled(buttonId) {
    const button = document.getElementById(buttonId);
    return Boolean(button && !button.disabled);
}

function selectAllVisibleGridRows() {
    if (!state.sheetData) return;
    const visibleRows = getFilteredRows(state.sheetData.rows);
    state.selectedRows = new Set(
        visibleRows.map(row => getRowUid(row)).filter(Boolean)
    );
    renderGrid(state.sheetData, visibleRows);
    keyboardMessage('נבחרו כל השורות המוצגות');
}

function clearSelectionFromShortcut() {
    clearRowSelection();
    keyboardMessage('הבחירה נוקתה');
}

function deleteSelectedRowsFromShortcut() {
    if (state.selectedRows.size === 0) {
        keyboardMessage('לא נבחרו שורות למחיקה');
        return;
    }
    deleteSelectedRows();
}

function exportCurrentSheetFromShortcut() {
    if (!state.sessionId || !state.currentSheet || !isButtonEnabled('export-sheet-btn')) {
        keyboardMessage('אין גיליון נבחר לייצוא');
        return;
    }
    exportCurrentSheet();
}

const keyboardShortcutRegistry = [
    {
        id: 'normalize',
        label: 'הרצת סטנדרטיזציה',
        match: event => (event.ctrlKey || event.metaKey) && event.key === 'Enter',
        enabled: () => Boolean(state.sessionId && isButtonEnabled('normalize-btn')),
        action: () => runstandardization(),
    },
    {
        id: 'undo',
        label: 'בטל שינוי אחרון',
        match: event => (event.ctrlKey || event.metaKey) && !event.shiftKey && event.key.toLowerCase() === 'z',
        enabled: () => Boolean(state.sessionId),
        action: () => undoLastGridEdit(),
    },
    {
        id: 'delete-selected',
        label: 'מחיקת שורות נבחרות',
        match: event => !event.ctrlKey && !event.metaKey && !event.altKey && event.key === 'Delete',
        enabled: event => Boolean(state.sessionId && isGridShortcutContext(event)),
        action: () => deleteSelectedRowsFromShortcut(),
    },
    {
        id: 'clear-selection',
        label: 'נקה בחירה',
        match: event => !event.ctrlKey && !event.metaKey && !event.altKey && event.key === 'Escape',
        enabled: event => Boolean(isOverlayOpen() || (state.selectedRows.size > 0 && isGridShortcutContext(event))),
        action: () => {
            if (isOverlayOpen()) closeGridOverlay();
            else clearSelectionFromShortcut();
        },
    },
    {
        id: 'select-all-visible',
        label: 'בחר את כל השורות המוצגות',
        match: event => (event.ctrlKey || event.metaKey) && !event.shiftKey && event.key.toLowerCase() === 'a',
        enabled: event => Boolean(isGridShortcutContext(event) && hasVisibleGridRows()),
        action: () => selectAllVisibleGridRows(),
    },
    {
        id: 'export-workbook',
        label: 'ייצא קובץ',
        match: event => (event.ctrlKey || event.metaKey) && !event.shiftKey && event.key.toLowerCase() === 's',
        enabled: () => Boolean(state.sessionId && isButtonEnabled('export-btn')),
        action: () => exportWorkbook(),
    },
    {
        id: 'export-current-sheet',
        label: 'ייצא גיליון',
        match: event => (event.ctrlKey || event.metaKey) && event.shiftKey && event.key.toLowerCase() === 'e',
        enabled: () => Boolean(state.sessionId),
        action: () => exportCurrentSheetFromShortcut(),
    },
];

function handleKeyboardShortcut(event) {
    if (isKeyboardEditableTarget(event.target)) return;
    const shortcut = keyboardShortcutRegistry.find(item => item.match(event));
    if (!shortcut) return;
    if (!shortcut.enabled(event)) return;
    event.preventDefault();
    shortcut.action(event);
}

function initKeyboardShortcuts() {
    if (window.__keyboardShortcutsInitialized) return;
    window.__keyboardShortcutsInitialized = true;

    document.addEventListener('pointerdown', event => {
        state.gridShortcutActive = Boolean(
            event.target?.closest && event.target.closest('#grid-section, #grid-overlay')
        );
    }, true);
    document.addEventListener('focusin', event => {
        if (event.target?.closest && event.target.closest('#grid-section, #grid-overlay')) {
            state.gridShortcutActive = true;
        }
    });
    document.addEventListener('keydown', handleKeyboardShortcut);
}

function readValue(id) {
    const el = document.getElementById(id);
    return el ? el.value || '' : '';
}

function writeValue(id, value) {
    const el = document.getElementById(id);
    if (el) el.value = value || '';
}

function serializeSessions() {
    return [...sessions.values()].map(session => ({
        sessionId: session.sessionId,
        filename: session.filename,
        sheetNames: session.sheetNames || [],
        lastSheet: session.lastSheet || null,
        isNormalized: Boolean(session.isNormalized),
        hasEdits: Boolean(session.hasEdits),
        sheetStats: session.sheetStats || {},
    }));
}

function serializeColumnFilters() {
    return [...state.columnFilters.entries()].map(([column, values]) => [
        column,
        [...values],
    ]);
}

function collectFormState() {
    return {
        instId: readValue('inst-id'),
        instName: readValue('inst-name'),
        instType1: readValue('inst-type-1'),
        instType2: readValue('inst-type-2'),
        instType3: readValue('inst-type-3'),
        instTypeApply: readValue('inst-type-apply-select'),
        instScope: readValue('inst-scope-select'),
        instSheet: readValue('inst-sheet-select'),
        uploadStatus: document.getElementById('upload-status')?.textContent || '',
        fileLabel: document.querySelector('.file-label')?.textContent || '',
    };
}

function restoreFormState(formState) {
    if (!formState) return;
    writeValue('inst-id', formState.instId);
    writeValue('inst-name', formState.instName);
    writeValue('inst-type-1', formState.instType1);
    writeValue('inst-type-2', formState.instType2);
    writeValue('inst-type-3', formState.instType3);
    writeValue('inst-scope-select', formState.instScope || 'workbook');

    const status = document.getElementById('upload-status');
    if (status) status.textContent = formState.uploadStatus || '';
    const fileLabel = document.querySelector('.file-label');
    if (fileLabel && formState.fileLabel) fileLabel.textContent = formState.fileLabel;

    updateMosadTypeDropdown();
    writeValue('inst-type-apply-select', formState.instTypeApply);
    updateInstSheetSelector();
    writeValue('inst-sheet-select', formState.instSheet);
    onScopeChange();
}

function saveStandardizationUiState() {
    try {
        const payload = {
            version: 1,
            savedAt: new Date().toISOString(),
            sessions: serializeSessions(),
            currentSessionId: state.sessionId,
            currentSheet: state.currentSheet,
            selectedRows: [...state.selectedRows],
            columnFilters: serializeColumnFilters(),
            form: collectFormState(),
            scrollY: window.scrollY,
        };
        sessionStorage.setItem(STANDARDIZATION_STATE_KEY, JSON.stringify(payload));
    } catch (err) {
        console.warn('Failed to save standardization UI state', err);
    }
}

function isSessionNotFoundError(err) {
    return err && err.status === 404 && err.message.includes('Session');
}

function clearStaleSessionState(message) {
    sessions.clear();
    state.sessionId = null;
    state.currentSheet = null;
    state.sheetData = null;
    state.selectedRows.clear();
    state.columnFilters.clear();

    renderSessionSwitcher();

    const sheetSelector = document.getElementById('sheet-selector');
    if (sheetSelector) sheetSelector.classList.add('hidden');

    const actionBar = document.getElementById('action-bar');
    if (actionBar) actionBar.classList.add('hidden');

    const gridSection = document.getElementById('grid-section');
    if (gridSection) gridSection.classList.add('hidden');

    const reportSection = document.getElementById('processing-report-section');
    if (reportSection) reportSection.classList.add('hidden');

    const gridContainer = document.getElementById('grid-container');
    if (gridContainer) gridContainer.innerHTML = '';

    const gridTitle = document.getElementById('grid-title');
    if (gridTitle) gridTitle.textContent = '';

    const uploadStatus = document.getElementById('upload-status');
    if (uploadStatus) uploadStatus.textContent = message || '';

    sessionStorage.removeItem(STANDARDIZATION_STATE_KEY);
}

async function restoreStandardizationUiState() {
    const raw = sessionStorage.getItem(STANDARDIZATION_STATE_KEY);
    if (!raw || sessions.size > 0) return;

    let payload;
    try {
        payload = JSON.parse(raw);
    } catch (_) {
        sessionStorage.removeItem(STANDARDIZATION_STATE_KEY);
        return;
    }

    if (!payload || !Array.isArray(payload.sessions) || payload.sessions.length === 0) {
        restoreFormState(payload?.form);
        return;
    }

    payload.sessions.forEach(session => {
        if (!session.sessionId) return;
        sessions.set(session.sessionId, {
            sessionId: session.sessionId,
            filename: session.filename || session.sessionId,
            sheetNames: session.sheetNames || [],
            lastSheet: session.lastSheet || null,
            isNormalized: Boolean(session.isNormalized),
            hasEdits: Boolean(session.hasEdits),
            sheetStats: session.sheetStats || {},
        });
    });

    const activeSessionId = payload.currentSessionId && sessions.has(payload.currentSessionId)
        ? payload.currentSessionId
        : [...sessions.keys()][0];

    if (!activeSessionId) {
        restoreFormState(payload.form);
        return;
    }

    try {
        await apiCall('GET', `/api/workbook/${activeSessionId}/summary`);
    } catch (err) {
        if (isSessionNotFoundError(err)) {
            clearStaleSessionState('הסשן הסתיים. יש להעלות קובץ מחדש.');
            return;
        }
        throw err;
    }

    renderSessionSwitcher();

    const activeSession = sessions.get(activeSessionId);
    if (payload.currentSheet && activeSession) activeSession.lastSheet = payload.currentSheet;
    await activateSession(activeSessionId);

    state.selectedRows = new Set(payload.selectedRows || []);
    state.columnFilters = new Map((payload.columnFilters || []).map(([column, values]) => [
        column,
        new Set(values || []),
    ]));

    restoreFormState(payload.form);

    if (state.sheetData) {
        renderGrid(state.sheetData, getFilteredRows(state.sheetData.rows));
    }

    if (typeof payload.scrollY === 'number') {
        window.scrollTo(0, payload.scrollY);
    }
}



document.addEventListener('DOMContentLoaded', initApp);

Object.assign(window, {
    saveStandardizationUiState,
    restoreStandardizationUiState,
});
