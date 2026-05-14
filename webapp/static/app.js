/**
 * Excel standardization Web App - Frontend JavaScript wiring.
 * Global behavior lives in static/js/*.js modules loaded before this file.
 */

const STANDARDIZATION_STATE_KEY = 'excelStandardizationState:v1';

function initApp() {
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

    // Keyboard shortcuts.
    document.addEventListener('keydown', e => {
        const mod = e.ctrlKey || e.metaKey;
        if (!mod) return;

        if (e.key === 'Enter') {
            e.preventDefault();
            if (state.sessionId) runstandardization();
        } else if (e.key === 's') {
            e.preventDefault();
            if (state.sessionId) exportWorkbook();
        }
    });

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
        }).catch(err => showError(`Failed to save institution metadata: ${err.message}`));
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

    restoreStandardizationUiState().catch(err => {
        console.warn('Failed to restore standardization UI state', err);
    });

    window.addEventListener('pagehide', saveStandardizationUiState);
    window.addEventListener('popstate', handleEngineManagementPopState);
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

    renderSessionSwitcher();

    const activeSessionId = payload.currentSessionId && sessions.has(payload.currentSessionId)
        ? payload.currentSessionId
        : [...sessions.keys()][0];
    if (activeSessionId) {
        const activeSession = sessions.get(activeSessionId);
        if (payload.currentSheet) activeSession.lastSheet = payload.currentSheet;
        await activateSession(activeSessionId);
    }

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

function ensureEngineManagementOverlay() {
    let overlay = document.getElementById('engine-management-overlay');
    if (overlay) return overlay;

    overlay = document.createElement('div');
    overlay.id = 'engine-management-overlay';
    overlay.className = 'engine-management-overlay hidden';
    overlay.setAttribute('role', 'dialog');
    overlay.setAttribute('aria-modal', 'true');
    overlay.setAttribute('aria-label', '\u05de\u05de\u05e9\u05e7 \u05e0\u05d9\u05d4\u05d5\u05dc \u05de\u05e0\u05d5\u05e2\u05d9\u05dd');
    overlay.innerHTML = `
        <div class="engine-management-overlay-bar">
            <strong>&#1502;&#1502;&#1513;&#1511; &#1504;&#1497;&#1492;&#1493;&#1500; &#1502;&#1504;&#1493;&#1506;&#1497;&#1501;</strong>
            <button type="button" class="btn btn-secondary" onclick="closeEngineManagement()">&#1495;&#1494;&#1512;&#1492; &#1500;&#1502;&#1502;&#1513;&#1511; &#1492;&#1514;&#1511;&#1504;&#1493;&#1503;</button>
        </div>
        <iframe id="engine-management-frame" class="engine-management-frame" title="&#1502;&#1502;&#1513;&#1511; &#1504;&#1497;&#1492;&#1493;&#1500; &#1502;&#1504;&#1493;&#1506;&#1497;&#1501;"></iframe>
    `;
    document.body.appendChild(overlay);
    return overlay;
}

function openEngineManagement() {
    saveStandardizationUiState();
    const overlay = ensureEngineManagementOverlay();
    const frame = document.getElementById('engine-management-frame');
    if (frame && !frame.src) {
        frame.src = '/engine-management';
    }
    overlay.classList.remove('hidden');
    document.body.classList.add('engine-management-open');

    if (window.location.pathname !== '/engine-management') {
        history.pushState({ view: 'engine-management' }, '', '/engine-management');
    }
}

function closeEngineManagement(updateHistory = true) {
    const overlay = document.getElementById('engine-management-overlay');
    if (overlay) overlay.classList.add('hidden');
    document.body.classList.remove('engine-management-open');
    saveStandardizationUiState();

    if (updateHistory && window.location.pathname === '/engine-management') {
        history.pushState({ view: 'standardization' }, '', '/');
    }
}

function handleEngineManagementPopState(event) {
    const wantsEngineManagement = event.state && event.state.view === 'engine-management';
    if (wantsEngineManagement) {
        openEngineManagement();
        return;
    }
    closeEngineManagement(false);
}

document.addEventListener('DOMContentLoaded', initApp);

Object.assign(window, {
    openEngineManagement,
    closeEngineManagement,
    saveStandardizationUiState,
    restoreStandardizationUiState,
});
