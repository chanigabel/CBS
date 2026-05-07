async function runstandardization() {
    if (!state.sessionId) return;
    dismissError();

    // U-01: Warn the user if they have unsaved manual edits that will be
    // overwritten by re-standardization.
    const session = sessions.get(state.sessionId);
    if (session && session.hasEdits) {
        const confirmed = confirm(
            'Running standardization will discard your manual cell edits.\n\n' +
            'Continue?'
        );
        if (!confirmed) return;
    }

    const btn = document.getElementById('normalize-btn');
    btn.disabled = true;
    btn.innerHTML = '⏳ standardizing... <span class="loading"></span>';

    try {
        // Normalize the entire workbook (all sheets) — this is the default and
        // correct behavior.  The backend supports ?sheet=<name> for single-sheet
        // standardization but the UI button must always normalize all sheets so
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
            `standardization complete (${result.sheets_processed} sheet${result.sheets_processed !== 1 ? 's' : ''}) — ${stats}`;
    } catch (err) {
        showError(`standardization failed: ${err.message}`);
    } finally {
        btn.disabled = false;
        btn.innerHTML = '▶ Run standardization';
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
        const report = await apiCall('GET', `/api/workbook/${state.sessionId}/processing-report`);
        document.getElementById('grid-stats').textContent =
            `Export complete (${report.status}) - ${formatProcessingReportSummary(report)}`;
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
    // Ctrl+Enter (or Cmd+Enter on Mac) = Run standardization
    // Ctrl+S (or Cmd+S on Mac) = Export / Download
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

    // Institution bar: save on blur for all inputs
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
        // Validate each mosad_type before saving
        for (const t of types) {
            const tErr = validateNumericMin3(t, 'סוג מוסד');
            if (tErr) { showError(tErr); return; }
        }
        apiCall('PATCH', `/api/workbook/${state.sessionId}/institution`, {
            mosad_id:    rawId || undefined,
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

Object.assign(window, { runstandardization, exportWorkbook, exportBulk, exportSelected, _downloadFile });
