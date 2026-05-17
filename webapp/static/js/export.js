async function runstandardization() {
    if (!state.sessionId) return;
    dismissError();

    const session = sessions.get(state.sessionId);

    const btn = document.getElementById('normalize-btn');
    btn.disabled = true;
    btn.innerHTML = '⏳ standardizing... <span class="loading"></span>';

    try {
        // Normalize the full workbook from the main action.
        const result = await apiCall('POST',
            `/api/workbook/${state.sessionId}/normalize`);

        // Store the returned sheet stats on the session.
        if (session) {
            session.isNormalized = true;
            session.hasEdits = false;
            if (!session.sheetStats) session.sheetStats = {};
            result.per_sheet_stats.forEach(s => {
                session.sheetStats[s.sheet_name] = s;
            });
            renderSessionSwitcher();
            _highlightActiveSession();
            // Re-render sheet tabs with updated stats.
            renderSheetSelector(session.sheetNames, session.sheetStats);
            if (state.currentSheet) setActiveSheetTab(state.currentSheet);
        }

        // Reload the current sheet.
        if (state.currentSheet) await loadSheet(state.currentSheet);

        const stats = result.per_sheet_stats
            .map(s => `${s.sheet_name}: ${s.rows} rows (${(s.success_rate * 100).toFixed(1)}% success)`)
            .join(' | ');
        document.getElementById('grid-stats').textContent =
            `standardization complete (${result.sheets_processed} sheet${result.sheets_processed !== 1 ? 's' : ''}) — ${stats}`;
        await refreshProcessingReport();
    } catch (err) {
        showError(`standardization failed: ${err.message}`);
    } finally {
        btn.disabled = false;
        btn.innerHTML = '▶ Run standardization';
    }
}

// ---------------------------------------------------------------------------
// Single-file export
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
        await refreshProcessingReport();
    } catch (err) {
        showError(`Export failed: ${err.message}`);
    } finally {
        btn.disabled = false;
        btn.innerHTML = '⬇ Export / Download';
    }
}

// ---------------------------------------------------------------------------
// Bulk export (ZIP)
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

// Export only the sessions whose file-tab checkboxes are checked.
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

Object.assign(window, { runstandardization, exportWorkbook, exportBulk, exportSelected, _downloadFile });
