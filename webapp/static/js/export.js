async function runstandardization() {
    if (!state.sessionId) return;
    dismissError();

    const session = sessions.get(state.sessionId);

    const btn = document.getElementById('normalize-btn');
    btn.disabled = true;
    btn.innerHTML = '⏳ מריץ סטנדרטיזציה... <span class="loading"></span>';

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
            .map(s => `${s.sheet_name}: ${s.rows} שורות (${(s.success_rate * 100).toFixed(1)}% מוצלח)`)
            .join(' | ');
        document.getElementById('grid-stats').textContent =
            `תקנון הושלם (${result.sheets_processed} גיליון${result.sheets_processed !== 1 ? 'ים' : ''}) — ${stats}`;
        await refreshProcessingReport();
    } catch (err) {
        showError(`הרצת הסטנדרטיזציה נכשלה: ${err.message}`);
    } finally {
        btn.disabled = false;
        btn.innerHTML = '▶ הרצת סטנדרטיזציה <span class="shortcut-hint">Ctrl+Enter</span>';
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
    btn.innerHTML = '⏳ מייצא... <span class="loading"></span>';

    try {
        await _downloadFile(`/api/workbook/${state.sessionId}/export`, 'POST', 'normalized.xlsx');
        const report = await apiCall('GET', `/api/workbook/${state.sessionId}/processing-report`);
        document.getElementById('grid-stats').textContent =
            `ייצוא הושלם (${report.status}) - ${formatProcessingReportSummary(report)}`;
        await refreshProcessingReport();
    } catch (err) {
        showError(`הייצוא נכשל: ${err.message}`);
    } finally {
        btn.disabled = false;
        btn.innerHTML = '⬇ ייצוא קובץ <span class="shortcut-hint">Ctrl+S</span>';
    }
}

async function exportCurrentSheet() {
    if (!state.sessionId || !state.currentSheet) return;
    dismissError();

    const btn = document.getElementById('export-sheet-btn');
    if (btn) {
        btn.disabled = true;
        btn.innerHTML = 'מייצא גיליון...';
    }

    try {
        await _downloadFile(
            `/api/workbook/${state.sessionId}/sheet/${encodeURIComponent(state.currentSheet)}/export`,
            'POST',
            `${state.currentSheet}.xlsx`
        );
        const stats = document.getElementById('grid-stats');
        if (stats) stats.textContent = `ייצוא הגיליון ${state.currentSheet} הושלם`;
        await refreshProcessingReport();
    } catch (err) {
        showError(`ייצוא הגיליון נכשל: ${err.message}`);
    } finally {
        if (btn) {
            btn.disabled = false;
            btn.innerHTML = 'ייצוא גיליון <span class="shortcut-hint">Ctrl+Shift+E</span>';
        }
    }
}

// ---------------------------------------------------------------------------
// Bulk export (ZIP)
// ---------------------------------------------------------------------------

async function exportBulk(sessionIds) {
    if (!sessionIds.length) return;
    dismissError();

    const btn = document.getElementById('export-all-btn') || document.querySelector('.bulk-export-bar .btn');
    if (btn) { btn.disabled = true; btn.innerHTML = '⏳ מייצא... <span class="loading"></span>'; }

    try {
        const response = await fetch('/api/export/bulk', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ session_ids: sessionIds }),
        });

        if (!response.ok) {
            let detail = `HTTP ${response.status}`;
            try {
                const err = await response.json();
                detail = typeof formatApiErrorDetail === 'function'
                    ? formatApiErrorDetail(err.detail || detail)
                    : (err.detail || detail);
            } catch (_) {}
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
        showError(`ייצוא הקבצים נכשל: ${err.message}`);
    } finally {
        if (btn) { btn.disabled = false; btn.innerHTML = '⬇ ייצא הכל כ-ZIP'; }
    }
}

// Export only the sessions whose file-tab checkboxes are checked.
async function exportSelected() {
    const checked = [...document.querySelectorAll('.file-tab-check:checked')]
        .map(cb => cb.dataset.sessionId);
    if (!checked.length) {
        showError('בחר לפחות קובץ אחד לייצוא.');
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
        try {
            const err = await response.json();
            detail = typeof formatApiErrorDetail === 'function'
                ? formatApiErrorDetail(err.detail || detail)
                : (err.detail || detail);
        } catch (_) {}
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

Object.assign(window, { runstandardization, exportWorkbook, exportCurrentSheet, exportBulk, exportSelected, _downloadFile });
