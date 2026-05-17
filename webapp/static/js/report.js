/**
 * Processing report UI helpers.
 */

let cachedProcessingReport = null;

function escapeHtml(value) {
    return String(value ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

function reportEndpoint(includeDetails = false) {
    const suffix = includeDetails ? '?include_details=true' : '';
    return `/api/workbook/${state.sessionId}/report${suffix}`;
}

async function fetchProcessingReport(includeDetails = false) {
    if (!state.sessionId) return null;
    const report = await apiCall('GET', reportEndpoint(includeDetails));
    cachedProcessingReport = report;
    return report;
}

function formatProcessingReportSummary(report) {
    if (!report) return '';
    const summary = report.summary || {};
    const parts = [];
    const missingRequired = (report.missing_required_fields || [])
        .map(item => `${item.field}: ${item.count}`)
        .join(', ');
    if (missingRequired) parts.push(`Missing required fields: ${missingRequired}`);
    if (summary.total_rows != null) parts.push(`Rows: ${summary.total_rows}`);
    if (summary.rows_with_warnings != null) parts.push(`Warnings: ${summary.rows_with_warnings}`);
    if (summary.rows_with_errors != null) parts.push(`Errors: ${summary.rows_with_errors}`);
    if (summary.corrected_fields != null) parts.push(`Corrected fields: ${summary.corrected_fields}`);
    if (report.export_blocked_reason) parts.push(report.export_blocked_reason);
    return parts.join(' | ');
}

function renderWorkbookReport(report) {
    const section = document.getElementById('processing-report-section');
    const status = document.getElementById('report-status');
    const summary = document.getElementById('report-summary');
    const details = document.getElementById('report-details');
    if (!section || !status || !summary || !details || !report) return;

    const s = report.summary || {};
    status.textContent = report.dirty || report.stale
        ? 'נדרש רענון'
        : (report.export_ready ? 'מוכן לייצוא' : 'לא מוכן לייצוא');
    status.className = `report-status ${report.export_ready && !report.dirty ? 'ready' : 'blocked'}`;

    summary.innerHTML = `
        <div class="report-meta">
            <span class="report-meta-label">קובץ</span>
            <strong>${escapeHtml(report.file_name || '—')}</strong>
        </div>
        <div class="report-meta">
            <span class="report-meta-label">מצב</span>
            <strong>${escapeHtml(report.status || '—')}</strong>
        </div>
        <div class="report-meta">
            <span class="report-meta-label">ייצוא</span>
            <strong>${report.export_ready ? 'מוכן' : 'חסום'}</strong>
        </div>
        <div class="report-meta">
            <span class="report-meta-label">מסומן כלא מעודכן</span>
            <strong>${report.dirty ? 'כן' : 'לא'}</strong>
        </div>
        <div class="report-meta">
            <span class="report-meta-label">גיליונות</span>
            <strong>${s.total_sheets || 0}</strong>
        </div>
        <div class="report-meta">
            <span class="report-meta-label">שורות</span>
            <strong>${s.total_rows || 0}</strong>
        </div>
        <div class="report-meta">
            <span class="report-meta-label">עריכות</span>
            <strong>${s.edited_cells || 0}</strong>
        </div>
        <div class="report-meta">
            <span class="report-meta-label">אזהרות</span>
            <strong>${s.rows_with_warnings || 0}</strong>
        </div>
        <div class="report-meta">
            <span class="report-meta-label">שגיאות</span>
            <strong>${s.rows_with_errors || 0}</strong>
        </div>
        <div class="report-meta">
            <span class="report-meta-label">שדות מתוקנים</span>
            <strong>${s.corrected_fields || 0}</strong>
        </div>
    `;

    const sheetRows = (report.sheets || []).map(sheet => `
        <div class="report-sheet-row">
            <span>${escapeHtml(sheet.sheet_name)}</span>
            <span>${sheet.row_count} שורות</span>
            <span>${sheet.column_count} עמודות</span>
            <span>${sheet.rows_with_warnings} אזהרות</span>
            <span>${sheet.rows_with_errors} שגיאות</span>
            <span>${sheet.corrected_fields} מתוקנים</span>
        </div>
    `).join('');

    const blocked = report.export_blocked_reason
        ? `<div class="report-blocked">${escapeHtml(report.export_blocked_reason)}</div>`
        : '';

    details.innerHTML = `
        ${blocked}
        <div class="report-sheet-list">
            <div class="report-subtitle">סיכום גיליונות</div>
            ${sheetRows || '<div class="report-empty">אין נתוני גיליונות להצגה.</div>'}
        </div>
    `;
    section.classList.remove('hidden');
}

function renderProcessingReportPreview(report) {
    const container = document.getElementById('processing-report-preview');
    if (!container || !report) return;

    const summary = report.summary || {};
    const manual = report.manual_edits || {};
    const sheetBlocks = (report.sheets || []).map(sheet => {
        const statusBlocks = Object.entries(sheet.status_counts || {}).map(([field, counts]) => {
            const values = Object.entries(counts || {})
                .map(([value, count]) => `<li><span>${escapeHtml(value)}</span><strong>${count}</strong></li>`)
                .join('');
            return `
                <div class="report-status-group">
                    <div class="report-status-group-title">${escapeHtml(field)}</div>
                    <ul class="report-status-list">${values || '<li><span>אין</span><strong>0</strong></li>'}</ul>
                </div>
            `;
        }).join('');

        return `
            <section class="report-preview-sheet">
                <div class="report-preview-sheet-head">
                    <strong>${escapeHtml(sheet.sheet_name)}</strong>
                    <span>${sheet.row_count} שורות | ${sheet.column_count} עמודות | ${sheet.issues_count} בעיות</span>
                </div>
                <div class="report-preview-metrics">
                    <span>אזהרות ${sheet.rows_with_warnings}</span>
                    <span>שגיאות ${sheet.rows_with_errors}</span>
                    <span>שדות מתוקנים ${sheet.corrected_fields}</span>
                </div>
                <div class="report-status-groups">${statusBlocks || '<div class="report-empty">אין סטטוסים חריגים.</div>'}</div>
            </section>
        `;
    }).join('');

    const issuesRows = (report.issues || []).map(issue => `
        <tr>
            <td>${escapeHtml(issue.severity)}</td>
            <td>${escapeHtml(issue.sheet_name)}</td>
            <td>${escapeHtml(issue.row_uid || String(issue.row_number || ''))}</td>
            <td>${escapeHtml(issue.status_field || issue.field_name)}</td>
            <td>${escapeHtml(issue.status_message)}</td>
        </tr>
    `).join('');

    container.innerHTML = `
        <div class="report-preview-summary">
            <div><span>קובץ</span><strong>${escapeHtml(report.file_name || '—')}</strong></div>
            <div><span>מצב</span><strong>${escapeHtml(report.status || '—')}</strong></div>
            <div><span>מוכן לייצוא</span><strong>${report.export_ready ? 'כן' : 'לא'}</strong></div>
            <div><span>לא מעודכן</span><strong>${report.dirty ? 'כן' : 'לא'}</strong></div>
            <div><span>גיליונות</span><strong>${summary.total_sheets || 0}</strong></div>
            <div><span>שורות</span><strong>${summary.total_rows || 0}</strong></div>
            <div><span>עריכות</span><strong>${summary.edited_cells || 0}</strong></div>
            <div><span>אזהרות</span><strong>${summary.rows_with_warnings || 0}</strong></div>
            <div><span>שגיאות</span><strong>${summary.rows_with_errors || 0}</strong></div>
            <div><span>שדות מתוקנים</span><strong>${summary.corrected_fields || 0}</strong></div>
        </div>
        ${report.export_blocked_reason ? `<div class="report-blocked">${escapeHtml(report.export_blocked_reason)}</div>` : ''}
        <div class="report-preview-section">
            <div class="report-subtitle">מידע כללי</div>
            <div class="report-kv-list">
                <div><span>session_id</span><strong>${escapeHtml(report.session_id)}</strong></div>
                <div><span>status</span><strong>${escapeHtml(report.status || '—')}</strong></div>
                <div><span>dirty</span><strong>${report.dirty ? 'true' : 'false'}</strong></div>
                <div><span>stale</span><strong>${report.stale ? 'true' : 'false'}</strong></div>
                <div><span>export_ready</span><strong>${report.export_ready ? 'true' : 'false'}</strong></div>
            </div>
        </div>
        <div class="report-preview-section">
            <div class="report-subtitle">עריכות ידניות</div>
            <div class="report-kv-list">
                <div><span>cells</span><strong>${manual.edited_cells || 0}</strong></div>
                <div><span>sheets</span><strong>${(manual.edited_sheets || []).join(', ') || '—'}</strong></div>
                <div><span>fields</span><strong>${(manual.edited_fields || []).join(', ') || '—'}</strong></div>
            </div>
        </div>
        <div class="report-preview-section">
            <div class="report-subtitle">סיכום גיליונות</div>
            <div class="report-preview-sheet-list">${sheetBlocks || '<div class="report-empty">אין נתוני גיליונות להצגה.</div>'}</div>
        </div>
        <div class="report-preview-section">
            <div class="report-subtitle">אזהרות ושגיאות</div>
            <div class="report-preview-table-wrap">
                <table class="report-preview-table">
                    <thead>
                        <tr>
                            <th>חומרה</th>
                            <th>גיליון</th>
                            <th>שורה</th>
                            <th>שדה</th>
                            <th>הודעה</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${issuesRows || '<tr><td colspan="5">אין בעיות להצגה.</td></tr>'}
                    </tbody>
                </table>
            </div>
        </div>
    `;
}

async function refreshProcessingReport() {
    if (!state.sessionId) return;
    try {
        const report = await fetchProcessingReport(false);
        renderWorkbookReport(report);
        return report;
    } catch (err) {
        showError(`טעינת דוח העיבוד נכשלה: ${err.message}`);
        return null;
    }
}

async function openProcessingReportPreview() {
    if (!state.sessionId) return;
    try {
        const report = await fetchProcessingReport(true);
        renderWorkbookReport(report);
        renderProcessingReportPreview(report);
        const modal = document.getElementById('processing-report-modal');
        if (modal) modal.classList.remove('hidden');
    } catch (err) {
        showError(`תצוגת הדוח נכשלה: ${err.message}`);
    }
}

function closeProcessingReportPreview() {
    const modal = document.getElementById('processing-report-modal');
    if (modal) modal.classList.add('hidden');
}

async function downloadProcessingReport() {
    if (!state.sessionId) return;
    try {
        const report = cachedProcessingReport || await fetchProcessingReport(false);
        const fallback = report?.file_name
            ? `processing_report_${report.file_name}.xlsx`
            : 'processing_report.xlsx';
        await _downloadFile(`/api/workbook/${state.sessionId}/report/export`, 'GET', fallback);
    } catch (err) {
        showError(`ייצוא דוח העיבוד נכשל: ${err.message}`);
    }
}

Object.assign(window, {
    formatProcessingReportSummary,
    refreshProcessingReport,
    openProcessingReportPreview,
    closeProcessingReportPreview,
    downloadProcessingReport,
    renderWorkbookReport,
    renderProcessingReportPreview,
});
