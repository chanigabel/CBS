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

function formatNumber(value) {
    const n = Number(value ?? 0);
    return Number.isFinite(n) ? new Intl.NumberFormat('he-IL').format(n) : '0';
}

function getStatusDescriptor(kind, text) {
    const safeKind = ['success', 'warning', 'error', 'pending', 'info'].includes(kind) ? kind : 'info';
    return {
        kind: safeKind,
        text: text || '—',
    };
}

function renderStatusBadgeHtml(kind, text) {
    const descriptor = getStatusDescriptor(kind, text);
    return `<span class="status-badge status-badge--${descriptor.kind}">${escapeHtml(descriptor.text)}</span>`;
}

function reportStatusDescriptor(report) {
    const summary = report?.summary || {};
    const warningCount = Number(summary.rows_with_warnings || 0);
    const errorCount = Number(summary.rows_with_errors || 0);
    const rawStatus = String(report?.status || '').toLowerCase();

    if (rawStatus === 'standardized') {
        if (errorCount > 0) return getStatusDescriptor('error', '✕ נכשל');
        if (warningCount > 0) return getStatusDescriptor('warning', '⚠ בוצע עם אזהרות');
        return getStatusDescriptor('success', '✓ בוצע');
    }
    if (rawStatus === 'failed') return getStatusDescriptor('error', '✕ נכשל');
    if (rawStatus === 'partial_success') {
        return errorCount > 0
            ? getStatusDescriptor('error', '✕ נכשל')
            : getStatusDescriptor('warning', '⚠ בוצע עם אזהרות');
    }
    return getStatusDescriptor('pending', '⏳ ממתין לעיבוד');
}

function sheetStatusDescriptor(sheet) {
    const text = String(sheet?.status || '').trim();
    const lower = text.toLowerCase();
    if (!text || lower === 'unknown') return getStatusDescriptor('pending', '⏳ ממתין לעיבוד');
    if (text.includes('נכשל') || lower.includes('failed')) return getStatusDescriptor('error', '✕ נכשל');
    if (text.includes('אזהרות') || lower.includes('warning')) return getStatusDescriptor('warning', '⚠ בוצע עם אזהרות');
    if (text.includes('בוצע') || lower.includes('completed') || lower === 'success') return getStatusDescriptor('success', '✓ בוצע');
    return getStatusDescriptor('pending', '⏳ ממתין לעיבוד');
}

function renderMetricCard(label, value, options = {}) {
    const html = Boolean(options.html);
    const className = options.className ? ` ${options.className}` : '';
    return `
        <div class="report-meta report-metric-card${className}">
            <span class="report-meta-label">${escapeHtml(label)}</span>
            <strong>${html ? value : escapeHtml(value)}</strong>
        </div>
    `;
}

function renderRowsList(rowNumbers) {
    const values = Array.isArray(rowNumbers) ? rowNumbers.filter(v => v !== null && v !== undefined && String(v).trim() !== '') : [];
    if (!values.length) return '<span class="report-issue-rows-empty">—</span>';
    return values
        .map(value => `<span class="report-row-pill">${escapeHtml(String(value))}</span>`)
        .join('');
}

function renderIssueGroup(group) {
    if (!group) return '';
    const severity = ['error', 'warning', 'info'].includes(group.severity) ? group.severity : 'warning';
    const rowNumbers = Array.isArray(group.row_numbers) ? group.row_numbers : [];
    const fieldNames = Array.isArray(group.field_names) ? group.field_names : [];
    const count = Number(group.count || rowNumbers.length || 0);
    const label = group.label || '—';
    const rowText = rowNumbers.length ? rowNumbers.join(', ') : '—';

    return `
        <div class="report-issue-group report-issue-group--${severity}">
            <div class="report-issue-group-head">
                <div class="report-issue-title">${escapeHtml(label)}</div>
                <div class="report-issue-count">${formatNumber(count)}</div>
            </div>
            <div class="report-issue-meta">
                <span>שורות:</span>
                <strong>${escapeHtml(rowText)}</strong>
            </div>
            ${fieldNames.length ? `
                <div class="report-issue-meta">
                    <span>שדות:</span>
                    <strong>${escapeHtml(fieldNames.join(', '))}</strong>
                </div>
            ` : ''}
        </div>
    `;
}

function renderSheetCard(sheet, index = 0) {
    const status = sheetStatusDescriptor(sheet);
    const issueGroups = Array.isArray(sheet.issue_groups) ? sheet.issue_groups : [];
    const hasIssues = issueGroups.length > 0;
    const openAttr = hasIssues || index === 0 ? ' open' : '';
    const summarySource = formatNumber(sheet.source_row_count ?? sheet.row_count ?? 0);
    const summaryCurrent = formatNumber(sheet.current_row_count ?? sheet.row_count ?? 0);
    const summaryDeleted = formatNumber(sheet.rows_deleted ?? 0);
    const summaryManual = formatNumber(sheet.rows_changed_manually ?? 0);
    const summaryAuto = formatNumber(sheet.rows_changed_automatically ?? 0);

    return `
        <details class="report-sheet-card"${openAttr}>
            <summary class="report-sheet-summary">
                <div class="report-sheet-summary-main">
                    <strong>${escapeHtml(sheet.sheet_name || '—')}</strong>
                    ${renderStatusBadgeHtml(status.kind, status.text)}
                </div>
                <div class="report-sheet-summary-sub">
                    <span>מקור: ${summarySource}</span>
                    <span>נוכחיות: ${summaryCurrent}</span>
                    ${Number(sheet.rows_deleted || 0) > 0 ? `<span>נמחקו: ${summaryDeleted}</span>` : ''}
                    <span>ידניות: ${summaryManual}</span>
                    <span>אוטומטיות: ${summaryAuto}</span>
                </div>
            </summary>
            <div class="report-sheet-body">
                <div class="report-sheet-metrics">
                    ${renderMetricCard('שורות בקובץ המקור', summarySource)}
                    ${renderMetricCard('שורות נוכחיות', summaryCurrent)}
                    ${Number(sheet.rows_deleted || 0) > 0 ? renderMetricCard('שורות שנמחקו', summaryDeleted) : ''}
                    ${renderMetricCard('שורות שנערכו ידנית', summaryManual)}
                    ${renderMetricCard('פעולות עריכה ידניות', formatNumber(sheet.rows_changed_manually || 0))}
                    ${renderMetricCard('שורות שעודכנו אוטומטית', summaryAuto)}
                    ${renderMetricCard('בעיות', formatNumber((sheet.rows_with_warnings || 0) + (sheet.rows_with_errors || 0)))}
                </div>
                <div class="report-issue-list">
                    ${hasIssues
                        ? issueGroups.map(renderIssueGroup).join('')
                        : '<div class="report-empty">אין אזהרות או שגיאות בגיליון זה.</div>'}
                </div>
            </div>
        </details>
    `;
}

function renderIssuesTable(report) {
    const issues = Array.isArray(report?.issues) ? report.issues : [];
    const rows = issues.map(issue => `
        <tr>
            <td>${escapeHtml(issue.severity || '')}</td>
            <td>${escapeHtml(issue.sheet_name || '')}</td>
            <td>${escapeHtml(issue.row_number ? String(issue.row_number) : '—')}</td>
            <td>${escapeHtml(issue.field_name || issue.status_field || '')}</td>
            <td>${escapeHtml(issue.status_message || '')}</td>
        </tr>
    `).join('');
    return `
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
                    ${rows || '<tr><td colspan="5">אין בעיות להצגה.</td></tr>'}
                </tbody>
            </table>
        </div>
    `;
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
    if (missingRequired) parts.push(`שדות חסרים: ${missingRequired}`);
    if (summary.source_rows != null) parts.push(`שורות בקובץ המקור: ${summary.source_rows}`);
    if (summary.current_rows != null) parts.push(`שורות נוכחיות: ${summary.current_rows}`);
    if (summary.rows_deleted) parts.push(`שורות שנמחקו: ${summary.rows_deleted}`);
    if (summary.manual_edit_rows != null) parts.push(`שורות שנערכו ידנית: ${summary.manual_edit_rows}`);
    if (summary.manual_edit_actions != null) parts.push(`פעולות עריכה ידניות: ${summary.manual_edit_actions}`);
    if (summary.rows_with_warnings != null) parts.push(`אזהרות: ${summary.rows_with_warnings}`);
    if (summary.rows_with_errors != null) parts.push(`שגיאות: ${summary.rows_with_errors}`);
    if (summary.corrected_fields != null) parts.push(`שדות מתוקנים: ${summary.corrected_fields}`);
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
    const currentSheet = state.currentSheet || '';
    const statusDescriptor = reportStatusDescriptor(report);
    const issueCount = Number(s.rows_with_warnings || 0) + Number(s.rows_with_errors || 0);

    status.className = 'report-status';
    status.innerHTML = renderStatusBadgeHtml(statusDescriptor.kind, statusDescriptor.text);
    status.title = statusDescriptor.text;
    status.setAttribute('aria-label', statusDescriptor.text);

    summary.innerHTML = `
        <div class="report-summary-grid">
            ${renderMetricCard('סטטוס', renderStatusBadgeHtml(statusDescriptor.kind, statusDescriptor.text), { html: true })}
            ${renderMetricCard('גיליון נוכחי', currentSheet || 'לא נבחר')}
            ${renderMetricCard('גיליונות', formatNumber(s.total_sheets || 0))}
            ${renderMetricCard('שורות בקובץ המקור', formatNumber(s.source_rows ?? s.total_rows ?? 0))}
            ${renderMetricCard('שורות נוכחיות', formatNumber(s.current_rows ?? 0))}
            ${Number(s.rows_deleted || 0) > 0 ? renderMetricCard('שורות שנמחקו', formatNumber(s.rows_deleted || 0)) : ''}
            ${renderMetricCard('שורות שנערכו ידנית', formatNumber(s.manual_edit_rows || 0))}
            ${renderMetricCard('פעולות עריכה ידניות', formatNumber(s.manual_edit_actions || 0))}
            ${renderMetricCard('שורות שעודכנו אוטומטית', formatNumber(s.rows_changed_automatically || 0))}
            ${renderMetricCard('אזהרות', formatNumber(s.rows_with_warnings || 0))}
            ${renderMetricCard('שגיאות', formatNumber(s.rows_with_errors || 0))}
            ${renderMetricCard('בעיות', formatNumber(issueCount))}
        </div>
    `;

    const blocked = report.export_blocked_reason
        ? `<div class="report-blocked">${escapeHtml(report.export_blocked_reason)}</div>`
        : '';
    const warnings = (issueCount > 0 || report.dirty || report.stale) && !report.export_blocked_reason
        ? `<div class="report-note">הדוח מסונכרן עם מצב הקובץ הנוכחי. ניתן לפתוח כל גיליון כדי לראות פירוט שורות ואזהרות.</div>`
        : '';

    details.innerHTML = `${blocked}${warnings}${(report.sheets || []).map((sheet, index) => renderSheetCard(sheet, index)).join('')}`;
    section.classList.remove('hidden');
}

function renderProcessingReportPreview(report) {
    const container = document.getElementById('processing-report-preview');
    if (!container || !report) return;

    const summary = report.summary || {};
    const manual = report.manual_edits || {};
    const sheetBlocks = (report.sheets || []).map((sheet, index) => renderSheetCard(sheet, index)).join('');

    container.innerHTML = `
        <div class="report-preview-summary">
            ${renderMetricCard('קובץ', report.file_name || '—')}
            ${renderMetricCard('סטטוס', renderStatusBadgeHtml(reportStatusDescriptor(report).kind, reportStatusDescriptor(report).text), { html: true })}
            ${renderMetricCard('מוכן לייצוא', report.export_ready ? 'כן' : 'לא')}
            ${renderMetricCard('עריכות לאחר תקנון', report.dirty ? 'כן' : 'לא')}
            ${renderMetricCard('גיליונות', formatNumber(summary.total_sheets || 0))}
            ${renderMetricCard('שורות בקובץ המקור', formatNumber(summary.source_rows || summary.total_rows || 0))}
            ${renderMetricCard('שורות נוכחיות', formatNumber(summary.current_rows || 0))}
            ${renderMetricCard('שורות שנערכו ידנית', formatNumber(summary.manual_edit_rows || 0))}
            ${renderMetricCard('פעולות עריכה ידניות', formatNumber(summary.manual_edit_actions || 0))}
            ${renderMetricCard('אזהרות', formatNumber(summary.rows_with_warnings || 0))}
            ${renderMetricCard('שגיאות', formatNumber(summary.rows_with_errors || 0))}
            ${renderMetricCard('שורות שנמחקו', formatNumber(summary.rows_deleted || 0))}
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
                <div><span>rows</span><strong>${manual.edited_rows || 0}</strong></div>
                <div><span>actions</span><strong>${manual.edited_actions || 0}</strong></div>
                <div><span>sheets</span><strong>${(manual.edited_sheets || []).join(', ') || '—'}</strong></div>
                <div><span>fields</span><strong>${(manual.edited_fields || []).join(', ') || '—'}</strong></div>
            </div>
        </div>
        <div class="report-preview-section">
            <div class="report-subtitle">סיכום גיליונות</div>
            <div class="report-preview-sheet-list">${sheetBlocks || '<div class="report-empty">אין נתוני גיליונות להצגה.</div>'}</div>
        </div>
        ${report.issues && report.issues.length ? `
            <div class="report-preview-section">
                <div class="report-subtitle">אזהרות ושגיאות</div>
                ${renderIssuesTable(report)}
            </div>
        ` : ''}
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
    renderStatusBadgeHtml,
    getStatusDescriptor,
});
