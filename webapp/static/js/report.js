/**
 * Processing report UI helpers.
 * Kept separate from the main app shell so summary formatting can evolve
 * without dragging the grid and upload code with it.
 */

function formatProcessingReportSummary(report) {
    const parts = [];
    const missingRequired = (report.missing_required_fields || [])
        .map(item => `${item.field}: ${item.count}`)
        .join(', ');
    const dateSummary = (report.date_summary || [])
        .map(item => `${item.message}: ${item.count}`)
        .join(', ');
    const identifierSummary = (report.identifier_summary || [])
        .map(item => `${item.message}: ${item.count}`)
        .join(', ');

    if (missingRequired) parts.push(`Missing required fields: ${missingRequired}`);
    if (dateSummary) parts.push(`Dates: ${dateSummary}`);
    if (identifierSummary) parts.push(`Identifiers: ${identifierSummary}`);

    return parts.length ? parts.join(' | ') : report.status_reason;
}
