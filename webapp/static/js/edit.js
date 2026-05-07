async function deleteSingleRow(rowUid) { await _deleteRows([rowUid]); }

async function deleteSelectedRows() {
    if (state.selectedRows.size === 0) return;
    const n = state.selectedRows.size;
    // Confirm multi-row deletes.
    if (n > 1) {
        const confirmed = confirm(
            `Delete ${n} rows?\n\nThis cannot be undone without re-running standardization.`
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

        // Clear deleted rows from the selection set.
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
            // Update the cached row data.
            const editedRow = state.sheetData?.rows.find(r => r._row_uid === rowUid);
            if (editedRow) editedRow[fieldName] = newValue;
            td.textContent = newValue;
            td.className = td.className.replace(' editing', '');
            if (fieldName.endsWith('_corrected')) {
                // Compare normalized values to avoid false highlights.
                const origVal = editedRow ? editedRow[fieldName.replace(/_corrected$/, '')] : null;
                const origStr = (origVal !== null && origVal !== undefined) ? String(origVal).trim() : '';
                td.className = (newValue.trim() !== '' && newValue.trim() !== origStr)
                    ? 'corrected-changed' : 'corrected-cell';
            }
            // Mark the session as dirty.
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
// Editing helpers
// ---------------------------------------------------------------------------


Object.assign(window, { deleteSingleRow, deleteSelectedRows, _deleteRows, makeEditable });
