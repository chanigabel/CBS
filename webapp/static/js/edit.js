async function deleteSingleRow(rowUid) { await _deleteRows([rowUid]); }

async function deleteSelectedRows() {
    if (state.selectedRows.size === 0) return;
    const n = state.selectedRows.size;
    // Confirm multi-row deletes.
    if (n > 1) {
        const confirmed = confirm(
            `למחוק ${n} שורות?\n\nלא ניתן לבטל פעולה זו ללא טעינה מחדש של הקובץ.`
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
        state.sheetData.rows = state.sheetData.rows.filter(r => !uidSet.has(getRowUid(r)));

        // Clear deleted rows from the selection set.
        uidSet.forEach(uid => state.selectedRows.delete(uid));
        const session = sessions.get(state.sessionId);
        if (session) session.hasEdits = true;

        const filtered = getFilteredRows(state.sheetData.rows);
        renderGrid(state.sheetData, filtered);
        document.getElementById('grid-stats').textContent =
            `נמחקו ${result.deleted_count} שורות. נותרו ${result.remaining_rows} שורות.`;
    } catch (err) {
        showError(`מחיקת השורות נכשלה: ${err.message}`);
        updateDeleteButton();
    }
}

// ---------------------------------------------------------------------------
// Inline Cell Editing
// ---------------------------------------------------------------------------

function makeEditable(td, rowUid, fieldName) {
    if (td.querySelector('input[type=text]')) return;

    const currentValue = td.textContent;
    state.focusedEditColumn = fieldName;
    refreshMultiEditHighlights();
    td.className = (td.className || '') + ' editing';
    td.textContent = '';

    const input = document.createElement('input');
    input.type = 'text';
    input.value = currentValue;
    td.appendChild(input);
    input.focus();
    input.select();

    function selectedEditRowUids() {
        if (!state.selectedRows.has(rowUid) || state.selectedRows.size <= 1) return [rowUid];
        return [...state.selectedRows];
    }

    function rowByUid(uid) {
        return state.sheetData?.rows.find(r => getRowUid(r) === uid);
    }

    function setGridMessage(message) {
        const stats = document.getElementById('grid-stats');
        if (stats) stats.textContent = message;
    }

    async function commitEdit() {
        const newValue = input.value;
        if (newValue === currentValue) {
            td.textContent = currentValue;
            td.className = td.className.replace(' editing', '');
            return;
        }
        const affectedUids = selectedEditRowUids();
        const oldValues = {};
        affectedUids.forEach(uid => {
            const row = rowByUid(uid);
            oldValues[uid] = row ? row[fieldName] : '';
        });
        try {
            if (affectedUids.length > 1) {
                const response = await apiCall(
                    'PATCH',
                    `/api/workbook/${state.sessionId}/sheet/${encodeURIComponent(state.currentSheet)}/multi-edit`,
                    { row_uids: affectedUids, field_name: fieldName, new_value: newValue }
                );
                Object.entries(response.updated_rows || {}).forEach(([uid, updatedRow]) => {
                    const row = rowByUid(uid);
                    if (row) Object.assign(row, updatedRow);
                });
                setGridMessage(`עודכנו ${response.edited_count || affectedUids.length} תאים בעמודה ${fieldName}`);
            } else {
                const response = await apiCall(
                    'PATCH',
                    `/api/workbook/${state.sessionId}/cell`,
                    {
                        sheet_name: state.currentSheet,
                        row_uid: rowUid,
                        field: fieldName,
                        value: newValue
                    }
                );
                const editedRow = rowByUid(rowUid);
                const updatedRow = response && response.updated_row ? response.updated_row : null;
                if (editedRow) {
                    if (updatedRow) Object.assign(editedRow, updatedRow);
                    else editedRow[fieldName] = newValue;
                }
                setGridMessage(`עודכן תא בעמודה ${fieldName}`);
            }
            state.undoStack.push({ sheetName: state.currentSheet, rowUids: affectedUids, fieldName, oldValues });
            updateUndoButton();
            renderGrid(state.sheetData, getFilteredRows(state.sheetData.rows));
            markUpdatedCells(affectedUids, fieldName);
            const session = sessions.get(state.sessionId);
            if (session) session.hasEdits = true;
        } catch (err) {
            showError(`עריכת התא נכשלה: ${err.message}`);
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

async function undoLastGridEdit() {
    if (!state.sessionId) return;
    if (state.undoStack.length === 0) {
        const stats = document.getElementById('grid-stats');
        if (stats) stats.textContent = 'אין שינוי לביטול';
        updateUndoButton();
        return;
    }
    const active = document.activeElement;
    if (active && active.matches && active.matches('input[type="text"], textarea')) return;

    const action = state.undoStack.pop();
    try {
        await Promise.all(action.rowUids.map(uid => apiCall(
            'PATCH',
            `/api/workbook/${state.sessionId}/cell`,
            {
                sheet_name: action.sheetName,
                row_uid: uid,
                field: action.fieldName,
                value: action.oldValues[uid] === null || action.oldValues[uid] === undefined
                    ? ''
                    : String(action.oldValues[uid])
            }
        ).then(response => {
            const row = state.sheetData?.rows.find(r => getRowUid(r) === uid);
            if (row && response?.updated_row) Object.assign(row, response.updated_row);
            else if (row) row[action.fieldName] = action.oldValues[uid];
        })));
        renderGrid(state.sheetData, getFilteredRows(state.sheetData.rows));
        markUpdatedCells(action.rowUids, action.fieldName);
        const stats = document.getElementById('grid-stats');
        if (stats) stats.textContent = 'השינוי האחרון בוטל';
        updateUndoButton();
    } catch (err) {
        showError(`ביטול השינוי נכשל: ${err.message}`);
        state.undoStack.push(action);
        updateUndoButton();
    }
}

function updateUndoButton() {
    const btn = document.getElementById('undo-edit-btn');
    if (!btn) return;
    const enabled = Boolean(state.sessionId && state.undoStack.length > 0);
    btn.disabled = !enabled;
    btn.title = enabled
        ? 'בטל את השינוי הידני האחרון (Ctrl+Z)'
        : 'אין שינוי לביטול (Ctrl+Z)';
    btn.setAttribute('aria-label', btn.title);
}

// ---------------------------------------------------------------------------
// Editing helpers
// ---------------------------------------------------------------------------


Object.assign(window, { deleteSingleRow, deleteSelectedRows, _deleteRows, makeEditable, undoLastGridEdit, updateUndoButton });
