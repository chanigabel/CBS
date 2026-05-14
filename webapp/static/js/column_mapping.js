// ---------------------------------------------------------------------------
// Inline column header mapping
// ---------------------------------------------------------------------------

let columnMappingSchema = null;

function isSourceColumnName(name) {
    if (!name || name.startsWith('_')) return false;
    if (name.endsWith('_corrected') || name.endsWith('_status')) return false;
    if (name === 'identifier_status' || name === 'gender_status') return false;
    if (name === 'birth_date_status' || name === 'entry_date_status') return false;
    if (name === 'MosadID' || name === 'SugMosad') return false;
    return true;
}

async function loadColumnMappingSchema(force = false) {
    if (columnMappingSchema && !force) return columnMappingSchema;
    columnMappingSchema = await apiCall('GET', '/api/workbook/column-schema');
    return columnMappingSchema;
}

async function refreshColumnMappingControls(sheetData) {
    await loadColumnMappingSchema(true);
    document.querySelectorAll('.editable-column-label').forEach(label => {
        label.title = 'Click to map this header to a standard field';
    });
}

function closeHeaderMappingPopover() {
    document.querySelectorAll('.column-mapping-popover').forEach(popover => popover.remove());
}

async function beginHeaderMappingEdit(labelEl, columnName) {
    if (!state.sessionId || !state.currentSheet || !isSourceColumnName(columnName)) return;
    closeHeaderMappingPopover();

    let schema;
    try {
        schema = await loadColumnMappingSchema();
    } catch (err) {
        showError(`Column schema failed to load: ${err.message}`);
        return;
    }

    const canonicalFields = schema.fields || [];
    if (!canonicalFields.length) {
        showError('No column mapping schema is configured.');
        return;
    }

    const popover = document.createElement('div');
    popover.className = 'column-mapping-popover';
    popover.addEventListener('mousedown', event => event.stopPropagation());

    const title = document.createElement('div');
    title.className = 'column-mapping-title';
    title.textContent = `Map "${columnName}"`;
    popover.appendChild(title);

    const select = document.createElement('select');
    select.className = 'column-mapping-select';
    canonicalFields.forEach(field => {
        const option = document.createElement('option');
        option.value = field;
        option.textContent = field;
        if (field === columnName) option.selected = true;
        select.appendChild(option);
    });
    popover.appendChild(select);

    const actions = document.createElement('div');
    actions.className = 'column-mapping-actions';

    const applyBtn = document.createElement('button');
    applyBtn.type = 'button';
    applyBtn.className = 'column-mapping-apply';
    applyBtn.textContent = 'Apply Mapping';
    applyBtn.addEventListener('click', async () => {
        try {
            await saveInlineColumnMapping(columnName, select.value);
            closeHeaderMappingPopover();
        } catch (err) {
            showError(`Column mapping failed: ${err.message}`);
        }
    });

    const cancelBtn = document.createElement('button');
    cancelBtn.type = 'button';
    cancelBtn.className = 'column-mapping-cancel';
    cancelBtn.textContent = 'Cancel';
    cancelBtn.addEventListener('click', closeHeaderMappingPopover);

    actions.append(applyBtn, cancelBtn);
    popover.appendChild(actions);
    document.body.appendChild(popover);

    const rect = labelEl.getBoundingClientRect();
    popover.style.top = `${rect.bottom + window.scrollY + 6}px`;
    popover.style.left = `${Math.max(12, rect.left + window.scrollX)}px`;
    select.focus();

    function onOutsideClick(event) {
        if (!popover.contains(event.target) && event.target !== labelEl) {
            closeHeaderMappingPopover();
            document.removeEventListener('mousedown', onOutsideClick, true);
        }
    }
    document.addEventListener('mousedown', onOutsideClick, true);
}

async function saveInlineColumnMapping(oldName, targetName) {
    const result = await apiCall(
        'POST',
        `/api/workbook/${state.sessionId}/sheet/${encodeURIComponent(state.currentSheet)}/column-mapping`,
        { old_name: oldName, new_name: targetName }
    );

    const newName = result.new_name;
    if (state.sheetData) {
        state.sheetData.field_names = result.field_names;
        state.sheetData.column_mappings = result.column_mappings;
        state.sheetData.rows.forEach(row => {
            if (Object.prototype.hasOwnProperty.call(row, oldName)) {
                row[newName] = row[oldName];
                delete row[oldName];
            }
        });
    }
    if (state.columnFilters.has(oldName)) {
        const activeFilter = state.columnFilters.get(oldName);
        state.columnFilters.delete(oldName);
        state.columnFilters.set(newName, activeFilter);
    }
    await loadSheet(state.currentSheet);
}

async function reloadColumnMappingForCurrentSheet() {
    if (!state.sessionId || !state.currentSheet) return;
    await apiCall(
        'POST',
        `/api/workbook/${state.sessionId}/sheet/${encodeURIComponent(state.currentSheet)}/reload-mapping`
    );
    await loadColumnMappingSchema(true);
    await loadSheet(state.currentSheet);
}

Object.assign(window, {
    isSourceColumnName,
    loadColumnMappingSchema,
    refreshColumnMappingControls,
    beginHeaderMappingEdit,
    saveInlineColumnMapping,
    reloadColumnMappingForCurrentSheet,
});
