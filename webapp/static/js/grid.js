function getFilteredRows(rows) {
    if (state.columnFilters.size === 0) return rows;
    return rows.filter(row => {
        for (const [col, values] of state.columnFilters) {
            if (values.size === 0) continue;
            const cell = row[col];
            const cellStr = (cell !== null && cell !== undefined) ? String(cell).trim() : '';
            if (!values.has(cellStr)) return false;
        }
        return true;
    });
}

/**
 * Get sorted distinct string values for a column across all (unfiltered) rows.
 */
function getDistinctValues(col) {
    if (!state.sheetData) return [];
    const seen = new Set();
    state.sheetData.rows.forEach(row => {
        const v = row[col];
        seen.add((v !== null && v !== undefined) ? String(v).trim() : '');
    });
    return [...seen].sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
}

/**
 * Open (or close) the filter dropdown for a column.
 * Closes any other open dropdown first.
 */
function openFilterDropdown(col, btnEl) {
    // Close any existing open dropdown
    const existing = document.querySelector('.col-filter-dropdown');
    if (existing) {
        const wasForSameCol = existing.dataset.col === col;
        existing.remove();
        if (wasForSameCol) return;   // toggle: clicking again closes it
    }

    const values = getDistinctValues(col);
    const activeSet = state.columnFilters.get(col) || new Set();

    const dropdown = document.createElement('div');
    dropdown.className = 'col-filter-dropdown';
    dropdown.dataset.col = col;

    // Search box (only shown when there are many values)
    if (values.length > 8) {
        const search = document.createElement('input');
        search.type = 'text';
        search.placeholder = 'חיפוש...';
        search.className = 'col-filter-search';
        search.addEventListener('input', () => {
            const q = search.value.toLowerCase();
            dropdown.querySelectorAll('.col-filter-item').forEach(item => {
                item.style.display = item.dataset.val.toLowerCase().includes(q) ? '' : 'none';
            });
        });
        dropdown.appendChild(search);
    }

    // Value list
    const list = document.createElement('div');
    list.className = 'col-filter-list';

    values.forEach(val => {
        const label = document.createElement('label');
        label.className = 'col-filter-item';
        label.dataset.val = val;

        const cb = document.createElement('input');
        cb.type = 'checkbox';
        cb.checked = activeSet.has(val);
        cb.addEventListener('change', () => {
            let set = state.columnFilters.get(col);
            if (!set) { set = new Set(); state.columnFilters.set(col, set); }
            if (cb.checked) set.add(val); else set.delete(val);
            if (set.size === 0) state.columnFilters.delete(col);
            applyFilters();
            // Update the filter button active state without closing the dropdown
            updateFilterButtonState(col);
        });

        const text = document.createElement('span');
        text.textContent = val === '' ? '(ריק)' : val;
        if (val === '') text.style.fontStyle = 'italic';

        label.appendChild(cb);
        label.appendChild(text);
        list.appendChild(label);
    });
    dropdown.appendChild(list);

    // Footer: clear this column's filter
    const footer = document.createElement('div');
    footer.className = 'col-filter-footer';
    const clearBtn = document.createElement('button');
    clearBtn.textContent = 'נקה סינון';
    clearBtn.className = 'col-filter-clear-btn';
    clearBtn.addEventListener('click', () => {
        state.columnFilters.delete(col);
        applyFilters();
        updateFilterButtonState(col);
        dropdown.querySelectorAll('input[type=checkbox]').forEach(cb => cb.checked = false);
    });
    footer.appendChild(clearBtn);
    dropdown.appendChild(footer);

    // Position below the button
    document.body.appendChild(dropdown);
    const rect = btnEl.getBoundingClientRect();
    dropdown.style.top = (rect.bottom + window.scrollY + 2) + 'px';
    // RTL: align right edge of dropdown to right edge of button
    const dropW = dropdown.offsetWidth;
    dropdown.style.left = (rect.right + window.scrollX - dropW) + 'px';

    // Close on outside click
    function onOutsideClick(e) {
        if (!dropdown.contains(e.target) && e.target !== btnEl) {
            dropdown.remove();
            document.removeEventListener('mousedown', onOutsideClick, true);
        }
    }
    document.addEventListener('mousedown', onOutsideClick, true);
}

function updateFilterButtonState(col) {
    const btn = document.querySelector(`.col-filter-btn[data-col="${CSS.escape(col)}"]`);
    if (!btn) return;
    const active = state.columnFilters.has(col) && state.columnFilters.get(col).size > 0;
    btn.classList.toggle('col-filter-active', active);
    btn.title = active ? 'סינון פעיל — לחץ לעריכה' : 'סנן לפי עמודה זו';
}

/**
 * Re-render only the tbody rows based on current filters.
 * Avoids a full table rebuild for performance.
 */
function applyFilters() {
    if (!state.sheetData) return;
    const filtered = getFilteredRows(state.sheetData.rows);

    // Update stats
    const statsDiv = document.getElementById('grid-stats');
    const total = state.sheetData.rows.length;
    const shown = filtered.length;
    const cols = state.sheetData.field_names.length;
    if (state.columnFilters.size > 0) {
        statsDiv.textContent = `מציג ${shown} מתוך ${total} שורות × ${cols} עמודות`;
    } else {
        statsDiv.textContent = `${total} rows × ${cols} columns`;
    }

    // Re-render the grid with filtered rows
    renderGrid(state.sheetData, filtered);

    // Update "clear all" button visibility
    const clearAllBtn = document.getElementById('clear-all-filters-btn');
    if (clearAllBtn) clearAllBtn.classList.toggle('hidden', state.columnFilters.size === 0);
}

function clearAllFilters() {
    state.columnFilters.clear();
    // Close any open dropdown
    const existing = document.querySelector('.col-filter-dropdown');
    if (existing) existing.remove();
    applyFilters();
}

// ---------------------------------------------------------------------------
// Full-screen grid overlay
// ---------------------------------------------------------------------------

function openGridOverlay() {
    const overlay          = document.getElementById('grid-overlay');
    const overlayContainer = document.getElementById('grid-overlay-container');
    const overlayTitle     = document.getElementById('grid-overlay-title');
    const overlayStats     = document.getElementById('grid-overlay-stats');
    if (!overlay || !overlayContainer || !state.sheetData) return;

    // Copy current title and stats into the overlay bar
    overlayTitle.textContent = document.getElementById('grid-title')?.textContent || '';
    const total = state.sheetData.rows.length;
    const filtered = getFilteredRows(state.sheetData.rows);
    const shown = filtered.length;
    overlayStats.textContent = state.columnFilters.size > 0
        ? `מציג ${shown} מתוך ${total} שורות`
        : `${total} שורות`;

    // Render the full interactive grid into the overlay container
    renderGrid(state.sheetData, filtered, overlayContainer);

    overlay.classList.remove('hidden');
    document.body.classList.add('grid-overlay-open');
    document.addEventListener('keydown', _overlayEscHandler);
}

function closeGridOverlay() {
    const overlay = document.getElementById('grid-overlay');
    if (!overlay) return;
    overlay.classList.add('hidden');
    document.body.classList.remove('grid-overlay-open');
    document.removeEventListener('keydown', _overlayEscHandler);

    // Sync any changes made inside the overlay back to the normal grid
    if (state.sheetData) {
        renderGrid(state.sheetData, getFilteredRows(state.sheetData.rows));
    }
}

function _overlayEscHandler(e) {
    if (e.key === 'Escape') closeGridOverlay();
}

function renderGrid(sheetData, rows, targetContainer) {
    const container = targetContainer || document.getElementById('grid-container');
    const statsDiv  = targetContainer ? null : document.getElementById('grid-stats');

    const displayRows = rows !== undefined ? rows : sheetData.rows;

    if (!sheetData.rows || sheetData.rows.length === 0) {
        container.innerHTML = '<p style="padding:20px;text-align:center">לא נמצאו שורות בגיליון זה.</p>';
        if (statsDiv) statsDiv.textContent = '';
        if (!targetContainer) updateDeleteButton();
        return;
    }

    const displayColumns = sheetData.field_names;

    function colClass(col) {
        if (col.endsWith('_corrected')) return 'corrected';
        if (col.endsWith('_status'))   return 'status';
        return 'original';
    }

    function isEditableDataColumn(col) {
        if (!col) return false;
        if (col === '_row_uid' || col === 'row_uid') return false;
        // Keep only true internal fields blocked.
        if (col === '_standardization_failures') return false;
        // Allow visible validation/status fields (e.g. '_validation_status')
        // to be editable if they are included in the displayed columns.
        if (col.startsWith('_')) {
            if (col === '_validation_status') return true;
            return false;
        }
        // Allow source fields, *_corrected and *_status fields to be editable.
        return true;
    }

    const table = document.createElement('table');
    table.className = 'data-grid';

    // Header
    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');

    const thCheck = document.createElement('th');
    thCheck.className = 'col-select';
    const selectAll = document.createElement('input');
    selectAll.type = 'checkbox';
    selectAll.title = 'Select all rows';
    selectAll.addEventListener('change', () => toggleSelectAll(selectAll.checked, displayRows));
    thCheck.appendChild(selectAll);
    headerRow.appendChild(thCheck);

    const thDel = document.createElement('th');
    thDel.className = 'col-del';
    headerRow.appendChild(thDel);

    displayColumns.forEach(col => {
        const th = document.createElement('th');
        const cls = colClass(col);
        if (cls === 'corrected') th.className = 'corrected-header';
        else if (cls === 'status') th.className = 'status-header';

        // Column label + filter button wrapper
        const headerWrap = document.createElement('div');
        headerWrap.className = 'col-header-wrap';

        const label = document.createElement('span');
        const displayName = sheetData.column_display_names && sheetData.column_display_names[col]
            ? sheetData.column_display_names[col]
            : col;
        label.textContent = displayName;
        if (typeof isSourceColumnName === 'function' && isSourceColumnName(col)) {
            label.className = 'editable-column-label';
            label.title = 'Click to map this header to a standard field';
            label.addEventListener('click', e => {
                e.stopPropagation();
                if (typeof beginHeaderMappingEdit === 'function') {
                    beginHeaderMappingEdit(label, col);
                }
            });
        }
        headerWrap.appendChild(label);

        const filterBtn = document.createElement('button');
        filterBtn.className = 'col-filter-btn';
        filterBtn.dataset.col = col;
        const isActive = state.columnFilters.has(col) && state.columnFilters.get(col).size > 0;
        if (isActive) filterBtn.classList.add('col-filter-active');
        filterBtn.title = isActive ? 'סינון פעיל — לחץ לעריכה' : 'סנן לפי עמודה זו';
        filterBtn.textContent = '▾';
        filterBtn.addEventListener('click', e => {
            e.stopPropagation();
            openFilterDropdown(col, filterBtn);
        });
        headerWrap.appendChild(filterBtn);

        th.appendChild(headerWrap);
        headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);
    table.appendChild(thead);

    // Body
    const tbody = document.createElement('tbody');
    displayRows.forEach((row) => {
        const tr = document.createElement('tr');
        tr.dataset.rowUid = row._row_uid;

        const tdCheck = document.createElement('td');
        tdCheck.className = 'col-select';
        const cb = document.createElement('input');
        cb.type = 'checkbox';
        cb.checked = state.selectedRows.has(row._row_uid);
        cb.addEventListener('change', () => {
            if (cb.checked) { state.selectedRows.add(row._row_uid); tr.classList.add('row-selected'); }
            else            { state.selectedRows.delete(row._row_uid); tr.classList.remove('row-selected'); }
            updateDeleteButton();
            const visibleRows = document.querySelectorAll('.data-grid tbody tr').length;
            selectAll.checked = state.selectedRows.size === visibleRows;
            selectAll.indeterminate = state.selectedRows.size > 0 && state.selectedRows.size < visibleRows;
        });
        if (cb.checked) tr.classList.add('row-selected');
        tdCheck.appendChild(cb);
        tr.appendChild(tdCheck);

        const changedFields = displayColumns.filter(col => {
            if (!col.endsWith('_corrected')) return false;
            const origVal = row[col.replace(/_corrected$/, '')];
            const corrVal = row[col];
            if (corrVal === null || corrVal === undefined) return false;
            const origStr = (origVal !== null && origVal !== undefined) ? String(origVal).trim() : '';
            const corrStr = String(corrVal).trim();
            return corrStr !== '' && corrStr !== origStr;
        });

        const tdDel = document.createElement('td');
        tdDel.className = 'col-del';
        const delBtn = document.createElement('button');
        delBtn.className = 'btn-row-delete';
        delBtn.textContent = '✕';
        delBtn.title = 'Delete this row';
        delBtn.addEventListener('click', () => deleteSingleRow(row._row_uid));
        tdDel.appendChild(delBtn);

        if (changedFields.length > 0) {
            tr.classList.add('row-has-changes');
            const badge = document.createElement('span');
            badge.className = 'change-badge';
            badge.textContent = changedFields.length;
            const fieldNames = changedFields.map(f => f.replace(/_corrected$/, '')).join(', ');
            badge.title = `${changedFields.length} field(s) corrected: ${fieldNames}`;
            tdDel.appendChild(badge);
        }

        tr.appendChild(tdDel);

        displayColumns.forEach(col => {
            const td = document.createElement('td');
            const value = row[col];
            td.textContent = value !== null && value !== undefined ? String(value) : '';

            const cls = colClass(col);
            if (cls === 'corrected') {
                // type mismatches (e.g. gender original="ז" str vs corrected=1 int).
                const origVal = row[col.replace(/_corrected$/, '')];
                const origStr = (origVal !== null && origVal !== undefined) ? String(origVal).trim() : '';
                const corrStr = (value !== null && value !== undefined) ? String(value).trim() : '';
                td.className = (corrStr !== '' && corrStr !== origStr)
                    ? 'corrected-changed' : 'corrected-cell';
            } else if (cls === 'status') {
                // Highlight non-empty status cells.
                const statusText = String(value || '').trim();
                td.className = statusText !== '' ? 'status-cell status-error' : 'status-cell status-ok';
            }

            if (isEditableDataColumn(col)) {
                td.classList.add('editable-cell');
                td.addEventListener('click', () => makeEditable(td, row._row_uid, col));
            }
            tr.appendChild(td);
        });

        tbody.appendChild(tr);
    });

    table.appendChild(tbody);
    container.innerHTML = '';
    container.appendChild(table);

    const total = sheetData.rows.length;
    const shown = displayRows.length;
    if (statsDiv) {
        if (state.columnFilters.size > 0) {
            statsDiv.textContent = `מציג ${shown} מתוך ${total} שורות × ${displayColumns.length} עמודות`;
        } else {
            statsDiv.textContent = `${total} rows × ${displayColumns.length} columns`;
        }
    }
    if (!targetContainer) updateDeleteButton();
}

// ---------------------------------------------------------------------------
// Row Selection
// ---------------------------------------------------------------------------

function toggleSelectAll(checked, displayRows) {
    state.selectedRows.clear();
    document.querySelectorAll('.data-grid tbody tr').forEach((tr) => {
        const cb = tr.querySelector('input[type=checkbox]');
        if (!cb) return;
        cb.checked = checked;
        const rowUid = tr.dataset.rowUid;
        if (checked) { state.selectedRows.add(rowUid); tr.classList.add('row-selected'); }
        else         { tr.classList.remove('row-selected'); }
    });
    updateDeleteButton();
}

function updateDeleteButton() {
    const btn = document.getElementById('delete-rows-btn');
    if (!btn) return;
    const n = state.selectedRows.size;
    btn.disabled = n === 0;
    btn.textContent = n > 0 ? `🗑 Delete ${n} row${n > 1 ? 's' : ''}` : '🗑 Delete rows';
}

// ---------------------------------------------------------------------------
// ---------------------------------------------------------------------------

Object.assign(window, { getFilteredRows, getDistinctValues, openFilterDropdown, updateFilterButtonState, applyFilters, clearAllFilters, openGridOverlay, closeGridOverlay, _overlayEscHandler, renderGrid, toggleSelectAll, updateDeleteButton });
