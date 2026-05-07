/**
 * Excel standardization Web App - Frontend JavaScript wiring.
 * Global behavior lives in static/js/*.js modules loaded before this file.
 */

function initApp() {
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

    // Keyboard shortcuts.
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

    // Institution metadata form.
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
        for (const t of types) {
            const tErr = validateNumericMin3(t, 'סוג מוסד');
            if (tErr) { showError(tErr); return; }
        }
        apiCall('PATCH', `/api/workbook/${state.sessionId}/institution`, {
            mosad_id: rawId || undefined,
            mosad_name: instName ? instName.value.trim() || undefined : undefined,
            mosad_types: types,
        }).catch(err => showError(`Failed to save institution metadata: ${err.message}`));
    }

    [instId, instName, instType1, instType2, instType3].forEach(el => {
        if (el) el.addEventListener('blur', saveInstitution);
    });

    const applySelect = document.getElementById('inst-type-apply-select');
    if (applySelect) {
        applySelect.addEventListener('change', () => {
            const v = applySelect.value;
            if (instType1 && !instType1.value) instType1.value = v;
            else if (instType2 && !instType2.value) instType2.value = v;
            else if (instType3 && !instType3.value) instType3.value = v;
        });
    }

    // Initialize dependent UI controls.
    updateMosadTypeDropdown();
    updateInstSheetSelector();
    onScopeChange();
}

document.addEventListener('DOMContentLoaded', initApp);
