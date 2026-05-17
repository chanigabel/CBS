// ---------------------------------------------------------------------------
// Error Handling
// ---------------------------------------------------------------------------

function showError(message) {
    const banner = document.getElementById('error-banner');
    document.getElementById('error-message').textContent = message;
    banner.classList.remove('hidden');
}

function dismissError() {
    document.getElementById('error-banner').classList.add('hidden');
}

// ---------------------------------------------------------------------------
// API Helpers
// ---------------------------------------------------------------------------

async function apiCall(method, url, body = null) {
    const options = { method, headers: {} };
    if (body && !(body instanceof FormData)) {
        options.headers['Content-Type'] = 'application/json';
        options.body = JSON.stringify(body);
    } else if (body instanceof FormData) {
        options.body = body;
    }

    const response = await fetch(url, options);

    if (!response.ok) {
        let detail = `HTTP ${response.status}`;
        try { const err = await response.json(); detail = err.detail || detail; } catch (_) {}
        const error = new Error(detail);
        error.status = response.status;
        throw error;
    }

    const ct = response.headers.get('content-type') || '';
    if (ct.includes('application/zip') || ct.includes('application/vnd.openxmlformats')) {
        return response;
    }
    return response.json();
}

Object.assign(window, { showError, dismissError, apiCall });
