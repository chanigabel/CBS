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

function formatApiErrorDetail(detail) {
    if (Array.isArray(detail)) {
        const messages = detail
            .map(item => {
                if (typeof item === 'string') return item;
                if (!item || typeof item !== 'object') return '';
                const field = Array.isArray(item.loc)
                    ? item.loc.filter(part => part !== 'body').join('.')
                    : '';
                const message = item.msg || item.message || '';
                return field && message ? `${field}: ${message}` : message;
            })
            .filter(Boolean);
        return messages.length
            ? `הבקשה לא תקינה: ${messages.join(' | ')}`
            : 'הבקשה לא תקינה. בדוק את הערכים שהוזנו ונסה שוב.';
    }
    if (detail && typeof detail === 'object') {
        return detail.message || detail.msg || 'הבקשה לא תקינה. בדוק את הערכים שהוזנו ונסה שוב.';
    }
    return detail || '';
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
        try {
            const err = await response.json();
            detail = formatApiErrorDetail(err.detail || detail);
        } catch (_) {}
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

Object.assign(window, { showError, dismissError, apiCall, formatApiErrorDetail });
