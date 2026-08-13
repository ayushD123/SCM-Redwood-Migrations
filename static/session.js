window.sessionHelper = (function () {
  const APP_BASE_PATH = String(window.REDWOOD_BASE_PATH || '').replace(/\/$/, '');
  const SESSION_CHECK_URL = '/api/session/status';
  const MIN_FETCH_OPTIONS = { credentials: 'same-origin', headers: {} };

  function appUrl(url) {
    if (!url || /^(?:https?:|data:|blob:)/i.test(url)) return url;
    const normalized = url.startsWith('/') ? url : `/${url}`;
    if (APP_BASE_PATH && (normalized === APP_BASE_PATH || normalized.startsWith(`${APP_BASE_PATH}/`))) {
      return normalized;
    }
    return `${APP_BASE_PATH}${normalized}` || '/';
  }

  async function checkSession(redirectUrl) {
    try {
      const response = await fetchWithSession(SESSION_CHECK_URL, { method: 'GET' });
      if (!response.ok) throw new Error('no session');
      const payload = await response.json();
      if (!payload || !payload.session_id) {
        throw new Error('session missing');
      }
      return payload;
    } catch (error) {
      window.location.href = appUrl(redirectUrl || '/');
      return null;
    }
  }

  function fetchWithSession(url, options = {}) {
    const merged = {
      ...MIN_FETCH_OPTIONS,
      ...options,
      headers: {
        ...MIN_FETCH_OPTIONS.headers,
        ...options.headers,
      },
    };

    return fetch(appUrl(url), merged);
  }

  async function errorMessage(response) {
    try {
      const payload = await response.clone().json();
      if (payload?.detail) return payload.detail;
      if (payload?.message) return payload.message;
    } catch (_) {
      // Fall through to plain text.
    }
    const text = await response.text();
    return text || `Request failed with HTTP ${response.status}`;
  }

  async function saveSession(formData) {
    const response = await fetchWithSession('/api/save-session', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(formData),
    });
    return response.json();
  }

  return {
    appUrl,
    checkSession,
    errorMessage,
    fetchWithSession,
    saveSession,
    requireSession: function (redirectUrl) {
      document.addEventListener('DOMContentLoaded', () => {
        checkSession(redirectUrl);
      });
    },
  };
})();
