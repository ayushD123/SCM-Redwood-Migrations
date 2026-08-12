window.sessionHelper = (function () {
  const SESSION_CHECK_URL = '/api/session/status';
  const MIN_FETCH_OPTIONS = { credentials: 'same-origin', headers: {} };

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
      window.location.href = redirectUrl || '/';
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

    return fetch(url, merged);
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
    checkSession,
    fetchWithSession,
    saveSession,
    requireSession: function (redirectUrl) {
      document.addEventListener('DOMContentLoaded', () => {
        checkSession(redirectUrl);
      });
    },
  };
})();