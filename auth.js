/**
 * auth.js — Google OAuth2 Implicit Grant Flow
 * Handles sign-in, token storage, and sign-out.
 * Tokens are stored in sessionStorage only (not persisted across tabs).
 */

const SCOPES = [
  'https://www.googleapis.com/auth/spreadsheets',
  'https://www.googleapis.com/auth/drive.file',
].join(' ');

const TOKEN_KEY   = 'xls2gsheets_token';
const USER_KEY    = 'xls2gsheets_user';
const EXPIRY_KEY  = 'xls2gsheets_expiry';

// ─── Public API ──────────────────────────────────────────

/**
 * Launch the Google OAuth2 popup and resolve with the access token.
 * @param {string} clientId  Google OAuth2 Client ID
 * @returns {Promise<string>}
 */
function signIn(clientId) {
  return new Promise((resolve, reject) => {
    if (!clientId || !clientId.trim()) {
      reject(new Error('Google Client ID is not set. Click ⚙️ to configure it.'));
      return;
    }

    const redirectUri = window.location.origin + window.location.pathname;
    const params = new URLSearchParams({
      client_id:     clientId,
      redirect_uri:  redirectUri,
      response_type: 'token',
      scope:         SCOPES,
      prompt:        'select_account',
      include_granted_scopes: 'true',
    });

    const authUrl = `https://accounts.google.com/o/oauth2/v2/auth?${params}`;
    const popup   = openPopup(authUrl);

    if (!popup) {
      reject(new Error('Popup was blocked. Please allow popups for this page.'));
      return;
    }

    let settled = false;
    const POLL_INTERVAL = 300;

    const poll = setInterval(() => {
      try {
        if (popup.closed) {
          if (!settled) {
            clearInterval(poll);
            reject(new Error('Sign-in window was closed before completing.'));
          }
          return;
        }

        const url = popup.location.href;
        if (!url || url === 'about:blank') return;

        const hash = new URL(url).hash.slice(1);
        const p    = new URLSearchParams(hash);

        if (p.get('access_token')) {
          clearInterval(poll);
          settled = true;
          popup.close();

          const token   = p.get('access_token');
          const expiresIn = parseInt(p.get('expires_in') || '3600', 10);
          const expiry  = Date.now() + expiresIn * 1000;

          sessionStorage.setItem(TOKEN_KEY,  token);
          sessionStorage.setItem(EXPIRY_KEY, String(expiry));

          // Fetch user info
          fetchUserInfo(token).then(user => {
            if (user) sessionStorage.setItem(USER_KEY, JSON.stringify(user));
            resolve(token);
          }).catch(() => resolve(token));

        } else if (p.get('error')) {
          clearInterval(poll);
          settled = true;
          popup.close();
          reject(new Error('Google sign-in was denied: ' + p.get('error')));
        }
      } catch (_) {
        // Cross-origin access on accounts.google.com — normal, keep polling
      }
    }, POLL_INTERVAL);
  });
}

/**
 * Sign out: clear all tokens from sessionStorage.
 */
function signOut() {
  sessionStorage.removeItem(TOKEN_KEY);
  sessionStorage.removeItem(USER_KEY);
  sessionStorage.removeItem(EXPIRY_KEY);
}

/**
 * Return the stored access token if still valid, else null.
 * @returns {string|null}
 */
function getToken() {
  const token  = sessionStorage.getItem(TOKEN_KEY);
  const expiry = parseInt(sessionStorage.getItem(EXPIRY_KEY) || '0', 10);
  if (!token) return null;
  // Expire 60s early to avoid edge cases
  if (Date.now() > expiry - 60_000) {
    signOut();
    return null;
  }
  return token;
}

/**
 * Return cached user info object or null.
 * @returns {{name:string, email:string, picture:string}|null}
 */
function getUser() {
  const raw = sessionStorage.getItem(USER_KEY);
  try { return raw ? JSON.parse(raw) : null; } catch { return null; }
}

/**
 * True if user is currently signed in with a valid token.
 */
function isSignedIn() {
  return getToken() !== null;
}

// ─── Helpers ─────────────────────────────────────────────

function openPopup(url) {
  const w = 520, h = 620;
  const left = Math.max(0, (screen.width  - w) / 2);
  const top  = Math.max(0, (screen.height - h) / 2);
  return window.open(
    url,
    'GoogleSignIn',
    `width=${w},height=${h},left=${left},top=${top},toolbar=no,menubar=no,scrollbars=yes`
  );
}

async function fetchUserInfo(token) {
  try {
    const res  = await fetch('https://www.googleapis.com/oauth2/v3/userinfo', {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (!res.ok) return null;
    return await res.json();
  } catch {
    return null;
  }
}

// Export
window.Auth = { signIn, signOut, getToken, getUser, isSignedIn };
