// Microsoft Entra ID sign-in: OAuth 2.0 authorization-code flow with PKCE (public client, no secret).
import { getSettings } from './settings.js';
import { encrypt, decrypt } from './secure.js';

export class AuthError extends Error {
  constructor(message) { super(message); this.name = 'AuthError'; }
}

const SCOPES = 'openid profile offline_access User.Read Files.ReadWrite.All';

const b64url = (buf) =>
  btoa(String.fromCharCode(...new Uint8Array(buf))).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');

const authority = (s) =>
  `https://login.microsoftonline.com/${encodeURIComponent(s.tenant || 'organizations')}/oauth2/v2.0`;

async function tokenRequest(s, params) {
  const res = await fetch(`${authority(s)}/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({ client_id: s.clientId, scope: SCOPES, ...params })
  });
  const json = await res.json().catch(() => ({}));
  if (!res.ok) {
    const message = json.error_description || `Sign-in failed (${res.status})`;
    if (['invalid_grant', 'interaction_required', 'consent_required'].includes(json.error)) throw new AuthError(message);
    throw new Error(message);
  }
  return json;
}

function decodeJwt(token) {
  const part = token.split('.')[1].replace(/-/g, '+').replace(/_/g, '/');
  const bytes = Uint8Array.from(atob(part), (c) => c.charCodeAt(0));
  return JSON.parse(new TextDecoder().decode(bytes));
}

async function saveTokens(tok) {
  await chrome.storage.session.set({ at: { token: tok.access_token, exp: Date.now() + tok.expires_in * 1000 } });
  if (tok.refresh_token) await chrome.storage.local.set({ rt: await encrypt(tok.refresh_token) });
  if (tok.id_token) {
    try {
      const p = decodeJwt(tok.id_token);
      await chrome.storage.local.set({ account: { name: p.name || '', username: p.preferred_username || p.email || '' } });
    } catch { /* display name is optional */ }
  }
}

export async function signIn() {
  const s = await getSettings();
  if (!s.clientId) throw new Error('Enter your Application (client) ID in Settings first.');

  const verifier = b64url(crypto.getRandomValues(new Uint8Array(48)));
  const challenge = b64url(await crypto.subtle.digest('SHA-256', new TextEncoder().encode(verifier)));
  const state = b64url(crypto.getRandomValues(new Uint8Array(16)));
  const redirectUri = chrome.identity.getRedirectURL();

  const url = `${authority(s)}/authorize?` + new URLSearchParams({
    client_id: s.clientId,
    response_type: 'code',
    redirect_uri: redirectUri,
    response_mode: 'query',
    scope: SCOPES,
    state,
    code_challenge: challenge,
    code_challenge_method: 'S256',
    prompt: 'select_account'
  });

  const callback = await chrome.identity.launchWebAuthFlow({ url, interactive: true });
  const q = new URL(callback).searchParams;
  if (q.get('error')) throw new Error(q.get('error_description') || q.get('error'));
  if (q.get('state') !== state) throw new Error('Sign-in response did not match the request. Try again.');

  const tok = await tokenRequest(s, {
    grant_type: 'authorization_code',
    code: q.get('code'),
    redirect_uri: redirectUri,
    code_verifier: verifier
  });
  await saveTokens(tok);
  return getAccount();
}

let refreshing = null;

export async function getAccessToken() {
  const { at } = await chrome.storage.session.get('at');
  if (at && at.exp - Date.now() > 60_000) return at.token;

  refreshing ??= (async () => {
    const s = await getSettings();
    const { rt } = await chrome.storage.local.get('rt');
    if (!rt || !s.clientId) throw new AuthError('Not signed in. Open Settings and sign in with Microsoft.');
    let refreshToken;
    try { refreshToken = await decrypt(rt); } catch { throw new AuthError('Stored sign-in is unreadable. Sign in again.'); }
    try {
      const tok = await tokenRequest(s, { grant_type: 'refresh_token', refresh_token: refreshToken });
      await saveTokens(tok);
      return tok.access_token;
    } catch (e) {
      if (e instanceof AuthError) await signOut();
      throw e;
    }
  })().finally(() => { refreshing = null; });
  return refreshing;
}

export const clearAccessToken = () => chrome.storage.session.remove('at');

export async function signOut() {
  await chrome.storage.session.remove('at');
  await chrome.storage.local.remove(['rt', 'account']);
}

export async function getAccount() {
  const { account, rt } = await chrome.storage.local.get(['account', 'rt']);
  return rt ? (account || { name: '', username: 'Signed in' }) : null;
}
