// Encrypts secrets (the refresh token) with a non-extractable AES-GCM key kept in IndexedDB,
// so the token is never stored as plain text in chrome.storage.
import { dbGet, dbPut } from './db.js';

let keyPromise;
function getKey() {
  keyPromise ??= (async () => {
    let key = await dbGet('keys', 'aes');
    if (!key) {
      key = await crypto.subtle.generateKey({ name: 'AES-GCM', length: 256 }, false, ['encrypt', 'decrypt']);
      await dbPut('keys', key, 'aes');
    }
    return key;
  })();
  return keyPromise;
}

export async function encrypt(text) {
  const key = await getKey();
  const iv = crypto.getRandomValues(new Uint8Array(12));
  const data = await crypto.subtle.encrypt({ name: 'AES-GCM', iv }, key, new TextEncoder().encode(text));
  return { iv: Array.from(iv), data: Array.from(new Uint8Array(data)) };
}

export async function decrypt(payload) {
  const key = await getKey();
  const plain = await crypto.subtle.decrypt(
    { name: 'AES-GCM', iv: new Uint8Array(payload.iv) },
    key,
    new Uint8Array(payload.data)
  );
  return new TextDecoder().decode(plain);
}
