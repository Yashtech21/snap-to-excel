// A "site" is protocol + hostname (ports are ignored by match patterns).
export function siteKey(url) {
  const u = new URL(url);
  if (u.protocol !== 'https:' && u.protocol !== 'http:') return null;
  return `${u.protocol}//${u.hostname}`;
}
export const sitePattern = (key) => `${key}/*`;
export const scriptId = (key) => 'site-' + key.replace(/[^a-z0-9]/gi, '-');
