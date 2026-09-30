// Microsoft Graph helpers: resolve the workbook, upload the image, add an Excel table row.
import { getAccessToken, clearAccessToken, AuthError } from './auth.js';
import { setSettings } from './settings.js';

const G = 'https://graph.microsoft.com/v1.0';
const SMALL_UPLOAD_LIMIT = 3.5 * 1024 * 1024;
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

export class GraphError extends Error {
  constructor(message, status) { super(message); this.name = 'GraphError'; this.status = status; }
}

async function gfetch(path, init = {}, attempt = 0) {
  const token = await getAccessToken();
  const res = await fetch(path.startsWith('http') ? path : G + path, {
    ...init,
    headers: { Authorization: `Bearer ${token}`, ...(init.headers || {}) }
  });
  if ((res.status === 429 || res.status === 503) && attempt < 3) {
    const wait = Number(res.headers.get('Retry-After')) || 2 ** attempt;
    await sleep(Math.min(wait, 10) * 1000);
    return gfetch(path, init, attempt + 1);
  }
  if (res.status === 401) {
    await clearAccessToken();
    if (attempt < 1) return gfetch(path, init, attempt + 1);
    throw new AuthError('Session expired. Sign in again in Settings.');
  }
  if (!res.ok) {
    let message = '';
    try { message = (await res.json()).error?.message || ''; } catch { /* no body */ }
    throw new GraphError(message || `Microsoft Graph request failed (${res.status})`, res.status);
  }
  return res.status === 204 ? null : res.json();
}

const json = (method, body) => ({ method, headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) });
const seg = (s) => encodeURIComponent(s);

// ---- Workbook -------------------------------------------------------------

export async function resolveWorkbook(url) {
  const id = 'u!' + btoa(url).replace(/=+$/, '').replace(/\//g, '_').replace(/\+/g, '-');
  let item;
  try {
    item = await gfetch(`/shares/${id}/driveItem?$select=id,name,parentReference`);
  } catch (e) {
    if (e instanceof GraphError && [400, 403, 404].includes(e.status)) {
      throw new GraphError('Could not open that workbook. In Excel/SharePoint choose Share → Copy link and paste it in Settings.', e.status);
    }
    throw e;
  }
  return {
    url,
    driveId: item.parentReference.driveId,
    itemId: item.id,
    parentId: item.parentReference.id || null,
    name: item.name
  };
}

export async function getWorkbookRef(s) {
  if (s.workbookRef && s.workbookRef.url === s.workbookUrl) return s.workbookRef;
  const ref = await resolveWorkbook(s.workbookUrl);
  await setSettings({ workbookRef: ref });
  return ref;
}

const wbPath = (ref) => `/drives/${ref.driveId}/items/${ref.itemId}/workbook`;
export const TABLE_HEADERS = ['Screenshot', 'Timestamp', 'Page Title', 'Page URL', 'File Name'];

// Creates the worksheet + table if they do not exist yet.
export async function ensureTable(ref, s) {
  const wb = wbPath(ref);
  try {
    await gfetch(`${wb}/tables/${seg(s.tableName)}?$select=id`);
    return { created: false };
  } catch (e) {
    if (!(e instanceof GraphError) || e.status !== 404) throw e;
  }
  try {
    await gfetch(`${wb}/worksheets/${seg(s.worksheet)}?$select=id`);
  } catch (e) {
    if (!(e instanceof GraphError) || e.status !== 404) throw e;
    await gfetch(`${wb}/worksheets/add`, json('POST', { name: s.worksheet }));
  }
  const address = `A1:${String.fromCharCode(64 + TABLE_HEADERS.length)}1`;
  await gfetch(`${wb}/worksheets/${seg(s.worksheet)}/range(address='${address}')`, json('PATCH', { values: [TABLE_HEADERS] }));
  const table = await gfetch(`${wb}/worksheets/${seg(s.worksheet)}/tables/add`, json('POST', { address, hasHeaders: true }));
  await gfetch(`${wb}/tables/${table.id}`, json('PATCH', { name: s.tableName }));
  return { created: true };
}

// values are matched to the table's own column headers, so users can reorder or add columns.
export async function addRow(ref, s, record) {
  const wb = wbPath(ref);
  const t = seg(s.tableName);
  let cols;
  try {
    cols = await gfetch(`${wb}/tables/${t}/columns?$select=name`);
  } catch (e) {
    if (e instanceof GraphError && e.status === 404) {
      throw new GraphError(`Table "${s.tableName}" was not found. Use "Set up destination" in Settings.`, 404);
    }
    throw e;
  }
  const names = cols.value.map((c) => c.name.trim().toLowerCase());
  if (!names.some((n) => n in record)) {
    throw new GraphError('The table has no Screenshot, Timestamp, Page Title, Page URL or File Name column.', 400);
  }
  const row = names.map((n) => record[n] ?? '');
  const res = await gfetch(`${wb}/tables/${t}/rows/add`, json('POST', { index: null, values: [row] }));
  return (res.index ?? 0) + 1; // 1-based data row number
}

// ---- Image upload ---------------------------------------------------------

export async function uploadImage(ref, folder, fileName, blob) {
  const base = ref.parentId ? `/drives/${ref.driveId}/items/${ref.parentId}:` : `/drives/${ref.driveId}/root:`;
  const folderPath = folder.split('/').map((p) => p.trim()).filter(Boolean).map(seg).join('/');
  const path = `${base}/${folderPath ? folderPath + '/' : ''}${seg(fileName)}`;

  if (blob.size <= SMALL_UPLOAD_LIMIT) {
    return gfetch(`${path}:/content?@microsoft.graph.conflictBehavior=rename`, {
      method: 'PUT',
      headers: { 'Content-Type': 'image/png' },
      body: blob
    });
  }

  const session = await gfetch(`${path}:/createUploadSession`, json('POST', {
    item: { '@microsoft.graph.conflictBehavior': 'rename' }
  }));
  const res = await fetch(session.uploadUrl, {
    method: 'PUT',
    headers: { 'Content-Range': `bytes 0-${blob.size - 1}/${blob.size}` },
    body: blob
  });
  if (!res.ok) throw new GraphError(`Image upload failed (${res.status})`, res.status);
  return res.json();
}
