// Upload queue: every screenshot is stored locally first, then uploaded + logged in Excel.
// Failed items stay in the queue and are retried automatically.
import { dbAll, dbPut, dbDel, dbGet } from './db.js';
import { getSettings } from './settings.js';
import { AuthError } from './auth.js';
import { getWorkbookRef, uploadImage, addRow } from './graph.js';
import { appendCaptureToWorkbook } from './local-workbook.js';

let running = null;

export async function pendingCount() {
  return (await dbAll('pending')).length;
}

export async function updateBadge() {
  const n = await pendingCount();
  await chrome.action.setBadgeBackgroundColor({ color: '#c2410c' });
  await chrome.action.setBadgeText({ text: n ? String(n) : '' });
}

const esc = (v) => String(v).replace(/"/g, '""');

function linkCell(url) {
  // Excel string literals in formulas are limited to 255 characters.
  return url.length <= 250 ? `=HYPERLINK("${esc(url)}","Open image")` : url;
}

export function processQueue() {
  running ??= run().finally(() => { running = null; });
  return running;
}

async function run() {
  const done = [];
  let error = null;
  const s = await getSettings();

  if (s.destinationMode === 'local') {
    if (!(await dbGet('handles', 'workbook'))) {
      error = { message: 'Choose an Excel workbook in Settings.', config: true };
      await updateBadge();
      return { done, error };
    }
  } else if (!s.clientId || !s.workbookUrl) {
    error = { message: 'Open Settings to connect Microsoft and choose a workbook.', config: true };
    await updateBadge();
    return { done, error };
  }

  const skipped = new Set();
  for (;;) {
    const items = (await dbAll('pending'))
      .filter((i) => !skipped.has(i.id))
      .sort((a, b) => a.createdAt - b.createdAt);
    if (!items.length) break;
    const item = items[0];

    try {
      let row;
      if (s.destinationMode === 'local') {
        ({ row } = await appendCaptureToWorkbook(s, item));
      } else {
        const ref = await getWorkbookRef(s);
        if (!item.upload) {
          const uploaded = await uploadImage(ref, s.folder, item.fileName, item.blob);
          item.upload = { webUrl: uploaded.webUrl };
          await dbPut('pending', item); // don't upload twice if the Excel step fails
        }
        row = await addRow(ref, s, {
          screenshot: linkCell(item.upload.webUrl),
          timestamp: s.includeTimestamp ? item.timestamp : '',
          'page title': s.includePageTitle ? item.pageTitle : '',
          'page url': s.includePageUrl ? item.pageUrl : '',
          'file name': item.fileName
        });
      }
      await dbDel('pending', item.id);
      done.push({ id: item.id, row });
    } catch (e) {
      skipped.add(item.id);
      item.attempts = (item.attempts || 0) + 1;
      item.lastError = e.message;
      await dbPut('pending', item);
      error = { message: e.message, auth: e instanceof AuthError };
      // Stop on problems that will affect every item; otherwise try the next one.
      if (e instanceof AuthError || e instanceof TypeError || e.status >= 500) break;
    }
  }
  await updateBadge();
  return { done, error };
}
