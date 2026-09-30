import { getSettings, setSettings } from '../lib/settings.js';
import { signIn, signOut, getAccount } from '../lib/auth.js';
import { processQueue, updateBadge } from '../lib/pipeline.js';
import { getWorkbookRef, ensureTable } from '../lib/graph.js';
import { cropBlob, sha256Hex, timestamps } from '../lib/capture.js';
import { dbPut, dbGet, dbAll, dbClear } from '../lib/db.js';
import { siteKey, sitePattern, scriptId } from '../lib/sites.js';

// ---- Lifecycle ------------------------------------------------------------

chrome.runtime.onInstalled.addListener(async (details) => {
  await syncContentScripts();
  await updateBadge();
  chrome.alarms.create('retry-uploads', { periodInMinutes: 1 });
  if (details.reason === 'install') chrome.runtime.openOptionsPage();
});

chrome.runtime.onStartup.addListener(async () => {
  await syncContentScripts();
  await updateBadge();
  chrome.alarms.create('retry-uploads', { periodInMinutes: 1 });
  processQueue().catch(() => {});
});

chrome.alarms.onAlarm.addListener((alarm) => {
  if (alarm.name === 'retry-uploads') processQueue().catch(() => {});
});

chrome.commands.onCommand.addListener(async (command) => {
  if (command !== 'capture-screenshot') return;
  const [tab] = await chrome.tabs.query({ active: true, lastFocusedWindow: true });
  if (tab) captureFlow(tab);
});

// Keeps registered content scripts in sync with the sites the user enabled.
async function syncContentScripts() {
  const s = await getSettings();
  const registered = await chrome.scripting.getRegisteredContentScripts();
  const wanted = new Map();
  for (const site of s.sites) {
    if (await chrome.permissions.contains({ origins: [sitePattern(site)] })) wanted.set(scriptId(site), site);
  }
  const stale = registered.filter((r) => r.id.startsWith('site-') && !wanted.has(r.id)).map((r) => r.id);
  if (stale.length) await chrome.scripting.unregisterContentScripts({ ids: stale });
  const have = new Set(registered.map((r) => r.id));
  const add = [...wanted]
    .filter(([id]) => !have.has(id))
    .map(([id, site]) => ({
      id,
      matches: [sitePattern(site)],
      js: ['src/content/floating-button.js'],
      runAt: 'document_idle',
      persistAcrossSessions: true
    }));
  if (add.length) await chrome.scripting.registerContentScripts(add);
}

// ---- Capture --------------------------------------------------------------

let capturing = false;

async function notify(tabId, kind, text) {
  try {
    await chrome.tabs.sendMessage(tabId, { type: 'toast', kind, text });
  } catch {
    chrome.notifications.create({
      type: 'basic',
      iconUrl: 'icons/icon128.png',
      title: 'Snap to Excel',
      message: text
    });
  }
}

const safeName = (s) => s.replace(/[^a-z0-9.-]+/gi, '-').replace(/^-+|-+$/g, '').slice(0, 60) || 'page';

async function captureFlow(tab) {
  if (capturing) return; // ignore double clicks while a capture is in progress
  capturing = true;
  try {
    const s = await getSettings();

    let info;
    try {
      [{ result: info }] = await chrome.scripting.executeScript({
        target: { tabId: tab.id },
        func: () => ({ innerWidth: window.innerWidth, title: document.title, url: location.href })
      });
    } catch {
      throw new Error('This page cannot be captured. Browser pages and the Web Store are restricted.');
    }

    await chrome.tabs.sendMessage(tab.id, { type: 'prepare' }).catch(() => {}); // hides the floating button
    let dataUrl;
    try {
      dataUrl = await chrome.tabs.captureVisibleTab(tab.windowId, { format: 'png' });
    } finally {
      chrome.tabs.sendMessage(tab.id, { type: 'restore' }).catch(() => {});
    }

    const raw = await (await fetch(dataUrl)).blob();
    const blob = await cropBlob(raw, s, info.innerWidth);

    // Duplicate protection: same pixels captured again within the time window.
    const hash = await sha256Hex(blob);
    const { last } = await chrome.storage.session.get('last');
    if (s.dedupeSeconds > 0 && last && last.hash === hash && Date.now() - last.time < s.dedupeSeconds * 1000) {
      await notify(tab.id, 'info', 'Duplicate screenshot skipped.');
      return;
    }
    await chrome.storage.session.set({ last: { hash, time: Date.now() } });

    const ts = timestamps();
    let host = 'page';
    try { host = new URL(info.url).hostname; } catch { /* keep default */ }
    const id = crypto.randomUUID();
    await dbPut('pending', {
      id,
      blob,
      hash,
      fileName: `${safeName(host)}_${ts.file}.png`,
      timestamp: ts.display,
      pageTitle: info.title || '',
      pageUrl: info.url,
      createdAt: Date.now(),
      attempts: 0
    });
    await updateBadge();
    await notify(tab.id, 'info', s.destinationMode === 'local' ? 'Saving to the selected workbook…' : 'Uploading…');

    let out = await processQueue();
    if ((await dbGet('pending', id)) && !out.error) out = await processQueue(); // queue was busy

    const mine = out.done.find((d) => d.id === id);
    if (mine) {
      await notify(tab.id, 'success', `Saved to Excel · row ${mine.row}`);
    } else if (out.error?.config) {
      await notify(tab.id, 'error', s.destinationMode === 'local'
        ? 'Saved on this device. Choose a workbook in Settings to write captures.'
        : 'Saved on this device. Finish setup in Settings to upload it.');
    } else if (out.error?.auth) {
      await notify(tab.id, 'error', 'Saved on this device. Sign in again in Settings to upload it.');
    } else {
      await notify(tab.id, 'error', `Saved on this device and will retry. ${out.error?.message || ''}`.trim());
    }
  } catch (e) {
    await notify(tab.id, 'error', e.message || 'Capture failed.');
  } finally {
    capturing = false;
  }
}

// ---- Messages from popup / options / content scripts ----------------------

const handlers = {
  async 'capture-request'(msg, sender) {
    await captureFlow(sender.tab);
  },
  async 'capture-active'() {
    const [tab] = await chrome.tabs.query({ active: true, lastFocusedWindow: true });
    if (tab) await captureFlow(tab);
  },
  async signin() {
    const account = await signIn();
    processQueue().catch(() => {});
    return { account };
  },
  async signout() {
    await signOut();
  },
  async status() {
    const pending = await dbAll('pending');
    const s = await getSettings();
    const localWorkbook = s.destinationMode === 'local' ? await dbGet('handles', 'workbook') : null;
    return {
      account: await getAccount(),
      pending: pending.length,
      lastError: pending.find((p) => p.lastError)?.lastError || '',
      configured: s.destinationMode === 'local' ? Boolean(localWorkbook) : Boolean(s.clientId && s.workbookUrl),
      destinationMode: s.destinationMode,
      localWorkbookName: s.localWorkbookName || '',
      sites: s.sites
    };
  },
  async setup() {
    const s = await getSettings();
    if (!s.workbookUrl) throw new Error('Paste a link to your Excel workbook first.');
    await setSettings({ workbookRef: null });
    const ref = await getWorkbookRef(await getSettings());
    const { created } = await ensureTable(ref, s);
    return { workbook: ref.name, created };
  },
  async 'enable-site'(msg) {
    const key = siteKey(msg.url);
    if (!key) throw new Error('Only http and https pages are supported.');
    const s = await getSettings();
    if (!s.sites.includes(key)) await setSettings({ sites: [...s.sites, key] });
    await syncContentScripts();
    if (msg.tabId) {
      await chrome.scripting.executeScript({ target: { tabId: msg.tabId }, files: ['src/content/floating-button.js'] });
    }
  },
  async 'disable-site'(msg) {
    const s = await getSettings();
    await setSettings({ sites: s.sites.filter((x) => x !== msg.site) });
    await chrome.permissions.remove({ origins: [sitePattern(msg.site)] }).catch(() => {});
    await syncContentScripts();
  },
  async 'save-region'(msg, sender) {
    const { x, y, width, height } = msg.region;
    await setSettings({ captureMode: 'region', region: { x, y, width, height } });
    if (sender.tab) await notify(sender.tab.id, 'success', `Capture area saved (${width} × ${height}).`);
  },
  async retry() {
    const out = await processQueue();
    return { done: out.done.length, error: out.error?.message || '' };
  },
  async 'clear-pending'() {
    await dbClear('pending');
    await updateBadge();
  },
  async 'clear-data'() {
    await signOut();
    await dbClear('pending');
    await dbClear('handles');
    await chrome.storage.local.clear();
    await chrome.storage.session.clear();
    await syncContentScripts();
    await updateBadge();
  }
};

chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  const handler = handlers[msg?.type];
  if (!handler) return false;
  handler(msg, sender)
    .then((result) => sendResponse({ ok: true, ...(result || {}) }))
    .catch((e) => sendResponse({ ok: false, error: e.message || String(e) }));
  return true;
});
