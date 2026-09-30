import { siteKey, sitePattern } from '../src/lib/sites.js';

const $ = (id) => document.getElementById(id);
let tab;
let key = null;

const send = (msg) => chrome.runtime.sendMessage(msg);
const say = (t) => { $('message').textContent = t; };

async function init() {
  [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  try { key = tab?.url ? siteKey(tab.url) : null; } catch { key = null; }

  const status = await send({ type: 'status' });
  $('account').textContent = status.account
    ? `Signed in as ${status.account.username || status.account.name}`
    : 'Not signed in. Open Settings to connect Microsoft.';
  if (status.pending) {
    $('pending').hidden = false;
    $('pending').textContent = `${status.pending} screenshot${status.pending > 1 ? 's' : ''} waiting to save.`;
  }

  const { settings = {} } = await chrome.storage.local.get('settings');
  $('mode').value = settings.captureMode || 'viewport';

  if (!key) {
    $('siteToggle').disabled = true;
    $('captureBtn').disabled = true;
    $('regionBtn').disabled = true;
    $('siteHint').hidden = false;
    $('siteHint').textContent = 'This page can’t be captured. Open a regular website first.';
    return;
  }
  $('siteLabel').textContent = `Show button on ${new URL(tab.url).hostname}`;
  const granted = await chrome.permissions.contains({ origins: [sitePattern(key)] });
  $('siteToggle').checked = granted && (status.sites || []).includes(key);
}

$('captureBtn').addEventListener('click', async () => {
  await send({ type: 'capture-active' });
  window.close();
});

$('siteToggle').addEventListener('change', async (e) => {
  if (e.target.checked) {
    // permissions.request must run directly inside the click handler.
    const ok = await chrome.permissions.request({ origins: [sitePattern(key)] });
    if (!ok) { e.target.checked = false; say('Permission was not granted.'); return; }
    const res = await send({ type: 'enable-site', url: tab.url, tabId: tab.id });
    if (!res.ok) { e.target.checked = false; say(res.error); return; }
    say('Button added. You can drag it anywhere.');
  } else {
    await send({ type: 'disable-site', site: key });
    say('Button removed from this site. Reload the page to clear it.');
  }
});

$('mode').addEventListener('change', async (e) => {
  const { settings = {} } = await chrome.storage.local.get('settings');
  await chrome.storage.local.set({ settings: { ...settings, captureMode: e.target.value } });
});

$('regionBtn').addEventListener('click', async () => {
  try {
    await chrome.scripting.executeScript({ target: { tabId: tab.id }, files: ['src/content/region-picker.js'] });
    window.close();
  } catch {
    say('Could not open the area picker on this page.');
  }
});

$('settingsBtn').addEventListener('click', () => chrome.runtime.openOptionsPage());

init();
