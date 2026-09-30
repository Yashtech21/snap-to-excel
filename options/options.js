import { getSettings, setSettings, DEFAULTS } from '../src/lib/settings.js';
import { dbGet, dbPut, dbClear } from '../src/lib/db.js';

const $ = (id) => document.getElementById(id);
const send = (msg) => chrome.runtime.sendMessage(msg);

function flash(el, text, kind = '') {
  el.textContent = text;
  el.className = `status ${kind}`;
}

async function load() {
  const s = await getSettings();
  $('destinationMode').value = s.destinationMode;
  const localWorkbook = await dbGet('handles', 'workbook');
  $('localWorkbookName').textContent = localWorkbook ? (s.localWorkbookName || 'Excel workbook selected') : 'No workbook selected';
  updateDestinationMode();
  $('clientId').value = s.clientId;
  $('tenant').value = s.tenant;
  $('redirect').value = chrome.identity.getRedirectURL();
  $('workbookUrl').value = s.workbookUrl;
  $('worksheet').value = s.worksheet;
  $('tableName').value = s.tableName;
  $('folder').value = s.folder;
  document.querySelector(`input[name=captureMode][value=${s.captureMode}]`).checked = true;
  $('rx').value = s.region.x;
  $('ry').value = s.region.y;
  $('rw').value = s.region.width;
  $('rh').value = s.region.height;
  $('dedupeSeconds').value = s.dedupeSeconds;
  $('includeTimestamp').checked = s.includeTimestamp;
  $('includePageTitle').checked = s.includePageTitle;
  $('includePageUrl').checked = s.includePageUrl;
  $('showButton').checked = s.showButton;
  $('buttonSize').value = s.buttonSize;
  $('buttonOpacity').value = s.buttonOpacity;
  await refreshStatus();
}

async function refreshStatus() {
  const st = await send({ type: 'status' });
  $('account').textContent = st.account ? `Signed in as ${st.account.username || st.account.name}` : 'Not signed in';
  $('account').className = `status ${st.account ? 'ok' : ''}`;
  $('queue').textContent = st.pending
    ? `${st.pending} screenshot${st.pending > 1 ? 's' : ''} waiting to save.${st.lastError ? ' Last error: ' + st.lastError : ''}`
    : 'Nothing waiting. All screenshots are saved.';
  const list = $('sites');
  list.replaceChildren();
  if (!st.sites.length) {
    const li = document.createElement('li');
    li.textContent = 'None yet. Open a website and turn on “Show button on this site” in the toolbar popup.';
    list.append(li);
  }
  for (const site of st.sites) {
    const li = document.createElement('li');
    const name = document.createElement('span');
    name.textContent = site;
    const rm = document.createElement('button');
    rm.textContent = 'Remove';
    rm.addEventListener('click', async () => { await send({ type: 'disable-site', site }); refreshStatus(); });
    li.append(name, rm);
    list.append(li);
  }
}

async function save() {
  const prev = await getSettings();
  const num = (id, fallback) => { const n = Number($(id).value); return Number.isFinite(n) ? n : fallback; };
  const workbookUrl = $('workbookUrl').value.trim();
  await setSettings({
    destinationMode: $('destinationMode').value,
    clientId: $('clientId').value.trim(),
    tenant: $('tenant').value.trim() || 'organizations',
    workbookUrl,
    workbookRef: workbookUrl === prev.workbookUrl ? prev.workbookRef : null,
    worksheet: $('worksheet').value.trim() || DEFAULTS.worksheet,
    tableName: $('tableName').value.trim().replace(/\s+/g, '_') || DEFAULTS.tableName,
    folder: $('folder').value.trim(),
    captureMode: document.querySelector('input[name=captureMode]:checked').value,
    region: {
      x: Math.max(0, num('rx', 0)), y: Math.max(0, num('ry', 0)),
      width: Math.max(20, num('rw', 1200)), height: Math.max(20, num('rh', 700))
    },
    dedupeSeconds: Math.max(0, num('dedupeSeconds', 10)),
    includeTimestamp: $('includeTimestamp').checked,
    includePageTitle: $('includePageTitle').checked,
    includePageUrl: $('includePageUrl').checked,
    showButton: $('showButton').checked,
    buttonSize: num('buttonSize', 48),
    buttonOpacity: num('buttonOpacity', 0.9)
  });
}

function updateDestinationMode() {
  const local = $('destinationMode').value === 'local';
  $('localDestination').hidden = !local;
  $('cloudAccount').hidden = local;
  $('cloudDestination').hidden = local;
}

async function chooseLocalWorkbook() {
  if (!window.showOpenFilePicker) throw new Error('Your Chrome version does not support selecting a workbook for direct writing.');
  const [handle] = await window.showOpenFilePicker({
    multiple: false,
    types: [{
      description: 'Excel workbook',
      accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] }
    }]
  });
  const permission = await handle.requestPermission({ mode: 'readwrite' });
  if (permission !== 'granted') throw new Error('Write access was not granted for that workbook.');
  const file = await handle.getFile();
  const ExcelJS = globalThis.ExcelJS;
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await file.arrayBuffer());
  await dbPut('handles', handle, 'workbook');
  await setSettings({ localWorkbookName: file.name, destinationMode: 'local' });
  $('destinationMode').value = 'local';
  $('localWorkbookName').textContent = file.name;
  updateDestinationMode();
  flash($('localSetupStatus'), 'Workbook selected. Captures will be embedded in this file.', 'ok');
  await send({ type: 'retry' });
  await refreshStatus();
}

async function withBusy(btn, fn) {
  btn.disabled = true;
  try { await fn(); } finally { btn.disabled = false; }
}

$('save').addEventListener('click', () => withBusy($('save'), async () => {
  await save();
  flash($('saved'), 'Settings saved.', 'ok');
  setTimeout(() => flash($('saved'), ''), 2500);
}));

$('destinationMode').addEventListener('change', updateDestinationMode);
$('chooseWorkbook').addEventListener('click', () => withBusy($('chooseWorkbook'), async () => {
  try { await chooseLocalWorkbook(); }
  catch (error) {
    if (error.name !== 'AbortError') flash($('localSetupStatus'), error.message || 'Could not select that workbook.', 'err');
  }
}));

$('copyRedirect').addEventListener('click', async () => {
  await navigator.clipboard.writeText($('redirect').value);
  $('copyRedirect').textContent = 'Copied';
  setTimeout(() => { $('copyRedirect').textContent = 'Copy'; }, 1500);
});

$('signin').addEventListener('click', () => withBusy($('signin'), async () => {
  await save();
  flash($('account'), 'Waiting for Microsoft sign-in…');
  const res = await send({ type: 'signin' });
  if (!res.ok) flash($('account'), res.error, 'err');
  else await refreshStatus();
}));

$('signout').addEventListener('click', () => withBusy($('signout'), async () => {
  await send({ type: 'signout' });
  await refreshStatus();
}));

$('setup').addEventListener('click', () => withBusy($('setup'), async () => {
  await save();
  flash($('setupStatus'), 'Checking workbook…');
  const res = await send({ type: 'setup' });
  if (!res.ok) return flash($('setupStatus'), res.error, 'err');
  flash($('setupStatus'), res.created
    ? `Created table in ${res.workbook}. Ready to capture.`
    : `Connected to ${res.workbook}. The table already exists.`, 'ok');
}));

$('retry').addEventListener('click', () => withBusy($('retry'), async () => {
  flash($('queue'), 'Retrying…');
  const res = await send({ type: 'retry' });
  await refreshStatus();
  if (res.ok && res.error) flash($('queue'), res.error, 'err');
}));

$('discard').addEventListener('click', async () => {
  if (!confirm('Delete all screenshots that have not been saved yet? This cannot be undone.')) return;
  await send({ type: 'clear-pending' });
  refreshStatus();
});

$('resetPos').addEventListener('click', async () => {
  await setSettings({ buttonPos: null });
});

$('clearAll').addEventListener('click', async () => {
  if (!confirm('Sign out and delete all settings and waiting screenshots from this browser?')) return;
  await send({ type: 'clear-data' });
  await dbClear('handles');
  await load();
});

load();
