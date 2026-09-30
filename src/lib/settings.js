export const DEFAULTS = {
  destinationMode: 'local', // 'local' | 'cloud'
  // Microsoft connection
  clientId: '',
  tenant: 'organizations',
  // Destination
  workbookUrl: '',
  workbookRef: null, // cached { url, driveId, itemId, parentId, name }
  worksheet: 'Screenshots',
  tableName: 'ScreenshotTable',
  folder: 'Screenshots',
  // Capture
  captureMode: 'viewport', // 'viewport' | 'region'
  region: { x: 100, y: 120, width: 1200, height: 700 }, // CSS pixels, relative to the page viewport
  dedupeSeconds: 10,
  // Floating button
  showButton: true,
  buttonSize: 48,
  buttonOpacity: 0.9,
  buttonPos: null, // { left, top } once dragged
  // Metadata columns
  includeTimestamp: true,
  includePageTitle: true,
  includePageUrl: true,
  // Sites where the floating button is enabled, e.g. "https://example.com"
  sites: []
};

export async function getSettings() {
  const { settings = {} } = await chrome.storage.local.get('settings');
  const destinationMode = settings.destinationMode || (settings.clientId || settings.workbookUrl ? 'cloud' : DEFAULTS.destinationMode);
  return { ...DEFAULTS, ...settings, destinationMode, region: { ...DEFAULTS.region, ...(settings.region || {}) } };
}

export async function setSettings(patch) {
  const next = { ...(await getSettings()), ...patch };
  await chrome.storage.local.set({ settings: next });
  return next;
}
