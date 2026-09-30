# Snap to Excel

Chrome extension (Manifest V3). Capture the visible page or a fixed area and save it to an Excel workbook. Choose a local `.xlsx` workbook for offline use, or choose OneDrive/SharePoint. No developer backend or build step is required.

## Load locally

1. Open `chrome://extensions` and enable **Developer mode**.
2. Click **Load unpacked** and choose this folder (the one containing `manifest.json`).
3. Open Settings. For offline use, select **Excel workbook on this device** and choose an existing `.xlsx` file. The extension embeds screenshots in a `Screenshots` worksheet and adds timestamp, page title, page URL, and file name columns.
4. Keep the workbook closed in Excel while captures are being saved.
5. Open a website, click the extension, and enable **Show button on this site**. Capture using the floating camera button or **Alt+Shift+S**. You can also select a fixed area.

The extension needs write access to the workbook you select. If Chrome asks again later, reopen Settings and select the workbook to restore access. Keep the workbook closed in Excel while capturing. Use a simple workbook dedicated to screenshots: complex or advanced Excel features may not be preserved when the file is rewritten. Captures that cannot be written stay in the local queue and retry automatically. Identical captures inside the configured window are skipped.

## Optional OneDrive/SharePoint mode

Local workbook mode works without Microsoft app registration or internet access. To use cloud mode instead:

1. Copy the **Redirect URI** shown in Settings (looks like `https://<extension-id>.chromiumapp.org/`).
2. Go to entra.microsoft.com → **App registrations** → **New registration**.
3. Choose the account types you need.
4. Under **Authentication** → **Add a platform** → **Single-page application**, register the Redirect URI.
5. Under **API permissions**, add Microsoft Graph **Delegated** permissions: `Files.ReadWrite.All`, `User.Read`. Grant admin consent if your tenant requires it.
6. Paste the Application (client) ID in Settings, sign in, and select a workbook link.

No client secret is needed. The extension uses Microsoft Graph only in cloud mode.

## Permissions

- `activeTab`, `scripting`: capture the current tab and read its title/URL when requested.
- `storage`, `alarms`: save settings and retry queued captures.
- `identity`: Microsoft sign-in for optional cloud mode.
- `notifications`: show capture/save status when the floating button is unavailable.
- Microsoft Graph and sign-in hosts: used only in optional cloud mode.
- Optional site permissions: add the floating capture button only on sites where you enable it.

## Chrome Web Store

Zip this folder's contents with `manifest.json` at the ZIP root. Provide the privacy policy at `PRIVACY_POLICY.md` after publishing it at a public URL. The extension has no analytics or developer-operated backend.

## Implementation

- `src/background/service-worker.js` – capture, queue, and message handling
- `src/lib/local-workbook.js` – offline `.xlsx` writing with embedded PNG images
- `src/lib/graph.js` and `src/lib/auth.js` – optional OneDrive/SharePoint integration
- `src/lib/pipeline.js` – retry queue
- `src/content/` – floating button and fixed-area picker
- `popup/`, `options/` – user interface
- `vendor/exceljs.min.js` – bundled ExcelJS browser build (MIT; see `vendor/EXCELJS-LICENSE`)
