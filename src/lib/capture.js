// Crop and hash helpers (run in the service worker using OffscreenCanvas).

export async function cropBlob(blob, settings, viewportWidth) {
  if (settings.captureMode !== 'region') return blob;
  const bmp = await createImageBitmap(blob);
  const scale = bmp.width / viewportWidth; // device pixel ratio
  const r = settings.region;
  const sx = Math.max(0, Math.round(r.x * scale));
  const sy = Math.max(0, Math.round(r.y * scale));
  const sw = Math.min(bmp.width - sx, Math.round(r.width * scale));
  const sh = Math.min(bmp.height - sy, Math.round(r.height * scale));
  if (sw <= 0 || sh <= 0) {
    bmp.close();
    throw new Error('The capture region is outside the visible page. Update it in Settings.');
  }
  const canvas = new OffscreenCanvas(sw, sh);
  canvas.getContext('2d').drawImage(bmp, sx, sy, sw, sh, 0, 0, sw, sh);
  bmp.close();
  return canvas.convertToBlob({ type: 'image/png' });
}

export async function sha256Hex(blob) {
  const digest = await crypto.subtle.digest('SHA-256', await blob.arrayBuffer());
  return Array.from(new Uint8Array(digest)).map((b) => b.toString(16).padStart(2, '0')).join('');
}

const p2 = (n) => String(n).padStart(2, '0');

export function timestamps(d = new Date()) {
  const date = `${d.getFullYear()}-${p2(d.getMonth() + 1)}-${p2(d.getDate())}`;
  const time = `${p2(d.getHours())}:${p2(d.getMinutes())}:${p2(d.getSeconds())}`;
  return { display: `${date} ${time}`, file: `${date.replace(/-/g, '')}_${time.replace(/:/g, '')}` };
}
