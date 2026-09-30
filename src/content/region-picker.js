// One-time helper: drag a rectangle over the page to define the fixed capture area.
(() => {
  if (window.__snapRegionPicker) return;
  window.__snapRegionPicker = true;

  const overlay = document.createElement('div');
  overlay.style.cssText = 'position:fixed;inset:0;z-index:2147483647;cursor:crosshair;background:rgba(15,23,42,.35);' +
    'font:14px system-ui,sans-serif;color:#fff;user-select:none;';
  const hint = document.createElement('div');
  hint.textContent = 'Drag to choose the capture area. Press Esc to cancel.';
  hint.style.cssText = 'position:absolute;top:16px;left:50%;transform:translateX(-50%);background:#0f172a;' +
    'padding:8px 14px;border-radius:8px;';
  const box = document.createElement('div');
  box.style.cssText = 'position:absolute;border:2px solid #34d399;background:rgba(52,211,153,.15);display:none;';
  const size = document.createElement('div');
  size.style.cssText = 'position:absolute;right:0;bottom:-26px;background:#0f172a;padding:2px 8px;border-radius:4px;font-size:12px;white-space:nowrap;';
  box.appendChild(size);
  overlay.append(hint, box);
  document.documentElement.appendChild(overlay);

  let start = null;
  const rect = (e) => ({
    x: Math.min(start.x, e.clientX),
    y: Math.min(start.y, e.clientY),
    width: Math.abs(e.clientX - start.x),
    height: Math.abs(e.clientY - start.y)
  });
  const close = () => {
    overlay.remove();
    document.removeEventListener('keydown', onKey, true);
    window.__snapRegionPicker = false;
  };
  const onKey = (e) => { if (e.key === 'Escape') close(); };
  document.addEventListener('keydown', onKey, true);

  overlay.addEventListener('pointerdown', (e) => {
    overlay.setPointerCapture(e.pointerId);
    start = { x: e.clientX, y: e.clientY };
    box.style.display = 'block';
  });
  overlay.addEventListener('pointermove', (e) => {
    if (!start) return;
    const r = rect(e);
    Object.assign(box.style, { left: r.x + 'px', top: r.y + 'px', width: r.width + 'px', height: r.height + 'px' });
    size.textContent = `${Math.round(r.width)} × ${Math.round(r.height)}`;
  });
  overlay.addEventListener('pointerup', (e) => {
    if (!start) return;
    const r = rect(e);
    close();
    if (r.width < 20 || r.height < 20) return;
    chrome.runtime.sendMessage({
      type: 'save-region',
      region: { x: Math.round(r.x), y: Math.round(r.y), width: Math.round(r.width), height: Math.round(r.height) }
    });
  });
})();
