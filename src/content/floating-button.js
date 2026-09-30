// Floating capture button. Injected only on sites the user enabled.
(() => {
  if (window.top !== window || window.__snapToExcelLoaded) return;
  window.__snapToExcelLoaded = true;

  const DEFAULTS = { showButton: true, buttonSize: 48, buttonOpacity: 0.9, buttonPos: null };
  let cfg = { ...DEFAULTS };
  let drag = null;
  let toastTimer = null;

  const host = document.createElement('div');
  host.style.cssText = 'all:initial;position:fixed;inset:0;pointer-events:none;z-index:2147483647;';
  const root = host.attachShadow({ mode: 'closed' });
  root.innerHTML = `
    <style>
      *{box-sizing:border-box}
      button{position:fixed;right:24px;bottom:24px;border:0;border-radius:50%;background:#0f6b4f;color:#fff;
        display:grid;place-items:center;cursor:grab;pointer-events:auto;touch-action:none;
        box-shadow:0 2px 8px rgba(0,0,0,.35);transition:transform .12s,background .12s}
      button:hover{transform:scale(1.06);background:#0b5a42}
      button:focus-visible{outline:3px solid #93c5fd;outline-offset:2px}
      button.busy{cursor:progress;pointer-events:none}
      button.busy svg{display:none}
      button.busy::after{content:"";width:40%;height:40%;border:3px solid rgba(255,255,255,.45);
        border-top-color:#fff;border-radius:50%;animation:spin .8s linear infinite}
      button.hidden{display:none}
      svg{width:52%;height:52%;pointer-events:none}
      .toast{position:fixed;top:16px;right:16px;max-width:320px;padding:10px 14px;border-radius:8px;
        font:14px/1.4 system-ui,-apple-system,Segoe UI,sans-serif;color:#fff;background:#334155;
        box-shadow:0 4px 14px rgba(0,0,0,.3);pointer-events:none;opacity:0;transform:translateY(-6px);
        transition:opacity .15s,transform .15s}
      .toast.show{opacity:1;transform:none}
      .toast.success{background:#0f6b4f}
      .toast.error{background:#b42318}
      @keyframes spin{to{transform:rotate(360deg)}}
      @media (prefers-reduced-motion:reduce){button,.toast{transition:none}button.busy::after{animation-duration:2s}}
    </style>
    <button type="button" aria-label="Capture screenshot to Excel" title="Capture screenshot (Alt+Shift+S). Drag to move.">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
        <path d="M4 8h3l2-3h6l2 3h3v11H4z"/><circle cx="12" cy="13" r="3.5"/>
      </svg>
    </button>
    <div class="toast" role="status" aria-live="polite"></div>`;
  const btn = root.querySelector('button');
  const toastEl = root.querySelector('.toast');
  document.documentElement.appendChild(host);

  function clamp(left, top) {
    const size = cfg.buttonSize;
    return {
      left: Math.min(Math.max(0, left), Math.max(0, window.innerWidth - size)),
      top: Math.min(Math.max(0, top), Math.max(0, window.innerHeight - size))
    };
  }

  function place(left, top) {
    const p = clamp(left, top);
    btn.style.left = p.left + 'px';
    btn.style.top = p.top + 'px';
    btn.style.right = 'auto';
    btn.style.bottom = 'auto';
    return p;
  }

  function apply() {
    btn.style.width = btn.style.height = cfg.buttonSize + 'px';
    btn.style.opacity = String(cfg.buttonOpacity);
    btn.classList.toggle('hidden', !cfg.showButton);
    if (cfg.buttonPos) {
      place(cfg.buttonPos.left, cfg.buttonPos.top);
    } else {
      btn.style.left = btn.style.top = 'auto';
      btn.style.right = btn.style.bottom = '24px';
    }
  }

  function toast(kind, text) {
    toastEl.textContent = text;
    toastEl.className = `toast show ${kind}`;
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => toastEl.classList.remove('show'), kind === 'error' ? 6000 : 3500);
  }

  async function capture() {
    if (btn.classList.contains('busy')) return;
    btn.classList.add('busy');
    try {
      await chrome.runtime.sendMessage({ type: 'capture-request' });
    } catch {
      toast('error', 'Snap to Excel was updated. Reload this page to keep using it.');
    } finally {
      btn.classList.remove('busy');
    }
  }

  btn.addEventListener('pointerdown', (e) => {
    if (e.button !== 0) return;
    btn.setPointerCapture(e.pointerId);
    const r = btn.getBoundingClientRect();
    drag = { dx: e.clientX - r.left, dy: e.clientY - r.top, sx: e.clientX, sy: e.clientY, moved: false };
  });
  btn.addEventListener('pointermove', (e) => {
    if (!drag) return;
    if (!drag.moved && Math.hypot(e.clientX - drag.sx, e.clientY - drag.sy) > 4) drag.moved = true;
    if (drag.moved) place(e.clientX - drag.dx, e.clientY - drag.dy);
  });
  btn.addEventListener('pointerup', async (e) => {
    if (!drag) return;
    const d = drag;
    drag = null;
    if (d.moved) {
      const p = clamp(e.clientX - d.dx, e.clientY - d.dy);
      const { settings = {} } = await chrome.storage.local.get('settings');
      chrome.storage.local.set({ settings: { ...settings, buttonPos: p } });
    } else {
      capture();
    }
  });
  btn.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); capture(); }
  });
  window.addEventListener('resize', apply);

  chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
    if (msg.type === 'prepare') {
      host.style.visibility = 'hidden';
      // Let the browser repaint without the button before the screenshot is taken.
      requestAnimationFrame(() => requestAnimationFrame(() => setTimeout(() => sendResponse({}), 60)));
      return true;
    }
    if (msg.type === 'restore') { host.style.visibility = ''; sendResponse({}); }
    if (msg.type === 'toast') { toast(msg.kind, msg.text); sendResponse({}); }
    return false;
  });

  chrome.storage.onChanged.addListener((changes, area) => {
    if (area === 'local' && changes.settings) {
      cfg = { ...DEFAULTS, ...(changes.settings.newValue || {}) };
      apply();
    }
  });
  chrome.storage.local.get('settings').then(({ settings }) => {
    cfg = { ...DEFAULTS, ...(settings || {}) };
    apply();
  });
})();
