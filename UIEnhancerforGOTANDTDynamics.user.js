// ==UserScript==
// @name         UIEnhancerforGOTANDTDynamics
// @namespace    https://github.com/mtoy30/GoTandT
// @version      1.3.7.27
// @updateURL    https://raw.githubusercontent.com/mtoy30/GoTandT/main/UIEnhancerforGOTANDTDynamics.user.js
// @downloadURL  https://raw.githubusercontent.com/mtoy30/GoTandT/main/UIEnhancerforGOTANDTDynamics.user.js
// @description  Dynamics UI tweaks; Boomerang form autofill (clipboard → GM storage bridge → googleusercontent iframe); PowerApps Copy button for Leg Info overlay; Uber Health ride autofill from Excel.
// @author       Michael Toy
// @match        https://health.uber.com/*
// @match        https://*.powerapps.com/*
// @match        https://*.powerplatform.com/*
// @match        https://gotandt.crm.dynamics.com/*
// @match        https://gttqap2.crm.dynamics.com/*
// @match        https://boomerangtransport.net/ride-input-request*
// @match        https://*.googleusercontent.com/*
// @match        https://script.google.com/*
// @match        https://script.googleusercontent.com/*
// @include      /^https:\/\/[^/]+-script\.googleusercontent\.com\/.*/
// @grant        GM_setClipboard
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        GM_deleteValue
// @grant        GM_xmlhttpRequest
// @grant        unsafeWindow
// @connect      lowmargin.mtoysystems.com
// @run-at       document-start
// ==/UserScript==

(function () {
  'use strict';

  const host = location.hostname;
  const path = location.pathname;

  const isPowerApps      = /\.powerapps\.com$/i.test(host) || /\.powerplatform\.com$/i.test(host);
  const isDynamics       = /(?:^|\.)gotandt\.crm\.dynamics\.com$/i.test(host) || /(?:^|\.)gttqap2\.crm\.dynamics\.com$/i.test(host);
  const onBoomerang      = host === 'boomerangtransport.net';
  const onAppsScript     = /(?:^|\.)googleusercontent\.com$/i.test(host) || host === 'script.google.com' || host === 'script.googleusercontent.com';

  // Boomerang autofill frame detection
  const isGoogleScript     = host === 'script.google.com';
  const isGoogleContent    = host.endsWith('script.googleusercontent.com');
  const isUserCodeAppPanel = isGoogleContent && path.includes('userCodeAppPanel');

  /* =====================================================================================
     PART A — POWER APPS: "Copy" button above the date inside Leg Info overlay
     ===================================================================================== */
  if (isPowerApps) {
    const sleep = (ms) => new Promise(r => setTimeout(r, ms));
    const until = async (fn, { tries = 400, delay = 100 } = {}) => {
      for (let i = 0; i < tries; i++) { const v = fn(); if (v) return v; await sleep(delay); }
      return null;
    };
    const isVisible = (el) => {
      if (!el) return false;
      const cs = getComputedStyle(el);
      if (cs.display === 'none' || cs.visibility === 'hidden' || cs.opacity === '0') return false;
      const r = el.getBoundingClientRect();
      return r.width > 0 && r.height > 0;
    };
    const compact = (s) => (s || '').replace(/\s+/g, ' ').trim();

    const TEXT_SELECTORS = [
      '.appmagic-label-text',
      '.appmagic-html-text',
      '.appmagic-richTextContainer',
      '.appmagic-textarea-text',
      '.appmagic-textinput-inner',
      '[data-control-part="text"]'
    ].join(',');

    const DATE_RE = /\b\d{1,2}\/\d{1,2}\/\d{2,4}(?:\s+\d{1,2}:\d{2}\s*[AP]M?)?\b/i;
    const CITY_RE = /^[A-Za-z].+,\s+[A-Z][a-zA-Z]+(?:\s[A-Z][a-zA-Z]+)*\s+\d{5}(?:-\d{4})?$/;

    function findOverlayContainer() {
      let cands = Array.from(document.querySelectorAll(
        '[role="dialog"], [aria-modal="true"], [class*="dialog"], [class*="Dialog"], [class*="popup"], [class*="Popup"], [class*="overlay"], [class*="Overlay"]'
      )).filter(isVisible);

      if (!cands.length) {
        cands = Array.from(document.querySelectorAll('div')).filter(el => {
          if (!isVisible(el)) return false;
          const cs = getComputedStyle(el);
          if (!/^(fixed|absolute)$/.test(cs.position)) return false;
          const r = el.getBoundingClientRect();
          const vw = innerWidth, vh = innerHeight;
          return r.width > 600 && r.height > 300 &&
                 r.left > 20 && r.top > 60 && (r.right < vw - 20) && (r.bottom < vh - 20);
        });
      }
      if (!cands.length) return null;

      let best = null, area = 0;
      for (const c of cands) {
        const r = c.getBoundingClientRect();
        const a = r.width * r.height;
        if (a > area) { area = a; best = c; }
      }
      return best;
    }

    function findDateAnchorIn(container) {
      const labels = Array.from(container.querySelectorAll(
        'div.appmagic-label-text[data-control-part="text"], [data-control-part="text"]'
      )).filter(isVisible);

      for (const n of labels) {
        const t = compact(n.innerText || n.textContent || '');
        if (DATE_RE.test(t)) return n;
      }
      return labels[0] || null;
    }

    function domPrecedes(a, b) { return !!(a.compareDocumentPosition(b) & Node.DOCUMENT_POSITION_FOLLOWING); }

    function sameVisualRow(a, b, minOverlapRatio = 0.35) {
      const top = Math.max(a.rect.top, b.rect.top);
      const bottom = Math.min(a.rect.bottom, b.rect.bottom);
      const overlap = Math.max(0, bottom - top);
      const denom = Math.max(a.rect.height, b.rect.height) || 1;
      return (overlap / denom) >= minOverlapRatio;
    }

    let legCopyInProgress = false;

    function collectCurrentOverlayLines(container) {
      let nodes = Array.from(container.querySelectorAll(TEXT_SELECTORS)).filter(isVisible);
      nodes = nodes.filter(n => !n.closest('#mtoy-inline-copy') && !n.closest('#mtoy-copy-toast'));

      const set = new Set(nodes);
      nodes = nodes.filter(n => !Array.from(set).some(m => m !== n && n.contains(m)));

      const items = nodes.map(el => {
        const rect = el.getBoundingClientRect();
        const text = compact(el.innerText || el.textContent || '');
        return { el, rect, text };
      }).filter(i => i.text);

      if (!items.length) return [];

      items.sort((a, b) => {
        if (!sameVisualRow(a, b)) {
          const topDelta = a.rect.top - b.rect.top;
          if (topDelta !== 0) return topDelta;
          const leftDelta = a.rect.left - b.rect.left;
          if (leftDelta !== 0) return leftDelta;
        } else {
          const aCity = CITY_RE.test(a.text), bCity = CITY_RE.test(b.text);
          if (aCity && bCity) {
            const dx = a.rect.left - b.rect.left;
            if (dx !== 0) return dx;
          }
          if (a.el !== b.el) return domPrecedes(a.el, b.el) ? -1 : 1;
        }
        return 0;
      });

      return items.map(it => it.text);
    }

    function mergeOverlayLines(existing, incoming) {
      if (!existing.length) return incoming.slice();
      if (!incoming.length) return existing;

      const maxOverlap = Math.min(existing.length, incoming.length);
      for (let size = maxOverlap; size >= 2; size--) {
        let match = true;
        for (let i = 0; i < size; i++) {
          if (existing[existing.length - size + i] !== incoming[i]) {
            match = false;
            break;
          }
        }
        if (match) {
          existing.push(...incoming.slice(size));
          return existing;
        }
      }

      existing.push(...incoming);
      return existing;
    }

    /*
       The Copy button/date row lives INSIDE the same scrollable Leg Info panel
       as all of the legs. Walk upward from that row and use the nearest real
       vertical scroll container. This avoids scrolling the Dynamics page or
       unrelated Power Apps wrappers.
    */
    function findLegInfoScroller(container) {
      const anchor = findDateAnchorIn(container);
      if (!anchor) return null;

      let firstScrollable = null;
      let p = anchor.parentElement;

      while (p && p !== document.documentElement) {
        try {
          const clientHeight = p.clientHeight || 0;
          const scrollHeight = p.scrollHeight || 0;
          const range = scrollHeight - clientHeight;

          if (clientHeight >= 100 && range > 8) {
            if (!firstScrollable) firstScrollable = p;

            const cs = getComputedStyle(p);
            const overflowY = cs.overflowY || '';

            // Prefer the nearest ancestor that is explicitly the scroll viewport.
            if (/auto|scroll|overlay/i.test(overflowY)) return p;
          }
        } catch {}

        if (p === container) break;
        p = p.parentElement;
      }

      // Power Apps occasionally reports overflow:hidden while still using
      // scrollTop internally. In that case, the nearest ancestor with an actual
      // scroll range is still the correct Leg Info viewport.
      return firstScrollable;
    }

    function legNumberFromId(id) {
      const m = String(id || '').match(/-(\d+)\s*$/);
      return m ? parseInt(m[1], 10) : Number.MAX_SAFE_INTEGER;
    }

    function extractLegBlocks(lines) {
      const starts = [];

      for (let i = 0; i < lines.length - 1; i++) {
        if (/^Leg:\s*$/i.test(lines[i])) {
          const id = lines[i + 1] || '';
          if (/\S+-\d+\s*$/.test(id)) {
            starts.push({
              legIndex: i,
              start: (i > 0 && DATE_RE.test(lines[i - 1])) ? i - 1 : i,
              id: id.trim()
            });
          }
        }
      }

      const blocks = [];
      for (let i = 0; i < starts.length; i++) {
        const cur = starts[i];
        const nextStart = i + 1 < starts.length ? starts[i + 1].start : lines.length;
        const blockLines = lines.slice(cur.start, nextStart).filter(Boolean);
        if (blockLines.length >= 3) blocks.push({ id: cur.id, lines: blockLines });
      }

      return blocks;
    }

    async function waitForPowerAppsRender() {
      // Two animation frames + a short delay gives the virtualized gallery time
      // to recycle its rows after scrollTop changes.
      await new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)));
      await sleep(260);
    }

    async function scrollLegPanel(scroller, top) {
      const maxScroll = Math.max(0, scroller.scrollHeight - scroller.clientHeight);
      const target = Math.max(0, Math.min(top, maxScroll));

      try { scroller.scrollTop = target; } catch {}
      try { scroller.scrollTo?.(0, target); } catch {}
      try { scroller.dispatchEvent(new Event('scroll', { bubbles: true })); } catch {}

      await waitForPowerAppsRender();
    }

    async function collectOverlayText(container) {
      const scroller = findLegInfoScroller(container);
      const originalTop = scroller ? (scroller.scrollTop || 0) : 0;
      const originalBehavior = scroller ? scroller.style.scrollBehavior : '';

      const legMap = new Map();
      const seenOrder = new Map();
      let orderCounter = 0;
      let fallbackLines = [];

      const capture = () => {
        const lines = collectCurrentOverlayLines(container);
        if (!lines.length) return;

        fallbackLines = mergeOverlayLines(fallbackLines, lines);

        for (const block of extractLegBlocks(lines)) {
          if (!seenOrder.has(block.id)) seenOrder.set(block.id, orderCounter++);

          const existing = legMap.get(block.id);
          // A leg can be clipped at the top/bottom of the viewport. Keep the
          // snapshot containing the greatest number of fields for that leg.
          if (!existing || block.lines.length > existing.length) {
            legMap.set(block.id, block.lines.slice());
          }
        }
      };

      try {
        // Always capture what is currently visible first.
        capture();

        if (scroller) {
          try { scroller.style.scrollBehavior = 'auto'; } catch {}

          // Start at the top of the same white Leg Info frame shown in the UI.
          await scrollLegPanel(scroller, 0);
          capture();

          let lastTop = scroller.scrollTop || 0;
          let loops = 0;

          while (loops++ < 100) {
            const maxScroll = Math.max(0, scroller.scrollHeight - scroller.clientHeight);
            const current = scroller.scrollTop || 0;

            if (current >= maxScroll - 2) break;

            // Keep substantial overlap between captures so a leg split across
            // two screens is eventually captured in full.
            const step = Math.max(90, Math.floor((scroller.clientHeight || 300) * 0.42));
            const target = Math.min(current + step, maxScroll);

            await scrollLegPanel(scroller, target);
            capture();

            const after = scroller.scrollTop || 0;
            if (Math.abs(after - lastTop) < 1) break;
            lastTop = after;
          }

          // Explicit bottom capture is the important part for referrals where
          // leg 5+ is not rendered until the scrollbar reaches the bottom.
          const bottom = Math.max(0, scroller.scrollHeight - scroller.clientHeight);
          await scrollLegPanel(scroller, bottom);
          capture();
          await waitForPowerAppsRender();
          capture();
        }
      } finally {
        if (scroller) {
          try {
            scroller.scrollTop = originalTop;
            scroller.style.scrollBehavior = originalBehavior;
            scroller.dispatchEvent(new Event('scroll', { bubbles: true }));
          } catch {}
        }
      }

      if (legMap.size) {
        const ids = Array.from(legMap.keys()).sort((a, b) => {
          const na = legNumberFromId(a), nb = legNumberFromId(b);
          if (na !== nb) return na - nb;
          return (seenOrder.get(a) || 0) - (seenOrder.get(b) || 0);
        });

        return ids.flatMap(id => legMap.get(id)).join('\n');
      }

      // If Power Apps changes its Leg markup, never return an empty clipboard;
      // fall back to the text gathered from the panel while scrolling.
      return fallbackLines.join('\n');
    }

    function showToast(msg) {
      let t = document.getElementById('mtoy-copy-toast');
      if (!t) {
        t = document.createElement('div');
        t.id = 'mtoy-copy-toast';
        t.style.cssText = `
          position: fixed; left: 50%; top: 20px; transform: translateX(-50%);
          background: #ffe44d; color:#111; padding: 8px 12px; border-radius:10px;
          z-index: 2147483647; pointer-events: none; box-shadow: 0 4px 18px rgba(0,0,0,.25);
          font: 13px/1.3 system-ui,-apple-system,Segoe UI,Roboto,Arial;
        `;
        document.body.appendChild(t);
      }
      t.textContent = msg;
      t.style.opacity = '1';
      setTimeout(() => (t.style.opacity = '0'), 1400);
    }

    function injectInlineButton() {
      if (legCopyInProgress) return;

      const overlay = findOverlayContainer();
      const existing = document.getElementById('mtoy-inline-copy');

      if (!overlay) { if (existing) existing.remove(); return; }

      const anchor = findDateAnchorIn(overlay);
      if (!anchor) { if (existing) existing.remove(); return; }

      if (existing && existing.nextElementSibling === anchor) return;
      if (existing) existing.remove();

      const btn = document.createElement('button');
      btn.id = 'mtoy-inline-copy';
      btn.type = 'button';
      btn.textContent = 'Copy';
      btn.title = 'Copy all legs in this Leg Info panel';
      btn.style.cssText = `
        display: inline-block;
        margin: 0 8px 2px 0;
        padding: 6px 10px;
        border-radius: 999px;
        cursor: pointer;
        background: #ffe44d;
        color: #111;
        border: 1px solid rgba(0,0,0,.25);
        font: 13px/1.2 system-ui,-apple-system,Segoe UI,Roboto,Arial;
        box-shadow: 0 2px 10px rgba(0,0,0,.15);
      `;
      btn.addEventListener('click', async () => {
        if (legCopyInProgress) return;

        legCopyInProgress = true;
        btn.disabled = true;
        btn.textContent = 'Copying…';

        try {
          const text = await collectOverlayText(overlay);
          if (!text) { alert('Nothing visible to copy.'); return; }

          if (typeof GM_setClipboard === 'function') {
            GM_setClipboard(text, { type: 'text', mimetype: 'text/plain' });
          } else {
            await navigator.clipboard.writeText(text);
          }

          showToast('Copied all legs.');
        } catch (e) {
          console.error('Copy failed', e);
          alert('Copy failed. See console.');
        } finally {
          legCopyInProgress = false;
          if (btn.isConnected) {
            btn.disabled = false;
            btn.textContent = 'Copy';
          }
          injectInlineButton();
        }
      });

      anchor.parentNode.insertBefore(btn, anchor);
    }

    (async () => {
      await until(() => document.body, { tries: 600, delay: 50 });
      const tick = () => injectInlineButton();
      tick();
      const mo = new MutationObserver(tick);
      mo.observe(document.documentElement, { childList: true, subtree: true });
      setInterval(tick, 800);
    })();

    return;
  }

  /* =====================================================================================
     PART B — UI ENHANCER (Dynamics, Boomerang, Apps Script)
     ===================================================================================== */

  if (!(isDynamics || onBoomerang || onAppsScript)) return;

  /* ============================== BOOMERANG/APPS SCRIPT ============================== */
  if (onBoomerang || onAppsScript) {
    const TIME_FIELD_IDS  = new Set(['apptTime','pickupTime']);
    const TIME_FIELD_HINT = /(time|appt[_-]?time|pickup[_-]?time|appointment[_-]?time)$/i;

    const ALWAYS_UNLOCK_IDS   = new Set(['yourName','yourCompany','yourEmail','passName','refNumber']);
    const ALWAYS_UNLOCK_NAMES = new Set(['yourName','yourCompany','yourEmail','passName','refNumber']);

    const AUTOCOMPLETE_HINTS_BY_ID = {
      yourCompany: 'organization',
      yourEmail  : 'email',
      yourName   : 'name',
      passName   : 'passname',
      refNumber  : 'off'
    };
    const AUTOCOMPLETE_HINTS_BY_NAME = AUTOCOMPLETE_HINTS_BY_ID;

    const SOFT_PREFILL_BY_ID = {
      yourCompany: 'Go T&T',
      yourEmail  : 'driverdeveloper@gotandt.com'
    };
    const SOFT_PREFILL_BY_NAME = SOFT_PREFILL_BY_ID;

    const UNLOCK_WINDOW_MS = 30000;

    const YOURNAME_OPTIONS = [
      'Annejulia Villegas-Torres','Ashley Oliver','Christian Antunez','Christina Armstrong',
      'Cooper Turner','Damaris Olmeda','David Hobbs','Jeremy Rivera','Kevin Roberts','Melantony Burks',
      'Michael Toy','Naomi Picklesimer','Tonyjay Matias','Yuri Nichols'
    ];

    function isTimeField(el) {
      if (!el || !el.matches || el.tagName !== 'INPUT') return false;
      const id   = el.id || '';
      const name = el.getAttribute('name') || '';
      const cls  = el.className || '';
      const type = (el.getAttribute('type') || '').toLowerCase();
      return (
        type === 'time' ||
        TIME_FIELD_IDS.has(id) || TIME_FIELD_IDS.has(name) ||
        /validate-time|ui-timepicker-input|timepicker/i.test(cls) ||
        TIME_FIELD_HINT.test(id) || TIME_FIELD_HINT.test(name)
      );
    }
    function isAlwaysUnlockField(el) {
      const id = el.id || '';
      const name = el.getAttribute('name') || '';
      return ALWAYS_UNLOCK_IDS.has(id) || ALWAYS_UNLOCK_NAMES.has(name);
    }
    function getAutocompleteHint(el) {
      const id = el.id || '';
      const name = el.getAttribute('name') || '';
      return AUTOCOMPLETE_HINTS_BY_ID[id] ?? AUTOCOMPLETE_HINTS_BY_NAME[name] ?? null;
    }
    function getSoftPrefill(el) {
      const id = el.id || '';
      const name = el.getAttribute('name') || '';
      return SOFT_PREFILL_BY_ID[id] ?? SOFT_PREFILL_BY_NAME[name] ?? null;
    }
    function alwaysUnlock(el) {
      try {
        el.readOnly = false;
        el.removeAttribute('readonly');
        el.dataset._locked = 'always-unlocked';
        const sync = () => { el.dataset._lastSnapshot = el.value ?? ''; };
        el.addEventListener('input',  sync, true);
        el.addEventListener('change', sync, true);
        el.addEventListener('blur',   sync, true);
      } catch {}
    }
    function attachCompanyDatalist(el) {
      if (el.dataset._companyListAttached === '1') return;
      const id = 'companyChoices';
      let list = el.ownerDocument.getElementById(id);
      if (!list) {
        list = el.ownerDocument.createElement('datalist');
        list.id = id;
        const opt = el.ownerDocument.createElement('option');
        opt.value = 'Go T&T';
        list.appendChild(opt);
        el.ownerDocument.documentElement.appendChild(list);
      }
      el.setAttribute('list', id);
      el.dataset._companyListAttached = '1';
    }
    function attachNameDatalist(el) {
      if (el.dataset._nameListAttached === '1') return;
      const id = 'nameChoices';
      let list = el.ownerDocument.getElementById(id);
      if (!list) {
        list = el.ownerDocument.createElement('datalist');
        list.id = id;
        YOURNAME_OPTIONS.forEach(n => {
          const opt = el.ownerDocument.createElement('option');
          opt.value = n;
          list.appendChild(opt);
        });
        el.ownerDocument.documentElement.appendChild(list);
      }
      el.setAttribute('list', id);
      el.setAttribute('autocomplete', 'off');
      el.dataset._nameListAttached = '1';
    }
    function unlockAllInForm(scope) {
      const root = scope && scope.querySelectorAll ? scope : document;
      root.querySelectorAll?.('input, textarea, select').forEach(f => {
        f.readOnly = false;
        f.removeAttribute('readonly');
      });
    }
    function beginUnlockWithDebounce(el) {
      if (isTimeField(el) || isAlwaysUnlockField(el)) { alwaysUnlock(el); return; }
      el.readOnly = false;
      el.dataset._locked = 'unlocked';

      const refresh = () => {
        const token = Math.random().toString(36).slice(2);
        el.dataset._unlockToken = token;
        if (el._unlockTimer) clearTimeout(el._unlockTimer);
        el._unlockTimer = setTimeout(() => {
          if (el.dataset._unlockToken === token) {
            el.dataset._lastSnapshot = el.value ?? '';
            if ((el.value ?? '').trim() !== '') {
              el.readOnly = true; el.dataset._locked = 'hard';
            } else {
              el.readOnly = true; el.dataset._locked = 'soft';
            }
          }
        }, UNLOCK_WINDOW_MS);
      };

      if (!el._unlockHandlersAttached) {
        ['input','keydown','paste','pointerdown'].forEach(evt =>
          el.addEventListener(evt, refresh, true)
        );
        el._unlockHandlersAttached = true;
      }

      refresh();
    }

    const TEXTLIKE = new Set(['text','email','tel','search','url','number','password','date','datetime-local','time','month','week']);

    function mark(el) {
      if (!el || el.dataset._guardApplied === '1') return;

      if (el.tagName === 'FORM') {
        el.setAttribute('autocomplete','off');
        el.addEventListener('submit', () => { unlockAllInForm(el); }, { capture: true });
      }

      if (el.matches?.('input, textarea, select')) {
        const tag  = el.tagName.toLowerCase();
        const type = (el.getAttribute('type') || 'text').toLowerCase();

        if (!isAlwaysUnlockField(el)) {
          el.setAttribute('autocomplete','off');
          if (tag === 'input' && TEXTLIKE.has(type)) el.setAttribute('autocomplete','new-password');
        }
        el.setAttribute('autocorrect','off');
        el.setAttribute('autocapitalize','off');
        el.setAttribute('spellcheck','false');
        el.dataset._lastSnapshot = el.value ?? '';

        if (isAlwaysUnlockField(el)) {
          if ((el.id === 'yourName') || (el.getAttribute('name') === 'yourName')) {
            attachNameDatalist(el);
          } else {
            const hint = getAutocompleteHint(el);
            if (hint) el.setAttribute('autocomplete', hint);
          }
          el.removeAttribute('aria-autocomplete');
          alwaysUnlock(el);
          const defVal = getSoftPrefill(el);
          if (defVal && (el.value ?? '').trim() === '') {
            el.value = defVal;
            el.dispatchEvent(new Event('input',  {bubbles:true}));
            el.dispatchEvent(new Event('change', {bubbles:true}));
            el.dataset._lastSnapshot = el.value ?? '';
          }
          if ((el.id === 'yourCompany') || (el.getAttribute('name') === 'yourCompany')) {
            attachCompanyDatalist(el);
          }
          const keepUnlocked = () => { el.readOnly = false; el.removeAttribute('readonly'); };
          el.addEventListener('focus', keepUnlocked, true);
          el.addEventListener('blur',  () => { el.dataset._lastSnapshot = el.value ?? ''; }, true);
          el.dataset._guardApplied = '1';
          return;
        }

        if (isTimeField(el)) {
          alwaysUnlock(el);
          el.addEventListener('focus', () => alwaysUnlock(el), true);
          el.addEventListener('blur',  () => alwaysUnlock(el), true);
          el.dataset._guardApplied = '1';
          return;
        }

        const protectIfFilled = () => {
          if ((el.value ?? '').trim() !== '') {
            el.readOnly = true; el.dataset._locked = 'hard';
          } else {
            el.readOnly = true; el.dataset._locked = 'soft';
          }
        };
        protectIfFilled();

        let lastUserIntentTs = 0;
        const stampIntent = () => { lastUserIntentTs = Date.now(); };
        const tryUnlock = () => {
          if ((el.value ?? '').trim() === '') beginUnlockWithDebounce(el);
        };

        el.addEventListener('pointerdown', () => { stampIntent(); tryUnlock(); }, true);
        el.addEventListener('keydown',     () => { stampIntent(); tryUnlock(); }, true);
        el.addEventListener('paste',       () => { stampIntent(); beginUnlockWithDebounce(el); }, true);
        el.addEventListener('focus',       () => { tryUnlock(); }, true);

        const guardInput = () => {
          const now = Date.now();
          const userLikely = (now - lastUserIntentTs) <= 1000;
          const locked = el.readOnly === true;
          if (locked && !userLikely) {
            const snap = el.dataset._lastSnapshot ?? '';
            if (el.value !== snap) el.value = snap;
            return;
          }
          el.dataset._lastSnapshot = el.value ?? '';
        };
        el.addEventListener('input',  guardInput, true);
        el.addEventListener('change', guardInput, true);

        el.addEventListener('blur', () => {
          el.dataset._lastSnapshot = el.value ?? '';
          protectIfFilled();
        }, true);
      }

      el.dataset._guardApplied = '1';
    }

    function unlockAllInFormScope(doc) {
      doc.querySelectorAll('input, textarea, select').forEach(f => {
        f.readOnly = false;
        f.removeAttribute('readonly');
      });
    }

    function wireUnlockOnSubmitClicks(doc) {
      doc.addEventListener('click', (e) => {
        const t = e.target;
        if (!t) return;
        if (t.matches && (t.matches('button[type="submit"], input[type="submit"]') || /submit/i.test(t.getAttribute?.('type') || ''))) {
          const form = t.form || doc.querySelector('form');
          unlockAllInFormScope(form || doc);
        }
      }, true);
    }

    function wireUnlockForAppsScriptSubmit(doc) {
      doc.addEventListener('click', (e) => {
        const btn = e.target && e.target.closest ? e.target.closest('#buttonSubmit, .submit-button') : null;
        if (!btn) return;
        const container = btn.parentElement || doc;
        unlockAllInFormScope(container);
        container.querySelectorAll('input, textarea, select').forEach(f => { try { f.blur && f.blur(); } catch {} });
      }, true);
      doc.addEventListener('keydown', (e) => {
        if (!e) return;
        const active = doc.activeElement;
        if (!active) return;
        const isSubmitLink = active.matches?.('#buttonSubmit, .submit-button');
        if (!isSubmitLink) return;
        if (e.key === 'Enter' || e.key === ' ' || e.keyCode === 13 || e.keyCode === 32) {
          const container = active.parentElement || doc;
          unlockAllInFormScope(container);
          container.querySelectorAll('input, textarea, select').forEach(f => { try { f.blur && f.blur(); } catch {} });
        }
      }, true);
    }

    function apply(doc) {
      doc.querySelectorAll('form, input, textarea, select').forEach(mark);
      const mo = new doc.defaultView.MutationObserver(muts => {
        for (const m of muts) {
          if (m.type === 'childList') {
            m.addedNodes.forEach(n => {
              if (n.nodeType !== 1) return;
              mark(n);
              n.querySelectorAll?.('form, input, textarea, select').forEach(mark);
              n.querySelectorAll?.('iframe').forEach(handleFrame);
            });
          } else if (m.type === 'attributes' &&
                    (m.attributeName === 'autocomplete' || m.attributeName === 'readonly' || m.attributeName === 'name' || m.attributeName === 'id')) {
            mark(m.target);
          }
        }
      });
      mo.observe(doc.documentElement, { childList: true, subtree: true, attributes: true });
      doc.querySelectorAll('iframe').forEach(handleFrame);
      wireUnlockOnSubmitClicks(doc);
      wireUnlockForAppsScriptSubmit(doc);
    }

    function handleFrame(ifr) {
      const tryWire = () => {
        try {
          const idoc = ifr.contentDocument || ifr.contentWindow?.document;
          if (idoc && idoc.documentElement) {
            if (idoc.readyState === 'loading') {
              idoc.addEventListener('DOMContentLoaded', () => apply(idoc), { once: true });
            } else {
              apply(idoc);
            }
          }
        } catch {}
      };
      ifr.addEventListener('load', tryWire);
      tryWire();
    }

    if (document.readyState === 'loading') {
      document.addEventListener('DOMContentLoaded', () => apply(document));
    } else {
      apply(document);
    }
  }

  /* =====================================================================================
     PART C — BOOMERANG AUTOFILL (clipboard → GM storage → userCodeAppPanel → form)
     ===================================================================================== */

  const BTAF_GM_KEY        = 'btaf_ride_data';
  const BTAF_GM_TS         = 'btaf_ride_timestamp';
  const BTAF_GM_VISIBILITY = 'btaf_visible_ts';  // written only by the visible tab
  const btafSleep          = ms => new Promise(r => setTimeout(r, ms));

  // ── Main page: inject "Paste Ride Data" button ───────────────────────────────
  if (onBoomerang) {

    function showBtafToast(msg, color) {
      color = color || '#2e7d32';
      let t = document.getElementById('btaf-toast');
      if (!t) {
        t = document.createElement('div');
        t.id = 'btaf-toast';
        Object.assign(t.style, {
          position: 'fixed', top: '20px', left: '50%', transform: 'translateX(-50%)',
          padding: '12px 22px', borderRadius: '8px', fontSize: '14px',
          fontWeight: 'bold', color: '#fff', zIndex: '2147483647',
          maxWidth: '520px', boxShadow: '0 4px 14px rgba(0,0,0,.35)',
          transition: 'opacity .4s ease',
          fontFamily: 'system-ui,-apple-system,Segoe UI,Roboto,Arial',
          lineHeight: '1.5', whiteSpace: 'pre-wrap', pointerEvents: 'none',
          textAlign: 'center',
        });
        document.body.appendChild(t);
      }
      t.style.background = color;
      t.style.opacity    = '1';
      t.textContent      = msg;
      clearTimeout(t._timer);
      t._timer = setTimeout(() => { t.style.opacity = '0'; }, 5000);
    }

    async function btafSendData(data) {
      const now = Date.now();
      // Stamp this tab's visibility time — only the tab where you clicked
      // the button (which must be visible/active) writes this.
      // userCodeAppPanel frames use this to know which tab is the active one.
      await GM_setValue(BTAF_GM_VISIBILITY, now.toString());
      await GM_setValue(BTAF_GM_KEY, JSON.stringify(data));
      await GM_setValue(BTAF_GM_TS, now.toString());
      showBtafToast('📨 Data sent — filling fields now…', '#1565c0');
    }

    function showBtafModal() {
      const old = document.getElementById('btaf-modal');
      if (old) old.remove();

      const overlay = document.createElement('div');
      overlay.id = 'btaf-modal';
      Object.assign(overlay.style, {
        position: 'fixed', inset: '0', zIndex: '2147483645',
        background: 'rgba(0,0,0,0.55)', display: 'flex',
        alignItems: 'center', justifyContent: 'center',
      });

      const box = document.createElement('div');
      Object.assign(box.style, {
        background: '#fff', borderRadius: '12px', padding: '28px 32px',
        boxShadow: '0 8px 32px rgba(0,0,0,.35)', maxWidth: '440px', width: '90%',
        fontFamily: 'system-ui,-apple-system,Segoe UI,Roboto,Arial', textAlign: 'center',
      });

      const title = document.createElement('div');
      title.textContent = '📋 Paste Ride Data';
      Object.assign(title.style, { fontSize: '18px', fontWeight: 'bold', marginBottom: '8px', color: '#333' });

      const hint = document.createElement('div');
      hint.textContent = 'Press Ctrl+V to paste your Excel data, then click Paste & Fill.';
      Object.assign(hint.style, { fontSize: '13px', color: '#555', marginBottom: '14px' });

      const ta = document.createElement('textarea');
      ta.placeholder = 'Paste JSON here (Ctrl+V)…';
      Object.assign(ta.style, {
        width: '100%', height: '90px', borderRadius: '6px', padding: '8px',
        border: '1.5px solid #ccc', fontSize: '12px', fontFamily: 'monospace',
        resize: 'none', boxSizing: 'border-box', marginBottom: '12px', color: '#222',
      });

      const btnRow = document.createElement('div');
      Object.assign(btnRow.style, { display: 'flex', gap: '10px', justifyContent: 'center', marginBottom: '8px' });

      const pasteBtn = document.createElement('button');
      pasteBtn.textContent = '✅ Paste & Fill';
      Object.assign(pasteBtn.style, {
        background: '#f57c00', color: '#fff', border: 'none',
        padding: '10px 22px', borderRadius: '7px', fontSize: '14px',
        fontWeight: 'bold', cursor: 'pointer',
      });

      const cancelBtn = document.createElement('button');
      cancelBtn.textContent = 'Cancel';
      Object.assign(cancelBtn.style, {
        background: '#eee', color: '#333', border: 'none',
        padding: '10px 18px', borderRadius: '7px', fontSize: '14px', cursor: 'pointer',
      });

      const errMsg = document.createElement('div');
      Object.assign(errMsg.style, { color: '#c62828', fontSize: '12px', minHeight: '16px', textAlign: 'left' });

      function attempt(text) {
        text = (text || '').trim();
        if (!text) { errMsg.textContent = 'Nothing pasted — press Ctrl+V first.'; return; }
        let data;
        try { data = JSON.parse(text); }
        catch { errMsg.textContent = '⚠️ Not valid JSON. Check your VBA macro output.'; return; }
        overlay.remove();
        btafSendData(data);
      }

      pasteBtn.addEventListener('click', () => attempt(ta.value));
      ta.addEventListener('keydown', e => { if (e.key === 'Enter' && e.ctrlKey) { e.preventDefault(); attempt(ta.value); } });
      ta.addEventListener('paste', () => setTimeout(() => { if (ta.value.trim().startsWith('{')) attempt(ta.value); }, 80));
      cancelBtn.addEventListener('click', () => overlay.remove());
      overlay.addEventListener('click', e => { if (e.target === overlay) overlay.remove(); });

      btnRow.appendChild(pasteBtn);
      btnRow.appendChild(cancelBtn);
      box.appendChild(title);
      box.appendChild(hint);
      box.appendChild(ta);
      box.appendChild(btnRow);
      box.appendChild(errMsg);
      overlay.appendChild(box);
      document.body.appendChild(overlay);
      setTimeout(() => ta.focus(), 80);
    }

    function injectBtafButton() {
      if (document.getElementById('btaf-btn')) return;
      const btn = document.createElement('button');
      btn.id = 'btaf-btn';
      btn.textContent = '📋 Paste Ride Data';
      Object.assign(btn.style, {
        position: 'fixed', bottom: '20px', right: '20px', zIndex: '2147483646',
        background: '#f57c00', color: '#fff', border: 'none',
        padding: '12px 20px', borderRadius: '8px', fontSize: '15px',
        fontWeight: 'bold', cursor: 'pointer', boxShadow: '0 4px 14px rgba(0,0,0,.35)',
        fontFamily: 'system-ui,-apple-system,Segoe UI,Roboto,Arial',
      });
      btn.addEventListener('mouseenter', () => { btn.style.background = '#e65100'; });
      btn.addEventListener('mouseleave', () => { btn.style.background = '#f57c00'; });
      btn.addEventListener('click', async () => {
        btn.disabled    = true;
        btn.textContent = '⏳ Reading…';
        try {
          const raw = await navigator.clipboard.readText();
          if (!raw || !raw.trim()) {
            showBtafToast('⚠️ Clipboard is empty. Run your VBA macro first.', '#e65100');
          } else {
            let data;
            try { data = JSON.parse(raw.trim()); }
            catch { showBtafToast('⚠️ Clipboard is not valid JSON. Run your VBA macro first.', '#e65100'); data = null; }
            if (data) await btafSendData(data);
          }
        } catch {
          // Clipboard API blocked — fall back to paste modal
          showBtafModal();
        }
        btn.disabled    = false;
        btn.textContent = '📋 Paste Ride Data';
      });
      document.body.appendChild(btn);
    }

    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', injectBtafButton);
    else injectBtafButton();
  }

  // ── userCodeAppPanel: reach into userHtmlIFrame and fill the form ─────────────
  if (isUserCodeAppPanel) {

    const BTAF_FIELDS = [
      { key: 'date',            id: 'rideDate',      label: 'Ride Date'        },
      { key: 'pickupTime',      id: 'pickupTime',    label: 'Pickup Time'      },
      { key: 'appointmentTime', id: 'apptTime',      label: 'Appointment Time' },
      { key: 'pickupAddress',   id: 'pickupAddress', label: 'Pickup Address'   },
      { key: 'destination1',    id: 'destAddress1',  label: 'Destination 1'    },
      { key: 'destination2',    id: 'destAddress2',  label: 'Destination 2', optional: true },
    ];

    function findFormDoc() {
      if (document.getElementById('rideDate')) return document;
      for (const ifr of document.querySelectorAll('iframe')) {
        try {
          const d = ifr.contentDocument || ifr.contentWindow?.document;
          if (d && d.getElementById('rideDate')) return d;
        } catch {}
      }
      return null;
    }

    function waitForFormDoc(timeoutMs) {
      return new Promise(resolve => {
        const check = () => { const d = findFormDoc(); if (d) { resolve(d); return true; } return false; };
        if (check()) return;
        const obs  = new MutationObserver(() => { if (check()) { obs.disconnect(); clearInterval(poll); } });
        const poll = setInterval(() => { if (check()) { obs.disconnect(); clearInterval(poll); } }, 300);
        obs.observe(document.documentElement || document, { childList: true, subtree: true });
        setTimeout(() => { obs.disconnect(); clearInterval(poll); resolve(null); }, timeoutMs || 30000);
      });
    }

    async function btafSetField(el, doc, value) {
      const win = doc.defaultView || window;
      el.dispatchEvent(new win.MouseEvent('pointerdown',  { bubbles: true, cancelable: true }));
      el.dispatchEvent(new win.FocusEvent('focus',         { bubbles: true }));
      el.dispatchEvent(new win.KeyboardEvent('keydown',    { bubbles: true, cancelable: true, key: 'a' }));
      await btafSleep(80);
      el.readOnly = false;
      el.removeAttribute('readonly');
      const proto = el.tagName === 'TEXTAREA' ? win.HTMLTextAreaElement.prototype : win.HTMLInputElement.prototype;
      const desc  = Object.getOwnPropertyDescriptor(proto, 'value');
      if (desc && desc.set) desc.set.call(el, value);
      else el.value = value;
      el.dispatchEvent(new win.Event('input',  { bubbles: true }));
      el.dispatchEvent(new win.Event('change', { bubbles: true }));
      await btafSleep(80);
      el.dispatchEvent(new win.FocusEvent('blur', { bubbles: true }));
      await btafSleep(40);
      return el.value === value;
    }

    async function btafRunAutofill(data) {
      const formDoc = await waitForFormDoc(30000);
      if (!formDoc) return;

      await btafSleep(200);
      const filled = [], missed = [];

      for (const field of BTAF_FIELDS) {
        const val = data[field.key];
        if (!val && field.optional) continue;
        if (!val) { missed.push(field.label); continue; }
        const el = formDoc.getElementById(field.id);
        if (!el) { missed.push(field.label); continue; }
        const ok = await btafSetField(el, formDoc, val);
        if (ok || el.value.trim() !== '') filled.push(field.label);
        else missed.push(field.label);
        await btafSleep(150);
      }

      if (missed.length === 0) {
        await btafSleep(500);
        const btn = Array.from(formDoc.querySelectorAll('button, input[type="submit"], a'))
          .find(el => /check\s*availability/i.test(el.textContent || el.value || ''));
        if (btn) btn.click();
      }

      try {
        await GM_deleteValue(BTAF_GM_KEY);
        await GM_deleteValue(BTAF_GM_TS);
        await GM_deleteValue(BTAF_GM_VISIBILITY);
      } catch {}
    }

    const BTAF_FRAME_LOAD_TS = Date.now();
    let _btafTriggered = false;

    (async () => {
      while (!_btafTriggered) {
        try {
          const raw = await GM_getValue(BTAF_GM_KEY, null);
          if (raw) {
            const ts  = parseInt(await GM_getValue(BTAF_GM_TS, '0'), 10);
            const age = Date.now() - ts;

            if (ts > BTAF_FRAME_LOAD_TS && age < 300000) {
              // KEY CHECK: only act if this tab is currently visible.
              // document.hidden in an iframe mirrors the top-level tab visibility.
              // Active tab = document.hidden is false.
              // Background tabs = document.hidden is true → skip.
              if (document.hidden) {
                // This tab is in the background — don't act, keep polling
                // in case user switches to this tab and re-triggers
                await btafSleep(200);
                continue;
              }

              _btafTriggered = true;
              await GM_deleteValue(BTAF_GM_KEY);
              await GM_deleteValue(BTAF_GM_TS);
              await GM_deleteValue(BTAF_GM_VISIBILITY);
              btafRunAutofill(JSON.parse(raw));
              return;

            } else if (age >= 300000) {
              await GM_deleteValue(BTAF_GM_KEY);
              await GM_deleteValue(BTAF_GM_TS);
              await GM_deleteValue(BTAF_GM_VISIBILITY);
            }
          }
        } catch {}
        await btafSleep(200);
      }
    })();
  }

  /* =================================== DYNAMICS SECTION =================================== */
  if (isDynamics) {
    const statusText          = "Pending - RATE Authorization Requested";
    const itineraryChangeText = "Itinerary Change Requested";
    const headerSelector      = '[id^="formHeaderTitle_"]';
    const buttonSelector      = 'button[aria-label="Rate Approval Status"]';
    const itinerarySelector   = 'button[data-id="gtt_itinerarychange.fieldControl-option-set-select"]';

    function isInSearchUI(el) {
      if (!el) return false;
      const selectors = [
        '#GlobalSearchBox','.ms-SearchBox','[role="search"]','[role="searchbox"]',
        '[aria-label="Search"]','[aria-label*="search"]','[aria-label*="Search results"]',
        '.quickFind','.globalSearch','.searchResults','.ms-Panel','.ms-Callout','.ms-Layer',
        '[data-lp-id="globalQuickFind"]','[data-id="globalQuickFind"]','[id*="GlobalQuickFind"]','[id*="SearchBox"]'
      ];
      const hit = el.closest(selectors.join(','));
      if (hit) return true;
      let a = el;
      while (a) {
        if (a.id && /(GlobalQuickFind|quickFind|Search|SearchBox)/i.test(a.id)) return true;
        a = a.parentElement;
      }
      return false;
    }

    // Keep the form banners in the LEFT header section exactly between the
    // Save Status row ("- Saved" / "- Unsaved") and the breadcrumb row that
    // begins with "Referral".  Dynamics changed the header DOM, so anchoring to
    // the old formHeaderTitle parent can either put the banner inline with the
    // referral number or make the header flex container grow vertically.
    //
    // These two data-id values are much more stable and describe the exact
    // before/after position we want:
    //   [data-id="header_saveStatus"]   -> banner goes AFTER this row
    //   [data-id="entity_name_span"]    -> banner goes BEFORE this row

    function directChildUnder(ancestor, descendant) {
      if (!ancestor || !descendant) return null;
      let node = descendant;
      while (node && node.parentElement && node.parentElement !== ancestor) {
        node = node.parentElement;
      }
      return node && node.parentElement === ancestor ? node : null;
    }

    function findFormBannerInsertionPoint() {
      const saveStatus = document.querySelector('[data-id="header_saveStatus"]');
      const entityName = document.querySelector('[data-id="entity_name_span"]');
      if (!saveStatus || !entityName) return null;

      // Find the LOWEST common ancestor containing both the save-status line and
      // the Referral/Information breadcrumb line.  On the current Dynamics form
      // this is the left-side header column, not the whole page header.
      let common = saveStatus.parentElement;
      while (common && common !== document.body && !common.contains(entityName)) {
        common = common.parentElement;
      }
      if (!common || common === document.body) return null;

      const saveRow = directChildUnder(common, saveStatus);
      const referralRow = directChildUnder(common, entityName);
      if (!saveRow || !referralRow || saveRow === referralRow) return null;

      // Make sure the save-status row really comes before Referral in this build.
      const relationship = saveRow.compareDocumentPosition(referralRow);
      if (!(relationship & Node.DOCUMENT_POSITION_FOLLOWING)) return null;

      return { parent: common, before: referralRow };
    }

    function normalizeFormBannerStack(stack) {
      if (!stack) return;
      const banners = Array.from(stack.children);
      const compact = banners.length > 1;

      Object.assign(stack.style, {
        display: 'flex',
        flexDirection: 'column',
        width: '100%',
        maxWidth: '100%',
        boxSizing: 'border-box',
        margin: '2px 0 3px 0',
        padding: '0',
        gap: compact ? '2px' : '0',
        position: 'static',
        inset: 'auto',
        transform: 'none',
        zIndex: 'auto',
        pointerEvents: 'auto'
      });

      banners.forEach(banner => {
        Object.assign(banner.style, {
          display: 'block',
          width: '100%',
          maxWidth: '100%',
          boxSizing: 'border-box',
          margin: '0',
          minHeight: compact ? '19px' : '23px',
          padding: compact ? '2px 6px' : '3px 6px',
          fontSize: compact ? '11px' : '12px',
          lineHeight: compact ? '15px' : '17px'
        });
      });
    }

    function ensureFormBannerStack() {
      const point = findFormBannerInsertionPoint();
      if (!point) return null;

      let stack = document.getElementById('mtoy-form-banner-stack');
      if (!stack) {
        stack = document.createElement('div');
        stack.id = 'mtoy-form-banner-stack';
        stack.setAttribute('role', 'presentation');
      }

      // Always re-home it because Dynamics can rebuild the header when records,
      // tabs, save state, or form data change.
      if (stack.parentNode !== point.parent || stack.nextSibling !== point.before) {
        point.parent.insertBefore(stack, point.before);
      }

      normalizeFormBannerStack(stack);
      return stack;
    }

    function placeFormBanner(banner) {
      const stack = ensureFormBannerStack();
      if (!stack) return false;
      if (banner.parentNode !== stack) stack.appendChild(banner);
      normalizeFormBannerStack(stack);
      return true;
    }

    function cleanupFormBannerStack() {
      const stack = document.getElementById('mtoy-form-banner-stack');
      if (!stack) return;
      if (!stack.children.length) {
        stack.remove();
        return;
      }

      // Re-check the exact Saved/Unsaved -> banner -> Referral placement after
      // any Dynamics redraw.
      ensureFormBannerStack();
    }

    function insertBanner() {
      let banner = document.getElementById('rate-status-banner');
      if (!banner) {
        banner = document.createElement('div');
        banner.id = 'rate-status-banner';
        banner.textContent = "**PENDING RATES**";
        banner.style.backgroundColor = 'darkblue';
        banner.style.color = 'white';
        banner.style.padding = '5px';
        banner.style.fontWeight = 'bold';
        banner.style.textAlign = 'center';
        banner.style.borderRadius = '5px';
      }
      placeFormBanner(banner);
    }

    function removeBanner() {
      const existing = document.getElementById('rate-status-banner');
      if (existing) existing.remove();
      cleanupFormBannerStack();
    }

    function checkStatusAndInsertBanner() {
      const button = document.querySelector(buttonSelector);
      if (!button) return;
      if (button.textContent.includes(statusText)) insertBanner();
      else removeBanner();
    }

    function observeRateStatusChanges() {
      const button = document.querySelector(buttonSelector);
      if (!button) return;
      const observer = new MutationObserver(() => checkStatusAndInsertBanner());
      observer.observe(button, { childList: true, subtree: true, characterData: true });
    }

    function insertItineraryChangeBanner() {
      let banner = document.getElementById('itinerary-change-banner');
      if (!banner) {
        banner = document.createElement('div');
        banner.id = 'itinerary-change-banner';
        banner.textContent = 'STOP PENDING CHANGED DO NOT STAFF YET';
        banner.style.backgroundColor = '#d32f2f';
        banner.style.color = 'white';
        banner.style.padding = '5px';
        banner.style.fontWeight = 'bold';
        banner.style.textAlign = 'center';
        banner.style.borderRadius = '5px';
      }
      placeFormBanner(banner);
    }

    function removeItineraryChangeBanner() {
      const existing = document.getElementById('itinerary-change-banner');
      if (existing) existing.remove();
      cleanupFormBannerStack();
    }

    function checkItineraryChangeBanner() {
      const button = document.querySelector(itinerarySelector);
      if (!button) {
        removeItineraryChangeBanner();
        return;
      }

      const currentValue = (
        button.value ||
        button.getAttribute('value') ||
        button.textContent ||
        button.title ||
        ''
      ).trim();

      if (currentValue === itineraryChangeText) insertItineraryChangeBanner();
      else removeItineraryChangeBanner();
    }

    function highlightAllRowsGlobal() {
      const vipTerms = [
        "defense medical exam -  always vip !!","defense medical exam -  always vip!!",
        "dr.'s visit: 2nd opinion -  always vip!!","dr.'s visit: ime:  always vip !!",
        "pqme -   always vip!!","pqme - panel qualified medical examination - always vip!!",
        "qme - qualified medical exam) vip!!","ame - agreed medical evaluation",
        "evaluation","fce - long appt.!!!!!"
      ];
      const rows = document.querySelectorAll('div[role="row"]');
      const now = new Date();
      rows.forEach(row => {
        if (row.dataset.highlightedGlobal) return;
        const text = row.textContent?.toLowerCase() || '';
        if (text.includes("pending - rate authorization requested")) {
          row.style.backgroundColor = 'lightblue';
        } else if (text.includes("rates approved")) {
          row.style.backgroundColor = 'plum';
        } else if (vipTerms.some(term => text.includes(term))) {
          row.style.backgroundColor = 'gold';
        } else if (text.includes("first") || text.includes("surgery")) {
          row.style.backgroundColor = 'lightgreen';
        } else if (text.includes("airport pickup/dropoff") || text.includes("airport dropoff/pickup")) {
          row.style.backgroundColor = 'lightcoral';
        }
        const timeCols = ["gtt_localizedpickuptime", "gtt_localizedappttime"];
        const minutesDelay = 5;
        const threshold = new Date(now.getTime() - minutesDelay * 60000);
        for (const col of timeCols) {
          const timeCell = row.querySelector(`div[col-id="${col}"]`);
          if (timeCell) {
            const label = timeCell.querySelector('label[aria-label]');
            if (label) {
              const dateStr = label.getAttribute('aria-label');
              const dateVal = new Date(dateStr);
              if (!isNaN(dateVal.getTime()) && dateVal < threshold) {
                row.style.backgroundColor = '#FFDAB9';
                break;
              }
            }
          }
        }
        row.dataset.highlightedGlobal = "true";
      });
    }

    function isVisibleElement(el) {
      if (!el) return false;
      const cs = getComputedStyle(el);
      if (cs.display === 'none' || cs.visibility === 'hidden' || cs.opacity === '0') return false;
      const r = el.getBoundingClientRect();
      return r.width > 0 && r.height > 0;
    }

    function isNotificationPanelOpen() {
      const candidates = [
        ...document.querySelectorAll('[id*="notificationWrapper"]'),
        ...document.querySelectorAll('[id*="message"]'),
        ...document.querySelectorAll('[aria-label*="notification" i]'),
        ...document.querySelectorAll('[title*="notification" i]'),
        ...document.querySelectorAll('[role="dialog"]'),
        ...document.querySelectorAll('[aria-modal="true"]'),
        ...document.querySelectorAll('.ms-Panel, .ms-Layer')
      ];
      for (const el of candidates) {
        if (!isVisibleElement(el)) continue;
        const txt = (el.innerText || el.textContent || '').toLowerCase();
        if ((txt.includes('you have') && txt.includes('notification')) ||
            txt.includes('select to view') || txt.includes('claimant:') || txt.includes('payer:')) {
          return true;
        }
      }
      return false;
    }

    function isSupportedLegendPage() {
      const hasDashboardBar = !!document.querySelector('[data-lp-id="commandbar-Dashboard:null"]');
      const hasGridBar = !!document.querySelector('[data-id="commandBar_0"]') ||
                         !!document.querySelector('[data-lp-id^="commandbar-HomePageGrid:"]');
      const hasFormHeader = !!document.querySelector('[id^="formHeaderTitle_"]') ||
                            !!document.querySelector('[data-lp-id="form-header-title"]') ||
                            !!document.querySelector('[data-id="form-header-title"]');
      return (hasDashboardBar || hasGridBar) && !hasFormHeader;
    }

    function createLegendChip(text, bg) {
      const chip = document.createElement('span');
      chip.textContent = text;
      chip.style.display = 'inline-block';
      chip.style.backgroundColor = bg;
      chip.style.color = '#111';
      chip.style.padding = '4px 10px';
      chip.style.borderRadius = '999px';
      chip.style.fontSize = '13px';
      chip.style.lineHeight = '1.2';
      chip.style.whiteSpace = 'nowrap';
      chip.style.boxShadow = '0 1px 4px rgba(0,0,0,.08)';
      return chip;
    }

    function createLegendElement(type = "default") {
      const legend = document.createElement("div");
      legend.dataset.legend = "true";
      legend.dataset.legendType = type;
      legend.style.display = 'flex';
      legend.style.flexWrap = 'wrap';
      legend.style.justifyContent = 'left';
      legend.style.alignItems = 'left';
      legend.style.gap = '8px';
      legend.style.width = '100%';
      legend.style.margin = '0 auto';
      legend.style.padding = '2px 0';
      legend.style.boxSizing = 'border-box';
      legend.style.textAlign = 'left';
      legend.appendChild(createLegendChip('Evaluations', 'gold'));
      legend.appendChild(createLegendChip('First Time/Surgery', 'lightgreen'));
      legend.appendChild(createLegendChip('Airport', 'lightcoral'));
      if (type === "default") {
        legend.appendChild(createLegendChip('Pending Rate Approval', 'lightblue'));
        legend.appendChild(createLegendChip('Rates Approved', 'plum'));
      }
      if (type === "-confirm") {
        legend.appendChild(createLegendChip('Pickup/Appt Time Passed', '#FFDAB9'));
      }
      return legend;
    }

    function removeExistingLegendArtifacts() {
      document.querySelectorAll(
        '#mtoy-legend-section, #mtoy-legend-bar, [data-legend="true"], [data-legend-wrapper="true"], [data-legend-host="true"]'
      ).forEach(el => {
        const inBadArea = isInSearchUI(el) ||
          !!el.closest('.ms-Panel, .ms-Layer, [role="dialog"], [aria-modal="true"]') ||
          !!el.closest('[id*="notificationWrapper"]');
        if (inBadArea) el.remove();
      });
    }

    function findCommandBarElement() {
      return document.querySelector('[data-lp-id="commandbar-Dashboard:null"]')
          || document.querySelector('[data-id="commandBar_0"]')
          || document.querySelector('[data-lp-id^="commandbar-HomePageGrid:"]')
          || document.querySelector('ul[data-id="CommandBar"]');
    }

    function getCommandBarShell() {
      const bar = findCommandBarElement();
      if (!bar) return null;
      let node = bar;
      while (node && node !== document.body) {
        const hasCommandBar = !!node.querySelector?.('[data-id="CommandBar"], [data-id="commandBar_0"], [data-lp-id="commandbar-Dashboard:null"], [data-lp-id^="commandbar-HomePageGrid:"]');
        const hasShare = !!node.querySelector?.('#collaborationShareButton_0, button[aria-label="Share"]');
        if (hasCommandBar && hasShare) return node;
        node = node.parentElement;
      }
      node = bar;
      while (node && node !== document.body) {
        const parent = node.parentElement;
        if (!parent) break;
        const children = Array.from(parent.children || []).filter(el => el.nodeType === 1);
        const hasCommandBar = !!parent.querySelector?.('[data-id="CommandBar"], [data-id="commandBar_0"], [data-lp-id="commandbar-Dashboard:null"], [data-lp-id^="commandbar-HomePageGrid:"]');
        if (hasCommandBar && children.length >= 2) return parent;
        node = parent;
      }
      return bar.parentElement || bar;
    }

    function ensureLegendSection() {
      const shell = getCommandBarShell();
      if (!shell || !shell.parentNode || isInSearchUI(shell)) return null;
      let section = document.getElementById('mtoy-legend-section');
      if (!section) {
        section = document.createElement('div');
        section.id = 'mtoy-legend-section';
        section.dataset.legendSection = 'true';
        section.setAttribute('role', 'presentation');
        section.style.display = 'block';
        section.style.width = 'calc(100% - 16px)';
        section.style.margin = '6px 8px 10px 8px';
        section.style.padding = '10px 14px';
        section.style.boxSizing = 'border-box';
        section.style.background = '#fff';
        section.style.border = '1px solid rgba(0,0,0,.08)';
        section.style.borderRadius = '10px';
        section.style.boxShadow = '0 1px 2px rgba(0,0,0,.04)';
        section.style.clear = 'both';
      }
      if (section.previousElementSibling !== shell || section.parentNode !== shell.parentNode) {
        if (section.parentNode) section.remove();
        shell.parentNode.insertBefore(section, shell.nextSibling);
      }
      return section;
    }

    function getLegendBar() {
      const section = ensureLegendSection();
      if (!section) return null;
      let bar = document.getElementById('mtoy-legend-bar');
      if (!bar) {
        bar = document.createElement('div');
        bar.id = 'mtoy-legend-bar';
        bar.dataset.legendBar = 'true';
        bar.style.display = 'flex';
        bar.style.flexWrap = 'wrap';
        bar.style.justifyContent = 'left';
        bar.style.alignItems = 'left';
        bar.style.gap = '8px';
        bar.style.width = '100%';
        bar.style.boxSizing = 'border-box';
        bar.style.margin = '0';
        bar.style.padding = '0';
        section.appendChild(bar);
      } else if (bar.parentNode !== section) {
        bar.remove();
        section.appendChild(bar);
      }
      return bar;
    }

    function addLegend() {
      removeExistingLegendArtifacts();
      const existingSection = document.getElementById('mtoy-legend-section');
      if (isNotificationPanelOpen()) { if (existingSection) existingSection.remove(); return; }
      if (!isSupportedLegendPage()) { if (existingSection) existingSection.remove(); return; }
      const type = detectLegendType();
      if (!type) { if (existingSection) existingSection.remove(); return; }
      const bar = getLegendBar();
      if (!bar) return;
      bar.innerHTML = '';
      bar.appendChild(createLegendElement(type));
    }

    function detectLegendType() {
      const fullText = ((document.title || '') + '\n' + (document.body?.innerText || '')).toLowerCase();
      if (fullText.includes('same day confirmations') || fullText.includes('same day (oncall)') || fullText.includes('-confirm')) return '-confirm';
      if (fullText.includes('unassigned transportation') || fullText.includes('unassigned transport') ||
          fullText.includes('prev vendor search') || fullText.includes('uber') || fullText.includes('~transport')) return 'default';
      return null;
    }

    function adjustSpacing() {
      const headerTitle = document.querySelector('[data-lp-id="form-header-title"] h1');
      if (headerTitle) {
        headerTitle.style.marginTop = "0px";
        headerTitle.style.marginBottom = "0px";
        headerTitle.style.padding = "0px";
      }
    }

    function styleNotificationWrapper() {
      document.querySelectorAll('[id*="notificationWrapper"], [id*="message"]').forEach(element => {
        const text = (element.textContent || element.innerText || '').trim();
        element.style.fontWeight = 'bold';
        element.style.fontSize = '18px';
        element.style.backgroundColor = text.includes('~~') ? 'yellow' : 'lightgreen';
      });
    }

    function observeNotifications() {
      const observer = new MutationObserver(styleNotificationWrapper);
      observer.observe(document.body, { childList: true, subtree: true });
    }

    function insertJbaBannerIfNeeded() {
      const titleText = document.title;
      const providerDiv = Array.from(document.querySelectorAll('div[role="presentation"]'))
        .find(el => el.textContent.trim() === "Provider Assignment");
      const header = document.querySelector(headerSelector);
      if (!titleText.includes("7327-") || !providerDiv || document.getElementById("jba-banner")) return;
      const banner = document.createElement("div");
      banner.id = "jba-banner";
      banner.textContent = "JBA file usually needs QUOTE please provide full breakdown and totals";
      banner.style.backgroundColor = "#f8d7da";
      banner.style.color = "#721c24";
      banner.style.padding = "6px";
      banner.style.fontWeight = "bold";
      banner.style.textAlign = "center";
      banner.style.borderRadius = "5px";
      if (header) placeFormBanner(banner);
    }

    function insertVipBannerIfNeeded() {
      const titleText = document.title;
      const header = document.querySelector(headerSelector);
      const vipIds = [
        "4474-65549","4474-48338","4474-48380","202-46904","202-50715",
        "4474-64737","10837-61025","4474-66551","4474-63533","10530-68938"
      ];
      if (!vipIds.some(id => titleText.includes(id))) return;
      if (document.getElementById("vip-banner")) return;
      const banner = document.createElement("div");
      banner.id = "vip-banner";
      banner.textContent = "VIP file please read all Alerts and staff as early as possible";
      banner.style.backgroundColor = "#f8d7da";
      banner.style.color = "#721c24";
      banner.style.padding = "6px";
      banner.style.fontWeight = "bold";
      banner.style.textAlign = "center";
      banner.style.borderRadius = "5px";
      if (header) placeFormBanner(banner);
    }

    function isInSearchUIWrapper() {
      document.querySelectorAll('[data-legend="true"], [data-legend-wrapper="true"]').forEach(el => {
        if (isInSearchUI(el)) el.remove();
      });
    }

    function waitForMoniqueInIframe(retries = 20, delay = 1000) {
      const iframe = document.querySelector('#WebResource_RecipientSelector');
      if (!iframe) { if (retries > 0) setTimeout(() => waitForMoniqueInIframe(retries - 1, delay), delay); return; }
      const doc = iframe.contentDocument || iframe.contentWindow.document;
      if (!doc || !doc.body) { if (retries > 0) setTimeout(() => waitForMoniqueInIframe(retries - 1, delay), delay); return; }
      const td = [...doc.querySelectorAll("td")].find(td => td.textContent.trim().includes("Monique Jones"));
      if (td && !td.querySelector(".monique-message")) {
        td.style.fontSize = "12px";
        td.style.fontWeight = "bold";
        const note = document.createElement("div");
        note.className = "monique-message";
        note.textContent = "Please combine staffing and/or auth requests into one email (include multiple dates into one email).";
        note.style.marginTop = "5px";
        note.style.marginBottom = "15px";
        note.style.color = "darkred";
        note.style.fontWeight = "bold";
        td.appendChild(note);
      } else if (retries > 0) {
        setTimeout(() => waitForMoniqueInIframe(retries - 3, delay), delay);
      }
    }

    function waitForAUTHEMAIL(retries = 20, delay = 1000) {
      const iframe = document.querySelector('#WebResource_RecipientSelector');
      if (!iframe) { if (retries > 0) setTimeout(() => waitForAUTHEMAIL(retries - 1, delay), delay); return; }
      const doc = iframe.contentDocument || iframe.contentWindow.document;
      if (!doc || !doc.body) { if (retries > 0) setTimeout(() => waitForAUTHEMAIL(retries - 1, delay), delay); return; }
      const td = [...doc.querySelectorAll("td")].find(td => td.textContent.trim().includes("AUTH EMAIL"));
      if (td && !td.querySelector(".auth-message")) {
        td.style.fontSize = "12px";
        td.style.fontWeight = "bold";
        const note = document.createElement("div");
        note.className = "auth-message";
        note.textContent = "DO NOT SEND staffing or rate request here.";
        note.style.marginTop = "5px";
        note.style.marginBottom = "15px";
        note.style.color = "darkred";
        note.style.fontWeight = "bold";
        td.appendChild(note);
      } else if (retries > 0) {
        setTimeout(() => waitForAUTHEMAIL(retries - 3, delay), delay);
      }
    }

    /* ================= CAREWORKS JURISDICTION WARNING ================= */
    const CAREWORKS_JURISDICTION_API =
      'https://lowmargin.mtoysystems.com/api/get_email_list.php?list=CareWorks_Jurisdiction';

    let careWorksCurrentReferralKey = '';
    let careWorksDismissedThisVisit = false;
    let careWorksCheckInFlight = false;
    let careWorksLastCheckSignature = '';

    function careWorksNormalizeName(value) {
      return (value || '')
        .toLowerCase()
        .replace(/\([^)]*\)/g, ' ')
        .replace(/[^a-z0-9]+/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();
    }

    function careWorksNameVariants(value) {
      const normalized = careWorksNormalizeName(value);
      if (!normalized) return [];

      const variants = new Set([normalized]);
      const parts = normalized.split(' ').filter(Boolean);
      if (parts.length >= 2) {
        variants.add(`${parts[parts.length - 1]} ${parts.slice(0, -1).join(' ')}`);
      }
      return [...variants];
    }

    function careWorksGetClaimantName() {
      const claimant = document.querySelector(
        'div[data-id="gtt_claimantid.fieldControl-LookupResultsDropdown_gtt_claimantid_selected_tag_text"]'
      );

      return (
        claimant?.getAttribute('title') ||
        claimant?.textContent ||
        ''
      ).trim();
    }

    function careWorksGetPayerText() {
      const payer =
        document.querySelector('[data-id*="gtt_payerid"][data-id*="selected_tag_text"]') ||
        document.querySelector('[data-id*="gtt_payerid"][data-id*="selected_tag"]') ||
        document.querySelector('[data-id*="gtt_payerid"] input') ||
        document.querySelector('[aria-label="Payer"]');

      return (
        payer?.textContent ||
        payer?.value ||
        payer?.getAttribute?.('title') ||
        payer?.getAttribute?.('aria-label') ||
        ''
      ).trim();
    }

    function careWorksGetReferralKey() {
      const claimant = careWorksGetClaimantName();
      if (!claimant) return '';

      try {
        const url = new URL(location.href);
        const id = url.searchParams.get('id');
        if (id) return `id:${id.toLowerCase()}`;
      } catch {}

      const header = document.querySelector('[id^="formHeaderTitle"]');
      const headerText = (header?.textContent || '').trim();
      return headerText ? `header:${headerText}` : `claimant:${careWorksNormalizeName(claimant)}`;
    }

    function careWorksCollectNames(payload) {
      const names = [];
      const add = value => {
        if (typeof value !== 'string') return;
        const trimmed = value.trim();
        if (!trimmed) return;

        // Support plain names, "Name <email>", and email-list values.
        const displayName = trimmed.match(/^\s*([^<]+?)\s*<[^>]+>\s*$/)?.[1]?.trim();
        if (displayName) names.push(displayName);
        names.push(trimmed);

        if (trimmed.includes('@')) {
          const localPart = trimmed.split('@')[0].replace(/[._-]+/g, ' ').trim();
          if (localPart) names.push(localPart);
        }
      };

      const walk = value => {
        if (Array.isArray(value)) {
          value.forEach(walk);
        } else if (value && typeof value === 'object') {
          ['name', 'full_name', 'display_name', 'claimant', 'value', 'email'].forEach(key => add(value[key]));
        } else {
          add(value);
        }
      };

      ['names', 'members', 'items', 'entries', 'to', 'cc', 'bcc'].forEach(key => walk(payload?.[key]));
      return names;
    }

    function careWorksRequestList() {
      return new Promise((resolve, reject) => {
        if (typeof GM_xmlhttpRequest === 'function') {
          GM_xmlhttpRequest({
            method: 'GET',
            url: `${CAREWORKS_JURISDICTION_API}&_=${Date.now()}`,
            headers: { Accept: 'application/json' },
            timeout: 15000,
            onload: response => {
              try {
                if (response.status < 200 || response.status >= 300) {
                  reject(new Error(`API returned HTTP ${response.status}`));
                  return;
                }
                resolve(JSON.parse(response.responseText));
              } catch (error) {
                reject(error);
              }
            },
            onerror: () => reject(new Error('CareWorks API request failed')),
            ontimeout: () => reject(new Error('CareWorks API request timed out'))
          });
          return;
        }

        fetch(`${CAREWORKS_JURISDICTION_API}&_=${Date.now()}`, { cache: 'no-store' })
          .then(response => {
            if (!response.ok) throw new Error(`API returned HTTP ${response.status}`);
            return response.json();
          })
          .then(resolve, reject);
      });
    }

    function showCareWorksJurisdictionPopup() {
      if (document.getElementById('mtoy-careworks-jurisdiction-popup')) return;

      const backdrop = document.createElement('div');
      backdrop.id = 'mtoy-careworks-jurisdiction-popup';
      backdrop.style.cssText = `
        position:fixed; inset:0; z-index:2147483647;
        display:flex; align-items:center; justify-content:center;
        background:rgba(0,0,0,.48); padding:20px; box-sizing:border-box;
      `;

      const box = document.createElement('div');
      box.setAttribute('role', 'alertdialog');
      box.setAttribute('aria-modal', 'true');
      box.style.cssText = `
        width:min(560px, 92vw); background:#fff; color:#111;
        border:3px solid #b91c1c; border-radius:12px; padding:24px;
        box-shadow:0 18px 55px rgba(0,0,0,.4); text-align:center;
        font-family:system-ui,-apple-system,Segoe UI,Roboto,Arial,sans-serif;
      `;

      const message = document.createElement('div');
      message.textContent = 'Out of Jurisdiction claim - Please confirm rates on PO are the correct rates. If not please email coordinator for updated rates';
      message.style.cssText = 'font-size:22px;font-weight:700;line-height:1.35;margin-bottom:22px;';

      const ok = document.createElement('button');
      ok.type = 'button';
      ok.textContent = 'OK';
      ok.style.cssText = `
        min-width:110px; padding:10px 24px; border:0; border-radius:8px;
        background:#b91c1c; color:#fff; font-size:16px; font-weight:700; cursor:pointer;
      `;
      ok.addEventListener('click', () => {
        careWorksDismissedThisVisit = true;
        backdrop.remove();
      });

      box.append(message, ok);
      backdrop.appendChild(box);
      document.body.appendChild(backdrop);
      setTimeout(() => ok.focus(), 0);
    }

    async function checkCareWorksJurisdiction() {
      const referralKey = careWorksGetReferralKey();

      // Leaving the referral resets the warning, including for a later return to the same referral.
      if (!referralKey) {
        careWorksCurrentReferralKey = '';
        careWorksDismissedThisVisit = false;
        careWorksLastCheckSignature = '';
        return;
      }

      if (referralKey !== careWorksCurrentReferralKey) {
        careWorksCurrentReferralKey = referralKey;
        careWorksDismissedThisVisit = false;
        careWorksLastCheckSignature = '';
        document.getElementById('mtoy-careworks-jurisdiction-popup')?.remove();
      }

      if (careWorksDismissedThisVisit || careWorksCheckInFlight) return;

      const payer = careWorksGetPayerText();
      const claimant = careWorksGetClaimantName();
      if (!payer.toLowerCase().includes('careworks') || !claimant) return;

      const signature = `${referralKey}|${careWorksNormalizeName(payer)}|${careWorksNormalizeName(claimant)}`;
      if (signature === careWorksLastCheckSignature) return;
      careWorksLastCheckSignature = signature;
      careWorksCheckInFlight = true;

      try {
        const payload = await careWorksRequestList();
        if (!payload || payload.ok === false) throw new Error(payload?.error || 'CareWorks API returned an invalid response');

        const claimantVariants = new Set(careWorksNameVariants(claimant));
        const isMatch = careWorksCollectNames(payload).some(name =>
          careWorksNameVariants(name).some(variant => claimantVariants.has(variant))
        );

        if (isMatch && referralKey === careWorksCurrentReferralKey && !careWorksDismissedThisVisit) {
          showCareWorksJurisdictionPopup();
        }
      } catch (error) {
        console.warn('CareWorks jurisdiction check failed:', error);
        // Allow a later retry if the API was temporarily unavailable.
        careWorksLastCheckSignature = '';
      } finally {
        careWorksCheckInFlight = false;
      }
    }

    /* ================= ADDITIONAL SERVICES / ANCILLARY FEE WARNING ================= */
    let ancillaryCurrentReferralKey = '';
    let ancillaryCheckCompletedThisVisit = false;
    let ancillaryControlsFirstSeenAt = 0;
    let ancillaryLastSignature = '';
    let ancillaryStageControl = null;

    function getAncillaryReferralKey() {
      // Only treat an actual Referral form as eligible for this warning.
      const entityName = (
        document.querySelector('[data-id="entity_name_span"]')?.textContent || ''
      ).trim();
      const header = document.querySelector('[id^="formHeaderTitle"]');
      const headerText = (header?.textContent || '').trim();
      const looksLikeReferral = /referral/i.test(entityName) || /\b\d+-\d+-\d+\b/.test(headerText);
      if (!looksLikeReferral) return '';

      try {
        const url = new URL(location.href);
        const id = url.searchParams.get('id');
        if (id) return `id:${id.toLowerCase()}`;
      } catch {}

      return headerText ? `header:${headerText.toLowerCase()}` : '';
    }

    function findAdditionalServicesControl() {
      const exactRoot = document.querySelector('#dataSetRoot_Subgrid_2');
      if (exactRoot && /additional services/i.test(exactRoot.textContent || '')) return exactRoot;

      const roots = document.querySelectorAll('[id^="dataSetRoot_"], [data-id^="dataSetRoot_"]');
      return [...roots].find(root => {
        const heading = root.querySelector('h1, h2, h3, h4, [role="heading"]');
        return /^additional services$/i.test((heading?.textContent || '').trim());
      }) || null;
    }

    function findElementAcrossOpenShadowRoots(selector) {
      const roots = [document];
      const seenRoots = new Set();

      while (roots.length) {
        const root = roots.shift();
        if (!root || seenRoots.has(root)) continue;
        seenRoots.add(root);

        try {
          const match = root.querySelector(selector);
          if (match) return match;

          for (const element of root.querySelectorAll('*')) {
            if (element.shadowRoot && !seenRoots.has(element.shadowRoot)) {
              roots.push(element.shadowRoot);
            }
          }
        } catch {}
      }

      return null;
    }

    function getDynamicsHeaderStageText() {
      const pageWindows = [];

      try {
        if (typeof unsafeWindow !== 'undefined') {
          pageWindows.push(unsafeWindow);
          if (unsafeWindow.top && unsafeWindow.top !== unsafeWindow) pageWindows.push(unsafeWindow.top);
        }
      } catch {}

      pageWindows.push(window);
      try {
        if (window.top && window.top !== window) pageWindows.push(window.top);
      } catch {}

      for (const pageWindow of pageWindows) {
        try {
          const page = pageWindow?.Xrm?.Page;
          if (!page) continue;

          const readAttributeText = attribute => {
            if (!attribute) return '';
            return String(
              attribute.getText?.() ||
              attribute.getSelectedOption?.()?.text ||
              ''
            ).trim();
          };

          let text = readAttributeText(page.getAttribute?.('gtt_stage'));
          if (text) return text;

          text = readAttributeText(page.getControl?.('header_gtt_stage')?.getAttribute?.());
          if (text) return text;

          const controls = page.ui?.controls?.get?.() || [];
          for (const control of controls) {
            const name = String(control?.getName?.() || '');
            if (!name.startsWith('header_gtt_stage')) continue;
            text = readAttributeText(control.getAttribute?.());
            if (text) return text;
          }
        } catch {}
      }

      return '';
    }

    function isProviderAssignmentHeaderStage() {
      // Prefer the actual Dynamics field value. This avoids depending on how
      // the header web components and their shadow roots are rendered.
      const dynamicsStage = getDynamicsHeaderStageText();
      if (dynamicsStage) return dynamicsStage.toLowerCase() === 'provider assignment';

      // DOM fallback for cases where Xrm.Page has not initialized yet.
      if (!ancillaryStageControl?.isConnected) {
        ancillaryStageControl = findElementAcrossOpenShadowRoots(
          'uci-header-control-list-item[data-name="header_gtt_stage"]'
        );
      }
      if (!ancillaryStageControl) return false;

      // The value and label live inside this component's own open shadow root.
      const root = ancillaryStageControl.shadowRoot || ancillaryStageControl;
      const value = root.querySelector(
        'div.value-text[slot="value"][data-id="0"], div.value-text[slot="value"]'
      );
      const label = root.querySelector('div.label[slot="label"]');

      return (label?.textContent || '').trim().toLowerCase() === 'stage' &&
        (value?.textContent || '').trim().toLowerCase() === 'provider assignment';
    }

    function showAncillaryFeeWarning() {
      if (document.getElementById('mtoy-ancillary-fee-warning')) return;

      const backdrop = document.createElement('div');
      backdrop.id = 'mtoy-ancillary-fee-warning';
      backdrop.style.cssText = `
        position:fixed; inset:0; z-index:2147483647;
        display:flex; align-items:center; justify-content:center;
        background:rgba(0,0,0,.48); padding:20px; box-sizing:border-box;
      `;

      const box = document.createElement('div');
      box.setAttribute('role', 'alertdialog');
      box.setAttribute('aria-modal', 'true');
      box.style.cssText = `
        width:min(540px, 92vw); background:#fff; color:#111827;
        border:3px solid #d97706; border-radius:12px; padding:24px;
        box-shadow:0 18px 55px rgba(0,0,0,.4); text-align:center;
        font-family:system-ui,-apple-system,Segoe UI,Roboto,Arial,sans-serif;
      `;

      const message = document.createElement('div');
      message.textContent = 'Ancillary fees found please check PO to ensure items are listed';
      message.style.cssText = 'font-size:22px;font-weight:700;line-height:1.35;margin-bottom:22px;';

      const ok = document.createElement('button');
      ok.type = 'button';
      ok.textContent = 'OK';
      ok.style.cssText = `
        min-width:110px; padding:10px 24px; border:0; border-radius:8px;
        background:#d97706; color:#fff; font-size:16px; font-weight:700; cursor:pointer;
      `;
      ok.addEventListener('click', () => backdrop.remove());

      box.append(message, ok);
      backdrop.appendChild(box);
      document.body.appendChild(backdrop);
      setTimeout(() => ok.focus(), 0);
    }

    function checkAncillaryFeesOnReferralLoad() {
      const referralKey = getAncillaryReferralKey();

      if (!referralKey) {
        ancillaryCurrentReferralKey = '';
        ancillaryCheckCompletedThisVisit = false;
        ancillaryControlsFirstSeenAt = 0;
        ancillaryLastSignature = '';
        document.getElementById('mtoy-ancillary-fee-warning')?.remove();
        return;
      }

      if (referralKey !== ancillaryCurrentReferralKey) {
        ancillaryCurrentReferralKey = referralKey;
        ancillaryCheckCompletedThisVisit = false;
        ancillaryControlsFirstSeenAt = 0;
        ancillaryLastSignature = '';
        document.getElementById('mtoy-ancillary-fee-warning')?.remove();
      }

      if (ancillaryCheckCompletedThisVisit) return;

      // Only run when the top Stage field says Provider Assignment. In later
      // workflow stages this same field changes to values such as Active.
      if (!isProviderAssignmentHeaderStage()) return;

      // This reminder applies only to CareWorks referrals. Keep waiting while
      // the payer lookup is still rendering so the initial load is not missed.
      const payer = careWorksGetPayerText();
      if (!payer.toLowerCase().includes('careworks')) return;

      const root = findAdditionalServicesControl();
      if (!root) return;

      const loadFeeId = '0e98ce8f-7597-ed11-aad1-000d3a3412c9';
      const toggles = [...root.querySelectorAll('input.nncb-control[type="checkbox"], input[type="checkbox"]')]
        .filter(toggle => {
          const identity = `${toggle.id || ''} ${toggle.value || ''}`.toLowerCase();
          if (identity.includes(loadFeeId)) return false;

          // Label fallback in case Dynamics changes the Load Fee record GUID.
          const row = toggle.closest('div[style*="flex"]') || toggle.parentElement?.parentElement;
          const label = row?.querySelector('.nncb-switch-label');
          return (label?.textContent || '').trim().toLowerCase() !== 'load fee';
        });
      if (!toggles.length) return;

      const signature = toggles
        .map(toggle => `${toggle.id || toggle.value || 'toggle'}:${toggle.checked ? 1 : 0}`)
        .join('|');

      // Dynamics builds this custom control in stages. Wait until its values
      // have remained unchanged briefly so an initially unchecked render does
      // not hide a saved selection that appears a moment later.
      if (signature !== ancillaryLastSignature) {
        ancillaryLastSignature = signature;
        ancillaryControlsFirstSeenAt = Date.now();
        return;
      }

      if (!ancillaryControlsFirstSeenAt || Date.now() - ancillaryControlsFirstSeenAt < 1000) return;

      ancillaryCheckCompletedThisVisit = true;
      if (toggles.some(toggle => toggle.checked || toggle.getAttribute('aria-checked') === 'true')) {
        showAncillaryFeeWarning();
      }
    }

    // ─── Entry point: wait for Dynamics to render a reliable landmark before injecting.
    // A fixed delay isn't reliable — Dynamics can take anywhere from 1s to 10s+ depending
    // on load. We watch for #searchBoxLiveRegion (the nav search bar) which is one of the
    // first stable elements Dynamics renders. Only then do we wire up observers and run the
    // first pass of all enhancer functions.
    // DYNAMICS_INIT_DELAY = minimum wait (ms) before we start checking — gives Dynamics a
    //   head-start. Raise if things still inject too early on slow connections.
    // DYNAMICS_WAIT_CAP = hard give-up timeout (ms) — we run anyway after this.
    const DYNAMICS_INIT_DELAY = 4000; // ← adjust minimum wait here if needed (ms)
    const DYNAMICS_WAIT_CAP   = 30000; // ← hard give-up timeout (ms) — rarely needs changing

    function startDynamicsEnhancer() {
    let debounceTimer;
    const globalObserver = new MutationObserver(() => {
      clearTimeout(debounceTimer);
      debounceTimer = setTimeout(() => {
        checkStatusAndInsertBanner();
        checkItineraryChangeBanner();
        highlightAllRowsGlobal();
        addLegend();
        adjustSpacing();
        styleNotificationWrapper();
        observeRateStatusChanges();
        insertJbaBannerIfNeeded();
        insertVipBannerIfNeeded();
        checkCareWorksJurisdiction();
        checkAncillaryFeesOnReferralLoad();
        isInSearchUIWrapper();
      }, 200);
    });
    globalObserver.observe(document.body, { childList: true, subtree: true });

    let attempts = 0;
    const tryInit = setInterval(() => {
      checkStatusAndInsertBanner();
      checkItineraryChangeBanner();
      highlightAllRowsGlobal();
      addLegend();
      adjustSpacing();
      styleNotificationWrapper();
      observeRateStatusChanges();
      insertJbaBannerIfNeeded();
      insertVipBannerIfNeeded();
      checkAncillaryFeesOnReferralLoad();
      isInSearchUIWrapper();
      attempts++;
      if (attempts > 20) clearInterval(tryInit);
    }, 500);

    observeNotifications();
    document.addEventListener('click', (e) => {
      if (e.target?.closest?.(itinerarySelector) || e.target?.closest?.('[role="option"]')) {
        setTimeout(checkItineraryChangeBanner, 50);
        setTimeout(checkItineraryChangeBanner, 250);
      }
    }, true);
    setInterval(checkCareWorksJurisdiction, 1200);
    setInterval(checkAncillaryFeesOnReferralLoad, 500);

    if (document.title.includes('Email:')) {
      waitForMoniqueInIframe();
      waitForAUTHEMAIL();
    }

    const titleObserver = new MutationObserver(() => {
      if (document.title.includes("Email:")) {
        waitForMoniqueInIframe();
        waitForAUTHEMAIL();
      }
    });
    const titleNode = document.querySelector("title");
    if (titleNode) titleObserver.observe(titleNode, { childList: true });
    }

    function waitForDynamicsLandmark() {
      if (document.querySelector('#searchBoxLiveRegion')) { startDynamicsEnhancer(); return; }
      let done = false;
      const capTimer = setTimeout(() => { if (!done) { done = true; obs.disconnect(); startDynamicsEnhancer(); } }, DYNAMICS_WAIT_CAP - DYNAMICS_INIT_DELAY);
      const obs = new MutationObserver(() => {
        if (document.querySelector('#searchBoxLiveRegion')) {
          if (!done) { done = true; clearTimeout(capTimer); obs.disconnect(); startDynamicsEnhancer(); }
        }
      });
      obs.observe(document.documentElement, { childList: true, subtree: true });
    }

    setTimeout(waitForDynamicsLandmark, DYNAMICS_INIT_DELAY);
  }

})();


/* =====================================================================================
   PART D — UBER HEALTH: Paste Ride Data / Excel auto-start (from version 1.3)
   Isolated from the existing UIEnhancer features and limited to Uber Health.
   ===================================================================================== */
(function () {
    'use strict';

    if (location.hostname !== 'health.uber.com') return;

    function startUberHealth() {
    'use strict';

    const BUTTON_ID = 'gotandt-uber-paste-button';

    // ============================================================
// EXCEL AUTO-START
// Open Today's Activity -> Create New -> Single Ride -> Paste
// ============================================================

async function autoStartRideFromExcel() {

    const params =
        new URLSearchParams(
            window.location.search
        );

    if (
        params.get('gotandtAuto') !== '1'
    ) {
        return;
    }

    console.log(
        'Uber Health: Excel auto-start detected.'
    );

    /*
     * Remove the flag immediately so Uber SPA navigation or
     * a refresh does not cause the ride setup to run twice.
     */
    try {

        const url =
            new URL(
                window.location.href
            );

        url.searchParams.delete(
            'gotandtAuto'
        );

        window.history.replaceState(
            {},
            '',
            url.pathname +
            url.search +
            url.hash
        );

    } catch (e) {
        console.log(
            'Could not remove auto-start URL flag.',
            e
        );
    }

    try {

        // --------------------------------------------------------
        // WAIT FOR TODAY'S ACTIVITY PAGE
        // --------------------------------------------------------

        await sleep(800);

        // --------------------------------------------------------
        // CREATE NEW
        // --------------------------------------------------------

        let createNew = null;

        const createStart =
            Date.now();

        while (
            Date.now() - createStart <
            15000
        ) {

            createNew =
                document.querySelector(
                    'button[aria-label="Create new ride or delivery"]'
                );

            if (
                createNew &&
                isVisible(createNew)
            ) {
                break;
            }

            await sleep(150);
        }

        if (!createNew) {

            throw new Error(
                'Could not find Uber Create new button.'
            );
        }

        console.log(
            'Uber Health: clicking Create new.'
        );

        simulateRealClick(
            createNew
        );

        // --------------------------------------------------------
        // SINGLE RIDE
        // --------------------------------------------------------

        let singleRide = null;

        const singleStart =
            Date.now();

        while (
            Date.now() - singleStart <
            7000
        ) {

            singleRide =
                document.querySelector(
                    'li[aria-label="Single ride"][role="option"]'
                );

            if (
                singleRide &&
                isVisible(singleRide)
            ) {
                break;
            }

            await sleep(100);
        }

        if (!singleRide) {

            throw new Error(
                'Could not find Single ride option.'
            );
        }

        console.log(
            'Uber Health: selecting Single ride.'
        );

        simulateRealClick(
            singleRide
        );

        // --------------------------------------------------------
        // WAIT FOR NEW RIDE FORM
        // --------------------------------------------------------

        console.log(
            'Uber Health: waiting for New ride setup.'
        );

        await waitForElement(
            'input[data-testid="pickupAddress_0"]',
            15000
        );

        /*
         * Give Uber a little extra time to finish mounting
         * the form before starting our existing automation.
         */
        await sleep(700);

        // --------------------------------------------------------
        // CLICK OUR PASTE BUTTON
        // --------------------------------------------------------

        const pasteButton =
            document.getElementById(
                BUTTON_ID
            );

        if (!pasteButton) {

            throw new Error(
                'Could not find Paste Ride Data button.'
            );
        }

        console.log(
            'Uber Health: automatically clicking Paste Ride Data.'
        );

        pasteButton.click();

    } catch (err) {

        console.error(
            'Uber Health Excel auto-start error:',
            err
        );

        alert(
            'Uber Health Auto Start Error\n\n' +
            err.message
        );
    }
}

    // ============================================================
    // BASIC HELPERS
    // ============================================================

    function sleep(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }

    function isVisible(el) {
        if (!el) return false;

        const style = window.getComputedStyle(el);
        const rect = el.getBoundingClientRect();

        return (
            style.display !== 'none' &&
            style.visibility !== 'hidden' &&
            style.opacity !== '0' &&
            rect.width > 0 &&
            rect.height > 0
        );
    }

    function cleanText(text) {
        return (text || '')
            .replace(/\s+/g, ' ')
            .trim();
    }

    async function waitForElement(selector, timeout = 10000) {

        const start = Date.now();

        while (Date.now() - start < timeout) {

            const elements = [
                ...document.querySelectorAll(selector)
            ];

            const visible =
                elements.find(isVisible);

            if (visible) {
                return visible;
            }

            await sleep(100);
        }

        throw new Error(
            'Could not find Uber element:\n' +
            selector
        );
    }

    // ============================================================
    // REAL CLICK
    // ============================================================

    function simulateRealClick(el) {

        if (!el) return;

        el.scrollIntoView({
            block: 'nearest'
        });

        const rect =
            el.getBoundingClientRect();

        const clientX =
            rect.left + rect.width / 2;

        const clientY =
            rect.top + rect.height / 2;

        // GM grants put this combined script in Tampermonkey's sandbox.
        // Construct pointer/mouse events with the real page Window, not its sandbox proxy.
        const pageWindow = typeof unsafeWindow !== 'undefined'
            ? unsafeWindow
            : el.ownerDocument.defaultView;

        const common = {
            bubbles: true,
            cancelable: true,
            composed: true,
            view: pageWindow,
            clientX,
            clientY,
            button: 0
        };

        try {
            el.dispatchEvent(
                new pageWindow.PointerEvent(
                    'pointerdown',
                    {
                        ...common,
                        buttons: 1,
                        pointerId: 1,
                        pointerType: 'mouse',
                        isPrimary: true
                    }
                )
            );
        } catch (e) {}

        el.dispatchEvent(
            new pageWindow.MouseEvent(
                'mousedown',
                {
                    ...common,
                    buttons: 1
                }
            )
        );

        try {
            el.dispatchEvent(
                new pageWindow.PointerEvent(
                    'pointerup',
                    {
                        ...common,
                        buttons: 0,
                        pointerId: 1,
                        pointerType: 'mouse',
                        isPrimary: true
                    }
                )
            );
        } catch (e) {}

        el.dispatchEvent(
            new pageWindow.MouseEvent(
                'mouseup',
                {
                    ...common,
                    buttons: 0
                }
            )
        );

        el.click();
    }

    // ============================================================
    // REACT INPUT
    // ============================================================

    function setReactInputValue(input, value) {

        if (!input) return;

        const setter =
            Object.getOwnPropertyDescriptor(
                HTMLInputElement.prototype,
                'value'
            ).set;

        setter.call(input, value);

        input.dispatchEvent(
            new InputEvent(
                'input',
                {
                    bubbles: true,
                    cancelable: true,
                    inputType: 'insertText',
                    data: value
                }
            )
        );

        input.dispatchEvent(
            new Event(
                'change',
                {
                    bubbles: true
                }
            )
        );
    }

    function pressKey(input, key) {

        let code = 0;

        if (key === 'ArrowDown') code = 40;
        if (key === 'Enter') code = 13;
        if (key === 'Escape') code = 27;

        const options = {
            key,
            code: key,
            keyCode: code,
            which: code,
            bubbles: true,
            cancelable: true
        };

        input.dispatchEvent(
            new KeyboardEvent(
                'keydown',
                options
            )
        );

        input.dispatchEvent(
            new KeyboardEvent(
                'keyup',
                options
            )
        );
    }

    // ============================================================
    // TEXT FINDER
    // ============================================================

    function findClickableByText(text) {

        const wanted =
            cleanText(text).toLowerCase();

        const clickable = [
            ...document.querySelectorAll(
                'button, [role="button"], [role="tab"], [role="radio"], label'
            )
        ].filter(isVisible);

        for (const el of clickable) {

            const current =
                cleanText(
                    el.textContent
                ).toLowerCase();

            if (current === wanted) {
                return el;
            }
        }

        const inner = [
            ...document.querySelectorAll(
                'div, span'
            )
        ].filter(isVisible);

        for (const el of inner) {

            const current =
                cleanText(
                    el.textContent
                ).toLowerCase();

            if (current !== wanted) {
                continue;
            }

            const parent =
                el.closest(
                    'button, [role="button"], [role="tab"], [role="radio"], label'
                );

            if (parent) {
                return parent;
            }

            return el;
        }

        return null;
    }

    // ============================================================
    // ADDRESS NORMALIZATION
    // ============================================================

    function normalizeAddress(text) {

        return (text || '')
            .toLowerCase()
            .replace(/[.,#]/g, ' ')
            .replace(/\bstreet\b/g, 'st')
            .replace(/\bavenue\b/g, 'ave')
            .replace(/\broad\b/g, 'rd')
            .replace(/\bdrive\b/g, 'dr')
            .replace(/\bboulevard\b/g, 'blvd')
            .replace(/\blane\b/g, 'ln')
            .replace(/\bcourt\b/g, 'ct')
            .replace(/\bplace\b/g, 'pl')
            .replace(/\bparkway\b/g, 'pkwy')
            .replace(/\bhighway\b/g, 'hwy')
            .replace(/\bnorth\b/g, 'n')
            .replace(/\bsouth\b/g, 's')
            .replace(/\beast\b/g, 'e')
            .replace(/\bwest\b/g, 'w')
            .replace(/\s+/g, ' ')
            .trim();
    }

    function getAddressParts(address) {

        const normalized =
            normalizeAddress(address);

        const words =
            normalized
                .split(' ')
                .filter(Boolean);

        const houseMatch =
            normalized.match(
                /^\d+[a-z]?/i
            );

        const zipMatch =
            normalized.match(
                /\b\d{5}(?:-\d{4})?\b/
            );

        return {
            normalized,
            words,
            houseNumber:
                houseMatch
                    ? houseMatch[0]
                    : '',
            zip:
                zipMatch
                    ? zipMatch[0].substring(0, 5)
                    : ''
        };
    }

    // ============================================================
    // ADDRESS MATCH SCORING
    // ============================================================

    function scoreAddressOption(
        optionText,
        requestedAddress
    ) {

        const option =
            normalizeAddress(
                optionText
            );

        const requested =
            getAddressParts(
                requestedAddress
            );

        let score = 0;

        // House number
        if (requested.houseNumber) {

            const optionHouse =
                option.match(
                    /^\d+[a-z]?/i
                );

            if (
                optionHouse &&
                optionHouse[0] ===
                    requested.houseNumber
            ) {
                score += 100;

            } else if (
                option.includes(
                    requested.houseNumber + ' '
                )
            ) {
                score += 80;

            } else {
                score -= 100;
            }
        }

        // ZIP
        if (
            requested.zip &&
            option.includes(
                requested.zip
            )
        ) {
            score += 150;
        }

        // Words
        for (
            const word of requested.words
        ) {

            if (word.length <= 1) {
                continue;
            }

            if (/^\d{5}$/.test(word)) {
                continue;
            }

            if (
                option.includes(word)
            ) {
                score += 10;
            }
        }

        // Near exact
        if (
            option.includes(
                requested.normalized
            )
        ) {
            score += 250;
        }

        return score;
    }

    function chooseBestAddressOption(
        options,
        requestedAddress
    ) {

        let best = null;
        let bestScore = -Infinity;

        for (const option of options) {

            const text =
                cleanText(
                    option.textContent
                );

            const score =
                scoreAddressOption(
                    text,
                    requestedAddress
                );

            console.log(
                'Uber address suggestion:',
                text,
                'score:',
                score
            );

            if (score > bestScore) {
                best = option;
                bestScore = score;
            }
        }

        if (bestScore < 30) {
            return null;
        }

        console.log(
            'Uber best match:',
            best
                ? cleanText(best.textContent)
                : 'NONE',
            bestScore
        );

        return best;
    }

    // ============================================================
    // ADDRESS AUTOCOMPLETE
    // ============================================================

    function getAddressListbox(input) {

        const listId =
            input.getAttribute(
                'aria-controls'
            );

        if (listId) {

            const listbox =
                document.getElementById(
                    listId
                );

            if (
                listbox &&
                isVisible(listbox)
            ) {
                return listbox;
            }
        }

        const listboxes = [
            ...document.querySelectorAll(
                '[role="listbox"]'
            )
        ].filter(isVisible);

        if (listboxes.length) {

            return listboxes[
                listboxes.length - 1
            ];
        }

        return null;
    }

    function getAddressOptions(input) {

        const listbox =
            getAddressListbox(input);

        if (!listbox) {
            return [];
        }

        let options = [
            ...listbox.querySelectorAll(
                '[role="option"]'
            )
        ].filter(isVisible);

        if (!options.length) {

            options = [
                ...listbox.querySelectorAll(
                    'li, [data-baseweb="menu-item"]'
                )
            ].filter(isVisible);
        }

        return options;
    }

    async function waitForMatchingAddress(
        input,
        requestedAddress,
        timeout = 10000
    ) {

        const start = Date.now();

        while (
            Date.now() - start <
            timeout
        ) {

            const options =
                getAddressOptions(
                    input
                );

            if (options.length) {

                const best =
                    chooseBestAddressOption(
                        options,
                        requestedAddress
                    );

                if (best) {
                    return best;
                }
            }

            await sleep(200);
        }

        throw new Error(
            'Uber did not return a matching address for:\n\n' +
            requestedAddress
        );
    }

    async function addressWasCommitted(
        input,
        timeout = 5000
    ) {

        const start = Date.now();

        while (
            Date.now() - start <
            timeout
        ) {

            const label =
                (
                    input.getAttribute(
                        'aria-label'
                    ) || ''
                ).toLowerCase();

            const expanded =
                input.getAttribute(
                    'aria-expanded'
                );

            if (
                label.includes(
                    'selected'
                ) &&
                expanded === 'false'
            ) {
                return true;
            }

            await sleep(100);
        }

        return false;
    }

    async function enterAddress(
        input,
        address,
        fieldName
    ) {

        input.scrollIntoView({
            behavior: 'smooth',
            block: 'center'
        });

        await sleep(250);

        input.focus();

        setReactInputValue(
            input,
            ''
        );

        pressKey(
            input,
            'Escape'
        );

        await sleep(350);

        setReactInputValue(
            input,
            address
        );

        await sleep(700);

        const bestOption =
            await waitForMatchingAddress(
                input,
                address
            );

        const selectedText =
            cleanText(
                bestOption.textContent
            );

        console.log(
            fieldName,
            'selecting:',
            selectedText
        );

        simulateRealClick(
            bestOption
        );

        const committed =
            await addressWasCommitted(
                input
            );

        if (!committed) {

            throw new Error(
                'Uber found "' +
                selectedText +
                '" but did not accept it.'
            );
        }

        await sleep(900);
    }

    // ============================================================
    // ROUND TRIP
    // ============================================================

    async function selectRoundTrip() {

        let button =
            document.querySelector(
                'button[data-baseweb="tab"][id$="-tab-RoundTrip"]'
            );

        if (!button) {

            button =
                findClickableByText(
                    'Round-trip'
                );
        }

        if (!button) {

            throw new Error(
                'Could not find Round-trip.'
            );
        }

        if (
            button.getAttribute(
                'aria-selected'
            ) !== 'true'
        ) {

            simulateRealClick(
                button
            );

            await sleep(750);
        }
    }

    // ============================================================
    // FUTURE RIDE
    // ============================================================

    async function selectFutureRide() {

        const button =
            findClickableByText(
                'Future ride'
            );

        if (!button) {

            throw new Error(
                'Could not find Future ride.'
            );
        }

        button.scrollIntoView({
            behavior: 'smooth',
            block: 'center'
        });

        await sleep(250);

        simulateRealClick(
            button
        );

        await sleep(650);
    }

    // ============================================================
    // CHOOSE DATE & TIME
    // ============================================================

    async function selectChooseDateAndTime() {

        const target =
            findClickableByText(
                'Choose date & time'
            );

        if (!target) {

            throw new Error(
                'Could not find Choose date & time.'
            );
        }

        target.scrollIntoView({
            behavior: 'smooth',
            block: 'center'
        });

        await sleep(250);

        simulateRealClick(
            target
        );

        // Wait until the date input appears
        await waitForElement(
            'input[aria-label="Select a date."], input#date',
            5000
        );

        await sleep(500);
    }

    // ============================================================
    // DATE PARSING
    // ============================================================

    const MONTH_NAMES = [
        'January',
        'February',
        'March',
        'April',
        'May',
        'June',
        'July',
        'August',
        'September',
        'October',
        'November',
        'December'
    ];

    function parseExcelDate(dateText) {

        const parts =
            (dateText || '')
                .split('/');

        if (parts.length !== 3) {

            throw new Error(
                'Invalid Excel date: ' +
                dateText
            );
        }

        const month =
            parseInt(
                parts[0],
                10
            );

        const day =
            parseInt(
                parts[1],
                10
            );

        let year =
            parseInt(
                parts[2],
                10
            );

        if (year < 100) {
            year += 2000;
        }

        return {
            month,
            day,
            year
        };
    }

// ============================================================
// GET ALL VISIBLE UBER CALENDAR DATE CELLS
// ============================================================

function getVisibleCalendarCells() {

    return [
        ...document.querySelectorAll(
            '[role="gridcell"][aria-label]'
        )
    ].filter(el => {

        if (!isVisible(el)) {
            return false;
        }

        const label =
            el.getAttribute('aria-label') || '';

        /*
         * Do NOT require the label to start with "Choose".
         *
         * Normal dates may say:
         * Choose Thursday, September 3rd 2026. It's available.
         *
         * Today's/currently selected date may have different
         * wording such as Selected / Today.
         */
        return (
            /\b(January|February|March|April|May|June|July|August|September|October|November|December)\b/i.test(label) &&
            /\b20\d{2}\b/.test(label)
        );
    });
}


// ============================================================
// FIND CALENDAR POPUP
// ============================================================

function findCalendarRoot() {

    const cells =
        getVisibleCalendarCells();

    if (!cells.length) {
        return null;
    }

    /*
     * Start at the first date cell and walk upward until
     * we find a container containing a reasonable number
     * of Uber calendar gridcells.
     */
    let current =
        cells[0].parentElement;

    while (
        current &&
        current !== document.body
    ) {

        const count =
            current.querySelectorAll(
                '[role="gridcell"][aria-label]'
            ).length;

        if (count >= 7) {
            return current;
        }

        current =
            current.parentElement;
    }

    /*
     * Calendar definitely exists even if we couldn't
     * identify its exact outer wrapper.
     */
    return cells[0].parentElement;
}


// ============================================================
// WAIT FOR CALENDAR
// ============================================================

async function waitForCalendar(timeout = 5000) {

    const start =
        Date.now();

    while (
        Date.now() - start <
        timeout
    ) {

        const cells =
            getVisibleCalendarCells();

        if (cells.length) {

            console.log(
                'Uber Health: calendar opened.',
                cells.length,
                'date cells found.'
            );

            return findCalendarRoot();
        }

        await sleep(100);
    }

    return null;
}


// ============================================================
// OPEN CALENDAR
// ============================================================

async function openCalendar() {

    const dateInput =
        await waitForElement(
            'input#date, input[aria-label="Select a date."]'
        );

    dateInput.scrollIntoView({
        behavior: 'smooth',
        block: 'center'
    });

    await sleep(200);

    /*
     * Your actual Uber date input responds to click,
     * so use a normal click first.
     */
    dateInput.focus();
    dateInput.click();

    let calendar =
        await waitForCalendar(
            2000
        );

    if (calendar) {
        return calendar;
    }

    /*
     * Fallback to our full pointer/mouse click sequence.
     */
    simulateRealClick(
        dateInput
    );

    calendar =
        await waitForCalendar(
            3000
        );

    if (!calendar) {

        throw new Error(
            'Uber date calendar did not open.'
        );
    }

    return calendar;
}


// ============================================================
// READ CURRENT CALENDAR MONTH / YEAR
// ============================================================

function getCalendarMonthYear(
    calendar
) {

    /*
     * We can determine the displayed month/year directly
     * from Uber's actual aria-labels:
     *
     * Choose Thursday, September 3rd 2026. It's available.
     */

    const cells =
        getVisibleCalendarCells();

    if (!cells.length) {
        return null;
    }

    for (const cell of cells) {

        const label =
            cell.getAttribute(
                'aria-label'
            ) || '';

        const match =
            label.match(
                /,\s*(January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2}(?:st|nd|rd|th)?\s+(20\d{2})/i
            );

        if (match) {

            const month =
                MONTH_NAMES.findIndex(
                    m =>
                        m.toLowerCase() ===
                        match[1].toLowerCase()
                ) + 1;

            return {
                month: month,
                year:
                    parseInt(
                        match[2],
                        10
                    )
            };
        }
    }

    return null;
}


// ============================================================
// FIND CALENDAR NEXT / PREVIOUS ARROW
// ============================================================

function getCalendarArrow(
    calendar,
    direction
) {

    /*
     * First try semantic labels.
     */
    const labelled = [
        ...document.querySelectorAll(
            '[aria-label], [title]'
        )
    ].filter(isVisible);

    for (const el of labelled) {

        const label =
            (
                el.getAttribute(
                    'aria-label'
                ) ||
                el.getAttribute(
                    'title'
                ) ||
                ''
            ).toLowerCase();

        if (
            direction === 'next' &&
            (
                label.includes(
                    'next month'
                ) ||
                label === 'next'
            )
        ) {

            return (
                el.closest(
                    'button, [role="button"]'
                ) ||
                el
            );
        }

        if (
            direction === 'previous' &&
            (
                label.includes(
                    'previous month'
                ) ||
                label.includes(
                    'prev month'
                )
            )
        ) {

            return (
                el.closest(
                    'button, [role="button"]'
                ) ||
                el
            );
        }
    }

    /*
     * Fallback:
     * use the popup containing the date cells and find
     * its SVG/icon click targets.
     */
    const root =
        calendar ||
        findCalendarRoot();

    if (!root) {
        return null;
    }

    const iconTargets = [];

    const svgs = [
        ...root.querySelectorAll(
            'svg'
        )
    ].filter(isVisible);

    for (const svg of svgs) {

        const clickable =
            svg.closest(
                'button, [role="button"], [tabindex]'
            ) ||
            svg.parentElement;

        if (
            clickable &&
            isVisible(clickable) &&
            !iconTargets.includes(
                clickable
            )
        ) {
            iconTargets.push(
                clickable
            );
        }
    }

    if (!iconTargets.length) {
        return null;
    }

    iconTargets.sort(
        (a, b) =>
            a.getBoundingClientRect().left -
            b.getBoundingClientRect().left
    );

    if (
        direction === 'previous'
    ) {
        return iconTargets[0];
    }

    return iconTargets[
        iconTargets.length - 1
    ];
}


// ============================================================
// MOVE CALENDAR TO CORRECT MONTH / YEAR
// ============================================================

async function moveCalendarToMonth(
    targetMonth,
    targetYear
) {

    let calendar =
        await waitForCalendar();

    if (!calendar) {

        throw new Error(
            'Uber calendar disappeared.'
        );
    }

    for (
        let attempt = 0;
        attempt < 36;
        attempt++
    ) {

        const current =
            getCalendarMonthYear(
                calendar
            );

        if (!current) {

            throw new Error(
                'Could not determine Uber calendar month/year.'
            );
        }

        console.log(
            'Uber calendar:',
            current.month,
            current.year,
            'Target:',
            targetMonth,
            targetYear
        );

        if (
            current.month ===
                targetMonth &&
            current.year ===
                targetYear
        ) {

            return calendar;
        }

        const currentValue =
            current.year * 12 +
            current.month;

        const targetValue =
            targetYear * 12 +
            targetMonth;

        const direction =
            targetValue >
                currentValue
                ? 'next'
                : 'previous';

        const arrow =
            getCalendarArrow(
                calendar,
                direction
            );

        if (!arrow) {

            throw new Error(
                'Could not find Uber calendar ' +
                direction +
                ' arrow.'
            );
        }

        simulateRealClick(
            arrow
        );

        await sleep(500);

        calendar =
            await waitForCalendar(
                3000
            );

        if (!calendar) {

            throw new Error(
                'Uber calendar disappeared while changing months.'
            );
        }
    }

    throw new Error(
        'Could not navigate Uber calendar to requested month.'
    );
}


// ============================================================
// CREATE DAY ORDINAL REGEX
// ============================================================

function getDayOrdinalPattern(
    day
) {

    /*
     * Uber uses:
     *
     * 1st
     * 2nd
     * 3rd
     * 4th
     *
     * etc.
     */

    return (
        String(day) +
        '(?:st|nd|rd|th)?'
    );
}


// ============================================================
// FIND EXACT UBER DATE CELL
// ============================================================

function findCalendarDay(
    day,
    month,
    year
) {

    const monthName =
        MONTH_NAMES[
            month - 1
        ];

    const dayPattern =
        getDayOrdinalPattern(
            day
        );

    const dateRegex =
        new RegExp(
            '\\b' +
            monthName +
            '\\s+' +
            dayPattern +
            '\\s+' +
            year +
            '\\b',
            'i'
        );

    const cells =
        getVisibleCalendarCells();

    console.log(
        'Searching Uber calendar for:',
        monthName,
        day,
        year
    );

    for (const cell of cells) {

        const label =
            cell.getAttribute(
                'aria-label'
            ) || '';

        console.log(
            'Uber date cell:',
            label
        );

        if (
            dateRegex.test(
                label
            )
        ) {

            /*
             * Do not select an unavailable date.
             */
            if (
                /unavailable|not available/i.test(
                    label
                )
            ) {

                throw new Error(
                    'Uber shows ' +
                    monthName +
                    ' ' +
                    day +
                    ', ' +
                    year +
                    ' as unavailable.'
                );
            }

            return cell;
        }
    }

    return null;
}


// ============================================================
// SELECT B3 DATE
// ============================================================

async function selectRideDate(dateText) {

    const target =
        parseExcelDate(
            dateText
        );

    console.log(
        'Uber requested ride date:',
        target
    );

    // --------------------------------------------------------
    // CHECK WHETHER B3 IS TODAY
    // --------------------------------------------------------

    const now =
        new Date();

    const isToday =
        target.month ===
            (now.getMonth() + 1) &&
        target.day ===
            now.getDate() &&
        target.year ===
            now.getFullYear();

    const dateInput =
        await waitForElement(
            'input#date, input[aria-label="Select a date."]'
        );

    const currentUberValue =
        (dateInput.value || '')
            .trim()
            .toLowerCase();

    /*
     * Uber already defaults Future Ride to Today.
     *
     * If Excel B3 is today AND Uber currently says Today,
     * there is nothing to select. Leave it alone.
     */
    if (
        isToday &&
        currentUberValue === 'today'
    ) {

        console.log(
            'Uber Health: B3 is today and Uber is already set to Today.'
        );

        return;
    }

    // --------------------------------------------------------
    // OTHERWISE OPEN CALENDAR
    // --------------------------------------------------------

    let calendar =
        await openCalendar();

    // --------------------------------------------------------
    // MOVE TO REQUESTED MONTH/YEAR
    // --------------------------------------------------------

    calendar =
        await moveCalendarToMonth(
            target.month,
            target.year
        );

    // --------------------------------------------------------
    // FIND DATE CELL
    // --------------------------------------------------------

    const dayCell =
        findCalendarDay(
            target.day,
            target.month,
            target.year
        );

    if (!dayCell) {

        throw new Error(
            'Could not find ' +
            MONTH_NAMES[
                target.month - 1
            ] +
            ' ' +
            target.day +
            ', ' +
            target.year +
            ' in Uber calendar.'
        );
    }

    console.log(
        'Uber Health selecting date:',
        dayCell.getAttribute(
            'aria-label'
        )
    );

    // --------------------------------------------------------
    // CLICK DATE
    // --------------------------------------------------------

    simulateRealClick(
        dayCell
    );

    await sleep(700);

    /*
     * If the calendar remains open, try a regular click
     * on Uber's actual gridcell.
     */
    if (
        getVisibleCalendarCells()
            .length > 0
    ) {

        dayCell.click();

        await sleep(500);
    }

    console.log(
        'Uber Health date selection completed:',
        dateText
    );
}

// ============================================================
// UBER TIME SELECTION
// ============================================================

function normalizeUberTime(text) {

    const match =
        (text || '')
            .trim()
            .match(
                /(\d{1,2}):(\d{2})\s*(AM|PM)/i
            );

    if (!match) {
        return null;
    }

    return (
        parseInt(match[1], 10) +
        ':' +
        match[2] +
        ' ' +
        match[3].toUpperCase()
    );
}


// ============================================================
// GET FIRST LEG TIME INPUT
// ============================================================

async function getFirstLegTimeInput(
    timeout = 5000
) {

    const start = Date.now();

    while (
        Date.now() - start <
        timeout
    ) {

        const input =
            document.querySelector(
                'input[id="tripLegs.0.time"]'
            );

        if (
            input &&
            isVisible(input)
        ) {
            return input;
        }

        await sleep(100);
    }

    throw new Error(
        'Could not find Uber first-leg time field.'
    );
}


// ============================================================
// GET THE LISTBOX BELONGING TO THIS TIME FIELD
// ============================================================

function getTimeListbox(input) {

    const listId =
        input.getAttribute(
            'aria-controls'
        );

    if (listId) {

        const list =
            document.getElementById(
                listId
            );

        if (
            list &&
            isVisible(list)
        ) {
            return list;
        }
    }

    /*
     * Fallback if Uber changes the generated ID.
     */
    const visibleLists = [
        ...document.querySelectorAll(
            '[role="listbox"]'
        )
    ].filter(isVisible);

    if (visibleLists.length) {

        return visibleLists[
            visibleLists.length - 1
        ];
    }

    return null;
}


// ============================================================
// TYPE TIME LIKE A USER
// ============================================================

async function typeUberTime(
    input,
    requestedTime
) {

    const wanted =
        normalizeUberTime(
            requestedTime
        );

    if (!wanted) {

        throw new Error(
            'Invalid pickup time from Excel: ' +
            requestedTime
        );
    }

    input.scrollIntoView({
        behavior: 'smooth',
        block: 'center'
    });

    await sleep(200);

    input.focus();

    /*
     * Clear whatever is currently in the search input.
     */
    setReactInputValue(
        input,
        ''
    );

    await sleep(200);

    /*
     * Build the value a character at a time.
     *
     * This more closely resembles manually typing into
     * Uber's BaseWeb combobox and causes its filtering
     * logic to update correctly.
     */
    let typed = '';

    for (const character of wanted) {

        typed += character;

        const nativeSetter =
            Object.getOwnPropertyDescriptor(
                HTMLInputElement.prototype,
                'value'
            ).set;

        nativeSetter.call(
            input,
            typed
        );

        input.dispatchEvent(
            new InputEvent(
                'input',
                {
                    bubbles: true,
                    cancelable: true,
                    inputType:
                        'insertText',
                    data:
                        character
                }
            )
        );

        await sleep(35);
    }

    input.dispatchEvent(
        new Event(
            'change',
            {
                bubbles: true
            }
        )
    );

    await sleep(500);

    return wanted;
}


// ============================================================
// FIND EXACT UBER TIME OPTION
// ============================================================

async function waitForTimeOption(
    input,
    requestedTime,
    timeout = 7000
) {

    const wanted =
        normalizeUberTime(
            requestedTime
        );

    const start =
        Date.now();

    while (
        Date.now() - start <
        timeout
    ) {

        const listbox =
            getTimeListbox(
                input
            );

        if (listbox) {

            const options = [
                ...listbox.querySelectorAll(
                    '[role="option"]'
                )
            ].filter(isVisible);

            for (const option of options) {

                const optionText =
                    cleanText(
                        option.textContent
                    );

                const optionTime =
                    normalizeUberTime(
                        optionText
                    );

                console.log(
                    'Uber time option:',
                    optionText
                );

                /*
                 * This intentionally ignores the
                 * timezone suffix.
                 *
                 * Excel:
                 * 9:30 AM
                 *
                 * Uber:
                 * 9:30 AM EDT
                 */
                if (
                    optionTime === wanted
                ) {

                    return option;
                }
            }
        }

        await sleep(100);
    }

    return null;
}


// ============================================================
// VERIFY UBER ACCEPTED THE TIME
// ============================================================

async function waitForTimeCommitted(
    input,
    requestedTime,
    timeout = 5000
) {

    const wanted =
        normalizeUberTime(
            requestedTime
        );

    const start =
        Date.now();

    while (
        Date.now() - start <
        timeout
    ) {

        /*
         * After selection your HTML shows:
         *
         * aria-label="Selected 9:55 AM EDT. "
         */
        const label =
            input.getAttribute(
                'aria-label'
            ) || '';

        const selectedTime =
            normalizeUberTime(
                label
            );

        if (
            selectedTime === wanted
        ) {
            return true;
        }

        await sleep(100);
    }

    return false;
}


// ============================================================
// SELECT B5 PICKUP TIME
// ============================================================

async function selectRideTime(
    requestedTime
) {

    if (!requestedTime) {
        return;
    }

    const wanted =
        normalizeUberTime(
            requestedTime
        );

    if (!wanted) {

        throw new Error(
            'Invalid pickup time from Excel: ' +
            requestedTime
        );
    }

    console.log(
        'Uber requested pickup time:',
        wanted
    );

    // --------------------------------------------------------
    // GET ACTUAL UBER COMBOBOX
    // --------------------------------------------------------

    const input =
        await getFirstLegTimeInput();

    /*
     * Clicking the time field opens Uber's list.
     */
    input.scrollIntoView({
        behavior: 'smooth',
        block: 'center'
    });

    await sleep(200);

    simulateRealClick(
        input
    );

    await sleep(300);

    // --------------------------------------------------------
    // TYPE B5 INTO UBER
    // --------------------------------------------------------

    await typeUberTime(
        input,
        wanted
    );

    // --------------------------------------------------------
    // WAIT FOR EXACT MATCH
    // --------------------------------------------------------

    const option =
        await waitForTimeOption(
            input,
            wanted
        );

    if (!option) {

        throw new Error(
            'Uber did not return pickup time ' +
            wanted +
            '.'
        );
    }

    const optionText =
        cleanText(
            option.textContent
        );

    console.log(
        'Uber selecting pickup time:',
        optionText
    );

    // --------------------------------------------------------
    // SELECT ACTUAL LI role="option"
    // --------------------------------------------------------

    simulateRealClick(
        option
    );

    await sleep(500);

    // --------------------------------------------------------
    // VERIFY
    // --------------------------------------------------------

    const committed =
        await waitForTimeCommitted(
            input,
            wanted
        );

    if (!committed) {

        /*
         * A normal click is worth a second attempt
         * because the option itself is Uber's true
         * role="option" element.
         */
        option.click();

        await sleep(400);

        const secondCheck =
            await waitForTimeCommitted(
                input,
                wanted,
                2500
            );

        if (!secondCheck) {

            throw new Error(
                'Uber showed ' +
                optionText +
                ' but did not accept the time.'
            );
        }
    }

    console.log(
        'Uber Health pickup time selected:',
        optionText
    );
}

    // ============================================================
    // MAIN PASTE ROUTINE
    // ============================================================

    // Wait for foreground clipboard access; never retry the form-filling steps.
    async function readUberRideClipboard(timeoutMs = 15000) {
        const deadline = Date.now() + timeoutMs;
        const focused = () => document.visibilityState === 'visible' && document.hasFocus();
        while (Date.now() < deadline) {
            if (!focused()) {
                await new Promise(resolve => setTimeout(resolve, 150));
                continue;
            }
            try {
                return await navigator.clipboard.readText();
            } catch (error) {
                if (!focused() || /not focused|document.*focus/i.test(error.message || '')) {
                    await new Promise(resolve => setTimeout(resolve, 150));
                    continue;
                }
                if (error.name === 'NotAllowedError') {
                    throw new Error('Clipboard access was blocked. Allow clipboard access for Uber Health in your browser, then run the Excel macro again.');
                }
                throw error;
            }
        }
        throw new Error('Uber Health did not stay focused. Close any Excel notice, keep the Uber Health page active, and run the Excel macro again.');
    }
    async function pasteRide() {

        const pasteButton =
            document.getElementById(
                BUTTON_ID
            );

        try {

            pasteButton.disabled = true;

            pasteButton.textContent =
                'Reading Clipboard...';

            // ----------------------------------------------------
            // READ EXCEL JSON
            // ----------------------------------------------------

            const clipboardText =
                await readUberRideClipboard();

            if (!clipboardText) {

                throw new Error(
                    'Clipboard is empty.'
                );
            }

            let ride;

            try {

                ride =
                    JSON.parse(
                        clipboardText
                    );

            } catch (e) {

                throw new Error(
                    'Clipboard does not contain valid Uber ride data. ' +
                    'Click the Uber button in Excel again.'
                );
            }

            console.log(
                'Uber Health ride data:',
                ride
            );

            if (
                !ride.pickupAddress ||
                !ride.dropoffAddress
            ) {

                throw new Error(
                    'Pickup or dropoff address is missing from Excel.'
                );
            }

            if (!ride.date) {

                throw new Error(
                    'Ride date is missing from Excel B3.'
                );
            }

            if (!ride.pickupTime) {

                throw new Error(
                    'Pickup time is missing from Excel B5.'
                );
            }

            // ----------------------------------------------------
            // ROUND TRIP
            // ----------------------------------------------------

            pasteButton.textContent =
                'Selecting Round Trip...';

            await selectRoundTrip();

            // ----------------------------------------------------
            // PICKUP ADDRESS
            // ----------------------------------------------------

            pasteButton.textContent =
                'Entering Pickup...';

            const pickup =
                await waitForElement(
                    'input[data-testid="pickupAddress_0"]'
                );

            await enterAddress(
                pickup,
                ride.pickupAddress,
                'pickup'
            );

            // ----------------------------------------------------
            // DROPOFF ADDRESS
            // ----------------------------------------------------

            pasteButton.textContent =
                'Entering Dropoff...';

            const dropoff =
                await waitForElement(
                    'input[data-testid="dropoffAddress_0_0"]'
                );

            await enterAddress(
                dropoff,
                ride.dropoffAddress,
                'dropoff'
            );

            // ----------------------------------------------------
            // FUTURE RIDE
            // ----------------------------------------------------

            pasteButton.textContent =
                'Selecting Future Ride...';

            await selectFutureRide();

            // ----------------------------------------------------
            // CHOOSE DATE & TIME
            // ----------------------------------------------------

            pasteButton.textContent =
                'Choosing Date & Time...';

            await selectChooseDateAndTime();

            // ----------------------------------------------------
            // B3 DATE
            // ----------------------------------------------------

            pasteButton.textContent =
                'Setting Ride Date...';

            await selectRideDate(
                ride.date
            );

            // ----------------------------------------------------
            // B5 PICKUP TIME
            // ----------------------------------------------------

            pasteButton.textContent =
                'Setting Pickup Time...';

            await selectRideTime(
                ride.pickupTime
            );

            // ----------------------------------------------------
            // DONE
            // ----------------------------------------------------

            pasteButton.textContent =
                'Ride Data Entered ✓';

            console.log(
                'Uber Health first leg completed.'
            );

        } catch (err) {

            console.error(
                'Uber Health Paste Ride error:',
                err
            );

            alert(
                'Uber Health Paste Ride Error\n\n' +
                err.message
            );

            pasteButton.textContent =
                'Paste Ride Data';

        } finally {

            pasteButton.disabled = false;
        }
    }

    // ============================================================
    // ADD BUTTON
    // ============================================================

function addPasteButton() {

    if (
        document.getElementById(
            BUTTON_ID
        )
    ) {
        return;
    }

    const button =
        document.createElement(
            'button'
        );

    button.id =
        BUTTON_ID;

    button.type =
        'button';

    button.textContent =
        'Paste Ride Data';

    // Keep button available for automation,
    // but hide it from the user.
    button.style.display =
        'none';

    button.addEventListener(
        'click',
        pasteRide
    );

    document.body.appendChild(
        button
    );
}

    // ============================================================
    // KEEP BUTTON PRESENT
    // ============================================================

    addPasteButton();

    const observer =
        new MutationObserver(
            function () {
                addPasteButton();
            }
        );

    observer.observe(
        document.documentElement,
        {
            childList: true,
            subtree: true
        }
    );

    // ============================================================
    // CHECK WHETHER EXCEL OPENED UBER
    // ============================================================

    setTimeout(
        function () {
            autoStartRideFromExcel();
        },
        500
    );


    }

    // UIEnhancer runs at document-start; Uber needs the page body before adding its button.
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', startUberHealth, { once: true });
    } else {
        startUberHealth();
    }
})();
