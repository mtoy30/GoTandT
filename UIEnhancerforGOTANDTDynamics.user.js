// ==UserScript==
// @name         UIEnhancerforGOTANDTDynamics
// @namespace    https://github.com/mtoy30/GoTandT
// @version      1.3.6.6
// @updateURL    https://raw.githubusercontent.com/mtoy30/GoTandT/main/UIEnhancerforGOTANDTDynamics.user.js
// @downloadURL  https://raw.githubusercontent.com/mtoy30/GoTandT/main/UIEnhancerforGOTANDTDynamics.user.js
// @description  Dynamics UI tweaks; Boomerang form autofill (clipboard → GM storage bridge → googleusercontent iframe); PowerApps Copy button for Leg Info overlay.
// @author       Michael Toy
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

    function collectOverlayText(container) {
      let nodes = Array.from(container.querySelectorAll(TEXT_SELECTORS)).filter(isVisible);
      nodes = nodes.filter(n => !n.closest('#mtoy-inline-copy') && !n.closest('#mtoy-copy-toast'));

      const set = new Set(nodes);
      nodes = nodes.filter(n => !Array.from(set).some(m => m !== n && n.contains(m)));

      const items = nodes.map(el => {
        const rect = el.getBoundingClientRect();
        const text = compact(el.innerText || el.textContent || '');
        return { el, rect, text };
      }).filter(i => i.text);

      if (!items.length) return '';

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

      const out = [];
      for (const it of items) out.push(it.text);
      return out.join('\n');
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
      btn.title = 'Copy all visible info in this Leg Info panel';
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
      btn.addEventListener('click', () => {
        try {
          const text = collectOverlayText(overlay);
          if (!text) { alert('Nothing visible to copy.'); return; }
          if (typeof GM_setClipboard === 'function') {
            GM_setClipboard(text, { type: 'text', mimetype: 'text/plain' });
          } else {
            navigator.clipboard.writeText(text).catch(() => {});
          }
          showToast('Copied.');
        } catch (e) {
          console.error('Copy failed', e);
          alert('Copy failed. See console.');
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
      'Alexandra Cirlan','Annejulia Villegas-Torres','Christian Antunez','Christina Armstrong',
      'Damaris Olmeda','David Hobbs','Jeremy Rivera','Kevin Roberts',
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
    const statusText     = "Pending - RATE Authorization Requested";
    const headerSelector = '[id^="formHeaderTitle_"]';
    const buttonSelector = 'button[aria-label="Rate Approval Status"]';

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

    function insertBanner() {
      const header = document.querySelector(headerSelector);
      if (!header || document.getElementById('rate-status-banner')) return;
      const banner = document.createElement('div');
      banner.id = 'rate-status-banner';
      banner.textContent = "**PENDING RATES**";
      banner.style.backgroundColor = 'lightblue';
      banner.style.color = 'black';
      banner.style.padding = '5px';
      banner.style.marginTop = '5px';
      banner.style.fontWeight = 'normal';
      banner.style.textAlign = 'center';
      banner.style.borderRadius = '5px';
      header.parentNode.insertBefore(banner, header.nextSibling);
    }

    function removeBanner() {
      const existing = document.getElementById('rate-status-banner');
      if (existing) existing.remove();
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

    function highlightAllRowsGlobal() {
      const vipTerms = [
        "defense medical exam -  always vip !!","defense medical exam -  always vip!!",
        "dr.'s visit: 2nd opinion -  always vip!!","dr.'s visit: ime:  always vip !!",
        "pqme -   always vip!!","pqme - panel qualified medical examination - always vip!!",
        "qme - qualified medical exam) vip!!","ame - agreed medical evaluation",
        "evaluation","fce - long appt!!!!!"
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
      banner.style.marginTop = "5px";
      banner.style.fontWeight = "bold";
      banner.style.textAlign = "center";
      banner.style.borderRadius = "5px";
      if (header) header.parentNode.insertBefore(banner, header.nextSibling);
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
      banner.style.marginTop = "5px";
      banner.style.fontWeight = "bold";
      banner.style.textAlign = "center";
      banner.style.borderRadius = "5px";
      if (header) header.parentNode.insertBefore(banner, header.nextSibling);
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

    function removeAuthElementsFromDoc(doc) {
      if (!doc) return 0;
      let removedCount = 0;
      const selectors = [
        '[data-id="gtt_attachauthemail"]','[data-control-name="gtt_attachauthemail"]',
        '[data-id="gtt_authorizationdocument"]','[data-control-name="gtt_authorizationdocument"]',
        '[data-id="gtt_authorizationrequired"]','[data-control-name="gtt_authorizationrequired"]',
        '[data-id="gtt_attachauthemail.fieldControl_container"]',
        '[data-id="gtt_authorizationdocument.fieldControl_container"]',
        '[data-id="gtt_authorizationrequired.fieldControl-pcf-container-id"]',
        '[data-id="gtt_attachauthemail-FieldSectionItemContainer"]',
        '[data-id="gtt_authorizationdocument-FieldSectionItemContainer"]',
        '[data-id="gtt_authorizationrequired-FieldSectionItemContainer"]'
      ];
      const seen = new Set();
      selectors.forEach(sel => {
        doc.querySelectorAll(sel).forEach(el => {
          // Walk up only to known field-level containers — removed the .pa-bz escalation
          // which was grabbing large page sections and causing the flash-then-blank issue.
          const target =
            el.closest('[data-id="gtt_attachauthemail-FieldSectionItemContainer"]') ||
            el.closest('[data-id="gtt_authorizationdocument-FieldSectionItemContainer"]') ||
            el.closest('[data-id="gtt_authorizationrequired-FieldSectionItemContainer"]') ||
            el.closest('[data-id="gtt_attachauthemail"]') ||
            el.closest('[data-control-name="gtt_attachauthemail"]') ||
            el.closest('[data-id="gtt_authorizationdocument"]') ||
            el.closest('[data-control-name="gtt_authorizationdocument"]') ||
            el.closest('[data-id="gtt_authorizationrequired"]') ||
            el.closest('[data-control-name="gtt_authorizationrequired"]') ||
            el;
          if (target && target.parentNode && !seen.has(target)) {
            seen.add(target);
            target.remove();
            removedCount++;
          }
        });
      });
      return removedCount;
    }

    function removeAuthElementsEverywhere() {
      let total = removeAuthElementsFromDoc(document);
      const iframe = document.querySelector('#WebResource_RecipientSelector');
      if (iframe) {
        try {
          const idoc = iframe.contentDocument || iframe.contentWindow?.document;
          total += removeAuthElementsFromDoc(idoc);
        } catch {}
      }
      return total;
    }

    function startAuthElementRemoval() {
      // Staggered initial passes to catch lazily-rendered fields — skip the immediate
      // call at t=0 which was running before Dynamics had anything rendered.
      setTimeout(removeAuthElementsEverywhere, 500);
      setTimeout(removeAuthElementsEverywhere, 1500);
      setTimeout(removeAuthElementsEverywhere, 3000);
      setTimeout(removeAuthElementsEverywhere, 6000);

      // MutationObserver catches fields injected dynamically after load.
      // Skips when body has very few children = Dynamics mid-navigation blank slate.
      const mo = new MutationObserver(() => {
        if (document.body && document.body.childElementCount < 3) return;
        removeAuthElementsEverywhere();
      });
      const startObserver = () => {
        if (!document.body) { setTimeout(startObserver, 100); return; }
        mo.observe(document.body, { childList: true, subtree: true });
      };
      startObserver();

      // Periodic sweep — self-stops after 60s (30 × 2000ms). Page is stable by then.
      // Raise sweepCap if you need longer coverage on slow machines.
      let sweepCount = 0;
      const sweepCap = 30;
      const sweepInterval = setInterval(() => {
        if (document.body && document.body.childElementCount >= 3) {
          removeAuthElementsEverywhere();
        }
        if (++sweepCount >= sweepCap) clearInterval(sweepInterval);
      }, 2000);
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
      const claimant = Array.from(document.querySelectorAll('a[aria-label][href*="etn=contact"]'))
        .find(el => isVisibleElement(el) && (el.textContent || '').trim());
      return (claimant?.textContent || '').trim();
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
      message.textContent = 'Out of Juristiction claim - Please confirm rates on PO are the contracted rates if not please email cordinator for updates contracted rates';
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
        highlightAllRowsGlobal();
        addLegend();
        adjustSpacing();
        styleNotificationWrapper();
        observeRateStatusChanges();
        insertJbaBannerIfNeeded();
        insertVipBannerIfNeeded();
        checkCareWorksJurisdiction();
        isInSearchUIWrapper();
      }, 200);
    });
    globalObserver.observe(document.body, { childList: true, subtree: true });

    let attempts = 0;
    const tryInit = setInterval(() => {
      checkStatusAndInsertBanner();
      highlightAllRowsGlobal();
      addLegend();
      adjustSpacing();
      styleNotificationWrapper();
      observeRateStatusChanges();
      insertJbaBannerIfNeeded();
      insertVipBannerIfNeeded();
      isInSearchUIWrapper();
      attempts++;
      if (attempts > 20) clearInterval(tryInit);
    }, 500);

    observeNotifications();
    startAuthElementRemoval();
    setInterval(checkCareWorksJurisdiction, 1200);

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
