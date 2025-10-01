// ==UserScript==
// @name         UIEnhancerforGOTANDTDynamics
// @namespace    https://github.com/mtoy30/GoTandT
// @version      1.2.10
// @updateURL    https://raw.githubusercontent.com/mtoy30/GoTandT/main/UIEnhancerforGOTANDTDynamics.user.js
// @downloadURL  https://raw.githubusercontent.com/mtoy30/GoTandT/main/UIEnhancerforGOTANDTDynamics.user.js
// @description  Dynamics UI tweaks; Boomerang form autofill behavior (iframe-safe). Time fields + key fields always unlocked; company/email soft-prefill; unlock-all-on-submit. Also adds a yellow Copy button in PowerApps Leg Info overlay that preserves on-screen order (including duplicate lines like city/state).
// @author       Michael Toy
// @match        https://*.powerapps.com/*
// @match        https://*.powerplatform.com/*
// @match        https://gotandt.crm.dynamics.com/*
// @match        https://gttqap2.crm.dynamics.com/*
// @match        https://boomerangtransport.net/ride-input-request/*
// @match        https://*.googleusercontent.com/*
// @match        https://script.google.com/*
// @match        https://script.googleusercontent.com/*
// @grant        GM_setClipboard
// @run-at       document-start
// ==/UserScript==

(function () {
  'use strict';

  const host = location.hostname;
  const isPowerApps = /\.powerapps\.com$/i.test(host) || /\.powerplatform\.com$/i.test(host);
  const isDynamics  = /(?:^|\.)gotandt\.crm\.dynamics\.com$/i.test(host) || /(?:^|\.)gttqap2\.crm\.dynamics\.com$/i.test(host);
  const onBoomerang = host === 'boomerangtransport.net';
  const onAppsScript =
    /(?:^|\.)googleusercontent\.com$/i.test(host) ||
    host === 'script.google.com' ||
    host === 'script.googleusercontent.com';

  /* =====================================================================================
     PART A — POWER APPS: “Copy” button above the date inside Leg Info overlay
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
    // City/state ZIP detector to keep two-city rows in visual (left→right) order
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

    // NEW: determine if two boxes visually share a row using vertical overlap
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

      // keep leaf-most nodes (avoid parent+child dupes)
      const set = new Set(nodes);
      nodes = nodes.filter(n => !Array.from(set).some(m => m !== n && n.contains(m)));

      const items = nodes.map(el => {
        const rect = el.getBoundingClientRect();
        const text = compact(el.innerText || el.textContent || '');
        return { el, rect, text };
      }).filter(i => i.text);

      if (!items.length) return '';

      items.sort((a, b) => {
        // First: by visual row (top→bottom) using overlap, not just top distance
        if (!sameVisualRow(a, b)) {
          const topDelta = a.rect.top - b.rect.top;
          if (topDelta !== 0) return topDelta;
          // rare exact tie: fall back to left→right
          const leftDelta = a.rect.left - b.rect.left;
          if (leftDelta !== 0) return leftDelta;
        } else {
          // Within the same visual row:
          const aCity = CITY_RE.test(a.text), bCity = CITY_RE.test(b.text);
          if (aCity && bCity) {
            // Two city/state/ZIPs → keep left→right on the row
            const dx = a.rect.left - b.rect.left;
            if (dx !== 0) return dx;
          }
          // Default: DOM order to keep label→value semantics
          if (a.el !== b.el) return domPrecedes(a.el, b.el) ? -1 : 1;
        }
        return 0;
      });

      // Keep duplicates on purpose (e.g., same city for pickup & drop)
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

      if (existing && existing.nextElementSibling === anchor) return; // already in place
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

      anchor.parentNode.insertBefore(btn, anchor); // sits right above the date
    }

    (async () => {
      await until(() => document.body, { tries: 600, delay: 50 });
      const tick = () => injectInlineButton();
      tick();
      const mo = new MutationObserver(tick);
      mo.observe(document.documentElement, { childList: true, subtree: true });
      setInterval(tick, 800);
    })();

    return; // don’t run the Dynamics code in Power Apps frames
  }

  /* =====================================================================================
     PART B — UI ENHANCER (Dynamics, Boomerang, Apps Script)
     ===================================================================================== */

  /* Early out: only run enhancer on Dynamics/Boomerang/Apps Script */
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

    const UNLOCK_WINDOW_MS = 10000;

    const YOURNAME_OPTIONS = [
      'Alexandra Cirlan',
      'Ashley Oliver',
      'Christian Antunez',
      'Christina Armstrong',
      'Chrissi Denson',
      'Damaris Olmeda',
      'David Hobbs',
      'Jasmine Allen',
      'Jeremy Rivera',
      'Kevin Roberts',
      'Michael Toy',
      'Naomi Picklesimer',
      'Tonyjay Matias',
      'Yuri Nichols'
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
        el.addEventListener('submit', () => {
          unlockAllInForm(el);
        }, { capture: true });
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
      const clickUnlock = (e) => {
        const t = e.target;
        if (!t) return;
        if (t.matches && (t.matches('button[type="submit"], input[type="submit"]') || /submit/i.test(t.getAttribute?.('type') || ''))) {
          const form = t.form || doc.querySelector('form');
          unlockAllInFormScope(form || doc);
        }
      };
      doc.addEventListener('click', clickUnlock, true);
    }

    function wireUnlockForAppsScriptSubmit(doc) {
      const onPreSubmit = (e) => {
        const btn = e.target && e.target.closest
          ? e.target.closest('#buttonSubmit, .submit-button')
          : null;
        if (!btn) return;
        const container = btn.parentElement || doc;
        unlockAllInFormScope(container);
        container.querySelectorAll('input, textarea, select').forEach(f => {
          try { f.blur && f.blur(); } catch {}
        });
      };
      doc.addEventListener('click', onPreSubmit, true);
      doc.addEventListener('keydown', (e) => {
        if (!e) return;
        const active = doc.activeElement;
        if (!active) return;
        const isSubmitLink = active.matches?.('#buttonSubmit, .submit-button');
        if (!isSubmitLink) return;
        if (e.key === 'Enter' || e.key === ' ' || e.keyCode === 13 || e.keyCode === 32) {
          const container = active.parentElement || doc;
          unlockAllInFormScope(container);
          container.querySelectorAll('input, textarea, select').forEach(f => {
            try { f.blur && f.blur(); } catch {}
          });
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
        } catch { /* cross-origin; also matched separately if same-origin */ }
      };
      ifr.addEventListener('load', tryWire);
      tryWire();
    }

    if (document.readyState === 'loading') {
      document.addEventListener('DOMContentLoaded', () => apply(document));
    } else {
      apply(document);
    }
    // Do not `return`; we may also be on Dynamics—handled below via isDynamics
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
    function removeLegendsInsideSearch() {
      document.querySelectorAll('[data-legend="true"]').forEach(el => {
        if (isInSearchUI(el)) el.remove();
      });
    }

    function insertBanner() {
      const header = document.querySelector(headerSelector);
      if (!header || document.getElementById('rate-status-banner')) return;
      const banner = document.createElement('div');
      banner.id = 'rate-status-banner';
      banner.textContent = statusText;
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
        "defense medical exam -  always vip !!",
        "defense medical exam -  always vip!!",
        "dr.'s visit: 2nd opinion -  always vip!!",
        "dr.'s visit: ime:  always vip !!",
        "pqme -   always vip!!",
        "pqme - panel qualified medical examination - always vip!!",
        "qme - qualified medical exam) vip!!",
        "ame - agreed medical evaluation",
        "evaluation",
        "fce - long appt.!!!!!"
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

    function createLegendElement(type = "default") {
      const legend = document.createElement("div");
      legend.style.marginTop = "5px";
      legend.style.fontSize = "14px";
      legend.dataset.legend = "true";
      let html = `
        <span style="background-color: gold; padding: 2px 6px; border-radius: 4px; margin-left: 20px; margin-right: 5px;">Evaluations</span>
        <span style="background-color: lightgreen; padding: 2px 6px; border-radius: 4px; margin-right: 5px;">First Time/Surgery</span>
        <span style="background-color: lightcoral; padding: 2px 6px; border-radius: 4px; margin-right: 5px;">Airport</span>
      `;
      if (type === "default") {
        html += `
          <span style="background-color: lightblue; padding: 2px 6px; border-radius: 4px; margin-right: 5px;">Pending Rate Approval</span>
          <span style="background-color: plum; padding: 2px 6px; border-radius: 4px; margin-right: 5px;">Rates Approved</span>
        `;
      }
      if (type === "-confirm") {
        html += `<span style="background-color: #FFDAB9; padding: 2px 6px; border-radius: 4px;">Pickup/Appt Time Passed</span>`;
      }
      legend.innerHTML = html;
      return legend;
    }
    function getMainRoot() {
      return document.querySelector('[data-id="fullPageContentRoot"]')
          || document.querySelector('[data-lp-id="fullPageContentRoot"]')
          || document.body;
    }
    function addLegend() {
      const root = getMainRoot();
      removeLegendsInsideSearch();
      root.querySelectorAll('[data-legend="true"]').forEach(el => el.remove());
      const targetSpans = root.querySelectorAll('span[id*="_text-value"]');
      targetSpans.forEach(span => {
        if (!span || isInSearchUI(span)) return;
        const label = (span.textContent || '').toLowerCase();
        if ((label.includes("unassigned transportation") ||
             label.includes("same day confirmations") ||
             label.includes("same day (oncall)")) &&
            !label.includes("delete")) {
          const type = (label.includes("-confirm") || label.includes("same day confirmations") || label.includes("same day (oncall)"))
            ? "-confirm" : "default";
          const legend = createLegendElement(type);
          span.parentElement?.appendChild(legend);
        }
      });
      const buttons = root.querySelectorAll('button[aria-label*="~Transport"], button[aria-label*="UBER"], button[aria-label*="Prev Vendor Search"], button[aria-label]');
      buttons.forEach(button => {
        if (!button || isInSearchUI(button)) return;
        const ariaLabel = (button.getAttribute('aria-label') || '').toLowerCase();
        const labelText = (button.querySelector('.ms-Button-label')?.textContent || '').toLowerCase();
        if (ariaLabel.includes("delete") || labelText.includes("delete")) return;
        const isConfirm = labelText.includes("-confirm")
                       || labelText.includes("same day confirmations")
                       || labelText.includes("same day (oncall)");
        if (ariaLabel.includes("unassigned transport") ||
            ariaLabel.includes("~transport") ||
            ariaLabel.includes("uber") ||
            ariaLabel.includes("prev vendor search") ||
            isConfirm) {
          const type = isConfirm ? "-confirm" : "default";
          const legend = createLegendElement(type);
          button.parentElement?.insertBefore(legend, button.nextSibling);
        }
      });
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
      const notificationElements = document.querySelectorAll('[id*="notificationWrapper"], [id*="message"]');
      notificationElements.forEach(element => {
        element.style.fontWeight = 'bold';
        element.style.fontSize = '18px';
        element.style.backgroundColor = 'lightgreen';
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
      banner.textContent = "JBA file please ensure a quote is not needed before staffing";
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
        "4474-64737","10837-61025","4474-66551","4474-63533"
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
      document.querySelectorAll('[data-legend="true"]').forEach(el => {
        if (isInSearchUI(el)) el.remove();
      });
    }

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

    function waitForMoniqueInIframe(retries = 20, delay = 1000) {
      const iframe = document.querySelector('#WebResource_RecipientSelector');
      if (!iframe) {
        if (retries > 0) setTimeout(() => waitForMoniqueInIframe(retries - 1, delay), delay);
        return;
      }
      const doc = iframe.contentDocument || iframe.contentWindow.document;
      if (!doc || !doc.body) {
        if (retries > 0) setTimeout(() => waitForMoniqueInIframe(retries - 1, delay), delay);
        return;
      }
      const td = [...doc.querySelectorAll("td")].find(td =>
        td.textContent.trim().includes("Monique Jones")
      );
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
      if (!iframe) {
        if (retries > 0) setTimeout(() => waitForAUTHEMAIL(retries - 1, delay), delay);
        return;
      }
      const doc = iframe.contentDocument || iframe.contentWindow.document;
      if (!doc || !doc.body) {
        if (retries > 0) setTimeout(() => waitForAUTHEMAIL(retries - 1, delay), delay);
        return;
      }
      const td = [...doc.querySelectorAll("td")].find(td =>
        td.textContent.trim().includes("AUTH EMAIL")
      );
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

    const titleObserver = new MutationObserver(() => {
      if (document.title.includes("Email:")) {
        waitForMoniqueInIframe();
        waitForAUTHEMAIL();
      }
    });
    const titleNode = document.querySelector("title");
    if (titleNode) titleObserver.observe(titleNode, { childList: true });
    if (document.title.includes('Email:')) {
        waitForMoniqueInIframe();
        waitForAUTHEMAIL();
    }
  }
})();
