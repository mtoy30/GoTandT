// ==UserScript==
// @name         sptab
// @namespace    https://github.com/mtoy30/GoTandT
// @version      0.4.1
// @updateURL    https://raw.githubusercontent.com/mtoy30/GoTandT/main/sptab.user.js
// @downloadURL  https://raw.githubusercontent.com/mtoy30/GoTandT/main/sptab.user.js
// @description  Hides specific Power Apps controls, repositions NoteText & Button4, and conditionally adds a "Portal Submission" button that sets the note and clicks Button4 to save.
// @author       Michael Toy
// @match        https://apps.powerapps.com/*
// @match        https://*.powerapps.com/*
// @match        https://runtime-app.powerplatform.com/*
// @match        https://*.powerplatform.com/*
// @grant        none
// @run-at       document-start
// ==/UserScript==

(function () {
  'use strict';

  // ---------- tiny utils ----------
  const sleep = (ms) => new Promise(r => setTimeout(r, ms));
  const until = async (testFn, { tries = 300, delay = 100 } = {}) => {
    for (let i = 0; i < tries; i++) {
      const v = testFn();
      if (v) return v;
      await sleep(delay);
    }
    return null;
  };

  function upsertStyle(css, id = 'pa-move-note-style') {
    let s = document.getElementById(id);
    if (!s) {
      s = document.createElement('style');
      s.id = id;
      s.type = 'text/css';
      document.documentElement.appendChild(s);
    }
    if (s.textContent !== css) s.textContent = css;
  }

  function inPowerAppsDom(doc = document) {
    return !!doc.querySelector('.canvasContentDiv, .appmagic-content-control-name, .player-app-frame, .appmagic-textbox, .appmagic-label');
  }

  // ---------- Layout targets (tweak as needed) ----------
  const notePos = { left: 507, top: 441, width: 407, height: 95, z: 50 };
  const btnPos  = { left: 507, top: 545, width: 232, height: 31, z: 51 };
  // Portal button sits to the right of Button4
  const portalBtnPos = {
    left: btnPos.left + btnPos.width + 10,
    top: btnPos.top,
    width: 180,
    height: btnPos.height,
    z: 52,
  };

  function buildCss() {
    const removeSelectors = [
      '[data-control-name="lblAccountInfo_RateAddInfo_1"]',
      '[data-control-name="lblFutureAppointmnets_1"]',
      '[data-control-name="tbAccountInfo_RateAddInfo_1"]',
      '[data-control-name="Label15_2"]',
      '[data-control-name="Label15_3"]',
      '[data-control-name="Future Appointments_1"]',
    ];

    return `
      /* ---------- Hide the unwanted controls entirely ---------- */
      ${removeSelectors.join(',\n')} {
        display: none !important;
        visibility: hidden !important;
        pointer-events: none !important;
      }

      /* ---------- Reposition NoteText ---------- */
      [data-control-name="NoteText"] {
        position: absolute !important;
        left: ${notePos.left}px !important;
        top: ${notePos.top}px !important;
        width: ${notePos.width}px !important;
        height: ${notePos.height}px !important;
        z-index: ${notePos.z} !important;
      }

      [data-control-name="NoteText"] textarea[appmagic-control="NoteTexttextarea"] {
        width: 100% !important;
        height: 100% !important;
      }

      /* ---------- Reposition Button4 ---------- */
      [data-control-name="Button4"] {
        position: absolute !important;
        left: ${btnPos.left}px !important;
        top: ${btnPos.top}px !important;
        width: ${btnPos.width}px !important;
        height: ${btnPos.height}px !important;
        z-index: ${btnPos.z} !important;
      }

      /* ---------- Our injected Portal button ---------- */
      #pa-portal-submission {
        position: absolute !important;
        left: ${portalBtnPos.left}px !important;
        top: ${portalBtnPos.top}px !important;
        width: ${portalBtnPos.width}px !important;
        height: ${portalBtnPos.height}px !important;
        z-index: ${portalBtnPos.z} !important;

        background: rgb(0,120,212) !important;
        color: #fff !important;
        border: none !important;
        padding: 5px !important;
        font-family: "Segoe UI","Open Sans",sans-serif !important;
        font-size: 10.5pt !important;
        font-weight: 600 !important;
        text-align: center !important;
        cursor: pointer !important;
        border-radius: 2px !important;

        display: none; /* toggled by JS */
      }

      #pa-portal-submission:disabled {
        opacity: 0.6 !important;
        cursor: not-allowed !important;
      }
    `;
  }

  // ---------- Label detection ----------
  const TARGET_LABEL = 'BOOMERANG TRANSPORT LLC (ACH)';

  function labelMatches() {
    const nodes = document.querySelectorAll('.appmagic-label-text,[data-control-part="text"].appmagic-label-text');
    for (const n of nodes) {
      const t = (n.textContent || '').trim().toUpperCase();
      if (t === TARGET_LABEL) return true;
    }
    return false;
  }

  // ---------- NoteText helpers ----------
  function getNoteTextarea() {
    return document.querySelector('[data-control-name="NoteText"] textarea[appmagic-control="NoteTexttextarea"]');
  }

  function setNoteText(val) {
    const ta = getNoteTextarea();
    if (!ta) return false;
    if (ta.hasAttribute('readonly')) return false; // respect view mode
    ta.value = val;
    ta.dispatchEvent(new Event('input', { bubbles: true }));
    ta.dispatchEvent(new Event('change', { bubbles: true }));
    return true;
  }

  // ---------- Get the real clickable Button4 element ----------
  function getButton4Clickable() {
    // Prefer the inner clickable button
    let btn = document.querySelector('[data-control-name="Button4"] button.appmagic-button-container');
    if (btn) return btn;
    // Fallback: click the whole control if inner button not found
    btn = document.querySelector('[data-control-name="Button4"]');
    return btn || null;
  }

  // ---------- Portal button injection (writes note, then clicks Button4) ----------
  function ensurePortalButton() {
    const btn4Clickable = getButton4Clickable();
    const shouldShow = !!btn4Clickable && labelMatches();

    // Keep it on same absolute layer as Button4 (so coords align)
    const btn4Host = document.querySelector('[data-control-name="Button4"]');
    const host = btn4Host ? btn4Host.parentElement : document.body;

    let portalBtn = document.getElementById('pa-portal-submission');

    if (!shouldShow) {
      if (portalBtn) portalBtn.style.display = 'none';
      return;
    }

    if (!portalBtn) {
      portalBtn = document.createElement('button');
      portalBtn.id = 'pa-portal-submission';
      portalBtn.type = 'button';
      portalBtn.textContent = 'Portal Submission';

      portalBtn.addEventListener('click', async () => {
        const wrote = setNoteText('Submitted on portal');

        // Give the framework a moment to bind changes before clicking save
        if (wrote && btn4Clickable) {
          // Some apps need a tiny delay to ensure model updates
          await sleep(50);
          btn4Clickable.click();
        }

        // brief visual feedback
        portalBtn.disabled = true;
        const original = portalBtn.textContent;
        portalBtn.textContent = wrote ? 'Saved ✓' : 'Not Editable';
        setTimeout(() => {
          portalBtn.textContent = original;
          portalBtn.disabled = false;
        }, 1200);
      });

      host.appendChild(portalBtn);
    } else if (portalBtn.parentElement !== host) {
      host.appendChild(portalBtn); // keep with Button4's layer
    }

    portalBtn.style.display = 'block';
  }

  // ---------- Main ----------
  async function init() {
    await until(() => inPowerAppsDom(document), { tries: 400, delay: 100 });

    const css = buildCss();
    upsertStyle(css);

    // Mutation observer: keep styles and button presence healthy
    const mo = new MutationObserver(() => {
      if (!document.getElementById('pa-move-note-style')) upsertStyle(css);
      ensurePortalButton();
    });
    mo.observe(document.documentElement, { childList: true, subtree: true });

    // Periodic keepalive (Power Apps re-writes inline styles frequently)
    setInterval(() => {
      upsertStyle(buildCss());
      ensurePortalButton();
    }, 1200);

    // First run
    ensurePortalButton();
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init, { once: true });
  } else {
    init();
  }
})();
