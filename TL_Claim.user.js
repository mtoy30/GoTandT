// ==UserScript==
// @name         TL_Claim
// @namespace    https://github.com/mtoy30/GoTandT
// @version      1.0.0
// @updateURL   https://raw.githubusercontent.com/mtoy30/GoTandT/main/TL_Claim.user.js
// @downloadURL https://raw.githubusercontent.com/mtoy30/GoTandT/main/TL_Claim.user.js
// @description  Adds one button that copies Claimant + Claim + DOS (Start Date) to clipboard. Nothing else.
// @match        https://gotandt.crm.dynamics.com/*
// @author       Michael Toy
// @grant        GM_setClipboard
// ==/UserScript==

(function () {
  "use strict";

  // ----------------------------
  // Small toast message
  // ----------------------------
  function showMessage(message, isSuccess = true) {
    const popup = document.createElement("div");
    popup.textContent = message;
    Object.assign(popup.style, {
      position: "fixed",
      top: "25px",
      left: "50%",
      transform: "translate(-50%, -50%)",
      backgroundColor: isSuccess ? "#28a745" : "#dc3545",
      color: "#fff",
      padding: "10px 12px",
      borderRadius: "6px",
      zIndex: "99999",
      fontWeight: "600",
      fontSize: "14px",
      boxShadow: "0 6px 16px rgba(0,0,0,0.2)",
      maxWidth: "80%",
      textAlign: "center",
      wordBreak: "break-word",
      transition: "opacity 0.5s",
    });

    document.body.appendChild(popup);

    setTimeout(() => {
      popup.style.opacity = "0";
      setTimeout(() => popup.remove(), 450);
    }, 2200);
  }

  // ----------------------------
  // Modern button
  // ----------------------------
  function createModernButton(text, gradientStart, gradientEnd, onClick) {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.textContent = text;

    btn.style.cssText = `
      margin-left: 10px;
      padding: 8px 10px;
      background: linear-gradient(135deg, ${gradientStart}, ${gradientEnd});
      color: black;
      border: none;
      border-radius: 10px;
      cursor: pointer;
      font-weight: 700;
      font-size: 13px;
      box-shadow: 0 4px 10px rgba(0,0,0,0.12);
      transition: transform .15s ease, box-shadow .15s ease;
      white-space: nowrap;
    `;

    btn.addEventListener("mouseenter", () => {
      btn.style.transform = "scale(1.04)";
      btn.style.boxShadow = "0 8px 16px rgba(0,0,0,0.16)";
    });
    btn.addEventListener("mouseleave", () => {
      btn.style.transform = "scale(1)";
      btn.style.boxShadow = "0 4px 10px rgba(0,0,0,0.12)";
    });

    if (typeof onClick === "function") btn.addEventListener("click", onClick);

    return btn;
  }

  // ----------------------------
  // Extract: Claimant, Claim, DOS
  // ----------------------------
  function getClaimantName() {
    const el = Array.from(document.querySelectorAll('a[aria-label][href*="etn=contact"]'))
      .find(a => (a.textContent || "").trim().length > 0);
    return el ? el.textContent.trim() : "";
  }

  function getClaimNumber() {
    const el = Array.from(document.querySelectorAll('a[aria-label][href*="etn=gtt_claim"]'))
      .find(a => (a.textContent || "").trim().length > 0);
    return el ? el.textContent.trim() : "";
  }

  function getDOS() {
    // This is what your existing script uses
    const startDateInput = document.querySelector('input[aria-label="Date of Start Date"]');
    const v1 = startDateInput ? (startDateInput.value || "").trim() : "";
    if (v1) return v1;

    // Fallbacks (Dynamics labels sometimes vary)
    const alt =
      document.querySelector('input[aria-label="Start Date"]') ||
      document.querySelector('input[aria-label*="Start Date"]') ||
      document.querySelector('input[aria-label*="Date of Start"]');
    return alt ? (alt.value || "").trim() : "";
  }

  function copyNameClaimDOS() {
    // Optional guard: only run on referral pages (safe)
    if (!document.title.includes("Referral: Information:")) {
      showMessage("Must be in a referral.", false);
      return;
    }

    const claimant = getClaimantName();
    const claim = getClaimNumber();
    const dos = getDOS();

    if (!claimant || !claim) {
      showMessage("Could not find Claimant and/or Claim #.", false);
      return;
    }

    const textToCopy = `Claimant: ${claimant} - Claim: ${claim} - DOS: ${dos}`;
    GM_setClipboard(textToCopy);
    showMessage(`Copied: "${textToCopy}"`, true);
  }

  // ----------------------------
  // Button injection near search box
  // ----------------------------
  function ensureButton() {
    const searchBox = document.querySelector("#searchBoxLiveRegion");
    if (!searchBox) return false;

    // Prevent duplicates
    if (document.getElementById("dd-copy-name-claim-dos-btn")) return true;

    const btn = createModernButton("Copy Name/Claim/DOS", "#3b82f6", "#60a5fa", copyNameClaimDOS);
    btn.id = "dd-copy-name-claim-dos-btn";

    // Add next to search box
    const container = document.createElement("div");
    container.id = "dd-copy-only-container";
    container.style.display = "inline-flex";
    container.style.alignItems = "center";
    container.style.marginLeft = "10px";

    container.appendChild(btn);

    // Insert after searchBox
    searchBox.parentNode.insertBefore(container, searchBox.nextSibling);
    return true;
  }

  // Try now
  ensureButton();

  // Observe for SPA navigation / dynamic rendering
  const obs = new MutationObserver(() => {
    ensureButton();
  });

  obs.observe(document.body, { childList: true, subtree: true });
})();
