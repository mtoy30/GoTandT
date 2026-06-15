// ==UserScript==
// @name         DD_Buttons_Admin
// @namespace    https://github.com/mtoy30/GoTandT
// @version      4.2.0
// @updateURL    https://raw.githubusercontent.com/mtoy30/GoTandT/main/DD_Buttons_Admin.user.js
// @downloadURL  https://raw.githubusercontent.com/mtoy30/GoTandT/main/DD_Buttons_Admin.user.js
// @description  Custom script for Dynamics 365 CRM page with multiple button functionalities
// @match        https://gotandt.crm.dynamics.com/*
// @match        https://gttqap2.crm.dynamics.com/*
// @author        Michael Toy
// @grant        GM_setClipboard
// @grant        GM_registerMenuCommand
// ==/UserScript==

(function() {
    'use strict';

    function init() {

    let processingTimeoutId = null;


    const DD_LAST_AUTH_RATES_STORAGE_KEY = "DD_Buttons_Admin_LastAuthorizedRates";
    const DD_LAST_AUTH_RATES_BACKUP_KEY = "DD_Buttons_Admin_LastAuthorizedRates_Backup";

    function ddSafeGetGMValue(key, fallbackValue = null) {
        try {
            if (typeof GM_getValue === "function") {
                const value = GM_getValue(key, fallbackValue);
                // Tampermonkey GM_getValue is synchronous. If another manager returns a Promise,
                // fall back to browser storage so we do not break the click flow.
                if (value && typeof value.then === "function") return fallbackValue;
                return value;
            }
        } catch (e) {
            console.warn("Unable to read GM value", key, e);
        }
        return fallbackValue;
    }

    function ddSafeSetGMValue(key, value) {
        try {
            if (typeof GM_setValue === "function") {
                GM_setValue(key, value);
                return true;
            }
        } catch (e) {
            console.warn("Unable to save GM value", key, e);
        }
        return false;
    }

    function ddSaveLastAuthorizedRatesGlobal(ratePayload) {
        // Do not erase the last good rates when a temporary/read failure returns blank.
        if (!ratePayload) return false;

        window.ddLastAuthorizedRatePayload = { ...ratePayload };
        window.ddLastAuthorizedRateSavedAt = new Date().toISOString();

        const savedObject = {
            savedAt: window.ddLastAuthorizedRateSavedAt,
            url: location.href,
            title: document.title || "",
            payload: ratePayload
        };

        try {
            const savedText = JSON.stringify(savedObject);
            const backupText = JSON.stringify(ratePayload);

            // GM storage is shared between all matching Dynamics tabs/windows in Tampermonkey.
            ddSafeSetGMValue(DD_LAST_AUTH_RATES_STORAGE_KEY, savedText);
            ddSafeSetGMValue(DD_LAST_AUTH_RATES_BACKUP_KEY, backupText);

            // Keep browser storage as an extra same-tab / same-origin fallback.
            localStorage.setItem(DD_LAST_AUTH_RATES_STORAGE_KEY, savedText);
            localStorage.setItem(DD_LAST_AUTH_RATES_BACKUP_KEY, backupText);
            sessionStorage.setItem(DD_LAST_AUTH_RATES_STORAGE_KEY, savedText);
            sessionStorage.setItem(DD_LAST_AUTH_RATES_BACKUP_KEY, backupText);

            console.log("Saved last Authorized Rates", savedObject);
            return true;
        } catch (e) {
            console.warn("Unable to save last authorized rates", e);
            return false;
        }
    }

    function ddParseSavedAuthorizedRates(raw, key) {
        try {
            if (!raw) return null;
            const saved = typeof raw === "string" ? JSON.parse(raw) : raw;
            if (saved && saved.payload) return saved.payload;
            if (saved && typeof saved === "object") return saved;
        } catch (e) {
            console.warn("Unable to parse saved authorized rates", key, e);
        }
        return null;
    }

    function ddReadSavedAuthorizedRatesFromStorage(storage, key) {
        try {
            const raw = storage.getItem(key);
            return ddParseSavedAuthorizedRates(raw, key);
        } catch (e) {
            console.warn("Unable to read saved authorized rates", key, e);
        }
        return null;
    }

    function ddLoadLastAuthorizedRatesGlobal() {
        if (window.ddLastAuthorizedRatePayload) return window.ddLastAuthorizedRatePayload;

        const payload =
            ddParseSavedAuthorizedRates(ddSafeGetGMValue(DD_LAST_AUTH_RATES_STORAGE_KEY, null), DD_LAST_AUTH_RATES_STORAGE_KEY) ||
            ddParseSavedAuthorizedRates(ddSafeGetGMValue(DD_LAST_AUTH_RATES_BACKUP_KEY, null), DD_LAST_AUTH_RATES_BACKUP_KEY) ||
            ddReadSavedAuthorizedRatesFromStorage(localStorage, DD_LAST_AUTH_RATES_STORAGE_KEY) ||
            ddReadSavedAuthorizedRatesFromStorage(localStorage, DD_LAST_AUTH_RATES_BACKUP_KEY) ||
            ddReadSavedAuthorizedRatesFromStorage(sessionStorage, DD_LAST_AUTH_RATES_STORAGE_KEY) ||
            ddReadSavedAuthorizedRatesFromStorage(sessionStorage, DD_LAST_AUTH_RATES_BACKUP_KEY);

        if (!payload) return null;

        window.ddLastAuthorizedRatePayload = { ...payload };
        return window.ddLastAuthorizedRatePayload;
    }
    window.ddLastAuthorizedRatePayload = window.ddLastAuthorizedRatePayload || null;
    window.ddLastAuthorizedRateSavedAt = window.ddLastAuthorizedRateSavedAt || null;


    function ddAuthNormalizeOptionText(text) {
        return (text || "")
            .replace(/\s+/g, " ")
            .replace(/[–—]/g, "-")
            .trim()
            .toLowerCase();
    }

    function ddAuthVisibleElement(el) {
        if (!el) return false;
        const rect = el.getBoundingClientRect();
        const style = window.getComputedStyle(el);
        return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
    }

    function ddAuthClickLikeUser(el) {
        if (!el) return false;
        el.focus?.();
        el.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true, cancelable: true, pointerType: 'mouse' }));
        el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
        el.dispatchEvent(new PointerEvent('pointerup', { bubbles: true, cancelable: true, pointerType: 'mouse' }));
        el.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true }));
        el.click();
        return true;
    }

    function ddAuthClickDynamicsTabByLabel(label, callback, delay = 900) {
        const wanted = ddAuthNormalizeOptionText(label);
        const tab = Array.from(document.querySelectorAll('li[role="tab"], [role="tab"]')).find(el => {
            const aria = ddAuthNormalizeOptionText(el.getAttribute('aria-label'));
            const title = ddAuthNormalizeOptionText(el.getAttribute('title'));
            const text = ddAuthNormalizeOptionText(el.textContent);
            return aria === wanted || title === wanted || text === wanted;
        });

        if (!tab) {
            showMessage(`Could not find ${label} tab.`, false);
            if (callback) callback(false);
            return;
        }

        ddAuthClickLikeUser(tab);
        setTimeout(() => callback && callback(true), delay);
    }

    function ddAuthSetReactInputValue(input, value) {
        if (!input || value === "" || value === null || value === undefined) return false;
        const valueSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value')?.set;
        valueSetter ? valueSetter.call(input, value) : input.value = value;
        input.dispatchEvent(new Event('input', { bubbles: true }));
        input.dispatchEvent(new Event('change', { bubbles: true }));
        input.dispatchEvent(new Event('blur', { bubbles: true }));
        return true;
    }

    function ddAuthFindFieldByAriaLabel(label) {
        return document.querySelector(`input[aria-label="${CSS.escape(label)}"], textarea[aria-label="${CSS.escape(label)}"]`);
    }

    function ddAuthFindFieldNearVisibleLabel(label) {
        const labelEls = Array.from(document.querySelectorAll('label, span, div')).filter(el => (el.textContent || '').trim() === label);
        for (const labelEl of labelEls) {
            const container = labelEl.closest('[data-id], .pa-gd, .flexbox, section, div') || labelEl.parentElement;
            const field = container?.querySelector('input, textarea, [role="combobox"] input');
            if (field) return field;
        }
        return null;
    }

    function ddAuthSetFieldByLabel(label, value) {
        if (value === "" || value === null || value === undefined) return false;
        const field = ddAuthFindFieldByAriaLabel(label) || ddAuthFindFieldNearVisibleLabel(label);
        if (!field) {
            console.warn(`Authorization field not found: ${label}`);
            return false;
        }
        return ddAuthSetReactInputValue(field, value);
    }

    function ddAuthGetDropdownControlByLabel(label) {
        const safeLabel = (label || "").replace(/"/g, '\\"');
        const direct = document.querySelector(
            `[aria-label="${safeLabel}"], input[aria-label="${safeLabel}"], [title="${safeLabel}"], input[title="${safeLabel}"]`
        );
        if (direct && ddAuthVisibleElement(direct)) return direct;

        const dataIdNeedle = label.toLowerCase().replace(/[^a-z0-9]/g, "");
        const dataIdMatch = Array.from(document.querySelectorAll('[data-id], input, button, [role="combobox"]')).find(el => {
            const txt = `${el.getAttribute('data-id') || ''} ${el.getAttribute('aria-label') || ''} ${el.getAttribute('title') || ''}`.toLowerCase().replace(/[^a-z0-9]/g, "");
            return txt.includes(dataIdNeedle) && ddAuthVisibleElement(el);
        });
        if (dataIdMatch) return dataIdMatch;

        const labelEls = Array.from(document.querySelectorAll('label, span, div')).filter(el =>
            ddAuthNormalizeOptionText(el.textContent) === ddAuthNormalizeOptionText(label)
        );

        for (const labelEl of labelEls) {
            let container = labelEl;
            for (let i = 0; i < 8 && container; i++, container = container.parentElement) {
                const field = container.querySelector('input[role="combobox"], [role="combobox"], input, button[aria-haspopup="listbox"], button[aria-expanded]');
                if (field && ddAuthVisibleElement(field)) return field;
            }
        }

        return ddAuthFindFieldNearVisibleLabel(label);
    }

    function ddAuthFindOpenDropdownOption(value) {
        const wanted = ddAuthNormalizeOptionText(value);
        const optionSelectors = [
            '[role="option"]',
            '[data-testid*="option"]',
            '[data-id*="option"]',
            '[data-id*="Option"]',
            'div[id*="fluent-listbox"] [role="option"]'
        ];

        let candidates = Array.from(document.querySelectorAll(optionSelectors.join(',')))
            .filter(ddAuthVisibleElement)
            .sort((a, b) => {
                const ar = a.getBoundingClientRect();
                const br = b.getBoundingClientRect();
                return (ar.width * ar.height) - (br.width * br.height);
            });

        let found = candidates.find(el => ddAuthNormalizeOptionText(el.textContent) === wanted) ||
                    candidates.find(el => ddAuthNormalizeOptionText(el.getAttribute('aria-label')) === wanted) ||
                    candidates.find(el => ddAuthNormalizeOptionText(el.getAttribute('title')) === wanted) ||
                    candidates.find(el => ddAuthNormalizeOptionText(el.textContent).includes(wanted));
        if (found) return found;

        candidates = Array.from(document.querySelectorAll('li, button, span'))
            .filter(ddAuthVisibleElement)
            .sort((a, b) => {
                const ar = a.getBoundingClientRect();
                const br = b.getBoundingClientRect();
                return (ar.width * ar.height) - (br.width * br.height);
            });

        return candidates.find(el => ddAuthNormalizeOptionText(el.textContent) === wanted) ||
               candidates.find(el => ddAuthNormalizeOptionText(el.getAttribute('aria-label')) === wanted) ||
               candidates.find(el => ddAuthNormalizeOptionText(el.getAttribute('title')) === wanted) ||
               candidates.find(el => ddAuthNormalizeOptionText(el.textContent).includes(wanted));
    }

    function ddAuthGetTransportRateTypeControl() {
        return document.querySelector('button[data-id="gtt_authtransportratetype.fieldControl-option-set-select"][aria-label="Transport Rate Type"]') ||
               document.querySelector('[data-id="gtt_authtransportratetype.fieldControl-option-set-select"]') ||
               ddAuthGetDropdownControlByLabel('Transport Rate Type');
    }

    function ddAuthSetTransportRateTypePerMile(callback) {
        const control = ddAuthGetTransportRateTypeControl();
        if (!control || !ddAuthVisibleElement(control)) {
            console.warn('Transport Rate Type dropdown not found.');
            if (callback) callback(false);
            return false;
        }

        const current = ddAuthNormalizeOptionText(control.textContent || control.getAttribute('title') || '');
        if (current === ddAuthNormalizeOptionText('Per Mile')) {
            if (callback) callback(true);
            return true;
        }

        ddAuthClickLikeUser(control);
        let attempts = 0;
        const maxAttempts = 20;

        const trySelectPerMile = () => {
            attempts++;

            if (attempts === 3 && control.getAttribute('aria-expanded') !== 'true') {
                ddAuthClickLikeUser(control);
            }

            const option = ddAuthFindOpenDropdownOption('Per Mile');
            if (option) {
                option.scrollIntoView({ block: 'center' });
                ddAuthClickLikeUser(option);
                control.dispatchEvent(new Event('change', { bubbles: true }));
                control.dispatchEvent(new Event('blur', { bubbles: true }));
                setTimeout(() => callback && callback(true), 350);
                return;
            }

            if (attempts === 8) {
                control.dispatchEvent(new KeyboardEvent('keydown', { key: 'p', code: 'KeyP', bubbles: true }));
                control.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', code: 'Enter', bubbles: true }));
            }

            if (attempts < maxAttempts) {
                setTimeout(trySelectPerMile, 250);
            } else {
                console.warn('Could not select Transport Rate Type = Per Mile.');
                if (callback) callback(false);
            }
        };

        setTimeout(trySelectPerMile, 400);
        return true;
    }

    function ddAuthSetLookupOrDropdownByLabel(label, value, callback) {
        if (label === 'Transport Rate Type' && ddAuthNormalizeOptionText(value) === ddAuthNormalizeOptionText('Per Mile')) {
            return ddAuthSetTransportRateTypePerMile(callback);
        }

        const control = ddAuthGetDropdownControlByLabel(label);
        if (!control) {
            console.warn(`Dropdown field not found: ${label}`);
            if (callback) callback(false);
            return false;
        }

        const input = control.tagName === 'INPUT' ? control : control.querySelector('input') || control;
        ddAuthClickLikeUser(input);

        if (input.tagName === 'INPUT') {
            ddAuthSetReactInputValue(input, value);
        }

        let attempts = 0;
        const maxAttempts = 12;
        const trySelect = () => {
            attempts++;
            const option = ddAuthFindOpenDropdownOption(value);

            if (option) {
                option.scrollIntoView({ block: 'center' });
                ddAuthClickLikeUser(option);
                input.dispatchEvent(new Event('change', { bubbles: true }));
                input.dispatchEvent(new Event('blur', { bubbles: true }));
                if (callback) setTimeout(() => callback(true), 250);
                return;
            }

            if (attempts === 4) {
                input.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', code: 'ArrowDown', bubbles: true }));
                input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', code: 'Enter', bubbles: true }));
            }

            if (attempts < maxAttempts) {
                setTimeout(trySelect, 250);
            } else {
                console.warn(`Dropdown option not found: ${label} = ${value}`);
                if (callback) callback(false);
            }
        };

        setTimeout(trySelect, 350);
        return true;
    }

    function ddAuthAuthorizationPayloadHasAtLeastOneRate(ratePayload) {
        if (!ratePayload) return false;
        return Object.entries(ratePayload).some(([label, value]) => label !== "Transport Rate Type" && value !== "" && value !== null && value !== undefined);
    }

    function ddAuthApplyHigherRatesToAuthorizationFields(savedPayload) {
        const ratePayload = savedPayload || ddLoadLastAuthorizedRatesGlobal();

        if (!ddAuthAuthorizationPayloadHasAtLeastOneRate(ratePayload)) {
            showCenteredOverlayMessage("No saved rates found. Open Margin Calculator and use Request Rates, Apply & Staff, Homelink, or Boomerang first.", false, 3500);
            return;
        }

        ddAuthClickDynamicsTabByLabel("General", () => {
            setTimeout(() => {
                ddAuthSetLookupOrDropdownByLabel("Rate Approval Status", "Pending - RATE Authorization Requested", () => {
                    setTimeout(() => {
                        ddAuthClickDynamicsTabByLabel("Authorized Rates", () => {
                            let updatedCount = 0;

                            const applyTextFields = () => {
                                Object.entries(ratePayload).forEach(([label, value]) => {
                                    if (!value || label === "Transport Rate Type") return;
                                    if (ddAuthSetFieldByLabel(label, value)) updatedCount++;
                                });

                                setTimeout(() => {
                                    showCenteredOverlayMessage("Please review rates. Notate and save the referral", true, 3000);
                                    console.log("Authorized Rates payload applied from saved rates", ratePayload, "Fields updated:", updatedCount);
                                }, 900);
                            };

                            if (ratePayload["Transport Rate Type"]) {
                                ddAuthSetLookupOrDropdownByLabel("Transport Rate Type", ratePayload["Transport Rate Type"], (selected) => {
                                    if (selected) updatedCount++;
                                    setTimeout(applyTextFields, 350);
                                });
                            } else {
                                applyTextFields();
                            }
                        }, 1200);
                    }, 500);
                });
            }, 700);
        }, 900);
    }

    window.ddApplyLastAuthorizedRates = function() {
        ddAuthApplyHigherRatesToAuthorizationFields(ddLoadLastAuthorizedRatesGlobal());
    };

    const style = document.createElement('style');
    style.innerHTML = `
       input[type="radio"] {
        width: 20px;
        height: 20px;
        accent-color: #3b82f6;
       }
       #tl-interpretation-calc-box input[type="number"],
       #tl-interpretation-calc-box input[type="text"] {
        box-sizing: border-box;
       }
       #tl-interpretation-calc-box input[type="number"]::-webkit-outer-spin-button,
       #tl-interpretation-calc-box input[type="number"]::-webkit-inner-spin-button {
        -webkit-appearance: none;
        margin: 0;
       }
       #tl-interpretation-calc-box input[type="number"] {
        -moz-appearance: textfield;
        appearance: textfield;
       }
    `;
    document.head.appendChild(style);

    function createModernButton(text, gradientStart, gradientEnd, onClick) {
        const btn = document.createElement("button");
        btn.innerText = text;
        btn.style.cssText = `
            margin-top: 5px;
            margin-right: 10px;
            padding: 10px 10px;
            background: linear-gradient(135deg, ${gradientStart}, ${gradientEnd});
            color: black;
            border: none;
            border-radius: 8px;
            cursor: pointer;
            font-weight: 600;
            font-size: 14px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
            transition: all 0.2s ease-in-out;
        `;
        btn.addEventListener("mouseenter", () => {
            btn.style.transform = "scale(1.05)";
            btn.style.boxShadow = "0 6px 12px rgba(0, 0, 0, 0.15)";
        });
        btn.addEventListener("mouseleave", () => {
            btn.style.transform = "scale(1)";
            btn.style.boxShadow = "0 4px 6px rgba(0, 0, 0, 0.1)";
        });

        if (typeof onClick === "function") {
            btn.addEventListener("click", onClick);
        }

        return btn;
    }

    function showMessage(message, isSuccess = true) {
        const popup = document.createElement('div');
        popup.textContent = message;
        popup.style.position = 'fixed';
        popup.style.top = '25px';
        popup.style.left = '50%';
        popup.style.transform = 'translate(-50%, -50%)';
        popup.style.backgroundColor = isSuccess ? '#28a745' : '#dc3545';
        popup.style.color = '#fff';
        popup.style.padding = '10px';
        popup.style.borderRadius = '5px';
        popup.style.zIndex = '9999';
        popup.style.transition = 'opacity 0.5s';
        document.body.appendChild(popup);

        setTimeout(() => {
            popup.style.opacity = '0';
            setTimeout(() => popup.remove(), 500);
        }, 2000);
    }

    function showCenteredOverlayMessage(message, isSuccess = true, duration = 1200) {
        const popup = document.createElement('div');
        popup.textContent = message;
        popup.style.position = 'fixed';
        popup.style.top = '50%';
        popup.style.left = '50%';
        popup.style.transform = 'translate(-50%, -50%)';
        popup.style.background = isSuccess ? 'rgba(0,0,0,0.8)' : 'rgba(220,53,69,0.92)';
        popup.style.color = '#fff';
        popup.style.padding = '15px 25px';
        popup.style.borderRadius = '8px';
        popup.style.zIndex = '10001';
        popup.style.fontSize = '18px';
        popup.style.fontWeight = 'bold';
        popup.style.textAlign = 'center';
        popup.style.maxWidth = '80%';
        popup.style.wordWrap = 'break-word';
        document.body.appendChild(popup);
        setTimeout(() => popup.remove(), duration);
    }

    function showCalculatorBox() {
        if (!document.title.includes("Referral: Information:")) {
            showMessage("Must be in a referral", false);
            return;
        }

        const billingTab = document.querySelector('li[role="tab"][title="Billing"]');
        if (billingTab) {
            billingTab.click();
            console.log('Clicked "Billing" tab before showing calculator.');
            setTimeout(showCalculatorUI, 1000);
        } else {
            console.warn('"Billing" tab not found.');
            showCalculatorUI();
        }
    }

    function showCalculatorUI() {
        const existing = document.getElementById("calcBox");
        if (existing) existing.remove();

        const box = document.createElement("div");
        box.id = "calcBox";
        box.style.position = "fixed";
        box.style.top = "5%";
        box.style.left = "80%";
        box.style.transform = "translateX(-50%)";
        box.style.background = "#fff";
        box.style.padding = "20px";
        box.style.border = "2px solid #000";
        box.style.borderRadius = "10px";
        box.style.zIndex = "10000";
        box.style.minWidth = "520px";
        box.style.maxWidth = "520px";
        box.style.color = "black";
        box.style.height = "800px";
        box.style.overflowY = "auto";

        const headerElement = document.querySelector('[id^="formHeaderTitle"]');
        const calcHeaderText = headerElement?.textContent?.trim() || "";

        const auditNumbers = [
            "-36229-","-36969-","-35155-","-35295-","-44564-",
            "-37377-","-39425-","-35591-","-36155-","-32948-",
            "-41729-","-40778-","-35233-","-37438-","-39764-",
            "-37003-","-40043-","-34333-","-57263-","-38733-",
            "-40689-","-40860-","-33578-","-38542-","-59196-",
            "-35645-","-35017-","-45203-","-38269-","-55659-",
            "-36354-","-37546-"
        ];

        const titleStr = document.title;
        let showAuditMessage = false;
        for (const auditNum of auditNumbers) {
            if (titleStr.includes(auditNum)) {
                showAuditMessage = true;
                break;
            }
        }

        if (showAuditMessage) {
            const auditMsg = document.createElement("div");
            auditMsg.innerText = "Approval Audit see Christina";
            auditMsg.style.fontWeight = "bold";
            auditMsg.style.color = "red";
            auditMsg.style.fontSize = "22px";
            auditMsg.style.marginBottom = "10px";
            box.appendChild(auditMsg);
        }

        const closeButton = document.createElement("button");
        closeButton.innerText = "X";
        closeButton.style.position = "absolute";
        closeButton.style.top = "5px";
        closeButton.style.right = "10px";
        closeButton.style.border = "none";
        closeButton.style.background = "transparent";
        closeButton.style.color = "#000";
        closeButton.style.fontSize = "20px";
        closeButton.style.fontWeight = "bold";
        closeButton.style.cursor = "pointer";
        closeButton.onclick = () => box.remove();

        const modeLabel = document.createElement("div");
        modeLabel.innerText = "Select Rate Type:";
        modeLabel.style.marginTop = "5px";
        modeLabel.style.marginBottom = "5px";
        modeLabel.style.fontWeight = "bold";
        modeLabel.style.fontSize = "22px";

        const flatRadio = document.createElement("input");
        flatRadio.type = "radio";
        flatRadio.name = "rateType";
        flatRadio.value = "flat";
        flatRadio.id = "rateFlat";
        flatRadio.checked = true;

        const flatLabel = document.createElement("label");
        flatLabel.htmlFor = "rateFlat";
        flatLabel.innerText = "Flat Rate";
        flatLabel.style.marginRight = "20px";

        const mileRadio = document.createElement("input");
        mileRadio.type = "radio";
        mileRadio.name = "rateType";
        mileRadio.value = "mile";
        mileRadio.id = "rateMile";

        const mileLabel = document.createElement("label");
        mileLabel.htmlFor = "rateMile";
        mileLabel.innerText = "Per Mile";

        const twoColumnWrapper = document.createElement("div");
        twoColumnWrapper.style.display = "flex";
        twoColumnWrapper.style.gap = "15px";
        twoColumnWrapper.style.marginTop = "15px";

        const providerWrapper = document.createElement("div");
        providerWrapper.style.flex = "1";

        const inputLabel = document.createElement("label");
        inputLabel.innerText = "Enter Provider Rate:";
        inputLabel.style.fontWeight = "bold";
        providerWrapper.appendChild(inputLabel);

        const input = document.createElement("input");
        input.type = "number";
        input.style.width = "100%";
        input.style.marginTop = "10px";
        input.style.marginBottom = "15px";
        providerWrapper.appendChild(input);

        const providerLoadFeeWrap = document.createElement("div");
        providerLoadFeeWrap.style.display = "none";

        const providerLoadFeeLabel = document.createElement("label");
        providerLoadFeeLabel.innerText = "Enter Provider Load Fee:";
        providerLoadFeeLabel.style.fontWeight = "bold";
        providerLoadFeeWrap.appendChild(providerLoadFeeLabel);

        const providerLoadFeeInput = document.createElement("input");
        providerLoadFeeInput.type = "number";
        providerLoadFeeInput.style.width = "100%";
        providerLoadFeeInput.style.marginTop = "6px";
        providerLoadFeeInput.style.marginBottom = "4px";
        providerLoadFeeInput.value = "";
        providerLoadFeeWrap.appendChild(providerLoadFeeInput);

        providerWrapper.appendChild(providerLoadFeeWrap);

        const waitWrapper = document.createElement("div");
        waitWrapper.style.flex = "1";

        const waitTimeLabel = document.createElement("label");
        waitTimeLabel.innerText = "Enter Wait Time:";
        waitTimeLabel.style.fontWeight = "bold";
        waitWrapper.appendChild(waitTimeLabel);

        const waitTimeInput = document.createElement("input");
        waitTimeInput.type = "number";
        waitTimeInput.style.width = "100%";
        waitTimeInput.style.marginTop = "10px";
        waitTimeInput.style.marginBottom = "15px";
        waitTimeInput.value = "";
        waitWrapper.appendChild(waitTimeInput);

        twoColumnWrapper.appendChild(providerWrapper);
        twoColumnWrapper.appendChild(waitWrapper);

        const result = document.createElement("div");
        result.style.marginTop = "3px";
        result.style.fontWeight = "bold";
        result.style.whiteSpace = "pre-line";

        const targetLabel = document.createElement("div");
        targetLabel.style.marginTop = "10px";
        targetLabel.style.fontWeight = "bold";

        const higherHeader = document.createElement("div");
        higherHeader.innerText = "Higher Rates Calculator";
        higherHeader.style.marginTop = "25px";
        higherHeader.style.fontWeight = "bold";
        higherHeader.style.fontSize = "20px";

        const higherInputsWrapper = document.createElement("div");
        higherInputsWrapper.style.display = "flex";
        higherInputsWrapper.style.flexWrap = "wrap";
        higherInputsWrapper.style.columnGap = "10px";

        const higherResult = document.createElement("div");
        higherResult.style.marginTop = "10px";
        higherResult.style.fontWeight = "bold";
        higherResult.style.whiteSpace = "pre-line";

        const productInputs = {};
        let noShowPreview = null;
        let transportPreview = null;

        const productsToTrack = [
            "Transport Ambulatory",
            "Transport Wheelchair",
            "Transport Stretcher, ALS & BLS",
            "Miscellaneous Dead Miles",
            "Load Fee",
            "One Way Surcharge",
            "Weekend Holiday",
            "After Hours Fee",
            "Additional Passenger",
            "Rush Fee",
            "Wheelchair Rental",
            "Airport Pickup Fee"
        ];

function getNoShowAmountFromRow() {
    const rows = document.querySelectorAll('div[row-index]');
    const transportProducts = [
        "Transport Ambulatory",
        "Transport Wheelchair",
        "Transport Stretcher, ALS & BLS"
    ];

    const headerElement = document.querySelector('[id^="formHeaderTitle"]');
    const headerText = headerElement?.textContent?.trim() || "";

    const groupFlat = ["4474-","11525-","8814-","10837-"];
    const group15 = ["133-","202-","9616-"];
    const group16 = [
        "9617-","145-","9548-","9337-","4234-","4403-","5219-",
        "6117-","6345-","10322-","10530-","10531-","4417-","6931-"
    ];

    for (const row of rows) {
        const productCell = row.querySelector('[col-id="gtt_accountproduct"]');
        const priceCell = row.querySelector('[col-id="gtt_price"]');

        if (!productCell || !priceCell) continue;

        const productName = productCell.innerText.trim();
        if (!transportProducts.includes(productName)) continue;

        const priceText = priceCell.innerText.trim().replace(/[^0-9.-]+/g, '');
        const priceValue = parseFloat(priceText);

        if (isNaN(priceValue) || priceValue <= 0) continue;

        if (groupFlat.some(prefix => headerText.startsWith(prefix))) {
            if (productName === "Transport Ambulatory") {
                return { amount: 35, rule: "Flat rule: Ambulatory = $35" };
            }
            if (productName === "Transport Wheelchair") {
                return { amount: 60, rule: "Flat rule: Wheelchair = $60" };
            }
            return { amount: 35, rule: "Flat rule fallback = $35" };
        }

        if (group15.some(prefix => headerText.startsWith(prefix))) {
            if (
                (headerText.startsWith("202-") || headerText.startsWith("9616-")) &&
                productName === "Transport Wheelchair"
            ) {
                return {
                    amount: (priceValue * 15) + 60,
                    rule: `Wheelchair rule: gtt_price × 15 + 60 (${priceValue.toFixed(2)} × 15 + 60)`
                };
            }

            return {
                amount: priceValue * 15,
                rule: `Price rule: gtt_price × 15 (${priceValue.toFixed(2)} × 15)`
            };
        }

        if (group16.some(prefix => headerText.startsWith(prefix))) {
            return {
                amount: priceValue * 16,
                rule: `Price rule: gtt_price × 16 (${priceValue.toFixed(2)} × 16)`
            };
        }

        return {
            amount: priceValue * 18,
            rule: `Default rule: gtt_price × 18 (${priceValue.toFixed(2)} × 18)`
        };
    }

    return { amount: 0, rule: "" };
}

function getTransportPreviewAmount() {
    const ambulatoryVal = parseFloat((productInputs["Transport Ambulatory"]?.value || "").trim());
    const wheelchairVal = parseFloat((productInputs["Transport Wheelchair"]?.value || "").trim());
    const loadFeeVal = parseFloat((productInputs["Load Fee"]?.value || "").trim()) || 0;

    let enteredRate = 0;
    let transportType = "";

    if (!isNaN(ambulatoryVal) && ambulatoryVal > 0) {
        enteredRate = ambulatoryVal;
        transportType = "Ambulatory";
    } else if (!isNaN(wheelchairVal) && wheelchairVal > 0) {
        enteredRate = wheelchairVal;
        transportType = "Wheelchair";
    } else {
        return { amount: 0, rule: "" };
    }

    const headerElement = document.querySelector('[id^="formHeaderTitle"]');
    const headerText = headerElement?.textContent?.trim() || "";

    const groupFlat = ["4474-","11525-","8814-","10837-"];
    const group15 = ["133-","202-","9616-"];
    const group16 = [
        "9617-","145-","9548-","9337-","4234-","4403-","5219-",
        "6117-","6345-","10322-","10530-","10531-","4417-","6931-"
    ];

    if (groupFlat.some(prefix => headerText.startsWith(prefix))) {
        if (transportType === "Wheelchair") {
            return { amount: 60, rule: "Flat rule: Wheelchair = $60" };
        }
        return { amount: 35, rule: "Flat rule: Ambulatory = $35" };
    }

    if (group15.some(prefix => headerText.startsWith(prefix))) {
        if (
            (headerText.startsWith("202-") || headerText.startsWith("9616-")) &&
            transportType === "Wheelchair"
        ) {
            return {
                amount: (enteredRate * 15) + loadFeeVal,
                rule: `Wheelchair preview rule: entered Transport Wheelchair × 15 + Load Fee (${enteredRate.toFixed(2)} × 15 + ${loadFeeVal.toFixed(2)})`
            };
        }

        return {
            amount: enteredRate * 15,
            rule: `Entered rate × 15 (${enteredRate.toFixed(2)} × 15)`
        };
    }

    if (group16.some(prefix => headerText.startsWith(prefix))) {
        return {
            amount: enteredRate * 16,
            rule: `Entered rate × 16 (${enteredRate.toFixed(2)} × 16)`
        };
    }

    return {
        amount: enteredRate * 18,
        rule: `Entered rate × 18 (${enteredRate.toFixed(2)} × 18)`
    };
}

        const resetButton = createModernButton("Reset", "#ef4444", "#f87171");
        resetButton.onclick = () => {
            input.value = "";
            waitTimeInput.value = "";
            providerLoadFeeInput.value = "";
            providerLoadFeeWrap.style.display = "none";
            flatRadio.checked = true;
            mileRadio.checked = false;
            result.innerHTML = "";
            targetLabel.innerHTML = "";
            higherResult.innerHTML = "";
            if (noShowPreview) {
                noShowPreview.innerText = "";
                noShowPreview.title = "";
            }
            if (transportPreview) {
                transportPreview.innerText = "";
                transportPreview.title = "";
            }

            Object.values(productInputs).forEach(field => {
                field.value = "";
            });
        };

        function getResultMargin() {
            const resultText = result.innerText || result.textContent || "";
            const marginMatch = resultText.match(/Margin:\s*(-?[0-9.]+)%/);
            return marginMatch ? marginMatch[1] : null;
        }

        const lowMarginButton = createModernButton("Low Margin OK", "#3b82f6", "#60a5fa");
        lowMarginButton.onclick = () => {
            if (!input.value || input.value.trim() === "") {
                showCenteredOverlayMessage("Enter a Provider Rate first to calculate margin.", false, 2000);
                return;
            }

            // Prefer higher-rates margin if it is displayed; fall back to standard margin
            let marginDisplay = null;
            const higherResultText = higherResult.innerText || higherResult.textContent || "";
            const higherMarginMatch = higherResultText.match(/Margin:\s*(-?[0-9.]+)%/);
            if (higherMarginMatch) {
                marginDisplay = higherMarginMatch[1];
            } else {
                marginDisplay = getResultMargin();
            }

            if (!marginDisplay) {
                showCenteredOverlayMessage("No margin calculated yet. Enter a Provider Rate first.", false, 2000);
                return;
            }

            const marginValue = parseFloat(marginDisplay);
            const marginType = marginValue < 0 ? "negative margin" : "low margin";
            const baseText = `Management aware of Transportation ${marginType}. Ok to staff approx. Margin ${marginDisplay}%`;

            const doCopy = (textToCopy) => {
                navigator.clipboard.writeText(textToCopy).then(() => {
                    const copiedMsg = document.createElement("div");
                    copiedMsg.innerText = `"${textToCopy}" copied!`;
                    copiedMsg.style.position = "fixed";
                    copiedMsg.style.top = "50%";
                    copiedMsg.style.left = "50%";
                    copiedMsg.style.transform = "translate(-50%, -50%)";
                    copiedMsg.style.background = "rgba(0,0,0,0.8)";
                    copiedMsg.style.color = "#fff";
                    copiedMsg.style.padding = "15px 25px";
                    copiedMsg.style.borderRadius = "8px";
                    copiedMsg.style.zIndex = "10001";
                    copiedMsg.style.fontSize = "18px";
                    copiedMsg.style.fontWeight = "bold";
                    copiedMsg.style.textAlign = "center";
                    copiedMsg.style.maxWidth = "80%";
                    copiedMsg.style.wordWrap = "break-word";
                    document.body.appendChild(copiedMsg);
                    setTimeout(() => copiedMsg.remove(), 1500);
                });
            };

            if (marginValue < 0) {
                showReasonPrompt((reason) => {
                    const textToCopy = reason ? `${baseText} Reason: ${reason}` : baseText;
                    doCopy(textToCopy);
                });
            } else {
                doCopy(baseText);
            }
        };

        const uberLMButton = createModernButton("UBER LM", "#f59e0b", "#fbbf24");
        uberLMButton.onclick = () => {
            if (!input.value || input.value.trim() === "") {
                showCenteredOverlayMessage("Enter a Provider Rate first to calculate margin.", false, 2000);
                return;
            }
            const resultText = result.innerText || result.textContent || "";
            const paidMatch = resultText.match(/Total Paid:\s*\$([0-9.,]+)/);
            const marginDisplay = getResultMargin();
            if (!paidMatch || !marginDisplay) {
                showCenteredOverlayMessage("No margin calculated yet. Enter a Provider Rate first.", false, 2000);
                return;
            }
            const paidDisplay = paidMatch[1];
            const textToCopy = `Management aware of Transportation low margin. Ok to staff with UBER based on round trip $${paidDisplay} approx margin ${marginDisplay}%`;
            navigator.clipboard.writeText(textToCopy).then(() => {
                const copiedMsg = document.createElement("div");
                copiedMsg.innerText = `"${textToCopy}" copied!`;
                copiedMsg.style.position = "fixed";
                copiedMsg.style.top = "50%";
                copiedMsg.style.left = "50%";
                copiedMsg.style.transform = "translate(-50%, -50%)";
                copiedMsg.style.background = "rgba(0,0,0,0.8)";
                copiedMsg.style.color = "#fff";
                copiedMsg.style.padding = "15px 25px";
                copiedMsg.style.borderRadius = "8px";
                copiedMsg.style.zIndex = "10001";
                copiedMsg.style.fontSize = "18px";
                copiedMsg.style.fontWeight = "bold";
                copiedMsg.style.textAlign = "center";
                copiedMsg.style.maxWidth = "80%";
                copiedMsg.style.wordWrap = "break-word";
                document.body.appendChild(copiedMsg);
                setTimeout(() => copiedMsg.remove(), 1500);
            });
        };

        const waittimeButton = createModernButton("Wait Time Request", "#3b82f6", "#60a5fa");
        waittimeButton.onclick = () => {
            const textToCopy = "Ok to request wait time";
            navigator.clipboard.writeText(textToCopy).then(() => {
                const copiedMsg = document.createElement("div");
                copiedMsg.innerText = `"${textToCopy}" copied!`;
                copiedMsg.style.position = "fixed";
                copiedMsg.style.top = "50%";
                copiedMsg.style.left = "50%";
                copiedMsg.style.transform = "translate(-50%, -50%)";
                copiedMsg.style.background = "rgba(0,0,0,0.8)";
                copiedMsg.style.color = "#fff";
                copiedMsg.style.padding = "15px 25px";
                copiedMsg.style.borderRadius = "8px";
                copiedMsg.style.zIndex = "10001";
                copiedMsg.style.fontSize = "18px";
                copiedMsg.style.fontWeight = "bold";
                copiedMsg.style.textAlign = "center";
                copiedMsg.style.maxWidth = "80%";
                copiedMsg.style.wordWrap = "break-word";
                document.body.appendChild(copiedMsg);
                setTimeout(() => copiedMsg.remove(), 1000);
            });
        };

        const WaitStaffButton = createModernButton("Wait Time Staff", "#3b82f6", "#60a5fa");
        WaitStaffButton.onclick = () => {
            const textToCopy = "Ok to apply wait time and include in staffing email due to mileage.";
            navigator.clipboard.writeText(textToCopy).then(() => {
                const copiedMsg = document.createElement("div");
                copiedMsg.innerText = `"${textToCopy}" copied!`;
                copiedMsg.style.position = "fixed";
                copiedMsg.style.top = "50%";
                copiedMsg.style.left = "50%";
                copiedMsg.style.transform = "translate(-50%, -50%)";
                copiedMsg.style.background = "rgba(0,0,0,0.8)";
                copiedMsg.style.color = "#fff";
                copiedMsg.style.padding = "15px 25px";
                copiedMsg.style.borderRadius = "8px";
                copiedMsg.style.zIndex = "10001";
                copiedMsg.style.fontSize = "18px";
                copiedMsg.style.fontWeight = "bold";
                copiedMsg.style.textAlign = "center";
                copiedMsg.style.maxWidth = "80%";
                copiedMsg.style.wordWrap = "break-word";
                document.body.appendChild(copiedMsg);
                setTimeout(() => copiedMsg.remove(), 1000);
            });
        };

        function buildPartsString(productInputs, quantities, miles, loadFeeQuantity) {
            if (productInputs["One Way Surcharge"] && quantities["One Way Surcharge"] !== undefined) {
                productInputs["One Way Surcharge"].value = quantities["One Way Surcharge"];
            }

            const parts = [];

            Object.entries(productInputs).forEach(([label, input]) => {
                const value = input.value.trim();
                if (value !== "") {
                    let normalizedLabel =
                        label === "Miscellaneous Dead Miles" ? "Dead Miles" :
                        label === "No Show" ? "No Show/Late Cancel" :
                        label;

                    if (["Transport Ambulatory", "Transport Wheelchair", "Transport Stretcher, ALS & BLS"].includes(label)) {
                        if (value.toLowerCase() === "contract rates") {
                            parts.push(`Contract rates/mile`);
                        } else if (!isNaN(parseFloat(value))) {
                            parts.push(`$${parseFloat(value).toFixed(2)}/mile`);
                        } else {
                            parts.push(`${value} ${normalizedLabel}`);
                        }
                    } else if (label === "Load Fee") {
                        if (!isNaN(parseFloat(value))) {
                            const fee = `$${parseFloat(value).toFixed(2)} Load Fee`;
                            const withQty = loadFeeQuantity ? `${fee} x ${loadFeeQuantity}` : fee;
                            parts.push(withQty);
                        } else {
                            parts.push(`${value} Load Fee`);
                        }
                    } else if (label === "Miscellaneous Dead Miles") {
                        if (!isNaN(parseFloat(value))) {
                            parts.push(`${parseFloat(value)} ${normalizedLabel}`);
                        } else {
                            parts.push(`${value} ${normalizedLabel}`);
                        }
                    } else if (label === "One Way Surcharge") {
                        if (!isNaN(parseFloat(value))) {
                            parts.push(`${parseFloat(value)} mile One Way Surcharge`);
                        } else {
                            parts.push(`${value} One Way Surcharge`);
                        }
                    } else {
                        if (value.toLowerCase() === "contract" || value.toLowerCase().includes("contract")) {
                            parts.push(`contract ${normalizedLabel}`);
                        } else if (!isNaN(parseFloat(value))) {
                            parts.push(`$${parseFloat(value).toFixed(2)} ${normalizedLabel}`);
                        } else {
                            parts.push(`${value} ${normalizedLabel}`);
                        }
                    }
                }
            });

            return parts.join(", ");
        }



        function getHigherRateInputValue(label) {
            return (productInputs[label]?.value || "").trim();
        }

        function getContractPriceForProduct(productName) {
            const rows = document.querySelectorAll('div[row-index], [role="row"]');
            for (const row of rows) {
                const productCell = row.querySelector('[col-id="gtt_accountproduct"]');
                const priceCell = row.querySelector('[col-id="gtt_price"]');
                if (!productCell || !priceCell) continue;
                if ((productCell.innerText || "").trim() !== productName) continue;
                const priceValue = parseFloat((priceCell.innerText || "").replace(/[^0-9.-]+/g, ''));
                if (!isNaN(priceValue) && priceValue > 0) return priceValue.toFixed(2);
            }
            return "";
        }

        function normalizeHigherRateValue(label) {
            const raw = getHigherRateInputValue(label);
            if (!raw) return "";
            if (/contract/i.test(raw)) {
                if (label === "No Show") {
                    const previewText = (noShowPreview?.innerText || "").replace(/[^0-9.-]+/g, '');
                    if (previewText) return parseFloat(previewText).toFixed(2);
                }
                return getContractPriceForProduct(label);
            }
            const num = parseFloat(raw.replace(/[^0-9.-]+/g, ''));
            return isNaN(num) ? raw : num.toFixed(2);
        }

        function normalizeHigherRateValueForAuth(label) {
            const raw = getHigherRateInputValue(label);
            if (!raw) return "";

            // For Authorized Rates, contract Wait Time should be skipped.
            // No Show / Cancel Fee still need the calculated green preview amount.
            if (/contract/i.test(raw) && label === "Wait Time") return "";

            return normalizeHigherRateValue(label);
        }

        function isContractTransportRateForAuth() {
            const transportLabels = [
                "Transport Ambulatory",
                "Transport Wheelchair",
                "Transport Stretcher, ALS & BLS"
            ];

            return transportLabels.some(label => /contract/i.test(getHigherRateInputValue(label)));
        }

        function getActiveTransportRateForAuth() {
            const transportLabels = [
                "Transport Ambulatory",
                "Transport Wheelchair",
                "Transport Stretcher, ALS & BLS"
            ];

            for (const label of transportLabels) {
                const value = normalizeHigherRateValue(label);
                if (value) return value;
            }
            return "";
        }

        function clickDynamicsTabByLabel(label, callback, delay = 900) {
            const tab = Array.from(document.querySelectorAll('li[role="tab"], [role="tab"]')).find(el => {
                const aria = (el.getAttribute('aria-label') || '').trim();
                const title = (el.getAttribute('title') || '').trim();
                const text = (el.textContent || '').trim();
                return aria === label || title === label || text === label;
            });

            if (!tab) {
                showMessage(`Could not find ${label} tab.`, false);
                if (callback) callback(false);
                return;
            }

            tab.click();
            setTimeout(() => callback && callback(true), delay);
        }

        function setReactInputValue(input, value) {
            if (!input || value === "" || value === null || value === undefined) return false;
            const valueSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value')?.set;
            valueSetter ? valueSetter.call(input, value) : input.value = value;
            input.dispatchEvent(new Event('input', { bubbles: true }));
            input.dispatchEvent(new Event('change', { bubbles: true }));
            input.dispatchEvent(new Event('blur', { bubbles: true }));
            return true;
        }

        function findFieldByAriaLabel(label) {
            return document.querySelector(`input[aria-label="${CSS.escape(label)}"], textarea[aria-label="${CSS.escape(label)}"]`);
        }

        function findFieldNearVisibleLabel(label) {
            const labelEls = Array.from(document.querySelectorAll('label, span, div')).filter(el => (el.textContent || '').trim() === label);
            for (const labelEl of labelEls) {
                const container = labelEl.closest('[data-id], .pa-gd, .flexbox, section, div') || labelEl.parentElement;
                const field = container?.querySelector('input, textarea, [role="combobox"] input');
                if (field) return field;
            }
            return null;
        }

        function setFieldByLabel(label, value) {
            if (value === "" || value === null || value === undefined) return false;
            const field = findFieldByAriaLabel(label) || findFieldNearVisibleLabel(label);
            if (!field) {
                console.warn(`Authorization field not found: ${label}`);
                return false;
            }
            return setReactInputValue(field, value);
        }

        function normalizeOptionText(text) {
            return (text || "")
                .replace(/\s+/g, " ")
                .replace(/[–—]/g, "-")
                .trim()
                .toLowerCase();
        }

        function visibleElement(el) {
            if (!el) return false;
            const rect = el.getBoundingClientRect();
            const style = window.getComputedStyle(el);
            return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
        }

        function getDropdownControlByLabel(label) {
            const safeLabel = (label || "").replace(/"/g, '\\"');
            const direct = document.querySelector(
                `[aria-label="${safeLabel}"], input[aria-label="${safeLabel}"], [title="${safeLabel}"], input[title="${safeLabel}"]`
            );
            if (direct && visibleElement(direct)) return direct;

            const dataIdNeedle = label.toLowerCase().replace(/[^a-z0-9]/g, "");
            const dataIdMatch = Array.from(document.querySelectorAll('[data-id], input, button, [role="combobox"]')).find(el => {
                const txt = `${el.getAttribute('data-id') || ''} ${el.getAttribute('aria-label') || ''} ${el.getAttribute('title') || ''}`.toLowerCase().replace(/[^a-z0-9]/g, "");
                return txt.includes(dataIdNeedle) && visibleElement(el);
            });
            if (dataIdMatch) return dataIdMatch;

            const labelEls = Array.from(document.querySelectorAll('label, span, div')).filter(el =>
                normalizeOptionText(el.textContent) === normalizeOptionText(label)
            );

            for (const labelEl of labelEls) {
                let container = labelEl;
                for (let i = 0; i < 8 && container; i++, container = container.parentElement) {
                    const field = container.querySelector('input[role="combobox"], [role="combobox"], input, button[aria-haspopup="listbox"], button[aria-expanded]');
                    if (field && visibleElement(field)) return field;
                }
            }

            return findFieldNearVisibleLabel(label);
        }

        function findOpenDropdownOption(value) {
            const wanted = normalizeOptionText(value);

            // Prefer real Fluent/Dynamics option rows first. Avoid broad parent DIVs that can contain
            // the option text but are not clickable selections.
            const optionSelectors = [
                '[role="option"]',
                '[data-testid*="option"]',
                '[data-id*="option"]',
                '[data-id*="Option"]',
                'div[id*="fluent-listbox"] [role="option"]'
            ];

            let candidates = Array.from(document.querySelectorAll(optionSelectors.join(',')))
                .filter(visibleElement)
                .sort((a, b) => {
                    const ar = a.getBoundingClientRect();
                    const br = b.getBoundingClientRect();
                    return (ar.width * ar.height) - (br.width * br.height);
                });

            let found = candidates.find(el => normalizeOptionText(el.textContent) === wanted) ||
                        candidates.find(el => normalizeOptionText(el.getAttribute('aria-label')) === wanted) ||
                        candidates.find(el => normalizeOptionText(el.getAttribute('title')) === wanted) ||
                        candidates.find(el => normalizeOptionText(el.textContent).includes(wanted));
            if (found) return found;

            // Fallback for Dynamics portals that render list items without role="option".
            candidates = Array.from(document.querySelectorAll('li, button, span'))
                .filter(visibleElement)
                .sort((a, b) => {
                    const ar = a.getBoundingClientRect();
                    const br = b.getBoundingClientRect();
                    return (ar.width * ar.height) - (br.width * br.height);
                });

            return candidates.find(el => normalizeOptionText(el.textContent) === wanted) ||
                   candidates.find(el => normalizeOptionText(el.getAttribute('aria-label')) === wanted) ||
                   candidates.find(el => normalizeOptionText(el.getAttribute('title')) === wanted) ||
                   candidates.find(el => normalizeOptionText(el.textContent).includes(wanted));
        }

        function clickLikeUser(el) {
            if (!el) return false;
            el.focus?.();
            el.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true, cancelable: true, pointerType: 'mouse' }));
            el.dispatchEvent(new MouseEvent('mousedown', { bubbles: true, cancelable: true }));
            el.dispatchEvent(new PointerEvent('pointerup', { bubbles: true, cancelable: true, pointerType: 'mouse' }));
            el.dispatchEvent(new MouseEvent('mouseup', { bubbles: true, cancelable: true }));
            el.click();
            return true;
        }

        function getTransportRateTypeControl() {
            return document.querySelector('button[data-id="gtt_authtransportratetype.fieldControl-option-set-select"][aria-label="Transport Rate Type"]') ||
                   document.querySelector('[data-id="gtt_authtransportratetype.fieldControl-option-set-select"]') ||
                   getDropdownControlByLabel('Transport Rate Type');
        }

        function setTransportRateTypePerMile(callback) {
            const control = getTransportRateTypeControl();
            if (!control || !visibleElement(control)) {
                console.warn('Transport Rate Type dropdown not found.');
                if (callback) callback(false);
                return false;
            }

            const current = normalizeOptionText(control.textContent || control.getAttribute('title') || '');
            if (current === normalizeOptionText('Per Mile')) {
                if (callback) callback(true);
                return true;
            }

            clickLikeUser(control);

            let attempts = 0;
            const maxAttempts = 20;

            const trySelectPerMile = () => {
                attempts++;

                // If it did not open, click again using the exact combobox button.
                if (attempts === 3 && control.getAttribute('aria-expanded') !== 'true') {
                    clickLikeUser(control);
                }

                const option = findOpenDropdownOption('Per Mile');
                if (option) {
                    option.scrollIntoView({ block: 'center' });
                    clickLikeUser(option);
                    control.dispatchEvent(new Event('change', { bubbles: true }));
                    control.dispatchEvent(new Event('blur', { bubbles: true }));

                    setTimeout(() => {
                        const after = normalizeOptionText(control.textContent || control.getAttribute('title') || '');
                        if (callback) callback(after.includes('per mile') || true);
                    }, 350);
                    return;
                }

                // Last-resort keyboard path for Fluent UI comboboxes.
                if (attempts === 8) {
                    control.dispatchEvent(new KeyboardEvent('keydown', { key: 'p', code: 'KeyP', bubbles: true }));
                    control.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', code: 'Enter', bubbles: true }));
                }

                if (attempts < maxAttempts) {
                    setTimeout(trySelectPerMile, 250);
                } else {
                    console.warn('Could not select Transport Rate Type = Per Mile.');
                    if (callback) callback(false);
                }
            };

            setTimeout(trySelectPerMile, 400);
            return true;
        }

        function setLookupOrDropdownByLabel(label, value, callback) {
            if (label === 'Transport Rate Type' && normalizeOptionText(value) === normalizeOptionText('Per Mile')) {
                return setTransportRateTypePerMile(callback);
            }

            const control = getDropdownControlByLabel(label);
            if (!control) {
                console.warn(`Dropdown field not found: ${label}`);
                if (callback) callback(false);
                return false;
            }

            const input = control.tagName === 'INPUT' ? control : control.querySelector('input') || control;
            clickLikeUser(input);

            // First try typing the exact value. Dynamics choice controls often filter options this way.
            if (input.tagName === 'INPUT') {
                setReactInputValue(input, value);
            }

            let attempts = 0;
            const maxAttempts = 12;

            const trySelect = () => {
                attempts++;
                const option = findOpenDropdownOption(value);

                if (option) {
                    option.scrollIntoView({ block: 'center' });
                    clickLikeUser(option);
                    input.dispatchEvent(new Event('change', { bubbles: true }));
                    input.dispatchEvent(new Event('blur', { bubbles: true }));
                    if (callback) setTimeout(() => callback(true), 250);
                    return;
                }

                if (attempts === 4) {
                    input.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', code: 'ArrowDown', bubbles: true }));
                    input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', code: 'Enter', bubbles: true }));
                }

                if (attempts < maxAttempts) {
                    setTimeout(trySelect, 250);
                } else {
                    console.warn(`Dropdown option not found: ${label} = ${value}`);
                    if (callback) callback(false);
                }
            };

            setTimeout(trySelect, 350);
            return true;
        }

        function buildAuthorizationRatePayload() {
            const isContractTransport = isContractTransportRateForAuth();
            const transportRate = isContractTransport ? "" : getActiveTransportRateForAuth();
            const noShowValue = normalizeHigherRateValueForAuth("No Show");

            return {
                // If transport is Contract Rates, do not touch Transport Rate Type or Transport Rate.
                "Transport Rate Type": (!isContractTransport && transportRate) ? "Per Mile" : "",
                "Transport Rate": (!isContractTransport && transportRate) ? transportRate : "",

                // Contract Wait Time should not be entered on Authorized Rates.
                "Wait Time Fee": normalizeHigherRateValueForAuth("Wait Time"),

                // Contract No Show still uses the green preview amount and also fills Cancel Fee.
                "No Show": noShowValue,
                "Load Fee": normalizeHigherRateValueForAuth("Load Fee"),
                "Cancel Fee": noShowValue,
                "Dead Miles": normalizeHigherRateValueForAuth("Miscellaneous Dead Miles"),
                "Tolls": normalizeHigherRateValueForAuth("Tolls")
            };
        }

        function authorizationPayloadHasAtLeastOneRate(ratePayload) {
            return Object.entries(ratePayload).some(([label, value]) => label !== "Transport Rate Type" && value !== "" && value !== null && value !== undefined);
        }

        function saveLastAuthorizedRatesFromCalculator() {
            const ratePayload = buildAuthorizationRatePayload();

            if (!authorizationPayloadHasAtLeastOneRate(ratePayload)) {
                console.warn("No authorized rates were found to save. Keeping any previously saved rates.", ratePayload);
                return false;
            }

            ddSaveLastAuthorizedRatesGlobal(ratePayload);
            console.log("Saved Authorized Rates payload", window.ddLastAuthorizedRatePayload);
            return true;
        }

        function applyHigherRatesToAuthorizationFields(savedPayload) {
            const ratePayload = savedPayload || buildAuthorizationRatePayload();

            if (!authorizationPayloadHasAtLeastOneRate(ratePayload)) {
                showCenteredOverlayMessage("Enter at least one higher rate before applying rates.", false, 2500);
                return;
            }

            clickDynamicsTabByLabel("General", () => {
                setTimeout(() => {
                    setLookupOrDropdownByLabel("Rate Approval Status", "Pending - RATE Authorization Requested", () => {
                        setTimeout(() => {
                            clickDynamicsTabByLabel("Authorized Rates", () => {
                                let updatedCount = 0;

                                const applyTextFields = () => {
                                    Object.entries(ratePayload).forEach(([label, value]) => {
                                        if (!value || label === "Transport Rate Type") return;
                                        if (setFieldByLabel(label, value)) updatedCount++;
                                    });

                                    setTimeout(() => {
                                        showCenteredOverlayMessage("Please proceed with notating the referral", true, 3000);
                                        console.log("Authorized Rates payload applied", ratePayload, "Fields updated:", updatedCount);
                                    }, 900);
                                };

                                if (ratePayload["Transport Rate Type"]) {
                                    setLookupOrDropdownByLabel("Transport Rate Type", ratePayload["Transport Rate Type"], (selected) => {
                                        if (selected) updatedCount++;
                                        setTimeout(applyTextFields, 350);
                                    });
                                } else {
                                    applyTextFields();
                                }
                            }, 1200);
                        }, 500);
                    });
                }, 700);
            }, 900);
        }

        window.ddApplyLastAuthorizedRates = function() {
            const savedPayload = ddLoadLastAuthorizedRatesGlobal();
            if (!savedPayload || !authorizationPayloadHasAtLeastOneRate(savedPayload)) {
                showCenteredOverlayMessage("No saved rates found. Open Margin Calculator and use Request Rates, Apply & Staff, Homelink, or Boomerang first.", false, 3500);
                return;
            }
            applyHigherRatesToAuthorizationFields(savedPayload);
        };

        const applySavedRatesButton = createModernButton("Apply Rates", "#16a34a", "#86efac");
        applySavedRatesButton.onclick = () => {
            calculateMargin();
            setTimeout(() => {
                if (saveLastAuthorizedRatesFromCalculator()) {
                    applyHigherRatesToAuthorizationFields(window.ddLastAuthorizedRatePayload);
                }
            }, 250);
        };

        const requestRatesButton = createModernButton("Request Rates", "#22c55e", "#4ade80");
        requestRatesButton.onclick = () => {
            let miles = 0;
            let loadFeeQuantity = 0;

            const rows = document.querySelectorAll('[role="row"]');
            rows.forEach(row => {
                const accountProductCell = row.querySelector('[col-id="gtt_accountproduct"]');
                const quantityCell = row.querySelector('[col-id="gtt_quantity"]');

                if (accountProductCell && quantityCell) {
                    const accountProductText = accountProductCell.innerText.trim();
                    const quantityText = quantityCell.innerText.trim();
                    const quantity = parseFloat(quantityText);

                    if (accountProductText.includes("Transport") && !isNaN(quantity)) miles = quantity;
                    if (accountProductText.includes("Load Fee") && !isNaN(quantity)) loadFeeQuantity = quantity;
                }
            });

            const finalParts = buildPartsString(productInputs, {}, miles, loadFeeQuantity);
            const finalText = "Request rates " + finalParts;

            calculateMargin();
            saveLastAuthorizedRatesFromCalculator();
            setTimeout(saveLastAuthorizedRatesFromCalculator, 600);

            navigator.clipboard.writeText(finalText).then(() => {
                const copiedMsg = document.createElement("div");
                copiedMsg.innerText = `"${finalText}" copied!`;
                copiedMsg.style.position = "fixed";
                copiedMsg.style.top = "50%";
                copiedMsg.style.left = "50%";
                copiedMsg.style.transform = "translate(-50%, -50%)";
                copiedMsg.style.background = "rgba(0,0,0,0.8)";
                copiedMsg.style.color = "#fff";
                copiedMsg.style.padding = "15px 25px";
                copiedMsg.style.borderRadius = "8px";
                copiedMsg.style.zIndex = "10001";
                copiedMsg.style.fontSize = "18px";
                copiedMsg.style.fontWeight = "bold";
                copiedMsg.style.textAlign = "center";
                copiedMsg.style.maxWidth = "80%";
                copiedMsg.style.wordWrap = "break-word";
                document.body.appendChild(copiedMsg);
                setTimeout(() => copiedMsg.remove(), 1500);
                applyHigherRatesToAuthorizationFields();
            });
        };

        const applyRatesButton = createModernButton("Apply & Staff", "#a855f7", "#c084fc");
        applyRatesButton.onclick = () => {
            let miles = 0;
            let loadFeeQuantity = 0;

            const rows = document.querySelectorAll('[role="row"]');
            rows.forEach(row => {
                const accountProductCell = row.querySelector('[col-id="gtt_accountproduct"]');
                const quantityCell = row.querySelector('[col-id="gtt_quantity"]');

                if (accountProductCell && quantityCell) {
                    const accountProductText = accountProductCell.innerText.trim();
                    const quantityText = quantityCell.innerText.trim();
                    const quantity = parseFloat(quantityText);

                    if (accountProductText.includes("Transport") && !isNaN(quantity)) miles = quantity;
                    if (accountProductText.includes("Load Fee") && !isNaN(quantity)) loadFeeQuantity = quantity;
                }
            });

            const finalParts = buildPartsString(productInputs, {}, miles, loadFeeQuantity);
            const finalText = "Apply rates " + finalParts + " // Advise in Staffing email";

            saveLastAuthorizedRatesFromCalculator();

            navigator.clipboard.writeText(finalText).then(() => {
                const copiedMsg = document.createElement("div");
                copiedMsg.innerText = `"${finalText}" copied!`;
                copiedMsg.style.position = "fixed";
                copiedMsg.style.top = "50%";
                copiedMsg.style.left = "50%";
                copiedMsg.style.transform = "translate(-50%, -50%)";
                copiedMsg.style.background = "rgba(0,0,0,0.8)";
                copiedMsg.style.color = "#fff";
                copiedMsg.style.padding = "15px 25px";
                copiedMsg.style.borderRadius = "8px";
                copiedMsg.style.zIndex = "10001";
                copiedMsg.style.fontSize = "18px";
                copiedMsg.style.fontWeight = "bold";
                copiedMsg.style.textAlign = "center";
                copiedMsg.style.maxWidth = "80%";
                copiedMsg.style.wordWrap = "break-word";
                document.body.appendChild(copiedMsg);
                setTimeout(() => copiedMsg.remove(), 1500);
                applyHigherRatesToAuthorizationFields();
            });
        };

        const homelinkButton = createModernButton("Request Homelink Rates", "#a855f7", "#c084fc");
        homelinkButton.onclick = () => {
            let higherTotal = 0;
            let alreadyCounted = new Set();

            const quantities = {};
            const rows = document.querySelectorAll('[role="row"]');
            rows.forEach(row => {
                const productCell = row.querySelector('[col-id="gtt_accountproduct"]');
                const qtyCell = row.querySelector('[col-id="gtt_quantity"]');

                if (productCell && qtyCell) {
                    const product = productCell.innerText.trim();
                    const qty = parseFloat(qtyCell.innerText.trim());
                    if (!isNaN(qty)) quantities[product] = qty;
                }
            });

            const foundProducts = Object.keys(productInputs);

            const activeTransportRate = parseFloat(
                productInputs["Transport Ambulatory"]?.value ||
                productInputs["Transport Wheelchair"]?.value ||
                productInputs["Transport Stretcher, ALS & BLS"]?.value
            );

            if (!activeTransportRate || isNaN(activeTransportRate)) {
                alert("Please enter a valid Transport rate.");
                return;
            }

            foundProducts.forEach(product => {
                if (product === "Wait Time" || product === "No Show") return;

                if (["Rush Fee", "Tolls", "Other", "Assistance Fee", "Passenger Fee", "Miscellaneous Dead Miles"].includes(product)) {
                    if (alreadyCounted.has(product)) return;
                    alreadyCounted.add(product);
                }

                const enteredValueRaw = productInputs[product]?.value;
                const enteredValue = parseFloat(enteredValueRaw);
                const qty = quantities[product] || 0;

                if (!isNaN(enteredValue)) {
                    if (product === "Miscellaneous Dead Miles" || product === "One Way Surcharge") {
                        higherTotal += enteredValue * (activeTransportRate / 2);
                    } else if (["Tolls", "Other", "Assistance Fee", "Passenger Fee", "Rush Fee"].includes(product)) {
                        higherTotal += enteredValue;
                    } else {
                        higherTotal += enteredValue * qty;
                    }
                }
            });

            let waitTimeText = "";
            let noShowText = "";
            if (productInputs["Wait Time"]) {
                const waitVal = productInputs["Wait Time"].value.trim();
                waitTimeText = waitVal.toLowerCase() === "contract rates" ? "Contract" : waitVal;
            }
            if (productInputs["No Show"]) {
                const noShowVal = productInputs["No Show"].value.trim();
                noShowText = noShowVal.toLowerCase() === "contract rates" ? "Contract" : noShowVal;
            }

            const partsString = buildPartsString(productInputs, quantities, 0, quantities["Load Fee"]);
            const goatString = `**Enter in Goat as ${partsString}`;

            let extras = [];
            if (waitTimeText) {
                const waitDisplay = isNaN(waitTimeText) ? waitTimeText : `$${waitTimeText}/hour`;
                extras.push(`${waitDisplay} wait time in addition to flat rate`);
            }
            if (noShowText) {
                const noShowDisplay = isNaN(noShowText) ? noShowText : `$${noShowText}`;
                extras.push(`${noShowDisplay} No Show/Late Cancel`);
            }

            const extrasText = extras.length > 0 ? `. ${extras.join(", ")}` : "";
            const homelinkText = `Request Flat Rate of $${higherTotal.toFixed(2)}${extrasText}. ${goatString}`;

            saveLastAuthorizedRatesFromCalculator();

            navigator.clipboard.writeText(homelinkText).then(() => {
                const copiedMsg = document.createElement("div");
                copiedMsg.innerText = `"${homelinkText}" copied!`;
                copiedMsg.style.position = "fixed";
                copiedMsg.style.top = "50%";
                copiedMsg.style.left = "50%";
                copiedMsg.style.transform = "translate(-50%, -50%)";
                copiedMsg.style.background = "rgba(0,0,0,0.8)";
                copiedMsg.style.color = "#fff";
                copiedMsg.style.padding = "15px 25px";
                copiedMsg.style.borderRadius = "8px";
                copiedMsg.style.zIndex = "10001";
                copiedMsg.style.fontSize = "18px";
                copiedMsg.style.fontWeight = "bold";
                copiedMsg.style.textAlign = "center";
                copiedMsg.style.maxWidth = "80%";
                copiedMsg.style.wordWrap = "break-word";
                document.body.appendChild(copiedMsg);
                setTimeout(() => copiedMsg.remove(), 1500);
                applyHigherRatesToAuthorizationFields();
            });
        };

        const boomerangButton = createModernButton("Boomerang Request & Staff", "#f97316", "#fb923c");
        boomerangButton.onclick = () => {
            const titleHeader = document.querySelector('[id^="formHeaderTitle"]');
            const titleText = titleHeader ? titleHeader.innerText.trim() : "";
            const isHomeLink = titleText.startsWith("212-");

            let miles = 0;
            let loadFeeQuantity = 0;
            const quantities = {};

            const rows = document.querySelectorAll('[role="row"]');
            rows.forEach(row => {
                const productCell = row.querySelector('[col-id="gtt_accountproduct"]');
                const qtyCell = row.querySelector('[col-id="gtt_quantity"]');

                if (productCell && qtyCell) {
                    const productText = productCell.innerText.trim();
                    const qty = parseFloat(qtyCell.innerText.trim());

                    if (productText && !isNaN(qty)) {
                        quantities[productText] = qty;
                        if (productText.includes("Transport")) miles = qty;
                        if (productText.includes("Load Fee")) loadFeeQuantity = qty;
                    }
                }
            });

            let flatRateValue = null;
            let waitTimeText = "";
            let noShowText = "";
            const partsString = buildPartsString(productInputs, quantities, miles, loadFeeQuantity);

            Object.entries(productInputs).forEach(([label, input]) => {
                let value = input.value.trim();

                if (["Transport Ambulatory", "Transport Wheelchair", "Transport Stretcher, ALS & BLS"].includes(label)) {
                    value = value.toLowerCase() === "contract rates" ? "Contract rates per mile" : value;
                }

                if (label === "Wait Time" && value !== "") {
                    waitTimeText = value.toLowerCase() === "contract rates" ? "Contract" : value;
                }

                if (label === "No Show" && value !== "") {
                    noShowText = value.toLowerCase() === "contract rates" ? "Contract" : value;
                }

                if (["Transport Ambulatory", "Transport Wheelchair", "Transport Stretcher, ALS & BLS"].includes(label)) {
                    flatRateValue = value === "Contract rates per mile" ? value : parseFloat(value);
                }
            });

            if ((!flatRateValue || isNaN(flatRateValue)) && flatRateValue !== "Contract rates per mile") {
                alert("Please enter a valid Transport rate.");
                return;
            }

            let higherTotal = 0;
            const alreadyCounted = new Set();

            Object.keys(productInputs).forEach(product => {
                if (product === "Wait Time" || product === "No Show") return;

                const enteredValue = parseFloat(productInputs[product]?.value);
                const qty = quantities[product] || 0;

                if (!isNaN(enteredValue)) {
                    if (product === "Miscellaneous Dead Miles") {
                        higherTotal += enteredValue * (flatRateValue / 2);
                    } else if (["Tolls", "Other", "Assistance Fee", "Passenger Fee", "Rush Fee"].includes(product)) {
                        if (!alreadyCounted.has(product)) {
                            alreadyCounted.add(product);
                            higherTotal += enteredValue;
                        }
                    } else {
                        higherTotal += enteredValue * qty;
                    }
                }
            });

            const flatTotal = higherTotal.toFixed(2);

            let extras = [];
            if (waitTimeText) {
                const waitDisplay = isNaN(waitTimeText) ? waitTimeText : `$${waitTimeText}/hour`;
                extras.push(`${waitDisplay} wait time in addition to flat rate`);
            }
            if (noShowText) {
                const noShowDisplay = isNaN(noShowText) ? noShowText : `$${noShowText}`;
                extras.push(`${noShowDisplay} No Show/Late Cancel`);
            }

            const extrasText = extras.length > 0 ? `. ${extras.join(", ")}` : "";
            let boomerangText = "";

            if (isHomeLink) {
                boomerangText = `Request Flat Rate of $${flatTotal}${extrasText}. **Enter in Goat as ${partsString}** Secure with Boomerang and leave in provider stage until rates approved`;
            } else {
                boomerangText = `Request Rates ${partsString}. **Secure with Boomerang and leave in provider stage until rates approved`;
            }

            saveLastAuthorizedRatesFromCalculator();

            navigator.clipboard.writeText(boomerangText).then(() => {
                const copiedMsg = document.createElement("div");
                copiedMsg.innerText = `"${boomerangText}" copied!`;
                copiedMsg.style.position = "fixed";
                copiedMsg.style.top = "50%";
                copiedMsg.style.left = "50%";
                copiedMsg.style.transform = "translate(-50%, -50%)";
                copiedMsg.style.background = "rgba(0,0,0,0.8)";
                copiedMsg.style.color = "#fff";
                copiedMsg.style.padding = "15px 25px";
                copiedMsg.style.borderRadius = "8px";
                copiedMsg.style.zIndex = "10001";
                copiedMsg.style.fontSize = "18px";
                copiedMsg.style.fontWeight = "bold";
                copiedMsg.style.textAlign = "center";
                copiedMsg.style.maxWidth = "80%";
                copiedMsg.style.wordWrap = "break-word";
                document.body.appendChild(copiedMsg);
                setTimeout(() => copiedMsg.remove(), 1500);
                applyHigherRatesToAuthorizationFields();
            });
        };

        const calculateMargin = () => {
            const rateType = document.querySelector('input[name="rateType"]:checked').value;
            const inputValue = parseFloat(input.value) || 0;
            const waitTimeValue = parseFloat(waitTimeInput.value) || 0;
            const noShowInfo = getNoShowAmountFromRow();

            if (noShowPreview) {
                noShowPreview.innerText = noShowInfo.amount > 0
                    ? `$${noShowInfo.amount.toFixed(2)}`
                    : "";
                noShowPreview.title = noShowInfo.rule || "";
            }

            if (transportPreview) {
                const transportInfo = getTransportPreviewAmount();
                transportPreview.innerText = transportInfo.amount > 0
                    ? `$${transportInfo.amount.toFixed(2)}`
                    : "";
                transportPreview.title = transportInfo.rule || "";
            }

            if (isNaN(inputValue) || inputValue <= 0) {
                result.innerText = "Please enter a valid amount.";
                result.style.color = "black";
                higherResult.innerText = "";
                return;
            }

            const rows = document.querySelectorAll('div[row-index]');
            let totalBilled = 0;
            let quantity = 0;
            let quantities = {};
            let foundProducts = new Set();

            rows.forEach(row => {
                const productCell = row.querySelector('[col-id="gtt_accountproduct"]');
                const totalCell = row.querySelector('[col-id="gtt_total"]');
                const qtyCell = row.querySelector('[col-id="gtt_quantity"]');

                if (productCell) {
                    const product = productCell.innerText.trim();
                    const totalText = totalCell?.innerText.trim().replace(/[^0-9.-]+/g, '') || "0";
                    const totalValue = parseFloat(totalText);

                    if (
                        (productsToTrack.includes(product) &&
                            product !== "Rush Fee" &&
                            product !== "Weekend Holiday" &&
                            product !== "Wheelchair Rental" &&
                            product !== "Airport Pickup Fee" &&
                            product !== "After Hours Fee") ||
                        product === "Wait Time"
                    ) {
                        if (!isNaN(totalValue)) totalBilled += totalValue;
                    }

                    if (qtyCell) {
                        const qtyVal = parseFloat(qtyCell.innerText.trim().replace(/[^0-9.-]+/g, ''));
                        if (!isNaN(qtyVal)) {
                            if (!quantities[product]) quantities[product] = 0;
                            quantities[product] += qtyVal;
                        }
                    }

                    if (productsToTrack.includes(product)) {
                        foundProducts.add(product);
                    }

                    if (
                        rateType === "mile" &&
                        quantity === 0 &&
                        qtyCell &&
                        ["Transport Ambulatory", "Transport Wheelchair", "Transport Stretcher, ALS & BLS"].includes(product)
                    ) {
                        const q = parseFloat(qtyCell.innerText.trim().replace(/[^0-9.-]+/g, ''));
                        if (!isNaN(q)) quantity = q;
                    }
                }
            });

            if (foundProducts.has("One Way Surcharge")) {
                foundProducts.add("One Way Surcharge");
            } else {
                foundProducts.add("Miscellaneous Dead Miles");
            }

            ["Tolls", "Other", "No Show", "Wait Time"].forEach(p => foundProducts.add(p));

            const preferredOrder = [
                "Transport Ambulatory",
                "Transport Wheelchair",
                "Transport Stretcher, ALS & BLS",
                "Miscellaneous Dead Miles",
                "One Way Surcharge",
                "Load Fee",
                "Tolls",
                "Other",
                "Wait Time",
                "Additional Passenger",
                "Rush Fee",
                "Assistance Fee",
                "After Hours Fee",
                "Weekend Holiday",
                "Wheelchair Rental",
                "Airport Pickup Fee",
                "No Show"
            ];

            preferredOrder.forEach(product => {
                if (foundProducts.has(product) && !productInputs[product]) {
                    const wrapper = document.createElement("div");
                    wrapper.style.marginTop = "10px";
                    wrapper.style.flex = "1 1 48%";

                    const labelRow = document.createElement("div");
                    labelRow.style.display = "flex";
                    labelRow.style.alignItems = "center";
                    labelRow.style.justifyContent = "space-between";
                    labelRow.style.marginBottom = "5px";

                    const leftGroup = document.createElement("div");
                    leftGroup.style.display = "flex";
                    leftGroup.style.alignItems = "center";
                    leftGroup.style.gap = "8px";

                    const rightGroup = document.createElement("div");
                    rightGroup.style.display = "flex";
                    rightGroup.style.alignItems = "center";

                    const label = document.createElement("label");
                    label.innerText = product;
                    label.style.fontWeight = "bold";
                    leftGroup.appendChild(label);

                    let inputField;

                    if (product === "No Show") {
                        noShowPreview = document.createElement("span");
                        noShowPreview.style.fontSize = "12px";
                        noShowPreview.style.fontWeight = "bold";
                        noShowPreview.style.color = "green";
                        noShowPreview.innerText = noShowInfo.amount > 0
                            ? `$${noShowInfo.amount.toFixed(2)}`
                            : "";
                        noShowPreview.title = noShowInfo.rule || "";
                        leftGroup.appendChild(noShowPreview);

                        const transportInfo = getTransportPreviewAmount();
                        transportPreview = document.createElement("span");
                        transportPreview.style.fontSize = "12px";
                        transportPreview.style.fontWeight = "bold";
                        transportPreview.style.color = "red";
                        transportPreview.style.marginLeft = "8px";
                        transportPreview.innerText = transportInfo.amount > 0
                            ? `$${transportInfo.amount.toFixed(2)}`
                            : "";
                        transportPreview.title = transportInfo.rule || "";
                        leftGroup.appendChild(transportPreview);
                    }

                    if (["Wait Time", "No Show", "Load Fee", "Additional Passenger", "Transport Ambulatory", "Transport Wheelchair", "Transport Stretcher, ALS & BLS", "After Hours Fee", "Weekend Holiday"].includes(product)) {
                        const contractBtn = document.createElement("button");
                        contractBtn.innerText = "Contract Rates";
                        contractBtn.style.marginLeft = "10px";
                        contractBtn.style.padding = "2px 6px";
                        contractBtn.style.fontSize = "12px";
                        contractBtn.style.cursor = "pointer";
                        contractBtn.onclick = () => {
                            inputField.value = "Contract Rates";
                            calculateMargin();
                        };
                        rightGroup.appendChild(contractBtn);
                    }

                    labelRow.appendChild(leftGroup);
                    labelRow.appendChild(rightGroup);

                    inputField = document.createElement("input");
                    inputField.type = "text";
                    inputField.style.width = "100%";
                    inputField.addEventListener("input", calculateMargin);

                    if (product === "One Way Surcharge" && quantities[product] !== undefined) {
                        inputField.value = quantities[product];
                        inputField.readOnly = true;
                        inputField.style.background = "#eee";
                        inputField.style.cursor = "not-allowed";
                    }

                    wrapper.appendChild(labelRow);
                    wrapper.appendChild(inputField);
                    higherInputsWrapper.appendChild(wrapper);

                    productInputs[product] = inputField;
                    if (!quantities[product]) quantities[product] = 0;
                }
            });

            if (totalBilled === 0) {
                result.innerText = "Could not find any billed total.";
                result.style.color = "black";
                higherResult.innerText = "";
            }

            const loadFeeQty = quantities["Load Fee"] || 0;

            if (rateType === "mile" && loadFeeQty > 0) {
                providerLoadFeeWrap.style.display = "block";
            } else {
                providerLoadFeeWrap.style.display = "none";
                providerLoadFeeInput.value = "";
            }

            const providerLoadFee = parseFloat(providerLoadFeeInput.value) || 0;

            let paidAmount = inputValue + waitTimeValue;
            if (rateType === "mile") {
                if (quantity === 0) {
                    result.innerText = "Could not find transport quantity.";
                    result.style.color = "black";
                    higherResult.innerText = "";
                    return;
                }
                paidAmount = (inputValue * quantity) + (providerLoadFee * loadFeeQty) + waitTimeValue;
            }

            const margin = 100 - ((paidAmount / totalBilled) * 100);
            let marginColor = "black";
            if (margin <= 24.99) marginColor = "red";
            else if (margin < 35) marginColor = "goldenrod";
            else marginColor = "green";

            const headerElement = document.querySelector('[id^="formHeaderTitle"]');
            const headerText = headerElement?.textContent?.trim() || "";

            let marginThreshold = 34.99;
            if (/^(133\-|202\-|9616\-)/.test(headerText)) {
                marginThreshold = 19.99;
            } else if (headerText.startsWith("999-")) {
                marginThreshold = 29.99;
            } else if (headerText.startsWith("4474-")) {
                marginThreshold = 19.99;
            } else if (headerText.startsWith("212-")) {
                marginThreshold = 49.99;
            }

            let approvalNote = margin <= marginThreshold
                ? `<br><span style="color: red; font-weight: bold;">Seek Management Approval</span>`
                : "";

            const milesLine =
                rateType === "mile"
                    ? `<span>Miles: ${quantity}</span>`
                    : "";

            const loadFeeLine =
                (rateType === "mile" && loadFeeQty > 0)
                    ? `<span>Load Fee: $${providerLoadFee.toFixed(2)} x ${loadFeeQty} = $${(providerLoadFee * loadFeeQty).toFixed(2)}</span>`
                    : "";

            result.innerHTML = `
${milesLine}
${loadFeeLine}
<span>Total Paid: $${paidAmount.toFixed(2)}</span>
<span>Total Billed: $${totalBilled.toFixed(2)}</span>
<span style="color:${marginColor};font-weight:bold;">Margin: ${margin.toFixed(2)}%</span>${approvalNote}
`;

            const target = totalBilled * (1 - 0.35);
            targetLabel.innerHTML = `<span>Target to pay this or less: $${target.toFixed(2)}</span>`;

            let higherTotal = 0;
            let alreadyCounted = new Set();

            foundProducts.forEach(product => {
                if (["Rush Fee", "Tolls", "Other", "Assistance Fee", "Passenger Fee", "Miscellaneous Dead Miles"].includes(product)) {
                    if (alreadyCounted.has(product)) return;
                    alreadyCounted.add(product);
                }

                let enteredValueRaw = productInputs[product]?.value;
                let enteredValue = parseFloat(enteredValueRaw);
                let qty = quantities[product] || 0;

                const contractProducts = {
                    "Wait Time": { forceQty: true },
                    "Transport Ambulatory": {},
                    "Transport Wheelchair": {},
                    "Transport Stretcher, ALS & BLS": {},
                    "Load Fee": {},
                    "After Hours Fee": {},
                    "Weekend Holiday": {},
                    "Additional Passenger": {}
                };

                let activeTransportRate = 0;
                const transportProducts = ["Transport Ambulatory", "Transport Wheelchair", "Transport Stretcher, ALS & BLS"];

                for (let transport of transportProducts) {
                    const transportInputValue = (productInputs[transport]?.value || "").toLowerCase();

                    if (transportInputValue.includes("contract")) {
                        const rows = document.querySelectorAll('div[row-index]');
                        for (let row of rows) {
                            const productCell = row.querySelector('[col-id="gtt_accountproduct"]');
                            const priceCell = row.querySelector('[col-id="gtt_price"]');
                            if (productCell && priceCell && productCell.innerText.trim() === transport) {
                                let priceText = priceCell.innerText.trim().replace(/[^0-9.-]+/g, '') || "0";
                                let priceValue = parseFloat(priceText);
                                if (!isNaN(priceValue)) activeTransportRate = priceValue;
                                break;
                            }
                        }
                        break;
                    } else if (!isNaN(parseFloat(transportInputValue))) {
                        activeTransportRate = parseFloat(transportInputValue);
                        break;
                    }
                }

                if (isNaN(activeTransportRate)) activeTransportRate = 0;

                if ((enteredValueRaw || "").toLowerCase().includes("contract") && contractProducts[product]) {
                    const rows = document.querySelectorAll('div[row-index]');
                    for (let row of rows) {
                        const productCell = row.querySelector('[col-id="gtt_accountproduct"]');
                        const priceCell = row.querySelector('[col-id="gtt_price"]');
                        if (productCell && priceCell && productCell.innerText.trim() === product) {
                            let valueText = priceCell.innerText.trim().replace(/[^0-9.-]+/g, '') || "0";
                            let value = parseFloat(valueText);
                            if (!isNaN(value)) {
                                enteredValue = value;
                                if (contractProducts[product].forceQty) qty = 1;
                            }
                            break;
                        }
                    }
                } else if (!isNaN(parseFloat(enteredValueRaw)) && enteredValueRaw.toLowerCase() !== "contract rates") {
                    enteredValue = parseFloat(enteredValueRaw);
                }

                if (!isNaN(enteredValue)) {
                    if (product === "Miscellaneous Dead Miles" || product === "One Way Surcharge") {
                        higherTotal += enteredValue * (activeTransportRate / 2);
                    } else if (["Tolls", "Other", "Assistance Fee", "Passenger Fee", "Rush Fee"].includes(product)) {
                        higherTotal += enteredValue;
                    } else if (product === "Wait Time") {
                        higherTotal += enteredValue;
                    } else {
                        higherTotal += enteredValue * qty;
                    }
                }
            });

            const higherMargin = 100 - ((paidAmount / higherTotal) * 100);
            let higherMarginColor = "black";
            if (higherMargin <= 24.99) higherMarginColor = "red";
            else if (higherMargin < 35) higherMarginColor = "goldenrod";
            else higherMarginColor = "green";

            let highermarginThreshold = 34.99;
            if (/^(133\-|202\-|9616\-)/.test(headerText)) {
                highermarginThreshold = 24.99;
            } else if (headerText.startsWith("999-")) {
                highermarginThreshold = 29.99;
            } else if (headerText.startsWith("4474-")) {
                highermarginThreshold = 19.99;
            } else if (headerText.startsWith("212-")) {
                highermarginThreshold = 49.99;
            }

            let higherApprovalNote = higherMargin <= highermarginThreshold
                ? `<br><span style="color: red; font-weight: bold;">Seek Management Approval</span>`
                : "";

            higherResult.innerHTML = `
                <span>Total Using Higher Rates: $${higherTotal.toFixed(2)}</span>
                <span style="color: ${higherMarginColor}; font-weight: bold;">Margin: ${higherMargin.toFixed(2)}%</span>${higherApprovalNote}
            `.trim();
        };

        input.addEventListener("input", calculateMargin);
        waitTimeInput.addEventListener("input", calculateMargin);
        providerLoadFeeInput.addEventListener("input", calculateMargin);
        flatRadio.addEventListener("change", () => {
            input.value = "";
            waitTimeInput.value = "";
            providerLoadFeeInput.value = "";
            providerLoadFeeWrap.style.display = "none";
            if (noShowPreview) {
                noShowPreview.innerText = "";
                noShowPreview.title = "";
            }
            if (transportPreview) {
                transportPreview.innerText = "";
                transportPreview.title = "";
            }
            calculateMargin();
        });
        mileRadio.addEventListener("change", () => {
            input.value = "";
            waitTimeInput.value = "";
            providerLoadFeeInput.value = "";
            if (noShowPreview) {
                noShowPreview.innerText = "";
                noShowPreview.title = "";
            }
            if (transportPreview) {
                transportPreview.innerText = "";
                transportPreview.title = "";
            }
            calculateMargin();
        });

        const headerElement2 = document.querySelector('[id^="formHeaderTitle"]');
        const headerText = headerElement2?.textContent?.trim() || "";

        box.appendChild(closeButton);
        box.appendChild(modeLabel);
        box.appendChild(flatRadio);
        box.appendChild(flatLabel);
        box.appendChild(mileRadio);
        box.appendChild(mileLabel);
        box.appendChild(twoColumnWrapper);
        box.appendChild(result);
        box.appendChild(targetLabel);
        box.appendChild(higherHeader);
        box.appendChild(higherInputsWrapper);
        box.appendChild(higherResult);
        box.appendChild(lowMarginButton);
        box.appendChild(uberLMButton);
        box.appendChild(waittimeButton);
        box.appendChild(WaitStaffButton);
        box.appendChild(boomerangButton);

        if (headerText.startsWith("212-")) {
            box.appendChild(homelinkButton);
        } else {
            box.appendChild(requestRatesButton);
            box.appendChild(applyRatesButton);
        }

        box.appendChild(resetButton);
        document.body.appendChild(box);

        let isDragging = false;
        let offsetX, offsetY;

        box.addEventListener('mousedown', function (e) {
            const isInteractive = ['INPUT', 'TEXTAREA', 'BUTTON'].includes(e.target.tagName);
            if (isInteractive) return;

            if (e.offsetY < 40) {
                isDragging = true;
                offsetX = e.clientX - box.getBoundingClientRect().left;
                offsetY = e.clientY - box.getBoundingClientRect().top;
                box.style.cursor = 'move';
            }
        });

        document.addEventListener('mousemove', function (e) {
            if (isDragging) {
                box.style.left = `${e.clientX - offsetX}px`;
                box.style.top = `${e.clientY - offsetY}px`;
                box.style.transform = '';
            }
        });

        document.addEventListener('mouseup', function () {
            isDragging = false;
            box.style.cursor = 'grab';
        });
    }

    const calculatorButton = document.querySelector('#yourCalculatorButtonSelector');
    if (calculatorButton) {
        calculatorButton.addEventListener('click', showCalculatorBox);
    }

    function copyClaimantName() {
        var elementToCopy = Array.from(document.querySelectorAll('a[aria-label][href*="etn=contact"]'))
            .find(el => el.textContent.trim().length > 0);

        if (elementToCopy) {
            var textToCopy = elementToCopy.textContent.trim();
            GM_setClipboard(textToCopy);
            showMessage(`Copied: "${textToCopy}" successfully.`);
            console.log('Copied to clipboard:', textToCopy);

            var serviceProviderTab = document.querySelector('li[role="tab"][title="Service Provider"]');
            if (serviceProviderTab) {
                serviceProviderTab.click();
                console.log('Clicked "Service Provider" tab.');
                waitForButtonAndClick();
            } else {
                console.error('"Service Provider" tab not found.');
            }
        } else {
            showMessage('Claimant Name not found. Please make sure you are in a referral.', false);
            console.error('Claimant element not found.');
        }
    }

    function waitForButtonAndClick() {
        var attempts = 0;
        const maxAttempts = 10;
        const interval = 1000;

        var intervalId = setInterval(() => {
            attempts++;
            console.log(`Checking for the button (Attempt ${attempts})...`);

            var targetButton = document.querySelector('#publishedCanvas > div > div.screen-animation.animated > div > div > div:nth-child(16) > div > div > div > div > button > div > div');
            if (targetButton) {
                targetButton.click();
                console.log('Clicked on the target button.');
                clearInterval(intervalId);
            }

            if (attempts >= maxAttempts) {
                console.error('Failed to find the button after multiple attempts.');
                clearInterval(intervalId);
            }
        }, interval);
    }

    function copyBoth() {
        var element1 = Array.from(document.querySelectorAll('a[aria-label][href*="etn=contact"]'))
            .find(el => el.textContent.trim().length > 0);

        var element2 = Array.from(document.querySelectorAll('a[aria-label][href*="etn=gtt_claim"]'))
            .find(el => el.textContent.trim().length > 0);

        var titleElement = document.querySelector('[id^="formHeaderTitle"]');
        var startDateInput =
            document.querySelector('input[aria-label="Start Date"]') ||
            document.querySelector('input[aria-label="Date of Start Date"]') ||
            document.querySelector('input[placeholder="---"][role="combobox"]');

        var startDateValue = startDateInput ? startDateInput.value.trim() : "";

        if (element1 && element2) {
            var text1 = element1.textContent.trim();
            var text2 = element2.textContent.trim();
            var headerTitle = titleElement ? titleElement.textContent.trim() : "";

            if (headerTitle.startsWith("4403-54316")) {
                alert("Please combine staffing and/or auth requests into one email (include multiple dates into one email).");
            }

            var referralDate = prompt("Please enter the referral date(s):", startDateValue);
            if (referralDate === null) {
                var textToCopy = `Claimant: ${text1} - Claim: ${text2} - on DOS:`;
                GM_setClipboard(textToCopy);
                showMessage(`Copied: "${textToCopy}" successfully.`);
                return;
            }

            if (!referralDate) referralDate = "[No Date Provided]";
            createDropdownMenu(text1, text2, referralDate, headerTitle);
        } else {
            showMessage('Claimant Name & Claim# not found. Please make sure you are in a referral.', false);
            console.error('Missing elements:', { element1, element2 });
        }
    }

    function createDropdownMenu(claimant, claim, referralDate, headerTitle) {
        var existingDropdown = document.getElementById("customDropdownContainer");
        if (existingDropdown) existingDropdown.remove();

        var dropdownContainer = document.createElement("div");
        dropdownContainer.id = "customDropdownContainer";
        dropdownContainer.style.position = "fixed";
        dropdownContainer.style.top = "30%";
        dropdownContainer.style.left = "50%";
        dropdownContainer.style.transform = "translate(-50%, -50%)";
        dropdownContainer.style.background = "#FFF";
        dropdownContainer.style.padding = "15px 15px";
        dropdownContainer.style.border = "1px solid #ccc";
        dropdownContainer.style.borderRadius = "8px";
        dropdownContainer.style.boxShadow = "0px 4px 6px rgba(0, 0, 0, 0.1)";
        dropdownContainer.style.zIndex = "10000";
        dropdownContainer.style.display = "flex";
        dropdownContainer.style.flexDirection = "column";
        dropdownContainer.style.alignItems = "stretch";

        var label = document.createElement("label");
        label.innerText = "Select which template you would like to apply:";
        label.style.display = "block";
        label.style.marginBottom = "16px";
        label.style.color = "#000";
        label.style.letterSpacing = "1.5px";
        label.style.fontWeight = "bold";
        dropdownContainer.appendChild(label);

        let fullOptions = [
            "Staffed Email",
            "Staffed UBER Health",
            "Staffed Revised at Approved Rates",
            "Standard Rate Request",
            "CareIQ Rate Request",
            "CareIQ Passenger Fee",
            "Convergence Higher Rate Request",
            "Homelink Rate Request",
            "Wait time request",
            "Request Demographics",
            "Other",
            "QUOTE"
        ];

const is4474 = headerTitle.startsWith("4474-");
const is11525 = headerTitle.startsWith("11525-");
const isHClaim = /^H/i.test((claim || "").trim());

// Prefer the selected text node, but fall back to selected tag/title/aria-label
const payerElement =
    document.querySelector('[data-id*="gtt_payerid"][data-id*="selected_tag_text"]') ||
    document.querySelector('[data-id*="gtt_payerid"][data-id*="selected_tag"]');

const payerText =
    payerElement?.textContent?.trim() ||
    payerElement?.getAttribute('title')?.trim() ||
    payerElement?.getAttribute('aria-label')?.trim() ||
    "";

const isJBSPacking = payerText.toLowerCase().includes("jbs packing");

// Only override 4474 when payer is JBS Packing
const use11525Rules = is11525 || (is4474 && isJBSPacking);

console.log("headerTitle:", headerTitle);
console.log("payerText:", payerText);
console.log("is4474:", is4474, "is11525:", is11525, "isJBSPacking:", isJBSPacking, "use11525Rules:", use11525Rules);

if (use11525Rules) {
    fullOptions.splice(6, 0, "JBS Request for Higher Rates");
} else if (is4474 && isHClaim) {
    fullOptions.splice(6, 0, "L-Orchid-CareWorks");
} else if (is4474 && !isHClaim) {
    fullOptions.splice(6, 0, "CareWorks Rate Request");
} else {
    fullOptions.splice(6, 0, "CareWorks Rate Request", "JBS Request for Higher Rates");
}

let exclusions = [];

if (headerTitle.startsWith("212-")) {
    exclusions = [
        "Standard Rate Request",
        "CareIQ Rate Request",
        "CareIQ Passenger Fee",
        "JBS Request for Higher Rates",
        "CareWorks Rate Request",
        "Convergence Higher Rate Request",
        "Staffed UBER Health",
        "Staffed Revised at Approved Rates"
    ];
} else if (use11525Rules) {
    exclusions = [
        "Standard Rate Request",
        "CareIQ Rate Request",
        "CareIQ Passenger Fee",
        "Convergence Higher Rate Request",
        "CareWorks Rate Request",
        "Homelink Rate Request",
        "Staffed UBER Health",
        "Staffed Revised at Approved Rates"
    ];
} else if (is4474) {
    exclusions = [
        "Standard Rate Request",
        "CareIQ Rate Request",
        "CareIQ Passenger Fee",
        "Convergence Higher Rate Request",
        "Homelink Rate Request",
        "JBS Request for Higher Rates"
    ];
} else if (headerTitle.startsWith("8814-")) {
    exclusions = [
        "Standard Rate Request",
        "CareIQ Rate Request",
        "CareIQ Passenger Fee",
        "JBS Request for Higher Rates",
        "CareWorks Rate Request",
        "Homelink Rate Request",
        "Staffed UBER Health",
        "Staffed Revised at Approved Rates"
    ];
} else if (headerTitle.startsWith("133-")) {
    exclusions = [
        "Standard Rate Request",
        "Homelink Rate Request",
        "JBS Request for Higher Rates",
        "CareWorks Rate Request",
        "Convergence Higher Rate Request",
        "Staffed UBER Health",
        "Staffed Revised at Approved Rates"
    ];
} else {
    exclusions = [
        "JBS Request for Higher Rates",
        "CareIQ Rate Request",
        "CareIQ Passenger Fee",
        "CareWorks Rate Request",
        "Homelink Rate Request",
        "Convergence Higher Rate Request",
        "Staffed UBER Health",
        "Staffed Revised at Approved Rates"
    ];
}

        const filteredOptions = fullOptions.filter(opt => !exclusions.includes(opt));

        filteredOptions.forEach(optionText => {
            const isStaff = /staff/i.test(optionText);
            const start = isStaff ? "#fde047" : "#3b82f6";
            const end   = isStaff ? "#facc15" : "#60a5fa";

            let buttonLabel = optionText;

            if (/passenger fee/i.test(optionText)) {
                buttonLabel = "👤 " + optionText;
            } else if (/wait time/i.test(optionText)) {
                buttonLabel = "🕒 " + optionText;
            } else if (/staff/i.test(optionText)) {
                buttonLabel = "✅ " + optionText;
            } else if (/demographics/i.test(optionText)) {
                buttonLabel = "📋 " + optionText;
            } else if (/quote/i.test(optionText)) {
                buttonLabel = "🧾 " + optionText;
            } else if (/other/i.test(optionText)) {
                buttonLabel = "⚙️ " + optionText;
            } else if (/close/i.test(optionText)) {
                buttonLabel = "❌ " + optionText;
            } else if (/rate request/i.test(optionText) || /higher rates/i.test(optionText)) {
                buttonLabel = "💵 " + optionText;
            }

            const button = createModernButton(buttonLabel, start, end, () => {
                finalizeCopy(claimant, claim, referralDate, optionText);
                dropdownContainer.remove();
            });
            button.style.width = "100%";
            dropdownContainer.appendChild(button);
        });

        const closeButton = createModernButton("❌ Close", "#7f1d1d", "#f87171", () => dropdownContainer.remove());
        closeButton.style.width = "100%";
        closeButton.style.marginTop = "10px";
        dropdownContainer.appendChild(closeButton);

        document.body.appendChild(dropdownContainer);
    }

    function finalizeCopy(claimant, claim, referralDate, selectedOption) {
        var textToCopy = `Claimant: ${claimant} - Claim: ${claim} on DOS: ${referralDate} `;
        GM_setClipboard(textToCopy);
        showMessage(`Copied: "${textToCopy}" successfully.`);

        var buttonToClick = document.querySelector('[id^="gtt_referral\\|NoRelationship\\|Form\\|gtt\\.gtt_referral\\.EmailConfirmation\\.Command"][id*="-button"]');
        if (buttonToClick) {
            buttonToClick.click();

            waitForPageToLoad().then(() => {
                setTimeout(() => {
                    var discardChangesButton = document.querySelector('button[title="Discard changes"]');
                    if (discardChangesButton) {
                        showMessage('Waiting for user to make choice. "Discard changes" works best.', false);

                        const observer = new MutationObserver((mutations, observerInstance) => {
                            if (!document.querySelector('button[title="Discard changes"]')) {
                                observerInstance.disconnect();
                                showMessage('Choice made, waiting for page reload...', false);

                                waitForPageToLoad().then(() => {
                                    setTimeout(() => {
                                        proceedWithRestOfFunction(claimant, claim, referralDate, selectedOption);
                                    }, 1500);
                                });
                            }
                        });

                        observer.observe(document.body, { childList: true, subtree: true });
                    } else {
                        proceedWithRestOfFunction(claimant, claim, referralDate, selectedOption);
                    }
                }, 2000);
            });
        } else {
            showMessage('EmailConfirmation button not found.', false);
        }
    }

    function waitForSavePrimary(retries = 10, delay = 1500) {
        return new Promise((resolve, reject) => {
            function tryFind() {
                var btn = document.querySelector('[id*="SavePrimary"]');
                if (btn) {
                    resolve(btn);
                } else if (retries > 0) {
                    setTimeout(() => {
                        retries--;
                        tryFind();
                    }, delay);
                } else {
                    reject('SavePrimary button not found.');
                }
            }
            tryFind();
        });
    }

    function proceedWithRestOfFunction(claimant, claim, referralDate, selectedOption) {
        showProcessingMessage();

        waitForSavePrimary()
            .then((savePrimaryButton) => {
                setTimeout(() => {
                    savePrimaryButton.click();
                }, 1500);

                setTimeout(() => {
                    var iframe = document.querySelector('#WebResource_AttachmentSelector');
                    if (iframe) {
                        try {
                            var iframeDoc = iframe.contentDocument || iframe.contentWindow.document;
                            var checkBox = iframeDoc.querySelector('#cbClientForm');

                            if (checkBox && !checkBox.checked) {
                                checkBox.click();
                            }

                            setTimeout(() => {
                                var templateButton = document.querySelector('[id*="Template"]');
                                if (templateButton) {
                                    templateButton.click();

                                    setTimeout(() => {
                                        selectCorrectRadioButton(selectedOption);
                                    }, 2500);
                                } else {
                                    showMessage('Template button not found.', false);
                                    hideProcessingMessage();
                                }
                            }, 2500);
                        } catch (e) {
                            console.error('Cannot access iframe content:', e);
                            hideProcessingMessage();
                        }
                    } else {
                        console.error('No iframe found.');
                        hideProcessingMessage();
                    }
                }, 5000);
            })
            .catch((error) => {
                showMessage(error, false);
                hideProcessingMessage();
            });
    }

    function showProcessingMessage() {
        hideProcessingMessage();

        const messages = [
            "Please be patient, I'm thinking...",
            "Crunching the numbers, hang tight!",
            "Just a moment, magic is happening...",
            "Working on it... almost there!",
            "Hold on, good things take time...",
            "Just like Windows Update... this might take a while.",
            "Loading... don't worry, it's not a blue screen!",
            "Consulting Clippy for advice...",
            "Rebooting in spirit only.",
            "Optimizing like it's Excel on a Friday afternoon...",
            "Dragging files across the desktop... metaphorically.",
            "Checking OneDrive... it’s probably syncing. Maybe.",
            "Running diagnostics... hope it’s not a Microsoft Teams call!",
            "Searching Bing... because someone has to.",
            "Applying 1,024 patches... just like Patch Tuesday.",
            "Installing updates you didn’t ask for...",
            "Waiting for Outlook to stop being 'Not Responding'...",
            "Wishing Word would stop autocorrecting me...",
            "Recompiling with love and Microsoft Paint.",
            "Consulting the Registry... send snacks.",
            "PowerPointing my way through the process...",
            "Unclogging the virtual printer queue...",
            "Restarting the restart...",
            "Syncing with the spirit of Windows 95..."
        ];
        const randomMessage = messages[Math.floor(Math.random() * messages.length)];

        if (!document.getElementById('processingMessageStyle')) {
            const style = document.createElement('style');
            style.id = 'processingMessageStyle';
            style.innerHTML = `
                .loader {
                    border: 4px solid #f3f3f3;
                    border-top: 4px solid #3498db;
                    border-radius: 50%;
                    width: 30px;
                    height: 30px;
                    animation: spin 1s linear infinite;
                    margin-bottom: 10px;
                }
                @keyframes spin {
                    0% { transform: rotate(0deg); }
                    100% { transform: rotate(360deg); }
                }
            `;
            document.head.appendChild(style);
        }

        const msg = document.createElement('div');
        msg.id = 'processingMessage';
        msg.innerHTML = `<div class="loader"></div><div>${randomMessage}</div>`;
        Object.assign(msg.style, {
            position: 'fixed',
            top: '50%',
            left: '50%',
            transform: 'translate(-50%, -50%)',
            backgroundColor: 'rgba(0, 0, 0, 0.8)',
            color: '#fff',
            padding: '20px 40px',
            borderRadius: '10px',
            fontSize: '18px',
            zIndex: '9999',
            textAlign: 'center',
            display: 'flex',
            flexDirection: 'column',
            alignItems: 'center',
            maxWidth: '80%',
            wordBreak: 'break-word'
        });

        const escHandler = (e) => {
            if (e.key === 'Escape') hideProcessingMessage();
        };
        msg.dataset.escHandler = 'true';
        document.addEventListener('keydown', escHandler, { once: true });
        msg._escHandler = escHandler;

        document.body.appendChild(msg);

        processingTimeoutId = window.setTimeout(() => {
            hideProcessingMessage();
        }, 14000);
    }

    function hideProcessingMessage() {
        if (processingTimeoutId !== null) {
            clearTimeout(processingTimeoutId);
            processingTimeoutId = null;
        }

        const msg = document.getElementById('processingMessage');
        if (msg) {
            if (msg._escHandler) {
                document.removeEventListener('keydown', msg._escHandler);
            }
            msg.remove();
        }
    }

    function waitForPageToLoad() {
        return new Promise((resolve) => {
            const observer = new MutationObserver((mutations, observerInstance) => {
                if (document.readyState === 'complete') {
                    observerInstance.disconnect();
                    resolve();
                }
            });

            observer.observe(document.body, { childList: true, subtree: true });
            setTimeout(resolve, 14000);
        });
    }

    function waitForVisibleButton(selector, timeout = 8000) {
        return new Promise((resolve, reject) => {
            const start = Date.now();

            function findButton() {
                const matches = Array.from(document.querySelectorAll(selector)).filter(btn => {
                    const style = window.getComputedStyle(btn);
                    return (
                        btn.offsetParent !== null &&
                        style.display !== 'none' &&
                        style.visibility !== 'hidden' &&
                        style.opacity !== '0' &&
                        !btn.disabled &&
                        btn.getAttribute('aria-disabled') !== 'true'
                    );
                });

                return matches.length ? matches[matches.length - 1] : null;
            }

            function check() {
                const btn = findButton();
                if (btn) {
                    resolve(btn);
                    return;
                }

                if (Date.now() - start >= timeout) {
                    reject(`Button not found: ${selector}`);
                    return;
                }

                setTimeout(check, 200);
            }

            check();
        });
    }

    function clickInsertSignatureButton(retries = 25, delay = 1000) {
        return new Promise((resolve, reject) => {
            function tryClick() {
                const buttons = Array.from(document.querySelectorAll('button[aria-label^="Insert Signature"]')).filter(btn => {
                    const style = window.getComputedStyle(btn);
                    return (
                        btn.offsetParent !== null &&
                        style.display !== 'none' &&
                        style.visibility !== 'hidden' &&
                        style.opacity !== '0' &&
                        !btn.disabled &&
                        btn.getAttribute('aria-disabled') !== 'true'
                    );
                });

                if (buttons.length) {
                    const btn = buttons[buttons.length - 1];
                    try { btn.focus(); } catch (e) {}
                    btn.click();
                    console.log('Clicked Insert Signature button:', btn);
                    resolve(btn);
                    return;
                }

                if (retries > 0) {
                    setTimeout(() => {
                        retries--;
                        tryClick();
                    }, delay);
                } else {
                    reject('Insert Signature button not found.');
                }
            }

            tryClick();
        });
    }

    function waitForDeleteReferralOutboxToFinish(timeout = 10000) {
        return new Promise((resolve) => {
            const start = Date.now();

            function check() {
                const deleteButton = document.querySelector('button[aria-label="Delete Referral Outbox"]');

                if (!deleteButton || deleteButton.offsetParent === null) {
                    resolve();
                    return;
                }

                if (Date.now() - start >= timeout) {
                    resolve();
                    return;
                }

                setTimeout(check, 250);
            }

            check();
        });
    }

    function clickDeleteReferralOutboxIfPresent(timeout = 5000) {
        return new Promise((resolve) => {
            const start = Date.now();

            function check() {
                const deleteButtons = Array.from(document.querySelectorAll('button[aria-label="Delete Referral Outbox"]')).filter(btn => {
                    const style = window.getComputedStyle(btn);
                    return (
                        btn.offsetParent !== null &&
                        style.display !== 'none' &&
                        style.visibility !== 'hidden' &&
                        style.opacity !== '0' &&
                        !btn.disabled &&
                        btn.getAttribute('aria-disabled') !== 'true'
                    );
                });

                if (deleteButtons.length) {
                    const btn = deleteButtons[deleteButtons.length - 1];
                    btn.click();
                    console.log('Clicked Delete Referral Outbox after signature.');
                    waitForDeleteReferralOutboxToFinish(10000).then(resolve);
                    return;
                }

                if (Date.now() - start >= timeout) {
                    resolve();
                    return;
                }

                setTimeout(check, 250);
            }

            check();
        });
    }

/* =================== CAREIQ PASSENGER CC WARNING =================== */

let mtoyCareIqPassengerShown = false;

function showCareIqPassengerCcPopup() {
    if (!window.mtoyCareIqPassengerWatchEnabled) return;
    if (mtoyCareIqPassengerShown) return;
    if (document.getElementById('mtoy-careiq-passenger-cc-popup')) return;

    mtoyCareIqPassengerShown = true;

    const box = document.createElement('div');
    box.id = 'mtoy-careiq-passenger-cc-popup';
    box.innerHTML = `
        <div style="font-weight:700; margin-bottom:6px;">Additional Passenger fee found</div>
        <div>Please make sure to CC <b>revisions@gotandt.com</b></div>
        <button id="mtoy-careiq-passenger-cc-ok" type="button" style="
            margin-top:10px;
            padding:6px 12px;
            border-radius:8px;
            border:1px solid rgba(0,0,0,.25);
            background:#fff;
            cursor:pointer;
            font-weight:600;
        ">OK</button>
    `;

    box.style.cssText = `
        position: fixed;
        left: 50%;
        top: 120px;
        transform: translateX(-50%);
        z-index: 2147483647;
        background: #fff3cd;
        color: #111;
        border: 2px solid #ff9800;
        border-radius: 12px;
        padding: 14px 18px;
        box-shadow: 0 8px 28px rgba(0,0,0,.30);
        font: 14px/1.35 system-ui,-apple-system,Segoe UI,Roboto,Arial;
        text-align: center;
        min-width: 310px;
    `;

    document.body.appendChild(box);
    box.querySelector('#mtoy-careiq-passenger-cc-ok').onclick = () => box.remove();
}

function scanCareIqPassengerInDocument(doc) {
    try {
        if (!doc || !doc.body) return;

        const editable = doc.querySelector(
            'body[contenteditable="true"], [contenteditable="true"], .cke_editable'
        );

        if (editable) {
            const txt = (editable.innerText || editable.textContent || '').toLowerCase();

            if (txt.includes('passenger')) {
                showCareIqPassengerCcPopup();
            }

            if (editable.dataset.mtoyCareIqPassengerWatch !== '1') {
                editable.dataset.mtoyCareIqPassengerWatch = '1';

                const check = () => scanCareIqPassengerInDocument(doc);

                new MutationObserver(check).observe(editable, {
                    childList: true,
                    subtree: true,
                    characterData: true
                });

                editable.addEventListener('input', check, true);
                editable.addEventListener('keyup', check, true);
                editable.addEventListener('paste', () => setTimeout(check, 100), true);
            }
        }

        doc.querySelectorAll('iframe').forEach(ifr => {
            try {
                const childDoc = ifr.contentDocument || ifr.contentWindow?.document;
                if (childDoc) scanCareIqPassengerInDocument(childDoc);
            } catch {}
        });

    } catch {}
}

function startCareIqPassengerWatch() {
    window.mtoyCareIqPassengerWatchEnabled = true;
    mtoyCareIqPassengerShown = false;

    let tries = 0;
    const timer = setInterval(() => {
        tries++;
        scanCareIqPassengerInDocument(document);

        if (mtoyCareIqPassengerShown || tries > 120) {
            clearInterval(timer);
        }
    }, 1000);
}

function showTemplateReminderPopup(message) {
    const existing = document.getElementById('mtoy-template-reminder-popup');
    if (existing) existing.remove();

    const box = document.createElement('div');
    box.id = 'mtoy-template-reminder-popup';
    box.innerHTML = `
        <div style="font-weight:700; margin-bottom:6px;">Email Reminder</div>
        <div>${message}</div>
        <button id="mtoy-template-reminder-ok" type="button" style="
            margin-top:10px;
            padding:6px 12px;
            border-radius:8px;
            border:1px solid rgba(0,0,0,.25);
            background:#fff;
            cursor:pointer;
            font-weight:600;
        ">OK</button>
    `;

    box.style.cssText = `
        position: fixed;
        left: 50%;
        top: 120px;
        transform: translateX(-50%);
        z-index: 2147483647;
        background: #fff3cd;
        color: #111;
        border: 2px solid #ff9800;
        border-radius: 12px;
        padding: 14px 18px;
        box-shadow: 0 8px 28px rgba(0,0,0,.30);
        font: 14px/1.35 system-ui,-apple-system,Segoe UI,Roboto,Arial;
        text-align: center;
        min-width: 360px;
        max-width: 620px;
    `;

    document.body.appendChild(box);
    box.querySelector('#mtoy-template-reminder-ok').onclick = () => box.remove();
}

    function selectCorrectRadioButton(selectedOption) {
        showProcessingMessage();

        var labelToFind = "";
        if (selectedOption === "Staffed Email") {
            labelToFind = "Staffed";
        } else if (selectedOption === "Staffed UBER Health") {
            labelToFind = "CareWorks Uber Staffed";
        } else if (selectedOption === "Staffed Revised at Approved Rates") {
            labelToFind = "Careworks Revision Staffed";
        } else if (selectedOption === "Standard Rate Request") {
            labelToFind = "Request for Higher Rates";
        } else if (selectedOption === "CareIQ Rate Request") {
            labelToFind = "CIQ Higher Rate Request";
            startCareIqPassengerWatch();
        } else if (selectedOption === "CareIQ Passenger Fee") {
            labelToFind = "CIQ - **Request for Additional Passenger Fee**";
            startCareIqPassengerWatch();
        } else if (selectedOption === "Homelink Rate Request") {
            labelToFind = "Homelink – Request for Higher Rates";
        } else if (selectedOption === "Convergence Higher Rate Request") {
            labelToFind = "Convergence Higher Rate Request";
            showTemplateReminderPopup("Please always CC <b>kfimbres@convergencecare.com</b>");
        } else if (selectedOption === "JBS Request for Higher Rates") {
            labelToFind = "JBS Request for Higher Rates (Default Rates)";
            showTemplateReminderPopup("Please send email to <b>AboveContractedRateRequest@sedgwick.com</b>");
        } else if (selectedOption === "Wait time request") {
            labelToFind = "Wait Time Request";
        } else if (selectedOption === "CareWorks Rate Request") {
            labelToFind = "CareWorks - Request for Higher Rates";
            showTemplateReminderPopup("Please send email to <b>AboveContractedRateRequest@sedgwick.com</b>");
        } else if (selectedOption === "Request Demographics") {
            labelToFind = "Request for Additional Information";
        } else if (selectedOption === "L-Orchid-CareWorks") {
            labelToFind = "L-Orchid-CareWorks";
            showTemplateReminderPopup("Please send email only to the rep that sent the order and to <b>Chandra.Thurman@careworks.com; savannah.lussier@careworks.com</b>");
        } else if (selectedOption === "QUOTE") {
            labelToFind = "Quote request";
        } else if (selectedOption === "Other") {
            labelToFind = "";
        }

        if (labelToFind === "") {
            hideProcessingMessage();
            return;
        }

        const isStaffTemplate =
            selectedOption === "Staffed Email" ||
            selectedOption === "Staffed UBER Health" ||
            selectedOption === "Staffed Revised at Approved Rates";

        try {
            var labels = document.querySelectorAll('label');
            var correctLabel = Array.from(labels).find(label => {
                var titleSpan = label.querySelector('.titleText');
                return titleSpan && titleSpan.textContent.trim() === labelToFind;
            });

            if (!correctLabel) {
                hideProcessingMessage();
                showMessage(`Template label not found: ${labelToFind}`, false);
                return;
            }

            var associatedRadio = document.getElementById(correctLabel.getAttribute('for'));
            if (!associatedRadio) {
                hideProcessingMessage();
                showMessage('Associated radio button not found.', false);
                return;
            }

            associatedRadio.click();

            waitForVisibleButton('button[title="Apply template"]', 8000)
                .then((applyTemplateButton) => {
                    applyTemplateButton.click();
                    return waitForVisibleButton('button[title="OK"]', 8000);
                })
                .then((okButton) => {
                    okButton.click();
                    return waitForVisibleButton('[id*="SavePrimary"]', 10000);
                })
                .then((savePrimaryButton) => {
                    savePrimaryButton.click();

                    setTimeout(() => {
                        clickInsertSignatureButton(
                            isStaffTemplate ? 20 : 30,
                            isStaffTemplate ? 800 : 1000
                        )
                            .then(() => {
                                if (isStaffTemplate) {
                                    hideProcessingMessage();

                                    var messageDiv = document.createElement("div");
                                    messageDiv.innerText = "Template applied and signature inserted. Please proceed with filling out the remainder of the information needed.";
                                    messageDiv.style.position = "fixed";
                                    messageDiv.style.top = "50%";
                                    messageDiv.style.left = "50%";
                                    messageDiv.style.transform = "translate(-50%, -50%)";
                                    messageDiv.style.backgroundColor = "#000";
                                    messageDiv.style.color = "#fff";
                                    messageDiv.style.padding = "20px";
                                    messageDiv.style.borderRadius = "10px";
                                    messageDiv.style.zIndex = "1000";
                                    document.body.appendChild(messageDiv);

                                    setTimeout(function () {
                                        if (document.body.contains(messageDiv)) {
                                            document.body.removeChild(messageDiv);
                                        }
                                    }, 1500);
                                    return;
                                }

                                setTimeout(() => {
                                    clickDeleteReferralOutboxIfPresent(5000).then(() => {
                                        setTimeout(() => {
                                            const saveAgainButtons = Array.from(document.querySelectorAll('[id*="SavePrimary"]')).filter(btn => {
                                                const style = window.getComputedStyle(btn);
                                                return (
                                                    btn.offsetParent !== null &&
                                                    style.display !== 'none' &&
                                                    style.visibility !== 'hidden' &&
                                                    style.opacity !== '0' &&
                                                    !btn.disabled &&
                                                    btn.getAttribute('aria-disabled') !== 'true'
                                                );
                                            });

                                            if (saveAgainButtons.length) {
                                                saveAgainButtons[saveAgainButtons.length - 1].click();
                                            }

                                            hideProcessingMessage();

                                            var messageDiv = document.createElement("div");
                                            messageDiv.innerText = "Template applied and signature inserted. Please proceed with filling out the remainder of the information needed.";
                                            messageDiv.style.position = "fixed";
                                            messageDiv.style.top = "50%";
                                            messageDiv.style.left = "50%";
                                            messageDiv.style.transform = "translate(-50%, -50%)";
                                            messageDiv.style.backgroundColor = "#000";
                                            messageDiv.style.color = "#fff";
                                            messageDiv.style.padding = "20px";
                                            messageDiv.style.borderRadius = "10px";
                                            messageDiv.style.zIndex = "1000";
                                            document.body.appendChild(messageDiv);

                                            setTimeout(function () {
                                                if (document.body.contains(messageDiv)) {
                                                    document.body.removeChild(messageDiv);
                                                }
                                            }, 1500);
                                        }, 800);
                                    });
                                }, 800);
                            })
                            .catch((error) => {
                                console.warn(error);
                                hideProcessingMessage();

                                var messageDiv = document.createElement("div");
                                messageDiv.innerText = "Template applied, but Insert Signature button was not found.";
                                messageDiv.style.position = "fixed";
                                messageDiv.style.top = "50%";
                                messageDiv.style.left = "50%";
                                messageDiv.style.transform = "translate(-50%, -50%)";
                                messageDiv.style.backgroundColor = "#000";
                                messageDiv.style.color = "#fff";
                                messageDiv.style.padding = "20px";
                                messageDiv.style.borderRadius = "10px";
                                messageDiv.style.zIndex = "1000";
                                document.body.appendChild(messageDiv);

                                setTimeout(function () {
                                    if (document.body.contains(messageDiv)) {
                                        document.body.removeChild(messageDiv);
                                    }
                                }, 2000);
                            });
                    }, isStaffTemplate ? 1200 : 1600);
                })
                .catch((error) => {
                    console.warn(error);
                    hideProcessingMessage();
                    showMessage(
                        error && error.message ? error.message : 'Template flow failed.',
                        false
                    );
                });

        } catch (error) {
            console.warn(error);
            hideProcessingMessage();
            showMessage('Template flow failed.', false);
        }
    }

    function extractAndCopyTitle() {
        const titleElement = document.querySelector('[id^="formHeaderTitle"]');
        if (titleElement) {
            const title = titleElement.textContent.trim();
            console.log('Title Found:', title);

            if (title.startsWith("4473-")) {
                alert("Rates up to $3.75/mile are ok to apply and include in staffing email.\n\n" +
                      "Wait time ok if 25 miles or more each way and not surgery.\n\n" +
                      "If scheduled the day before appointment, ok to proceed with staffing above approved rates if needed and include in staffing email.");
            }

            if (title.startsWith("5843-")) {
                alert("Rate Increases do not need AUTH - provide rate increase in staffed email");
            }

            if (title.startsWith("212-")) {
                alert("Homelink must be atleast 50% margin");
            }

            const hyphenIndexes = [...title].reduce((acc, char, idx) => {
                if (char === '-') acc.push(idx);
                return acc;
            }, []);

            if (hyphenIndexes.length >= 2) {
                const start = hyphenIndexes[0] + 1;
                const end = hyphenIndexes[1];
                let extractedText = title.slice(start, end).trim();
                extractedText = `*${extractedText}`;
                GM_setClipboard(extractedText);
                showMessage(`Copied: "${extractedText}" successfully.`);
                console.log('Extracted and copied to clipboard:', extractedText);
            } else {
                showMessage('Title does not contain two hyphens.', false);
                console.error('Title does not contain two hyphens.');
            }
        } else {
            showMessage('Claimant ID not found yet.', false);
            console.error('Claimant ID not found.');
        }
    }


    function detectCombinedMarginModes() {
        let hasTransport = false;
        let hasInterpretation = false;

        document.querySelectorAll('div[row-index], [role="row"]').forEach((row) => {
            const productCell = row.querySelector('[col-id="gtt_accountproduct"]');
            const serviceTypeCell = row.querySelector('[col-id="gtt_servicetype"]');

            const productText = (productCell?.innerText || productCell?.textContent || "").trim().toLowerCase();
            const serviceTypeText = (serviceTypeCell?.innerText || serviceTypeCell?.textContent || "").trim().toLowerCase();

            if (
                productText.includes("transport") ||
                serviceTypeText.includes("transport")
            ) {
                hasTransport = true;
            }

            if (
                productText.includes("interpretation") ||
                serviceTypeText.includes("interpretation")
            ) {
                hasInterpretation = true;
            }
        });

        return { hasTransport, hasInterpretation };
    }

    function showCombinedMarginChooser() {
        const existing = document.getElementById("combinedMarginChooser");
        if (existing) existing.remove();

        const chooser = document.createElement("div");
        chooser.id = "combinedMarginChooser";
        chooser.style.position = "fixed";
        chooser.style.top = "30%";
        chooser.style.left = "50%";
        chooser.style.transform = "translate(-50%, -50%)";
        chooser.style.background = "#fff";
        chooser.style.padding = "16px";
        chooser.style.border = "2px solid #000";
        chooser.style.borderRadius = "10px";
        chooser.style.zIndex = "10001";
        chooser.style.minWidth = "320px";
        chooser.style.boxShadow = "0 10px 24px rgba(0,0,0,0.2)";

        const title = document.createElement("div");
        title.innerText = "Select margin calculator";
        title.style.fontWeight = "bold";
        title.style.fontSize = "18px";
        title.style.marginBottom = "12px";

        const transportBtn = createModernButton("Transport", "#22c55e", "#4ade80", () => {
            chooser.remove();
            showCalculatorBox();
        });
        transportBtn.style.marginLeft = "0";

        const interpretationBtn = createModernButton("Interpretation", "#3b82f6", "#60a5fa", () => {
            chooser.remove();
            openInterpretationCalculator();
        });

        const closeBtn = createModernButton("Close", "#ef4444", "#f87171", () => chooser.remove());
        closeBtn.style.marginLeft = "0";
        closeBtn.style.marginTop = "10px";

        chooser.appendChild(title);
        chooser.appendChild(transportBtn);
        chooser.appendChild(interpretationBtn);
        chooser.appendChild(document.createElement("br"));
        chooser.appendChild(closeBtn);
        document.body.appendChild(chooser);
    }

    function openCombinedMarginSelector() {
        const tabSelectors = [
            'li[role="tab"][title="Billed"]',
            'li[role="tab"][title="Billing"]',
            'button[role="tab"][title="Billed"]',
            'button[role="tab"][title="Billing"]'
        ];

        const billedTab = tabSelectors
            .map(selector => document.querySelector(selector))
            .find(Boolean);

        const openAfterBillingLoads = (maxAttempts = 14, delay = 500) => {
            let attempts = 0;

            const check = () => {
                attempts += 1;
                const modes = detectCombinedMarginModes();

                if (modes.hasTransport && modes.hasInterpretation) {
                    showCombinedMarginChooser();
                    return;
                }
                if (modes.hasTransport) {
                    showCalculatorBox();
                    return;
                }
                if (modes.hasInterpretation) {
                    openInterpretationCalculator();
                    return;
                }

                if (attempts >= maxAttempts) {
                    showMessage("No transport or interpretation lines found.", false);
                    return;
                }

                setTimeout(check, delay);
            };

            setTimeout(check, 250);
        };

        if (billedTab) {
            billedTab.click();
            openAfterBillingLoads();
            return;
        }

        openAfterBillingLoads(1, 0);
    }


  function tlParseMoney(text) {
    const cleaned = String(text || "").replace(/[^0-9.-]+/g, "").trim();
    const value = parseFloat(cleaned);
    return Number.isFinite(value) ? value : 0;
  }

  function tlParseQty(text) {
    const cleaned = String(text || "").replace(/[^0-9.-]+/g, "").trim();
    const value = parseFloat(cleaned);
    return Number.isFinite(value) ? value : 0;
  }

  function tlGetGridRows() {
    return Array.from(document.querySelectorAll('div[row-index], [role="row"]'));
  }

  function tlGetGridData() {
    const billedProducts = [
      "Certified Interpretation",
      "Standard Interpretation",
      "Legal Interpretation",
      "Rush Fee",
      "Weekend Holiday",
      "Interpretation Travel Mileage",
      "Interpretation Travel Time",
      "Tolls",
      "Parking"
    ];

    const productMap = new Map();
    let billedTotal = 0;

    tlGetGridRows().forEach((row) => {
      const productCell = row.querySelector('[col-id="gtt_accountproduct"]');
      if (!productCell) return;

      const rawName = (productCell.innerText || productCell.textContent || "").trim();
      if (!rawName) return;

      const qtyCell = row.querySelector('[col-id="gtt_quantity"]');
      const priceCell = row.querySelector('[col-id="gtt_price"]');
      const totalCell = row.querySelector('[col-id="gtt_total"]');

      const qty = tlParseQty(qtyCell ? qtyCell.innerText : "");
      const price = tlParseMoney(priceCell ? priceCell.innerText : "");
      const total = tlParseMoney(totalCell ? totalCell.innerText : "");

      const existing = productMap.get(rawName) || { qty: 0, price: 0, total: 0, rows: 0 };
      existing.qty += qty;
      existing.total += total;
      existing.rows += 1;
      if (price > 0) existing.price = price;
      productMap.set(rawName, existing);

      if (billedProducts.includes(rawName)) {
        billedTotal += total;
      }
    });

    return { productMap, billedTotal };
  }

  function tlCalculateInterpreterMargin(state) {
    const mode = state.flatRadio.checked ? "flat" : "rates";

    const flatProviderRate = parseFloat(state.flatProviderRateInput.value) || 0;
    const perHour = parseFloat(state.perHourInput.value) || 0;
    const minHours = parseFloat(state.minHoursInput.value) || 0;
    const miles = parseFloat(state.milesInput.value) || 0;
    const mileageRate = parseFloat(state.mileageRateInput.value) || 0;

    const paid = mode === "flat"
      ? flatProviderRate
      : (perHour * minHours) + (miles * mileageRate);

    const { billedTotal, productMap } = tlGetGridData();

    const breakdown = [];
    const orderedBilledLines = [
      "Certified Interpretation",
      "Standard Interpretation",
      "Legal Interpretation",
      "Rush Fee",
      "Weekend Holiday",
      "Interpretation Travel Mileage",
      "Interpretation Travel Time",
      "Tolls",
      "Parking"
    ];

    orderedBilledLines.forEach((name) => {
      const item = productMap.get(name);
      if (!item) return;

      if (["Certified Interpretation", "Standard Interpretation", "Legal Interpretation", "Interpretation Travel Mileage", "Interpretation Travel Time"].includes(name)) {
        breakdown.push(`${name}: ${item.qty} x $${item.price.toFixed(2)} = $${item.total.toFixed(2)}`);
      } else {
        breakdown.push(`${name}: $${item.total.toFixed(2)}`);
      }
    });

    if (paid <= 0) {
      state.result.innerHTML = `<span>Please enter a valid paid amount.</span>`;
      state.result.style.color = "black";
      state.targetLabel.innerHTML = "";
      state.breakdown.innerHTML = breakdown.length
        ? `<div style="margin-top:10px;font-weight:bold;">Billed Breakdown</div><div>${breakdown.join("<br>")}</div>`
        : "<div style='margin-top:10px;'>No billed interpretation lines found.</div>";
      state.higherResult.innerHTML = "";
      return;
    }

    if (billedTotal <= 0) {
      state.result.innerHTML = `<span>Could not find any billed interpretation total.</span>`;
      state.result.style.color = "black";
      state.targetLabel.innerHTML = "";
      state.breakdown.innerHTML = "<div style='margin-top:10px;'>No billed interpretation lines found.</div>";
      state.higherResult.innerHTML = "";
      return;
    }

    const margin = 100 - ((paid / billedTotal) * 100);
    let marginColor = "green";
    if (margin < 45) marginColor = "red";
    else if (margin < 50) marginColor = "goldenrod";

    const approvalNote = margin < 45
      ? `<br><span style="color:red;font-weight:bold;">Seek Management Approval</span>`
      : "";

    const target = billedTotal * 0.5;

    const paidLines = mode === "flat"
      ? [`Flat Provider Rate: $${flatProviderRate.toFixed(2)}`]
      : [
          `Per Hour: $${perHour.toFixed(2)} x ${minHours.toFixed(2)} = $${(perHour * minHours).toFixed(2)}`,
          `Mileage: ${miles.toFixed(2)} x $${mileageRate.toFixed(2)} = $${(miles * mileageRate).toFixed(2)}`
        ];

    state.result.innerHTML = `
      ${paidLines.map(line => `<span>${line}</span>`).join("")}
      <span>Total Paid: $${paid.toFixed(2)}</span>
      <span>Total Billed: $${billedTotal.toFixed(2)}</span>
      <span style="color:${marginColor};font-weight:bold;">Margin: ${margin.toFixed(2)}%</span>${approvalNote}
    `;

    state.targetLabel.innerHTML = `<span>Target to pay this or less (50% margin): $${target.toFixed(2)}</span>`;
    state.breakdown.innerHTML = breakdown.length
      ? `<div style="margin-top:10px;font-weight:bold;">Billed Breakdown</div><div>${breakdown.join("<br>")}</div>`
      : "<div style='margin-top:10px;'>No billed interpretation lines found.</div>";

    tlCalculateHigherRatesMargin(state, paid);
  }

  function tlCalculateHigherRatesMargin(state, paidAmount) {
    const { productMap } = tlGetGridData();
    const fields = state.higherRateFields;

    const quantityProducts = new Set([
      "Certified Interpretation",
      "Standard Interpretation",
      "Legal Interpretation",
      "Interpretation Travel Mileage",
      "Interpretation Travel Time"
    ]);

    let higherTotal = 0;
    const lines = [];

    state.higherRateOrder.forEach((name) => {
      const input = fields[name];
      if (!input) return;

      const raw = String(input.value || "").trim();
      if (!raw) return;

      const entered = parseFloat(raw);
      if (!Number.isFinite(entered)) return;

      const item = productMap.get(name);
      if (quantityProducts.has(name)) {
        const qty = item ? item.qty : 0;
        const lineTotal = qty * entered;
        higherTotal += lineTotal;
        lines.push(`${name}: ${qty.toFixed(2)} x $${entered.toFixed(2)} = $${lineTotal.toFixed(2)}`);
      } else {
        higherTotal += entered;
        lines.push(`${name}: $${entered.toFixed(2)}`);
      }
    });

    if (higherTotal <= 0) {
      state.higherResult.innerHTML = "";
      return;
    }

    const higherMargin = 100 - ((paidAmount / higherTotal) * 100);
    let higherMarginColor = "green";
    if (higherMargin < 45) higherMarginColor = "red";
    else if (higherMargin < 50) higherMarginColor = "goldenrod";

    const higherApprovalNote = higherMargin < 45
      ? `<br><span style="color:red;font-weight:bold;">Seek Management Approval</span>`
      : "";

    state.higherResult.innerHTML = `
      <div style="margin-top:10px;font-weight:bold;">Higher Rates Calculator</div>
      <span>Total Using Higher Rates: $${higherTotal.toFixed(2)}</span>
      <span style="color:${higherMarginColor};font-weight:bold;">Margin: ${higherMargin.toFixed(2)}%</span>${higherApprovalNote}
      ${lines.length ? `<div style="margin-top:8px;">${lines.join("<br>")}</div>` : ""}
    `;
  }

  function tlGetRequestRatesParts(state) {
    const { productMap } = tlGetGridData();
    const fields = state.higherRateFields;
    const quantityProducts = new Set([
      "Certified Interpretation",
      "Standard Interpretation",
      "Legal Interpretation",
      "Interpretation Travel Mileage",
      "Interpretation Travel Time"
    ]);

    const parts = [];

    state.higherRateOrder.forEach((name) => {
      const input = fields[name];
      if (!input) return;

      const raw = String(input.value || "").trim();
      if (!raw) return;

      const entered = parseFloat(raw);
      if (!Number.isFinite(entered)) return;

      if (quantityProducts.has(name)) {
        const qty = productMap.get(name)?.qty || 0;
        parts.push(`$${entered.toFixed(2)} ${name}${qty ? ` x ${qty}` : ""}`);
      } else {
        parts.push(`$${entered.toFixed(2)} ${name}`);
      }
    });

    return parts.join(", ");
  }

  function tlCopyHigherRatesRequest(state, prefixText) {
    const finalParts = tlGetRequestRatesParts(state);

    if (!finalParts) {
      showCenteredOverlayMessage("Please enter at least one higher-rates value.", false, 1400);
      return;
    }

    const finalText = `${prefixText} ${finalParts}`;
    GM_setClipboard(finalText);
    showCenteredOverlayMessage(`Copied: "${finalText}"`, true, 1500);
  }

  function openInterpretationCalculator() {
    const tabSelectors = [
      'li[role="tab"][title="Billed"]',
      'li[role="tab"][title="Billing"]',
      'button[role="tab"][title="Billed"]',
      'button[role="tab"][title="Billing"]'
    ];

    const billedTab = tabSelectors
      .map(selector => document.querySelector(selector))
      .find(Boolean);

    const hasLoadedBillingData = () => {
      const { productMap, billedTotal } = tlGetGridData();
      if (billedTotal > 0) return true;

      const triggerProducts = [
        "Certified Interpretation",
        "Standard Interpretation",
        "Legal Interpretation",
        "Interpretation Travel Mileage",
        "Interpretation Travel Time",
        "Rush Fee",
        "Weekend Holiday",
        "No Show",
        "Tolls",
        "Parking"
      ];

      return triggerProducts.some((name) => productMap.has(name));
    };

    const waitForBillingDataThenOpen = (maxAttempts = 12, delay = 500) => {
      let attempts = 0;

      const check = () => {
        attempts += 1;

        if (hasLoadedBillingData() || attempts >= maxAttempts) {
          showTLInterpretationCalculator();
          return;
        }

        setTimeout(check, delay);
      };

      setTimeout(check, 250);
    };

    if (billedTab) {
      billedTab.click();
      waitForBillingDataThenOpen();
      return;
    }

    showTLInterpretationCalculator();
  }

  function showTLInterpretationCalculator() {
    const existing = document.getElementById("tl-interpretation-calc-box");
    if (existing) existing.remove();

    const box = document.createElement("div");
    box.id = "tl-interpretation-calc-box";
    Object.assign(box.style, {
      position: "fixed",
      top: "6%",
      left: "78%",
      transform: "translateX(-50%)",
      background: "#fff",
      padding: "20px",
      border: "2px solid #000",
      borderRadius: "10px",
      zIndex: "10000",
      minWidth: "560px",
      maxWidth: "560px",
      height: "820px",
      overflowY: "auto",
      color: "#000",
      boxShadow: "0 10px 24px rgba(0,0,0,0.2)"
    });

      const interpretationLowMarginButton = createModernButton(
  "Low Margin OK",
  "#3b82f6",
  "#60a5fa",
  () => {
    // Prefer higher-rates margin if displayed; fall back to standard margin
    let marginDisplay = null;
    const higherResultText = higherResult.innerText || higherResult.textContent || "";
    const higherMarginMatch = higherResultText.match(/Margin:\s*(-?[0-9.]+)%/);
    if (higherMarginMatch) {
      marginDisplay = higherMarginMatch[1];
    } else {
      const resultText = result.innerText || result.textContent || "";
      const marginMatch = resultText.match(/Margin:\s*(-?[0-9.]+)%/);
      if (marginMatch) marginDisplay = marginMatch[1];
    }

    if (!marginDisplay) {
      showCenteredOverlayMessage("No margin calculated yet. Enter a Provider Rate first.", false, 2000);
      return;
    }

    const marginValue = parseFloat(marginDisplay);
    const marginType = marginValue < 0 ? "negative margin" : "low margin";
    const baseText = `Management aware of Interpretation ${marginType}. Ok to staff approx. Margin ${marginDisplay}%`;

    if (marginValue < 0) {
      showReasonPrompt((reason) => {
        const textToCopy = reason ? `${baseText} Reason: ${reason}` : baseText;
        navigator.clipboard.writeText(textToCopy).then(() => {
          showCenteredOverlayMessage(`"${textToCopy}" copied!`, true, 1000);
        });
      });
    } else {
      navigator.clipboard.writeText(baseText).then(() => {
        showCenteredOverlayMessage(`"${baseText}" copied!`, true, 1000);
      });
    }
  }
);

    const closeButton = document.createElement("button");
    closeButton.innerText = "X";
    Object.assign(closeButton.style, {
      position: "absolute",
      top: "5px",
      right: "10px",
      border: "none",
      background: "transparent",
      color: "#000",
      fontSize: "20px",
      fontWeight: "bold",
      cursor: "pointer"
    });
    closeButton.onclick = () => box.remove();

    const header = document.createElement("div");
    header.innerText = "Interpretation Margin Calculator";
    header.style.fontWeight = "bold";
    header.style.fontSize = "22px";
    header.style.marginBottom = "12px";

    const modeLabel = document.createElement("div");
    modeLabel.innerText = "Select Payment Type:";
    modeLabel.style.marginBottom = "8px";
    modeLabel.style.fontWeight = "bold";

    const flatRadio = document.createElement("input");
    flatRadio.type = "radio";
    flatRadio.name = "tlInterpretationRateType";
    flatRadio.value = "flat";
    flatRadio.id = "tlInterpretationRateFlat";
    flatRadio.checked = true;

    const flatLabel = document.createElement("label");
    flatLabel.htmlFor = "tlInterpretationRateFlat";
    flatLabel.innerText = "Flat Fee";
    flatLabel.style.marginRight = "22px";
    flatLabel.style.marginLeft = "6px";

    const ratesRadio = document.createElement("input");
    ratesRadio.type = "radio";
    ratesRadio.name = "tlInterpretationRateType";
    ratesRadio.value = "rates";
    ratesRadio.id = "tlInterpretationRateRates";

    const ratesLabel = document.createElement("label");
    ratesLabel.htmlFor = "tlInterpretationRateRates";
    ratesLabel.innerText = "Rates";
    ratesLabel.style.marginLeft = "6px";

    const flatSection = document.createElement("div");
    flatSection.style.marginTop = "16px";

    const flatProviderRateLabel = document.createElement("label");
    flatProviderRateLabel.innerText = "Enter Provider Rate:";
    flatProviderRateLabel.style.fontWeight = "bold";

    const flatProviderRateInput = document.createElement("input");
    flatProviderRateInput.type = "number";
    flatProviderRateInput.step = "0.01";
    flatProviderRateInput.style.width = "100%";
    flatProviderRateInput.style.marginTop = "8px";
    flatProviderRateInput.style.marginBottom = "4px";
    flatProviderRateInput.style.padding = "8px";

    flatSection.appendChild(flatProviderRateLabel);
    flatSection.appendChild(flatProviderRateInput);

    const ratesSection = document.createElement("div");
    ratesSection.style.display = "none";
    ratesSection.style.marginTop = "16px";

    const ratesGrid = document.createElement("div");
    ratesGrid.style.display = "grid";
    ratesGrid.style.gridTemplateColumns = "1fr 1fr";
    ratesGrid.style.columnGap = "14px";
    ratesGrid.style.rowGap = "12px";
    ratesGrid.style.alignItems = "start";
    ratesSection.appendChild(ratesGrid);

    function buildField(labelText) {
      const wrap = document.createElement("div");
      const label = document.createElement("label");
      label.innerText = labelText;
      label.style.fontWeight = "bold";
      label.style.display = "block";
      const input = document.createElement("input");
      input.type = "number";
      input.step = "0.01";
      input.style.width = "100%";
      input.style.marginTop = "8px";
      input.style.padding = "8px";
      wrap.appendChild(label);
      wrap.appendChild(input);
      ratesGrid.appendChild(wrap);
      return input;
    }

    const perHourInput = buildField("Per Hour:");
    const minHoursInput = buildField("Min Hours:");
    const milesInput = buildField("Miles:");
    const mileageRateInput = buildField("Mileage Rate:");

    const result = document.createElement("div");
    result.style.marginTop = "12px";
    result.style.fontWeight = "bold";
    result.style.whiteSpace = "pre-line";
    result.style.display = "flex";
    result.style.flexDirection = "column";
    result.style.gap = "4px";

    const targetLabel = document.createElement("div");
    targetLabel.style.marginTop = "10px";
    targetLabel.style.fontWeight = "bold";

    const higherHeader = document.createElement("div");
    higherHeader.innerText = "Higher Rates Calculator";
    higherHeader.style.marginTop = "18px";
    higherHeader.style.marginBottom = "8px";
    higherHeader.style.fontWeight = "bold";
    higherHeader.style.fontSize = "20px";

    const higherInputsWrapper = document.createElement("div");
    higherInputsWrapper.style.display = "flex";
    higherInputsWrapper.style.flexWrap = "wrap";
    higherInputsWrapper.style.columnGap = "10px";
    higherInputsWrapper.style.rowGap = "12px";

    const higherResult = document.createElement("div");
    higherResult.style.marginTop = "12px";
    higherResult.style.fontWeight = "bold";
    higherResult.style.whiteSpace = "pre-line";
    higherResult.style.display = "flex";
    higherResult.style.flexDirection = "column";
    higherResult.style.gap = "4px";

    const { productMap } = tlGetGridData();

    const higherRateOrder = [
      "Certified Interpretation",
      "Standard Interpretation",
      "Legal Interpretation",
      "Rush Fee",
      "Weekend Holiday",
      "Interpretation Travel Mileage",
      "Interpretation Travel Time",
      "No Show"
    ];

    const interpretationTypes = [
      "Certified Interpretation",
      "Standard Interpretation",
      "Legal Interpretation"
    ];
    const chosenInterpretationType = interpretationTypes.find((name) => productMap.has(name)) || null;

    const higherRateFields = {};

    function createHigherRateField(name, hasQty) {
      const wrap = document.createElement("div");
      wrap.style.flex = "1 1 48%";
      wrap.style.border = "1px solid #cfcfcf";
      wrap.style.borderRadius = "10px";
      wrap.style.padding = "10px";
      wrap.style.background = "#fafafa";

      const label = document.createElement("label");
      label.innerText = name;
      label.style.fontWeight = "bold";
      label.style.display = "block";
      label.style.marginBottom = "8px";

      wrap.appendChild(label);

      let amountInput = null;
      let qtyInput = null;

      if (hasQty) {
        const row = document.createElement("div");
        row.style.display = "grid";
        row.style.gridTemplateColumns = "1fr 1fr";
        row.style.columnGap = "10px";
        row.style.alignItems = "start";

        const amountWrap = document.createElement("div");
        const amountLabel = document.createElement("div");
        amountLabel.innerText = "Price:";
        amountLabel.style.fontSize = "12px";
        amountLabel.style.fontWeight = "bold";
        amountLabel.style.marginBottom = "4px";

        amountInput = document.createElement("input");
        amountInput.type = "number";
        amountInput.step = "0.01";
        amountInput.style.width = "100%";
        amountInput.style.padding = "8px";

        amountWrap.appendChild(amountLabel);
        amountWrap.appendChild(amountInput);

        const qtyWrap = document.createElement("div");
        const qtyLabel = document.createElement("div");
        qtyLabel.innerText = "Qty:";
        qtyLabel.style.fontSize = "12px";
        qtyLabel.style.fontWeight = "bold";
        qtyLabel.style.marginBottom = "4px";

        qtyInput = document.createElement("input");
        qtyInput.type = "number";
        qtyInput.step = "0.01";
        qtyInput.style.width = "100%";
        qtyInput.style.padding = "8px";

        qtyWrap.appendChild(qtyLabel);
        qtyWrap.appendChild(qtyInput);

        row.appendChild(amountWrap);
        row.appendChild(qtyWrap);
        wrap.appendChild(row);
      } else {
        const amountLabel = document.createElement("div");
        amountLabel.innerText = "Price:";
        amountLabel.style.fontSize = "12px";
        amountLabel.style.fontWeight = "bold";
        amountLabel.style.marginBottom = "4px";

        amountInput = document.createElement("input");
        amountInput.type = "number";
        amountInput.step = "0.01";
        amountInput.style.width = "100%";
        amountInput.style.padding = "8px";

        wrap.appendChild(amountLabel);
        wrap.appendChild(amountInput);
      }

      higherInputsWrapper.appendChild(wrap);
      higherRateFields[name] = { wrap, amountInput, qtyInput, hasQty };
    }

    if (chosenInterpretationType) {
      createHigherRateField(chosenInterpretationType, true);
    }
    if (productMap.has("Rush Fee")) {
      createHigherRateField("Rush Fee", false);
    }
    if (productMap.has("Weekend Holiday")) {
      createHigherRateField("Weekend Holiday", false);
    }
    createHigherRateField("Interpretation Travel Mileage", true);
    createHigherRateField("Interpretation Travel Time", true);
    createHigherRateField("No Show", false);

    function tlCalculateHigherRatesMargin(paidAmount) {
      let higherTotal = 0;
      const lines = [];

      higherRateOrder.forEach((name) => {
        const field = higherRateFields[name];
        if (!field) return;

        const amount = parseFloat(String(field.amountInput?.value || "").trim()) || 0;
        const qty = parseFloat(String(field.qtyInput?.value || "").trim()) || 0;

        if (name === "No Show") {
          if (amount <= 0) return;
          lines.push(`No Show/Late Cancel: $${amount.toFixed(2)}`);
          return;
        }

        if (field.hasQty) {
          if (amount <= 0 || qty <= 0) return;
          const lineTotal = amount * qty;
          higherTotal += lineTotal;
          lines.push(`${name}: ${qty.toFixed(2)} x $${amount.toFixed(2)} = $${lineTotal.toFixed(2)}`);
        } else {
          if (amount <= 0) return;
          higherTotal += amount;
          lines.push(`${name}: $${amount.toFixed(2)}`);
        }
      });

      if (higherTotal <= 0) {
        higherResult.innerHTML = "";
        return;
      }

      const higherMargin = 100 - ((paidAmount / higherTotal) * 100);
      let higherMarginColor = "green";
      if (higherMargin < 45) higherMarginColor = "red";
      else if (higherMargin < 50) higherMarginColor = "goldenrod";

      const higherApprovalNote = higherMargin < 45
        ? `<br><span style="color:red;font-weight:bold;">Seek Management Approval</span>`
        : "";

      higherResult.innerHTML = `
        <div style="margin-top:10px;font-weight:bold;">Higher Rates Calculator</div>
        <span>Total Using Higher Rates: $${higherTotal.toFixed(2)}</span>
        <span style="color:${higherMarginColor};font-weight:bold;">Margin: ${higherMargin.toFixed(2)}%</span>${higherApprovalNote}
        ${lines.length ? `<div style="margin-top:8px;">${lines.join("<br>")}</div>` : ""}
      `;
    }

    function tlGetRequestRatesParts() {
      const parts = [];

      higherRateOrder.forEach((name) => {
        const field = higherRateFields[name];
        if (!field) return;

        const amount = parseFloat(String(field.amountInput?.value || "").trim()) || 0;
        const qty = parseFloat(String(field.qtyInput?.value || "").trim()) || 0;

        if (name === "Certified Interpretation" || name === "Standard Interpretation" || name === "Legal Interpretation") {
          if (amount <= 0 || qty <= 0) return;
          parts.push(`$${amount.toFixed(2)}/hour x ${qty}`);
          return;
        }

        if (name === "Interpretation Travel Mileage") {
          if (amount <= 0 || qty <= 0) return;
          parts.push(`${qty} Travel Miles x $${amount.toFixed(2)}`);
          return;
        }

        if (name === "Interpretation Travel Time") {
          if (amount <= 0 || qty <= 0) return;
          parts.push(`$${amount.toFixed(2)} Travel Time x ${qty}`);
          return;
        }

        if (name === "No Show") {
          if (amount <= 0) return;
          parts.push(`$${amount.toFixed(2)} No Show/Late Cancel`);
          return;
        }

        if (field.hasQty) {
          if (amount <= 0 || qty <= 0) return;
          parts.push(`$${amount.toFixed(2)} ${name} x ${qty}`);
        } else {
          if (amount <= 0) return;
          parts.push(`$${amount.toFixed(2)} ${name}`);
        }
      });

      return parts.join(", ");
    }

    function tlCopyHigherRatesRequest(prefixText) {
      const finalParts = tlGetRequestRatesParts();

      if (!finalParts) {
        showMessage("Please enter at least one higher-rates value.", false);
        return;
      }

      const finalText = `${prefixText} ${finalParts}`;
      GM_setClipboard(finalText);
      showMessage(`Copied: "${finalText}"`, true);
    }

    function recalc() {
      const mode = flatRadio.checked ? "flat" : "rates";

      const flatProviderRate = parseFloat(flatProviderRateInput.value) || 0;
      const perHour = parseFloat(perHourInput.value) || 0;
      const minHours = parseFloat(minHoursInput.value) || 0;
      const miles = parseFloat(milesInput.value) || 0;
      const mileageRate = parseFloat(mileageRateInput.value) || 0;

      const paid = mode === "flat"
        ? flatProviderRate
        : (perHour * minHours) + (miles * mileageRate);

      const { billedTotal, productMap: latestProductMap } = tlGetGridData();
      const breakdownLines = [];
      const billedOrder = [
        "Certified Interpretation",
        "Standard Interpretation",
        "Legal Interpretation",
        "Rush Fee",
        "Weekend Holiday",
        "Interpretation Travel Mileage",
        "Interpretation Travel Time",
        "Tolls",
        "Parking"
      ];

      billedOrder.forEach((name) => {
        const item = latestProductMap.get(name);
        if (!item) return;

        if ([
          "Certified Interpretation",
          "Standard Interpretation",
          "Legal Interpretation",
          "Interpretation Travel Mileage",
          "Interpretation Travel Time"
        ].includes(name)) {
          breakdownLines.push(`${name}: ${item.qty} x $${item.price.toFixed(2)} = $${item.total.toFixed(2)}`);
        } else {
          breakdownLines.push(`${name}: $${item.total.toFixed(2)}`);
        }
      });

      if (paid <= 0) {
        result.innerHTML = `<span>Please enter a valid paid amount.</span>`;
        targetLabel.innerHTML = "";

        higherResult.innerHTML = "";
        return;
      }

      if (billedTotal <= 0) {
        result.innerHTML = `<span>Could not find any billed interpretation total.</span>`;
        targetLabel.innerHTML = "";

        higherResult.innerHTML = "";
        return;
      }

      const margin = 100 - ((paid / billedTotal) * 100);
      let marginColor = "green";
      if (margin < 45) marginColor = "red";
      else if (margin < 50) marginColor = "goldenrod";

      const approvalNote = margin < 45
        ? `<br><span style="color:red;font-weight:bold;">Seek Management Approval</span>`
        : "";

      const target = billedTotal * 0.5;

      const paidLines = mode === "flat"
        ? [`Flat Provider Rate: $${flatProviderRate.toFixed(2)}`]
        : [
            `Per Hour: $${perHour.toFixed(2)} x ${minHours.toFixed(2)} = $${(perHour * minHours).toFixed(2)}`,
            `Mileage: ${miles.toFixed(2)} x $${mileageRate.toFixed(2)} = $${(miles * mileageRate).toFixed(2)}`
          ];

      result.innerHTML = `
        ${paidLines.map(line => `<span>${line}</span>`).join("")}
        <span>Total Paid: $${paid.toFixed(2)}</span>
        <span>Total Billed: $${billedTotal.toFixed(2)}</span>
        <span style="color:${marginColor};font-weight:bold;">Margin: ${margin.toFixed(2)}%</span>${approvalNote}
      `;

      targetLabel.innerHTML = `<span>Target to pay this or less: $${target.toFixed(2)}</span>`;


      tlCalculateHigherRatesMargin(paid);
    }

    const requestRatesButton = createModernButton("Request Rates", "#22c55e", "#4ade80", () => {
      tlCopyHigherRatesRequest("Request rates");
    });
    requestRatesButton.style.marginLeft = "0";
    requestRatesButton.style.marginTop = "12px";

    const applyRatesButton = createModernButton("Apply & Staff", "#a855f7", "#c084fc", () => {
      tlCopyHigherRatesRequest("Apply rates");
    });
    applyRatesButton.style.marginTop = "12px";

    const resetButton = createModernButton("Reset", "#ef4444", "#f87171", () => {
      flatRadio.checked = true;
      ratesRadio.checked = false;
      flatSection.style.display = "block";
      ratesSection.style.display = "none";
      flatProviderRateInput.value = "";
      perHourInput.value = "";
      minHoursInput.value = "";
      milesInput.value = "";
      mileageRateInput.value = "";
      Object.values(higherRateFields).forEach((field) => {
        if (field.amountInput) field.amountInput.value = "";
        if (field.qtyInput) field.qtyInput.value = "";
      });
      result.innerHTML = "";
      targetLabel.innerHTML = "";

      higherResult.innerHTML = "";
    });
    resetButton.style.marginTop = "12px";

    const allNumberInputs = [flatProviderRateInput, perHourInput, minHoursInput, milesInput, mileageRateInput];
    Object.values(higherRateFields).forEach((field) => {
      if (field.amountInput) allNumberInputs.push(field.amountInput);
      if (field.qtyInput) allNumberInputs.push(field.qtyInput);
    });

    allNumberInputs.forEach((input) => {
      input.addEventListener("input", recalc);
      input.addEventListener("wheel", (e) => {
        e.preventDefault();
        input.blur();
      }, { passive: false });
    });

    flatRadio.addEventListener("change", () => {
      flatSection.style.display = "block";
      ratesSection.style.display = "none";
      recalc();
    });

    ratesRadio.addEventListener("change", () => {
      flatSection.style.display = "none";
      ratesSection.style.display = "block";
      recalc();
    });

    const buttonRow = document.createElement("div");
    buttonRow.style.display = "flex";
    buttonRow.style.alignItems = "center";
    buttonRow.style.flexWrap = "wrap";
    buttonRow.appendChild(requestRatesButton);
    buttonRow.appendChild(applyRatesButton);
    buttonRow.appendChild(resetButton);

    box.appendChild(closeButton);
    box.appendChild(header);
    box.appendChild(modeLabel);
    box.appendChild(flatRadio);
    box.appendChild(flatLabel);
    box.appendChild(ratesRadio);
    box.appendChild(ratesLabel);
    box.appendChild(flatSection);
    box.appendChild(ratesSection);
    box.appendChild(result);
    box.appendChild(targetLabel);
    box.appendChild(higherHeader);
    box.appendChild(higherInputsWrapper);
    box.appendChild(higherResult);
    box.appendChild(interpretationLowMarginButton);
    box.appendChild(buttonRow);
    document.body.appendChild(box);

    let isDragging = false;
    let offsetX = 0;
    let offsetY = 0;

    box.addEventListener("mousedown", function (e) {
      const isInteractive = ["INPUT", "TEXTAREA", "BUTTON", "LABEL"].includes(e.target.tagName);
      if (isInteractive) return;

      if (e.offsetY < 40) {
        isDragging = true;
        offsetX = e.clientX - box.getBoundingClientRect().left;
        offsetY = e.clientY - box.getBoundingClientRect().top;
        box.style.cursor = "move";
      }
    });

    document.addEventListener("mousemove", function (e) {
      if (!isDragging) return;
      box.style.left = `${e.clientX - offsetX}px`;
      box.style.top = `${e.clientY - offsetY}px`;
      box.style.transform = "";
    });

    document.addEventListener("mouseup", function () {
      if (!isDragging) return;
      isDragging = false;
      box.style.cursor = "default";
    });

    recalc();
  }


    // Shows a styled prompt dialog asking for a reason, then calls callback(reason).
    // If the user cancels, callback is not called.
    function showReasonPrompt(onConfirm) {
        const overlay = document.createElement("div");
        overlay.style.cssText = `
            position:fixed;top:0;left:0;width:100%;height:100%;
            background:rgba(0,0,0,0.55);z-index:10100;
            display:flex;align-items:center;justify-content:center;
        `;

        const dialog = document.createElement("div");
        dialog.style.cssText = `
            background:#fff;color:#000;padding:24px 28px;border-radius:10px;
            box-shadow:0 8px 30px rgba(0,0,0,0.25);min-width:340px;max-width:480px;
            display:flex;flex-direction:column;gap:14px;
        `;

        const title = document.createElement("div");
        title.innerText = "Enter Reason for Negative Margin";
        title.style.cssText = "font-weight:bold;font-size:16px;";

        const textarea = document.createElement("textarea");
        textarea.placeholder = "e.g. to prevent a miss";
        textarea.style.cssText = `
            width:100%;box-sizing:border-box;height:80px;padding:8px;
            font-size:14px;border:1px solid #ccc;border-radius:6px;resize:vertical;
        `;

        const errorMsg = document.createElement("div");
        errorMsg.innerText = "Reason must be entered.";
        errorMsg.style.cssText = `
            color:#dc3545;font-size:13px;font-weight:600;display:none;margin-top:-6px;
        `;

        const btnRow = document.createElement("div");
        btnRow.style.cssText = "display:flex;justify-content:flex-end;gap:10px;";

        const cancelBtn = document.createElement("button");
        cancelBtn.innerText = "Cancel";
        cancelBtn.style.cssText = `
            padding:8px 18px;border-radius:6px;border:1px solid #ccc;
            background:#f0f0f0;cursor:pointer;font-size:14px;font-weight:600;
        `;
        cancelBtn.onclick = () => overlay.remove();

        const confirmBtn = document.createElement("button");
        confirmBtn.innerText = "Copy";
        confirmBtn.style.cssText = `
            padding:8px 18px;border-radius:6px;border:none;
            background:linear-gradient(135deg,#3b82f6,#60a5fa);
            color:#fff;cursor:pointer;font-size:14px;font-weight:600;
        `;
        confirmBtn.onclick = () => {
            const reason = textarea.value.trim();
            if (!reason) {
                errorMsg.style.display = "block";
                textarea.style.border = "1px solid #dc3545";
                textarea.focus();
                return;
            }
            overlay.remove();
            onConfirm(reason);
        };

        // Clear error state when user starts typing
        textarea.addEventListener("input", () => {
            if (textarea.value.trim()) {
                errorMsg.style.display = "none";
                textarea.style.border = "1px solid #ccc";
            }
        });

        // Allow Enter (without Shift) to confirm, Escape to cancel
        textarea.addEventListener("keydown", (e) => {
            if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); confirmBtn.click(); }
            if (e.key === "Escape") cancelBtn.click();
        });

        btnRow.appendChild(cancelBtn);
        btnRow.appendChild(confirmBtn);
        dialog.appendChild(title);
        dialog.appendChild(textarea);
        dialog.appendChild(errorMsg);
        dialog.appendChild(btnRow);
        overlay.appendChild(dialog);
        document.body.appendChild(overlay);
        setTimeout(() => textarea.focus(), 50);
    }

    function createButtons() {
        const searchBox = document.querySelector('#searchBoxLiveRegion');

        if (searchBox) {
            if (document.getElementById('custom-button-container')) {
                console.log('Buttons already created.');
                return;
            }

            const buttonContainer = document.createElement('div');
            buttonContainer.id = 'custom-button-container';
            buttonContainer.style.display = 'inline-flex';
            buttonContainer.style.alignItems = 'center';
            buttonContainer.style.marginLeft = '10px';

            const button1 = createModernButton('Payer Emails', '#3b82f6', '#60a5fa', copyBoth);
            const button2 = createModernButton('Copy Name/SP Tab', '#3b82f6', '#60a5fa', copyClaimantName);
            const button3 = createModernButton('ClaimantID', '#3b82f6', '#60a5fa', extractAndCopyTitle);
            const button4 = createModernButton('Margin', '#22c55e', '#4ade80', openCombinedMarginSelector);
            const button5 = createModernButton('Multi Day Rates', '#991b1b', '#ef4444', () => {
                if (typeof window.ddApplyLastAuthorizedRates === 'function') {
                    window.ddApplyLastAuthorizedRates();
                } else {
                    showCenteredOverlayMessage('Multi Rates helper is not ready. Refresh the page and try again.', false, 2500);
                }
            });

            buttonContainer.appendChild(button3);
            buttonContainer.appendChild(button2);
            buttonContainer.appendChild(button1);
            buttonContainer.appendChild(button4);
            buttonContainer.appendChild(button5);

            searchBox.parentNode.insertBefore(buttonContainer, searchBox.nextSibling);
            console.log('Buttons created and inserted next to #searchBoxLiveRegion.');
        } else {
            console.log('Search box not found, retrying...');
        }
    }

    function addTooltip(button, message) {
        const tooltip = document.createElement("div");
        tooltip.textContent = message;
        tooltip.style.position = "absolute";
        tooltip.style.backgroundColor = "#333";
        tooltip.style.color = "#fff";
        tooltip.style.padding = "5px";
        tooltip.style.borderRadius = "5px";
        tooltip.style.fontSize = "12px";
        tooltip.style.visibility = "hidden";
        tooltip.style.whiteSpace = "nowrap";
        tooltip.style.zIndex = "1000";
        tooltip.style.transition = "opacity 0.2s";
        tooltip.style.opacity = "0";
        document.body.appendChild(tooltip);

        button.addEventListener("mouseenter", (event) => {
            tooltip.style.top = `${event.clientY + 10}px`;
            tooltip.style.left = `${event.clientX + 10}px`;
            tooltip.style.visibility = "visible";
            tooltip.style.opacity = "1";
        });

        button.addEventListener("mouseleave", () => {
            tooltip.style.visibility = "hidden";
            tooltip.style.opacity = "0";
        });
    }

    createButtons();

    // ─── SPA navigation: re-inject buttons whenever Dynamics navigates.
    // window 'load' only fires once on the initial page load — it never fires
    // again on SPA navigations (which is how Dynamics moves between views).
    // Using a MutationObserver on document.body catches every navigation.
    function waitForSearchBox() {
        const observer = new MutationObserver(() => {
            const searchBox = document.querySelector('#searchBoxLiveRegion');
            if (searchBox && !document.getElementById('custom-button-container')) {
                createButtons();
            }
        });

        observer.observe(document.body, {
            childList: true,
            subtree: true,
        });
    }

    waitForSearchBox();

    } // end init()

    // ─── Entry point: wait for Dynamics to render #searchBoxLiveRegion before injecting.
    // A fixed delay isn't reliable — Dynamics can take anywhere from 1s to 10s+ depending
    // on load. Instead we watch for the landmark element that the buttons attach to.
    // DYNAMICS_INIT_DELAY is the minimum wait (ms) before we even start checking, giving
    // Dynamics a head-start before the observer fires. Raise this if buttons still inject
    // too early on slow connections. DYNAMICS_WAIT_CAP is the hard timeout (ms) after which
    // we give up waiting and run init anyway (catches edge cases where the landmark never appears).
    const DYNAMICS_INIT_DELAY = 4000; // ← adjust minimum wait here if needed (ms)
    const DYNAMICS_WAIT_CAP   = 30000; // ← hard give-up timeout (ms) — rarely needs changing

    function waitForLandmarkThenInit() {
        if (document.querySelector('#searchBoxLiveRegion')) { init(); return; }
        let done = false;
        const capTimer = setTimeout(() => { if (!done) { done = true; obs.disconnect(); init(); } }, DYNAMICS_WAIT_CAP - DYNAMICS_INIT_DELAY);
        const obs = new MutationObserver(() => {
            if (document.querySelector('#searchBoxLiveRegion')) {
                if (!done) { done = true; clearTimeout(capTimer); obs.disconnect(); init(); }
            }
        });
        obs.observe(document.documentElement, { childList: true, subtree: true });
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', () => setTimeout(waitForLandmarkThenInit, DYNAMICS_INIT_DELAY));
    } else {
        setTimeout(waitForLandmarkThenInit, DYNAMICS_INIT_DELAY);
    }

})();
