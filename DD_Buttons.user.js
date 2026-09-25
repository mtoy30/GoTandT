// ==UserScript==
// @name         DD_Buttons
// @namespace    https://github.com/mtoy30/GoTandT
// @version      4.2.31
// @updateURL    https://raw.githubusercontent.com/mtoy30/GoTandT/main/DD_Buttons.user.js
// @downloadURL  https://raw.githubusercontent.com/mtoy30/GoTandT/main/DD_Buttons.user.js
// @description  Custom script for Dynamics 365 CRM page with multiple button functionalities
// @match        https://gotandt.crm.dynamics.com/*
// @author       Michael Toy
// @grant        GM_setClipboard
// @grant        GM_registerMenuCommand
// @grant        GM_xmlhttpRequest
// @grant        unsafeWindow
// @connect      lowmargin.mtoysystems.com
// ==/UserScript==
//Moved to GitHub for 3.2.5+

(function() {
    'use strict';

    function init() {

    // One-shot timeout so the processing message never hangs
    let processingTimeoutId = null;

    const style = document.createElement('style');
    style.innerHTML = `
       input[type="radio"] {
        width: 20px;
        height: 20px;
        accent-color: #3b82f6; /* or brand color */
   }
   `;
   document.head.appendChild(style);

    // Utility function to create a modern styled button
function createModernButton(text, gradientStart, gradientEnd, onClick) {
    const btn = document.createElement("button");
    btn.innerText = text;
    btn.style.cssText = `
        margin-top: 5px;
        //margin-left: 10px;
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

    // Function to show a temporary message
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
        }, 3000);
    }

    // Center-screen status used by Blind Send, matching DD_Buttons_Admin.
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


// --- LMS Transport No Show API helpers ---
const LMS_TRANSPORT_NOSHOW_API_URL = "https://lowmargin.mtoysystems.com/api/submit_transport.php";
const LMS_LOGIN_URL = "https://lowmargin.mtoysystems.com/login.php";

function moneyTextToNumber(text) {
    const cleaned = String(text || "").replace(/[^0-9.-]+/g, "");
    const n = parseFloat(cleaned);
    return isNaN(n) ? 0 : n;
}

function formatMoneyValue(value) {
    const n = parseFloat(value);
    return isNaN(n) ? "" : n.toFixed(2);
}

function getReferralNumberFromHeader() {
    const headerElement = document.querySelector('[id^="formHeaderTitle"]');
    const headerText = headerElement?.textContent?.trim() || document.title || "";
    const m = headerText.match(/\b\d+-\d+-\d+\b/);
    return m ? m[0] : "";
}

function getProviderNameFromPage() {
    const selectors = [
        '[data-id*="gtt_serviceprovider"][data-id*="selected_tag_text"]',
        '[data-id*="gtt_serviceprovider"][data-id*="selected_tag"]',
        '[data-id*="provider"][data-id*="selected_tag_text"]',
        '[data-id*="provider"][data-id*="selected_tag"]',
        'a[aria-label][href*="etn=account"]'
    ];

    for (const sel of selectors) {
        const el = document.querySelector(sel);
        const txt =
            el?.textContent?.trim() ||
            el?.getAttribute("title")?.trim() ||
            el?.getAttribute("aria-label")?.trim() ||
            "";
        if (txt) return txt.replace(/^Lookup\s*:\s*/i, "").trim();
    }

    return "";
}

function getCurrentMarginFromResult(resultElement) {
    const txt = resultElement?.innerText || "";
    const m = txt.match(/Margin:\s*(-?\d+(?:\.\d+)?)%/i);
    return m ? m[1] : "";
}

function getBillingProductPrice(productName) {
    const rows = document.querySelectorAll('div[row-index]');
    for (const row of rows) {
        const productCell = row.querySelector('[col-id="gtt_accountproduct"]');
        const priceCell = row.querySelector('[col-id="gtt_price"]');
        if (!productCell || !priceCell) continue;
        if (productCell.innerText.trim() !== productName) continue;

        const value = moneyTextToNumber(priceCell.innerText);
        return value > 0 ? value : 0;
    }
    return 0;
}

async function ensureBillingTabOpen() {
    const billingTab = document.querySelector('li[role="tab"][title="Billing"]');
    if (!billingTab) {
        throw new Error('Could not find the Billing tab. Please open Billing and try again.');
    }

    const isSelected =
        billingTab.getAttribute('aria-selected') === 'true' ||
        billingTab.classList.contains('selected');

    if (!isSelected) {
        billingTab.click();
    }

    // Allow Dynamics time to activate and render the Billing product grid.
    const timeoutAt = Date.now() + 5000;
    while (Date.now() < timeoutAt) {
        const selectedNow = billingTab.getAttribute('aria-selected') === 'true';
        const billingRowsLoaded = document.querySelectorAll('div[row-index]').length > 0;

        if (selectedNow && billingRowsLoaded) {
            await new Promise(resolve => setTimeout(resolve, 350));
            return;
        }

        await new Promise(resolve => setTimeout(resolve, 150));
    }

    // Dynamics does not always update aria-selected consistently, so continue
    // after a final render delay once the click has been issued.
    await new Promise(resolve => setTimeout(resolve, 750));
}

function askYesNo(question) {
    return confirm(question + "\n\nOK = Yes\nCancel = No") ? "Yes" : "No";
}

function askUberOption() {
    const options = [
        "No Phone Number for claimant",
        "Claimant does not want rideshare",
        "No drivers in area",
        "Low availability",
        "Medium availability",
        "High availability",
        "Cost more than provider",
        "Referral has wait time",
        "To Large (Over 35miles one way)",
        "Not allowed for this payer",
        "This is a TECH",
        "ALERT shows this is PREFERRED provider"
    ];

    const msg = "Did you check UBER/Lyft?\n\n" +
        options.map((o, i) => `${i + 1}. ${o}`).join("\n") +
        "\n\nEnter option number:";

    while (true) {
        const ans = prompt(msg, "");
        if (ans === null) return null;
        const idx = parseInt(ans, 10);
        if (idx >= 1 && idx <= options.length) return options[idx - 1];
        alert("Please enter a valid option number.");
    }
}

function buildProviderRatesForLms(rateType, providerRate, waitTime, noShow) {
    const parts = [];
    const rate = formatMoneyValue(providerRate);

    if (rate) {
        if (rateType === "mile") {
            parts.push(`$${rate}/mile`);
        } else {
            parts.push(`Flat Rate $${rate}`);
        }
    }

    const wait = formatMoneyValue(waitTime);
    if (wait) parts.push(`Wait Time $${wait}/hr`);

    const ns = formatMoneyValue(noShow);
    if (ns) parts.push(`No Show $${ns}`);

    return parts.join(", ");
}

async function submitTransportNoShowToLms(payload) {
    // Use Tampermonkey's cross-site request so the existing LMS login cookie is sent reliably.
    // A normal browser fetch from Dynamics can miss the LMS session because it is cross-domain.
    return new Promise((resolve, reject) => {
        GM_xmlhttpRequest({
            method: "POST",
            url: LMS_TRANSPORT_NOSHOW_API_URL,
            headers: {
                "Content-Type": "application/json",
                "X-Requested-With": "XMLHttpRequest"
            },
            data: JSON.stringify(payload),
            withCredentials: true,
            anonymous: false,
            onload: function(res) {
                let data = {};
                try {
                    data = JSON.parse(res.responseText || "{}");
                } catch (e) {
                    reject(new Error("LMS returned a non-JSON response."));
                    return;
                }

                if (data.login_required) {
                    window.open(LMS_LOGIN_URL, "_blank");
                    reject(new Error(data.error || "Please sign into LMS first, then submit again."));
                    return;
                }

                if (data.duplicate) {
                    const latest = data.latest || {};
                    const ok = confirm(
                        "This referral already exists in LMS.\n\n" +
                        `Latest: ${latest.referral || ""} (${latest.status || ""})\n\n` +
                        "Submit again anyway?"
                    );
                    if (!ok) {
                        reject(new Error("Submit cancelled because referral already exists."));
                        return;
                    }

                    payload.force_resubmit = true;
                    submitTransportNoShowToLms(payload).then(resolve).catch(reject);
                    return;
                }

                if (res.status < 200 || res.status >= 300 || !data.ok) {
                    const msg = data.error || (data.errors ? Object.values(data.errors).join("\n") : "LMS submit failed.");
                    reject(new Error(msg));
                    return;
                }

                resolve(data);
            },
            onerror: function() {
                reject(new Error("Could not reach the LMS API."));
            },
            ontimeout: function() {
                reject(new Error("LMS API request timed out."));
            }
        });
    });
}

// Get DOS the same way the Payer Emails button gets its default date.
// If Dynamics has no Start Date value, return a blank string.
function getDateOfServiceForLmsApi() {
    const startDateInput =
        document.querySelector('input[aria-label="Start Date"]') ||
        document.querySelector('input[aria-label="Date of Start Date"]') ||
        document.querySelector('input[placeholder="---"][role="combobox"]');

    return startDateInput ? startDateInput.value.trim() : "";
}

function normalizeDosForDateInput(dosText) {
    const raw = String(dosText || "").trim();
    if (!raw) return "";

    // Already in ISO format.
    let m = raw.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if (m) {
        const y = parseInt(m[1], 10);
        const mo = parseInt(m[2], 10);
        const d = parseInt(m[3], 10);
        const check = new Date(y, mo - 1, d);
        if (
            check.getFullYear() === y &&
            check.getMonth() === mo - 1 &&
            check.getDate() === d
        ) {
            return `${String(y).padStart(4, "0")}-${String(mo).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
        }
        return "";
    }

    // Dynamics commonly displays M/D/YYYY or MM/DD/YYYY.
    m = raw.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (!m) return "";

    const mo = parseInt(m[1], 10);
    const d = parseInt(m[2], 10);
    const y = parseInt(m[3], 10);
    const check = new Date(y, mo - 1, d);

    if (
        check.getFullYear() !== y ||
        check.getMonth() !== mo - 1 ||
        check.getDate() !== d
    ) {
        return "";
    }

    return `${String(y).padStart(4, "0")}-${String(mo).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
}

function isDateOfServiceToday(dosText) {
    const raw = String(dosText || "").trim();
    if (!raw) return false;

    let year, month, day;

    // Most Dynamics date fields are shown as M/D/YYYY or MM/DD/YYYY.
    let m = raw.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (m) {
        month = parseInt(m[1], 10);
        day = parseInt(m[2], 10);
        year = parseInt(m[3], 10);
    } else {
        // Also support YYYY-MM-DD if Dynamics/browser returns ISO-style text.
        m = raw.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
        if (!m) return false;
        year = parseInt(m[1], 10);
        month = parseInt(m[2], 10);
        day = parseInt(m[3], 10);
    }

    const today = new Date();
    return (
        year === today.getFullYear() &&
        month === (today.getMonth() + 1) &&
        day === today.getDate()
    );
}

function showLmsNoShowModal() {
    return new Promise((resolve) => {
        const options = [
            "No Phone Number for claimant",
            "Claimant does not want rideshare",
            "No drivers in area",
            "Low availability",
            "Medium availability",
            "High availability",
            "Cost more than provider",
            "Referral has wait time",
            "To Large (Over 35miles one way)",
            "Not allowed for this payer",
            "This is a TECH",
            "ALERT shows this is PREFERRED provider"
        ];

        const overlay = document.createElement("div");
        overlay.style.position = "fixed";
        overlay.style.inset = "0";
        overlay.style.background = "rgba(0,0,0,0.35)";
        overlay.style.zIndex = "100000";
        overlay.style.display = "flex";
        overlay.style.alignItems = "center";
        overlay.style.justifyContent = "center";

        const modal = document.createElement("div");
        modal.style.width = "460px";
        modal.style.maxWidth = "calc(100vw - 40px)";
        modal.style.background = "#fff";
        modal.style.color = "#000";
        modal.style.borderRadius = "12px";
        modal.style.boxShadow = "0 15px 40px rgba(0,0,0,0.35)";
        modal.style.padding = "18px";
        modal.style.fontFamily = "Arial, sans-serif";

        const title = document.createElement("div");
        title.innerText = "Submit Rates to LMS";
        title.style.fontSize = "20px";
        title.style.fontWeight = "bold";
        title.style.marginBottom = "14px";
        modal.appendChild(title);

        function addLabel(text) {
            const lbl = document.createElement("label");
            lbl.innerText = text;
            lbl.style.display = "block";
            lbl.style.fontWeight = "bold";
            lbl.style.marginTop = "10px";
            lbl.style.marginBottom = "5px";
            modal.appendChild(lbl);
            return lbl;
        }

        addLabel("Provider Name");
        const providerInput = document.createElement("input");
        providerInput.type = "text";
        providerInput.value = "";
        providerInput.style.width = "100%";
        providerInput.style.boxSizing = "border-box";
        providerInput.style.padding = "8px";
        modal.appendChild(providerInput);

        addLabel("DOS");
        const dosInput = document.createElement("input");
        dosInput.type = "date";
        dosInput.required = true;
        dosInput.value = normalizeDosForDateInput(getDateOfServiceForLmsApi());
        dosInput.style.width = "100%";
        dosInput.style.boxSizing = "border-box";
        dosInput.style.padding = "8px";
        dosInput.style.border = "1px solid #777";
        dosInput.style.borderRadius = "2px";
        dosInput.style.background = "#fff";
        dosInput.style.fontFamily = "inherit";
        dosInput.style.fontSize = "14px";
        modal.appendChild(dosInput);

        addLabel("Additional Referrals (Optional)");
        const additionalReferralsInput = document.createElement("input");
        additionalReferralsInput.type = "text";
        additionalReferralsInput.value = "";
        additionalReferralsInput.placeholder = "Example: -10, -11, -12";
        additionalReferralsInput.style.width = "100%";
        additionalReferralsInput.style.boxSizing = "border-box";
        additionalReferralsInput.style.padding = "8px";
        additionalReferralsInput.style.border = "1px solid #777";
        additionalReferralsInput.style.borderRadius = "2px";
        additionalReferralsInput.style.background = "#fff";
        additionalReferralsInput.style.fontFamily = "inherit";
        additionalReferralsInput.style.fontSize = "14px";
        modal.appendChild(additionalReferralsInput);

        addLabel("Is it a rush?");
        const rushWrap = document.createElement("div");
        rushWrap.style.display = "flex";
        rushWrap.style.gap = "16px";

        const defaultRushYes = isDateOfServiceToday(dosInput.value);

        rushWrap.innerHTML = `
            <label style="font-weight:normal;"><input type="radio" name="lmsRush" value="Yes" ${defaultRushYes ? "checked" : ""}> Yes</label>
            <label style="font-weight:normal;"><input type="radio" name="lmsRush" value="No" ${defaultRushYes ? "" : "checked"}> No</label>
        `;
        modal.appendChild(rushWrap);

        // If the DOS is entered or corrected in the modal, update the Rush default too.
        dosInput.addEventListener("change", () => {
            const isToday = isDateOfServiceToday(dosInput.value);
            const yesRadio = modal.querySelector('input[name="lmsRush"][value="Yes"]');
            const noRadio = modal.querySelector('input[name="lmsRush"][value="No"]');
            if (yesRadio) yesRadio.checked = isToday;
            if (noRadio) noRadio.checked = !isToday;
        });

        addLabel("Did you check UBER/Lyft?");
        const uberSelect = document.createElement("select");
        uberSelect.style.width = "100%";
        uberSelect.style.boxSizing = "border-box";
        uberSelect.style.padding = "8px";
        uberSelect.innerHTML = `<option value="">-- Select an option --</option>` +
            options.map(o => `<option value="${o.replace(/&/g, "&amp;").replace(/"/g, "&quot;")}">${o}</option>`).join("");
        modal.appendChild(uberSelect);

        addLabel("Comments (Optional)");
        const commentsInput = document.createElement("textarea");
        commentsInput.rows = 4;
        commentsInput.value = "";
        commentsInput.placeholder = "Enter any additional comments...";
        commentsInput.style.width = "100%";
        commentsInput.style.boxSizing = "border-box";
        commentsInput.style.padding = "8px";
        commentsInput.style.border = "1px solid #777";
        commentsInput.style.borderRadius = "2px";
        commentsInput.style.background = "#fff";
        commentsInput.style.fontFamily = "inherit";
        commentsInput.style.fontSize = "14px";
        commentsInput.style.resize = "vertical";
        modal.appendChild(commentsInput);

        const uberCostWrap = document.createElement("div");
        uberCostWrap.style.display = "none";
        addLabel("UBER Cost Round Trip").style.display = "none";
        const uberCostLabel = modal.lastChild;
        const uberCostInput = document.createElement("input");
        uberCostInput.type = "number";
        uberCostInput.step = "0.01";
        uberCostInput.min = "0";
        uberCostInput.style.width = "100%";
        uberCostInput.style.boxSizing = "border-box";
        uberCostInput.style.padding = "8px";
        uberCostWrap.appendChild(uberCostInput);
        modal.appendChild(uberCostWrap);

        uberSelect.addEventListener("change", () => {
            const needsCost = uberSelect.value === "Medium availability" || uberSelect.value === "High availability";
            uberCostLabel.style.display = needsCost ? "block" : "none";
            uberCostWrap.style.display = needsCost ? "block" : "none";
            if (!needsCost) uberCostInput.value = "";
        });

        const error = document.createElement("div");
        error.style.color = "#dc2626";
        error.style.fontWeight = "bold";
        error.style.marginTop = "10px";
        error.style.minHeight = "18px";
        modal.appendChild(error);

        const buttons = document.createElement("div");
        buttons.style.display = "flex";
        buttons.style.justifyContent = "flex-end";
        buttons.style.gap = "10px";
        buttons.style.marginTop = "16px";

        const cancelBtn = document.createElement("button");
        cancelBtn.type = "button";
        cancelBtn.innerText = "Cancel";
        cancelBtn.style.padding = "8px 14px";

        const submitBtn = document.createElement("button");
        submitBtn.type = "button";
        submitBtn.innerText = "Submit";
        submitBtn.style.padding = "8px 14px";
        submitBtn.style.background = "#22c55e";
        submitBtn.style.color = "#fff";
        submitBtn.style.border = "none";
        submitBtn.style.borderRadius = "6px";
        submitBtn.style.fontWeight = "bold";

        buttons.appendChild(cancelBtn);
        buttons.appendChild(submitBtn);
        modal.appendChild(buttons);
        overlay.appendChild(modal);
        document.body.appendChild(overlay);
        providerInput.focus();
        providerInput.select();

        function close(value) {
            overlay.remove();
            resolve(value);
        }

        cancelBtn.onclick = () => close(null);
        overlay.addEventListener("click", (e) => { if (e.target === overlay) close(null); });

        submitBtn.onclick = () => {
            const providerName = providerInput.value.trim();
            const additionalReferrals = additionalReferralsInput.value.trim();
            const rush = modal.querySelector('input[name="lmsRush"]:checked')?.value || "No";
            const uberOption = uberSelect.value;
            const uberCost = uberCostInput.value.trim();

            if (!providerName) { error.innerText = "Provider Name is required."; return; }
            if (!dosInput.value) { error.innerText = "DOS is required."; dosInput.focus(); return; }
            if (additionalReferrals && !/^[0-9,\-\s]+$/.test(additionalReferrals)) {
                error.innerText = "Additional Referrals can only contain numbers, hyphens, commas, and spaces.";
                return;
            }
            if (!uberOption) { error.innerText = "Please choose a UBER/Lyft option."; return; }
            if ((uberOption === "Medium availability" || uberOption === "High availability") && (!uberCost || isNaN(parseFloat(uberCost)))) {
                error.innerText = "UBER Cost is required for Medium/High availability.";
                return;
            }

            close({
                providerName,
                dos: dosInput.value.trim(),
                additionalReferrals,
                rush,
                uberOption,
                uberCost,
                comments: commentsInput.value.trim()
            });
        };
    });
}

// Function to show calculator UI
function showCalculatorBox() {
    // Check if the page title contains 'Referral: Information:'
    if (!document.title.includes("Referral: Information:")) {
        showMessage("Must be in a referral", false);
        return;
    }

    // Look for the tab with title "Billing"
    const billingTab = document.querySelector('li[role="tab"][title="Billing"]');
    if (billingTab) {
        billingTab.click();
        console.log('Clicked "Billing" tab before showing calculator.');
        setTimeout(showCalculatorUI, 1000);
    } else {
        console.warn('"Billing" tab not found.');
        showCalculatorUI(); // fallback
    }
}

// LMS owns calculator thresholds. Each opening uses a fresh, validated snapshot.
const LMS_MARGIN_SETTINGS_URL = 'https://lowmargin.mtoysystems.com/api/margin_settings.php';
let lmsMarginOpenSequence = 0;

function validateLmsMarginSettings(data) {
    if (!data || data.ok !== true || data.schema_version !== 1 || !Array.isArray(data.rules) || !data.default) {
        throw new Error('LMS returned invalid margin settings.');
    }
    const prefixes = new Set();
    for (const rule of [data.default, ...data.rules]) {
        if (!rule || typeof rule.header_prefix !== 'string' || prefixes.has(rule.header_prefix)) {
            throw new Error('LMS returned invalid header rules.');
        }
        prefixes.add(rule.header_prefix);
        for (const mode of ['regular', 'higher']) {
            const range = rule[mode];
            if (!range || !['approval_max', 'red_max', 'green_min'].every(key =>
                typeof range[key] === 'number' && Number.isFinite(range[key]) && range[key] >= -10000 && range[key] <= 100
            ) || range.red_max >= range.green_min) {
                throw new Error('LMS returned invalid margin ranges.');
            }
        }
    }
    if (data.default.header_prefix !== '' || data.rules.some(rule => rule.header_prefix === '')) {
        throw new Error('LMS returned an invalid default rule.');
    }
    return data;
}

function fetchLmsMarginSettings() {
    return new Promise((resolve, reject) => {
        GM_xmlhttpRequest({
            method: 'GET',
            url: `${LMS_MARGIN_SETTINGS_URL}?_=${Date.now()}`,
            headers: { 'Accept': 'application/json', 'Cache-Control': 'no-cache' },
            withCredentials: true,
            anonymous: false,
            timeout: 15000,
            onload(response) {
                try {
                    if (response.status === 401) throw new Error('Sign into LMS, then reopen the margin calculator.');
                    if (response.status !== 200) throw new Error(`LMS margin settings request failed (HTTP ${response.status}).`);
                    resolve(validateLmsMarginSettings(JSON.parse(response.responseText)));
                } catch (error) {
                    reject(new Error(error instanceof SyntaxError ? 'LMS did not return settings. Sign into LMS and try again.' : error.message));
                }
            },
            onerror: () => reject(new Error('Could not reach LMS. Reopen the calculator to try again.')),
            ontimeout: () => reject(new Error('LMS settings request timed out. Reopen the calculator to try again.')),
            onabort: () => reject(new Error('LMS settings request was cancelled.'))
        });
    });
}

function findLmsMarginRule(settings, headerText) {
    return settings.rules.filter(rule => headerText.startsWith(rule.header_prefix))
        .sort((a, b) => b.header_prefix.length - a.header_prefix.length)[0] || settings.default;
}

function lmsMarginStatus(rawMargin, range) {
    if (!Number.isFinite(rawMargin)) return { color: 'black', approval: false };
    const margin = Number(rawMargin.toFixed(2));
    return {
        color: margin <= range.red_max ? 'red' : margin < range.green_min ? 'goldenrod' : 'green',
        approval: margin <= range.approval_max
    };
}

async function showCalculatorUI() {
    const sequence = ++lmsMarginOpenSequence;
    document.getElementById('calcBox')?.remove();
    document.getElementById('lmsMarginLoading')?.remove();
    const panel = document.createElement('div');
    panel.id = 'lmsMarginLoading';
    Object.assign(panel.style, { position: 'fixed', top: '10%', right: '20px', zIndex: '10001',
        background: 'white', color: 'black', border: '2px solid #333', borderRadius: '10px', padding: '20px', maxWidth: '420px' });
    const status = document.createElement('p');
    status.textContent = 'Loading current margin settings from LMS…';
    const close = document.createElement('button');
    close.type = 'button';
    close.textContent = 'Close';
    close.onclick = () => { if (sequence === lmsMarginOpenSequence) ++lmsMarginOpenSequence; panel.remove(); };
    panel.append(status, close);
    document.body.appendChild(panel);
    try {
        const settings = await fetchLmsMarginSettings();
        if (sequence !== lmsMarginOpenSequence || !panel.isConnected) return;
        if (!document.title.includes('Referral: Information:')) {
            status.textContent = 'Return to a referral and reopen the calculator.';
            return;
        }
        panel.remove();
        renderCalculatorUI(settings);
    } catch (error) {
        if (sequence !== lmsMarginOpenSequence || !panel.isConnected) return;
        status.textContent = `${error.message} Current thresholds are required before calculating.`;
        const login = document.createElement('a');
        login.href = LMS_LOGIN_URL;
        login.target = '_blank';
        login.rel = 'noopener noreferrer';
        login.textContent = 'Open LMS';
        login.style.marginLeft = '12px';
        panel.appendChild(login);
    }
}


function renderCalculatorUI(marginSettings) {
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
    box.style.minWidth = "500px";
    box.style.maxWidth = "500px";
    box.style.color = "black";

    // Auto height now that Higher Rates is collapsed by default; still scroll if content overflows.
    box.style.height = "auto";
    box.style.maxHeight = "calc(100vh - 40px)";
    box.style.overflowY = "auto"; // vertical scroll if content overflows

    // Rest of your existing code remains unchanged...

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

// Provider Rate column
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

// --- Provider Load Fee (only shows when Per Mile + Load Fee is present) ---
const providerLoadFeeWrap = document.createElement("div");
providerLoadFeeWrap.style.display = "none"; // hidden by default

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

// Wait Time column
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

// Add both columns to the row
twoColumnWrapper.appendChild(providerWrapper);
twoColumnWrapper.appendChild(waitWrapper);

// No Show box used for LMS API submit
const noShowTopWrapper = document.createElement("div");
noShowTopWrapper.style.marginTop = "15px";
noShowTopWrapper.style.padding = "10px";
noShowTopWrapper.style.border = "2px solid #facc15";
noShowTopWrapper.style.borderRadius = "8px";
noShowTopWrapper.style.background = "#fffef0";

const noShowTopLabel = document.createElement("label");
noShowTopLabel.innerText = "No Show:";
noShowTopLabel.style.fontWeight = "bold";
noShowTopLabel.style.display = "block";
noShowTopLabel.style.marginBottom = "6px";
noShowTopWrapper.appendChild(noShowTopLabel);

const noShowTopInput = document.createElement("input");
noShowTopInput.type = "number";
noShowTopInput.step = "0.01";
noShowTopInput.min = "0";
noShowTopInput.style.width = "100%";
noShowTopInput.value = "";
noShowTopInput.placeholder = "Enter provider no show amount";
let noShowTopTouched = false;
noShowTopInput.addEventListener("input", () => { noShowTopTouched = true; });
noShowTopWrapper.appendChild(noShowTopInput);


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

    // Higher Rates section is collapsed by default to keep the transport submit view clean.
    let higherRatesVisible = false;
    higherHeader.style.display = "none";
    higherInputsWrapper.style.display = "none";
    higherResult.style.display = "none";

    const toggleHigherRatesButton = document.createElement("button");
    toggleHigherRatesButton.type = "button";
    toggleHigherRatesButton.innerText = "▼ Show Higher Rates Calculator";
    toggleHigherRatesButton.style.marginTop = "18px";
    toggleHigherRatesButton.style.padding = "8px 12px";
    toggleHigherRatesButton.style.border = "1px solid #999";
    toggleHigherRatesButton.style.borderRadius = "6px";
    toggleHigherRatesButton.style.background = "#f3f4f6";
    toggleHigherRatesButton.style.color = "#111";
    toggleHigherRatesButton.style.fontWeight = "bold";
    toggleHigherRatesButton.style.cursor = "pointer";
    toggleHigherRatesButton.onclick = () => {
        higherRatesVisible = !higherRatesVisible;
        higherHeader.style.display = higherRatesVisible ? "block" : "none";
        higherInputsWrapper.style.display = higherRatesVisible ? "flex" : "none";
        higherResult.style.display = higherRatesVisible ? "block" : "none";
        toggleHigherRatesButton.innerText = higherRatesVisible
            ? "▲ Hide Higher Rates Calculator"
            : "▼ Show Higher Rates Calculator";
    };

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
        noShowTopInput.value = "";
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


// Submit Provider No Show to LMS API
const submitLmsNoShowButton = createModernButton("Submit to LMS", "#22c55e", "#4ade80");
submitLmsNoShowButton.style.float = "right";
submitLmsNoShowButton.style.marginLeft = "15px";
submitLmsNoShowButton.style.color = "#fff";
submitLmsNoShowButton.style.fontSize = "16px";
submitLmsNoShowButton.style.padding = "12px 18px";

submitLmsNoShowButton.onclick = async () => {
    try {
        await ensureBillingTabOpen();
    } catch (err) {
        alert(err.message || "Could not open the Billing tab.");
        return;
    }

    const referral = getReferralNumberFromHeader();
    if (!referral) {
        alert("Could not find referral number from the page header.");
        return;
    }

    const rateType = document.querySelector('input[name="rateType"]:checked')?.value || "flat";
    const providerRate = input.value.trim();
    const waitTime = waitTimeInput.value.trim();
    const noShowValue = noShowTopInput.value.trim();
    const billingWaitTime = getBillingProductPrice("Wait Time");

    if (!providerRate || isNaN(parseFloat(providerRate))) {
        alert("Please enter a valid Provider Rate.");
        return;
    }

    if (!noShowValue || isNaN(parseFloat(noShowValue))) {
        alert("Please enter a valid No Show amount.");
        return;
    }

    const noShowNumber = parseFloat(noShowValue);
    const waitTimeNumber = waitTime === "" ? 0 : parseFloat(waitTime);

    if (isNaN(waitTimeNumber) || waitTimeNumber < 0) {
        alert("Please enter a valid Wait Time amount.");
        return;
    }

    // The API silently decides whether the request qualifies for auto-approval.

    let margin = getCurrentMarginFromResult(result);
    const noBillingRatesFound = result.innerText.includes("No billing rates found.");

    if (!margin && noBillingRatesFound) {
        margin = "-999";
    } else if (!margin) {
        alert("Please calculate margin first by entering the provider rate.");
        return;
    }

    const modalAnswers = await showLmsNoShowModal();
    if (!modalAnswers) return;

    const providerName = modalAnswers.providerName;
    const dos = modalAnswers.dos || "";
    const additionalReferrals = modalAnswers.additionalReferrals || "";
    const rush = modalAnswers.rush;
    const uberOption = modalAnswers.uberOption;
    const uberCost = modalAnswers.uberCost || "";
    const userComments = modalAnswers.comments || "";

    const providerRates = buildProviderRatesForLms(rateType, providerRate, waitTime, noShowValue);

    const payload = {
        referral,
        dos,
        additional_referrals: additionalReferrals,
        rush,
        margin,
        uber_option: uberOption,
        uber_cost: uberCost,
        provider_name: providerName,
        provider_rates: providerRates,
        provider_wait_time: waitTimeNumber,
        billing_wait_time: billingWaitTime,
        provider_no_show: noShowNumber,
        comments: userComments
            ? "API submission from Margin Calc\n\n" + userComments
            : "API submission from Margin Calc"
    };

    submitLmsNoShowButton.disabled = true;
    submitLmsNoShowButton.innerText = "Submitting...";

    try {
        const data = await submitTransportNoShowToLms(payload);

        const waitTimeRateNotFound =
            waitTimeNumber > 0 &&
            billingWaitTime <= 0 &&
            data.auto_approval_enabled === true &&
            data.matched_enabled_rule === true;

        if (waitTimeRateNotFound) {
            alert("Wait time rate not found being sent for manual review");
            showMessage("Submitted for manual review.", true);
        } else {
            showMessage("Submitted successfully! Check LMS to view management response.", true);
        }
    } catch (err) {
        alert(err.message || "Submit failed.");
        showMessage(err.message || "Submit failed.", false);
    } finally {
        submitLmsNoShowButton.disabled = false;
        submitLmsNoShowButton.innerText = "Submit to LMS (No Show)";
    }
};


    // Create the "Request Rates" button
const requestRatesButton = createModernButton("Request Rates", "#22c55e", "#4ade80");

// Button click behavior
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

      if (accountProductText.includes("Transport") && !isNaN(quantity)) {
        miles = quantity;
      }

      if (accountProductText.includes("Load Fee") && !isNaN(quantity)) {
        loadFeeQuantity = quantity;
      }
    }
  });

  const finalParts = buildPartsString(productInputs, {}, miles, loadFeeQuantity);

  const finalText = "Request rates " + finalParts;

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

    setTimeout(() => {
      copiedMsg.remove();
    }, 1500);
  });
};

// Create the "Apply Rates" button
const applyRatesButton = createModernButton("Apply & Staff", "#a855f7", "#c084fc");

// Button click behavior
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

      if (accountProductText.includes("Transport") && !isNaN(quantity)) {
        miles = quantity;
      }

      if (accountProductText.includes("Load Fee") && !isNaN(quantity)) {
        loadFeeQuantity = quantity;
      }
    }
  });

  const finalParts = buildPartsString(productInputs, {}, miles, loadFeeQuantity);

  const finalText = "Apply rates " + finalParts + " // Advise in Staffing email";

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

    setTimeout(() => {
      copiedMsg.remove();
    }, 1500);
  });
};

//Homelink Button
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

  // Extract Wait Time and No Show like Boomerang button
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

  // Use the same buildPartsString function
  const partsString = buildPartsString(productInputs, quantities, 0, quantities["Load Fee"]);
  const goatString = `**Enter in Goat as ${partsString}`;

// Build extras text
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


  // Copy and notify
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

    setTimeout(() => {
      copiedMsg.remove();
    }, 1500);
  });
};


// Create the "Boomerang" button
const boomerangButton = createModernButton("Boomerang Request & Staff", "#f97316", "#fb923c");

// Button click behavior
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

  // Normalize Transport inputs for contract case
  if (
    ["Transport Ambulatory", "Transport Wheelchair", "Transport Stretcher, ALS & BLS"].includes(label)
  ) {
    value = value.toLowerCase() === "contract rates" ? "Contract rates per mile" : value;
  }

  // Handle Wait Time
  if (label === "Wait Time" && value !== "") {
    waitTimeText = value.toLowerCase() === "contract rates" ? "Contract" : value;
  }

  // Handle No Show
  if (label === "No Show" && value !== "") {
    noShowText = value.toLowerCase() === "contract rates" ? "Contract" : value;
  }

  // Set flatRateValue based on normalized Transport value
  if (
    ["Transport Ambulatory", "Transport Wheelchair", "Transport Stretcher, ALS & BLS"].includes(label)
  ) {
    flatRateValue = value === "Contract rates per mile" ? value : parseFloat(value);
  }
});


  if (
  (!flatRateValue || isNaN(flatRateValue)) &&
  flatRateValue !== "Contract rates per mile"
) {
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


  // Build extras text
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

  // Copy to clipboard and show confirmation
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

    setTimeout(() => {
      copiedMsg.remove();
    }, 1500);
  });
};

//Reusable partstring
function buildPartsString(productInputs, quantities, miles, loadFeeQuantity) {
  // Prefill "One Way Surcharge" input with the quantity found, if available
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
  // For transport, user wants ONLY the rate per mile (no "x ## miles")
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

        // Exclude "Rush Fee" from totalBilled
        if (productsToTrack.includes(product) &&
            product !== "Rush Fee" &&
            product !== "Weekend Holiday" &&
            product !== "Wheelchair Rental" &&
            product !== "Airport Pickup Fee" &&
            product !== "After Hours Fee"
            || product === "Wait Time"
           ) {
            if (!isNaN(totalValue)) {
                if (product === "Wait Time" && totalValue === 0) {
                    const priceCell = row.querySelector('[col-id="gtt_price"]');
                    const priceText = priceCell?.innerText.trim().replace(/[^0-9.-]+/g, '') || "0";
                    const priceValue = parseFloat(priceText);

                    // Wait Time always counts as quantity 1
                    if (!isNaN(priceValue) && priceValue > 0) {
                        totalBilled += priceValue;
                        quantities["Wait Time"] = 1;
                    }
                } else {
                    totalBilled += totalValue;
                }
            }
        }

        // Handle quantity per product
        if (qtyCell) {
            const qtyVal = parseFloat(qtyCell.innerText.trim().replace(/[^0-9.-]+/g, ''));
            if (!isNaN(qtyVal)) {
                if (!quantities[product]) quantities[product] = 0;
                quantities[product] += qtyVal;
            }
        }

        // Track found products
        if (productsToTrack.includes(product)) {
            foundProducts.add(product);
        }

        // Set quantity if rateType is "mile" and applicable product
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

        // Always add the remaining extras
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
        wrapper.style.flex = "1 1 48%"; // two columns with spacing

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

        const inputField = document.createElement("input");
        inputField.type = "text";
        inputField.style.width = "100%";
        inputField.addEventListener("input", calculateMargin);

        // Prefill "One Way Surcharge" input with the quantity found, if available
        if (product === "One Way Surcharge" && quantities[product] !== undefined) {
            inputField.value = quantities[product];
            inputField.readOnly = true; // Make it read-only
            inputField.style.background = "#eee"; // Optional: visually show it's disabled
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
            result.innerText = "No billing rates found. LMS submission will use a -999 margin.";
            result.style.color = "black";
            targetLabel.innerHTML = "";
            higherResult.innerText = "";
            return;
        }

        const loadFeeQty = quantities["Load Fee"] || 0;

        // Only show Provider Load Fee input when Per Mile + Load Fee exists
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
        const headerElement = document.querySelector('[id^="formHeaderTitle"]');
        const headerText = headerElement?.textContent?.trim() || "";
        const marginRule = findLmsMarginRule(marginSettings, headerText);
        const regularStatus = lmsMarginStatus(margin, marginRule.regular);
        const marginColor = regularStatus.color;
        const approvalNote = regularStatus.approval
            ? '<br><span style="color: red; font-weight: bold;">Seek Management Approval</span>' : '';

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

        const target = totalBilled * (1 - marginRule.regular.green_min / 100);
        targetLabel.innerHTML = `<span>Target to pay this or less: $${target.toFixed(2)}</span>`;

        const transportProducts = ["Transport Ambulatory", "Transport Wheelchair", "Transport Stretcher, ALS & BLS"];
        let activeTransportRate = 0;
        for (const tp of transportProducts) {
            const inputElem = productInputs[tp];
            if (inputElem && inputElem.value.trim() !== "") {
                activeTransportRate = parseFloat(inputElem.value);
                if (isNaN(activeTransportRate)) activeTransportRate = 0;
                break;
            }
        }

        let higherTotal = 0;
let alreadyCounted = new Set(); // Track items added once

foundProducts.forEach(product => {
    // We now INCLUDE Wait Time in higherTotal
    // Prevent duplicates for Rush Fee, Tolls, etc.
    if (["Rush Fee", "Tolls", "Other", "Assistance Fee", "Passenger Fee", "Miscellaneous Dead Miles"].includes(product)) {
        if (alreadyCounted.has(product)) return; // Skip if already counted
        alreadyCounted.add(product);
    }

    let enteredValueRaw = productInputs[product]?.value;
    let enteredValue = parseFloat(enteredValueRaw);
    let qty = product === "Wait Time" ? 1 : (quantities[product] || 0);

    //Enter Contract Rates
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

// Check transport inputs and set activeTransportRate
const transportProducts2 = ["Transport Ambulatory", "Transport Wheelchair", "Transport Stretcher, ALS & BLS"];

for (let transport of transportProducts2) {
    const transportInputValue = (productInputs[transport]?.value || "").toLowerCase();

    if (transportInputValue.includes("contract")) {
        // Find gtt_price for this transport
        const rows = document.querySelectorAll('div[row-index]');
        for (let row of rows) {
            const productCell = row.querySelector('[col-id="gtt_accountproduct"]');
            const priceCell = row.querySelector('[col-id="gtt_price"]');
            if (productCell && priceCell && productCell.innerText.trim() === transport) {
                let priceText = priceCell.innerText.trim().replace(/[^0-9.-]+/g, '') || "0";
                let priceValue = parseFloat(priceText);
                if (!isNaN(priceValue)) {
                    activeTransportRate = priceValue;
                }
                break;
            }
        }
        break; // Stop after first matching contract transport
    } else if (!isNaN(parseFloat(transportInputValue))) {
        // Use numeric value from input
        activeTransportRate = parseFloat(transportInputValue);
        break;
    }
}

// Default to 0 if no rate found
if (isNaN(activeTransportRate)) activeTransportRate = 0;

// Process each product
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
        higherTotal += enteredValue * qty;
    } else {
        higherTotal += enteredValue * qty;
    }
}
});

        if (!Number.isFinite(higherTotal) || higherTotal <= 0) {
            higherResult.innerText = 'Enter higher billing rates to calculate their margin.';
            return;
        }
        const higherMargin = 100 - ((paidAmount / higherTotal) * 100);
        const higherStatus = lmsMarginStatus(higherMargin, marginRule.higher);
        const higherMarginColor = higherStatus.color;
        const higherApprovalNote = higherStatus.approval
            ? '<br><span style="color: red; font-weight: bold;">Seek Management Approval</span>' : '';

        higherResult.innerHTML = `
            <span>Total Using Higher Rates: $${higherTotal.toFixed(2)}</span>
            <span style="color: ${higherMarginColor}; font-weight: bold;">Margin: ${higherMargin.toFixed(2)}%</span>${higherApprovalNote}
        `.trim();
    };

    input.addEventListener("input", calculateMargin);
    waitTimeInput.addEventListener("input", calculateMargin);
    noShowTopInput.addEventListener("input", calculateMargin);
    providerLoadFeeInput.addEventListener("input", calculateMargin);
    flatRadio.addEventListener("change", () => {
        input.value = "";
        waitTimeInput.value = "";
        providerLoadFeeInput.value = "";
        noShowTopInput.value = "";
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
        noShowTopInput.value = "";
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

    box.appendChild(closeButton);
    box.appendChild(modeLabel);
    box.appendChild(flatRadio);
    box.appendChild(flatLabel);
    box.appendChild(mileRadio);
    box.appendChild(mileLabel);
    box.appendChild(twoColumnWrapper);
    box.appendChild(noShowTopWrapper);
    box.appendChild(result);
    box.appendChild(targetLabel);
    box.appendChild(toggleHigherRatesButton);
    box.appendChild(higherHeader);
    box.appendChild(higherInputsWrapper);
    box.appendChild(higherResult);
    box.appendChild(resetButton);
    box.appendChild(submitLmsNoShowButton);
    // box.appendChild(requestRatesButton);


    document.body.appendChild(box);

// Make calcBox draggable (but not from inside inputs/buttons)
let isDragging = false;
let offsetX, offsetY;

box.addEventListener('mousedown', function (e) {
    const isInteractive = ['INPUT', 'TEXTAREA', 'BUTTON'].includes(e.target.tagName);
    if (isInteractive) return; // don't drag if inside input/button/textarea

    // Only drag if clicking on top 40px of the box
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
        box.style.transform = ''; // cancel translateX on drag
    }
});

document.addEventListener('mouseup', function () {
    isDragging = false;
    box.style.cursor = 'grab';
});


}

    // Add event listener for the calculator button
    const calculatorButton = document.querySelector('#yourCalculatorButtonSelector'); // Replace with actual button selector
    if (calculatorButton) {
        calculatorButton.addEventListener('click', showCalculatorBox);
    }

    // Function to copy claimant name
function copyClaimantName() {
    var elementToCopy = document.querySelector(
        'div[data-id="gtt_claimantid.fieldControl-LookupResultsDropdown_gtt_claimantid_selected_tag_text"]'
    );

    if (elementToCopy) {
        var textToCopy =
            (elementToCopy.getAttribute('title') || elementToCopy.textContent).trim();

        GM_setClipboard(textToCopy);
        showMessage(`Copied: "${textToCopy}" successfully.`);
        console.log('Copied to clipboard:', textToCopy);

        var serviceProviderTab =
            document.querySelector('li[role="tab"][title="Service Provider"]');

        if (serviceProviderTab) {
            serviceProviderTab.click();
            console.log('Clicked "Service Provider" tab.');
            waitForButtonAndClick();
        } else {
            console.error('"Service Provider" tab not found.');
        }
    } else {
        showMessage(
            'Claimant Name not found. Please make sure you are in a referral.',
            false
        );
        console.error('Claimant element not found.');
    }
}

    // Function to wait for the button to appear and click it
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

// Function to copy both claimant name and claim
function copyBoth() {
    var element1 = document.querySelector(
        'div[data-id="gtt_claimantid.fieldControl-LookupResultsDropdown_gtt_claimantid_selected_tag_text"]'
    );

    var element2 = document.querySelector(
        'div[data-id="gtt_claimid.fieldControl-LookupResultsDropdown_gtt_claimid_selected_tag_text"]'
    );

    var titleElement = document.querySelector('[id^="formHeaderTitle"]');

    var startDateInput =
        document.querySelector('input[aria-label="Start Date"]') ||
        document.querySelector('input[aria-label="Date of Start Date"]') ||
        document.querySelector('input[placeholder="---"][role="combobox"]');

    var startDateValue = startDateInput
        ? startDateInput.value.trim()
        : "";

    if (element1 && element2) {
        var text1 =
            (element1.getAttribute('title') || element1.textContent).trim();

        var text2 =
            (element2.getAttribute('title') || element2.textContent).trim();

        var headerTitle = titleElement
            ? titleElement.textContent.trim()
            : "";

        if (headerTitle.startsWith("4403-54316")) {
            alert(
                "Please combine staffing and/or auth requests into one email " +
                "(include multiple dates into one email)."
            );
        }

        var referralDate = prompt(
            "Please enter the referral date(s):",
            startDateValue
        );

        if (referralDate === null) {
            var textToCopy =
                `Claimant: ${text1} - Claim: ${text2} - on DOS:`;

            GM_setClipboard(textToCopy);
            showMessage(`Copied: "${textToCopy}" successfully.`);
            return;
        }

        if (!referralDate) {
            referralDate = "[No Date Provided]";
        }

        createDropdownMenu(
            text1,
            text2,
            referralDate,
            headerTitle
        );
    } else {
        showMessage(
            'Claimant Name & Claim# not found. Please make sure you are in a referral.',
            false
        );

        console.error('Missing elements:', {
            claimantElement: element1,
            claimElement: element2
        });
    }
}

//Create Options for Email templates
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

    // Base options (we will insert the 4474-related options dynamically)
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
const is11912 = headerTitle.startsWith("11912-");
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

// H claims for these three prefixes use L-Orchid
const isLOrchidPrefix = is4474 || is11525 || is11912;

console.log("headerTitle:", headerTitle);
console.log("payerText:", payerText);
console.log("is4474:", is4474, "is11525:", is11525, "isJBSPacking:", isJBSPacking, "use11525Rules:", use11525Rules);

if (isLOrchidPrefix && isHClaim) {
    fullOptions.splice(6, 0, "L-Orchid-CareWorks");
} else if (use11525Rules) {
    fullOptions.splice(6, 0, "JBS Request for Higher Rates");
} else if (is4474 || is11912) {
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
        "Convergence Higher Rate Request",
        "JBS Request for Higher Rates",
        "CareWorks Rate Request",
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
        "Convergence Higher Rate Request",
        "JBS Request for Higher Rates",
        "CareWorks Rate Request",
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
// If L-Orchid applies, do not show Standard Rate Request
if (isLOrchidPrefix && isHClaim) {
    exclusions.push("Standard Rate Request");
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
        } else if (/rate request/i.test(optionText) || /higher rates/i.test(optionText)) {
            buttonLabel = "💵 " + optionText;
        }

        const button = createModernButton(
            buttonLabel,   // only changes what user sees
            start, end,
            () => {
                finalizeCopy(claimant, claim, referralDate, optionText); // keeps real template name
                dropdownContainer.remove();
            }
        );

        button.style.width = "100%";
        dropdownContainer.appendChild(button);
    });

    const closeButton = createModernButton(
        "❌ Close",
        "#7f1d1d", "#f87171",
        () => dropdownContainer.remove()
    );
    closeButton.style.width = "100%";
    closeButton.style.marginTop = "10px";
    dropdownContainer.appendChild(closeButton);

    document.body.appendChild(dropdownContainer);
}

// Function to finalize the copy action
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

// Wait for SavePrimary button to appear
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


/* =================== EMAIL CONFIRMATION SAFEGUARDS =================== */

// Dynamics can render the referral-recipient rows and the confirmation-form
// checkboxes in different same-origin iframe documents.  Do NOT assume they
// live in one specific iframe.  Instead, walk every accessible document and
// search them independently.
function getAllAccessibleDynamicsDocuments() {
    const docs = [];
    const seen = new Set();

    function walk(doc) {
        if (!doc || seen.has(doc)) return;
        seen.add(doc);
        docs.push(doc);

        let frames = [];
        try {
            frames = Array.from(doc.querySelectorAll('iframe'));
        } catch (e) {
            return;
        }

        for (const frame of frames) {
            try {
                const childDoc = frame.contentDocument || frame.contentWindow?.document;
                if (childDoc) walk(childDoc);
            } catch (e) {
                // Cross-origin iframe; ignore it.
            }
        }
    }

    walk(document);
    return docs;
}

function findConfirmationControls() {
    const docs = getAllAccessibleDynamicsDocuments();

    for (const doc of docs) {
        try {
            const clientCheck = doc.getElementById('cbClientForm');
            const driverCheck = doc.getElementById('cbDriverForm');
            const interpreterCheck = doc.getElementById('cbInterpreterForm');

            if (clientCheck || driverCheck || interpreterCheck) {
                return { doc, clientCheck, driverCheck, interpreterCheck };
            }
        } catch (e) {}
    }

    return { doc: null, clientCheck: null, driverCheck: null, interpreterCheck: null };
}

function getSelectedReferralRecipients() {
    const recipients = [];
    const seenInputs = new Set();

    for (const doc of getAllAccessibleDynamicsDocuments()) {
        let checks = [];

        try {
            // Use the exact referral-recipient structure from the Dynamics web resource.
            // IDs vary, but AddOrRemoveRecipient is stable for these recipient rows.
            checks = Array.from(doc.querySelectorAll(
                'input[type="checkbox"][onclick*="AddOrRemoveRecipient"]'
            ));
        } catch (e) {
            continue;
        }

        for (const check of checks) {
            if (!check || seenInputs.has(check)) continue;
            seenInputs.add(check);

            // :checked should mirror .checked, but test both because Dynamics can
            // update the element during its inline onclick handler.
            let isChecked = !!check.checked;
            try { isChecked = isChecked || check.matches(':checked'); } catch (e) {}
            if (!isChecked) continue;

            let row = null;
            try { row = check.closest('tr'); } catch (e) {}
            if (!row) continue;

            let cells = [];
            try { cells = Array.from(row.querySelectorAll('td')); } catch (e) {}

            const visibleText = String(cells[1]?.textContent || cells[1]?.innerText || '')
                .replace(/\s+/g, ' ')
                .trim();

            const hiddenName = String(cells[2]?.textContent || cells[2]?.innerText || '')
                .replace(/\s+/g, ' ')
                .trim();

            // IMPORTANT: search the ENTIRE row.  Service-provider rows can be:
            // ACGE T&T LLC (MO) (Service Provider - Interpreter)
            // and the role text must not be lost by preferring only a name cell.
            const rowText = String(row.textContent || row.innerText || '')
                .replace(/\s+/g, ' ')
                .trim();

            const searchText = [visibleText, hiddenName, rowText]
                .filter(Boolean)
                .join(' ')
                .replace(/\s+/g, ' ')
                .trim();

            if (!searchText) continue;

            recipients.push({
                input: check,
                text: visibleText || rowText || hiddenName,
                visibleText,
                hiddenName,
                rowText,
                searchText
            });
        }
    }

    return recipients;
}

function setConfirmationUnchecked(checkBox) {
    if (!checkBox || !checkBox.checked) return;

    // Prefer a normal click so the web resource receives its native behavior.
    // Dynamics may show its own confirmation dialog after this click.  The
    // safeguard cleanup watcher below handles that dialog only when one of our
    // safety rules has fired.
    try { checkBox.click(); } catch (e) {}

    // Hard fallback in case Dynamics blocks or replaces the normal click behavior.
    if (checkBox.checked) {
        checkBox.checked = false;
        try { checkBox.dispatchEvent(new Event('input', { bubbles: true })); } catch (e) {}
        try { checkBox.dispatchEvent(new Event('change', { bubbles: true })); } catch (e) {}
    }
}

/* ======================= JACKIE EMAIL WORKFLOW ======================= */

const DD_JACKIE_ACCESS_API =
    'https://lowmargin.mtoysystems.com/api/get_email_list.php?list=DD_Buttons_Jackie';

let ddJackieWorkflowRunning = false;
let ddJackieAccessPromise = null;

function ddJackieDelay(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

function ddJackieWaitFor(getValue, timeout = 30000, interval = 250, errorMessage = 'Timed out waiting for Dynamics.') {
    return new Promise((resolve, reject) => {
        const startedAt = Date.now();

        function check() {
            let value = null;
            try { value = getValue(); } catch (e) {}

            if (value) {
                resolve(value);
                return;
            }

            if (Date.now() - startedAt >= timeout) {
                reject(new Error(errorMessage));
                return;
            }

            setTimeout(check, interval);
        }

        check();
    });
}

function ddJackieIdentityVariants(value) {
    const variants = new Set();

    function add(raw) {
        if (raw === null || raw === undefined) return;
        const text = String(raw).replace(/\s+/g, ' ').trim().toLowerCase();
        if (!text) return;
        variants.add(text);

        const angleEmail = text.match(/<\s*([^<>\s]+@[^<>\s]+)\s*>/)?.[1];
        if (angleEmail) add(angleEmail);
        if (text.includes('\\')) add(text.split('\\').pop());
        if (text.includes('@')) add(text.split('@')[0]);
    }

    add(value);
    return variants;
}

function ddJackieCollectAllowedIdentities(payload) {
    const values = [];
    const seen = new Set();

    function walk(value) {
        if (typeof value === 'string') {
            value.split(/[;,\r\n]+/).forEach(part => {
                if (part.trim()) values.push(part.trim());
            });
            return;
        }

        if (!value || typeof value !== 'object' || seen.has(value)) return;
        seen.add(value);

        if (Array.isArray(value)) value.forEach(walk);
        else Object.values(value).forEach(walk);
    }

    walk(payload);

    const identities = new Set();
    values.forEach(value => {
        ddJackieIdentityVariants(value).forEach(variant => identities.add(variant));
    });
    return identities;
}

function ddJackieRequestAccessList() {
    return new Promise((resolve, reject) => {
        GM_xmlhttpRequest({
            method: 'GET',
            url: `${DD_JACKIE_ACCESS_API}&_=${Date.now()}`,
            headers: { Accept: 'application/json' },
            timeout: 15000,
            onload: response => {
                try {
                    if (response.status < 200 || response.status >= 300) {
                        reject(new Error(`Blind Send access API returned HTTP ${response.status}`));
                        return;
                    }
                    resolve(JSON.parse(response.responseText));
                } catch (error) {
                    reject(error);
                }
            },
            onerror: () => reject(new Error('Blind Send access API request failed')),
            ontimeout: () => reject(new Error('Blind Send access API request timed out'))
        });
    });
}

function ddJackieGetXrm() {
    try {
        if (typeof unsafeWindow !== 'undefined' && unsafeWindow.Xrm) return unsafeWindow.Xrm;
    } catch (e) {}
    try {
        if (window.Xrm) return window.Xrm;
    } catch (e) {}
    return null;
}

function ddJackieGetSignedInEmailFromPage() {
    for (const doc of getAllAccessibleDynamicsDocuments()) {
        let accountEmail = null;
        try { accountEmail = doc.querySelector('#mectrl_currentAccount_secondary'); } catch (e) {}

        const text = String(accountEmail?.innerText || accountEmail?.textContent || '').trim();
        const email = text.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i)?.[0];
        if (email) return email;
    }
    return '';
}

async function ddJackieGetCurrentUserIdentities() {
    const source = await ddJackieWaitFor(
        () => {
            const accountEmail = ddJackieGetSignedInEmailFromPage();
            const xrm = ddJackieGetXrm();
            return accountEmail || xrm ? { accountEmail, xrm } : null;
        },
        15000,
        250,
        'Dynamics user information was not available.'
    );

    const values = [source.accountEmail, ddJackieGetSignedInEmailFromPage()].filter(Boolean);
    const context = source.xrm?.Utility?.getGlobalContext?.();
    const userSettings = context?.userSettings;
    if (userSettings?.userName) values.push(userSettings.userName);

    const userId = String(userSettings?.userId || '').replace(/[{}]/g, '');
    if (userId && source.xrm?.WebApi?.retrieveRecord) {
        try {
            const record = await source.xrm.WebApi.retrieveRecord(
                'systemuser',
                userId,
                '?$select=domainname,internalemailaddress,fullname'
            );
            ['domainname', 'internalemailaddress', 'fullname'].forEach(key => {
                if (record?.[key]) values.push(record[key]);
            });
        } catch (error) {
            console.warn('Blind Send could not read the full Dynamics user record.', error);
        }
    }

    const identities = new Set();
    values.forEach(value => {
        ddJackieIdentityVariants(value).forEach(variant => identities.add(variant));
    });
    return identities;
}

function ddJackieCurrentUserIsAllowed() {
    if (ddJackieAccessPromise) return ddJackieAccessPromise;

    ddJackieAccessPromise = Promise.all([
        ddJackieRequestAccessList(),
        ddJackieGetCurrentUserIdentities()
    ])
        .then(([payload, currentUserIdentities]) => {
            const allowedIdentities = ddJackieCollectAllowedIdentities(payload);
            return Array.from(currentUserIdentities).some(identity => allowedIdentities.has(identity));
        })
        .catch(error => {
            // Fail closed: if access cannot be verified, keep Blind Send hidden.
            console.warn('Blind Send button hidden because access could not be verified.', error);
            return false;
        });

    return ddJackieAccessPromise;
}

function ddJackieFindVisibleElement(selector) {
    const matches = [];
    for (const doc of getAllAccessibleDynamicsDocuments()) {
        let found = [];
        try { found = Array.from(doc.querySelectorAll(selector)); } catch (e) {}
        found.forEach(el => {
            if (elementIsActuallyVisible(el) && !el.disabled && el.getAttribute?.('aria-disabled') !== 'true') {
                matches.push(el);
            }
        });
    }
    return matches.length ? matches[matches.length - 1] : null;
}

function ddJackieFindSubjectInput(requireValue = false) {
    for (const doc of getAllAccessibleDynamicsDocuments()) {
        let candidates = [];
        try {
            candidates = Array.from(doc.querySelectorAll(
                'input[data-id="subject.fieldControl-text-box-text"], input[aria-label="Subject"]'
            ));
        } catch (e) {}

        for (let i = candidates.length - 1; i >= 0; i--) {
            const input = candidates[i];
            if (!elementIsActuallyVisible(input)) continue;
            if (requireValue && !String(input.value || '').trim()) continue;
            return input;
        }
    }
    return null;
}

function ddJackieFindEmailSaveButton() {
    const buttons = [];
    for (const doc of getAllAccessibleDynamicsDocuments()) {
        let found = [];
        try {
            found = Array.from(doc.querySelectorAll(
                '[id*="SavePrimary"], button[data-id$=".Save"], button[aria-label="Save"]'
            ));
        } catch (e) {}

        found.forEach(button => {
            if (!elementIsActuallyVisible(button) || button.disabled || button.getAttribute('aria-disabled') === 'true') return;
            const label = String(button.getAttribute('aria-label') || button.getAttribute('title') || button.textContent || '')
                .replace(/\s+/g, ' ').trim().toLowerCase();
            if (!label.includes('save & close') && !label.includes('save and close')) buttons.push(button);
        });
    }
    return buttons.length ? buttons[buttons.length - 1] : null;
}

function ddJackieFindTransportRecipient() {
    const matches = [];
    for (const doc of getAllAccessibleDynamicsDocuments()) {
        let checks = [];
        try {
            checks = Array.from(doc.querySelectorAll(
                'input[type="checkbox"][onclick*="AddOrRemoveRecipient"]'
            ));
        } catch (e) {}

        checks.forEach(input => {
            const row = input.closest?.('tr');
            const text = String(row?.innerText || row?.textContent || '').replace(/\s+/g, ' ').trim();
            const lower = text.toLowerCase();
            let score = 99;
            if (/service provider\s*-\s*transporter/i.test(text)) score = 0;
            else if (lower.includes('service provider') && lower.includes('transport')) score = 1;
            else if (/\btransport(?:er|ation)?\b/i.test(text)) score = 2;
            if (score < 99 && !lower.includes('interpreter')) matches.push({ input, text, score });
        });
    }
    matches.sort((a, b) => a.score - b.score);
    return matches[0] || null;
}

function ddJackieTransportIsSelected() {
    return getSelectedReferralRecipients().some(item =>
        /\btransport(?:er|ation)?\b/i.test(getRecipientSearchText(item))
    );
}

function ddJackieCheckboxIsClickable(checkBox) {
    if (!checkBox || !elementIsActuallyVisible(checkBox)) return false;
    if (checkBox.disabled || checkBox.getAttribute?.('aria-disabled') === 'true') return false;
    try {
        if (checkBox.ownerDocument?.defaultView?.getComputedStyle(checkBox)?.pointerEvents === 'none') return false;
    } catch (e) {}
    return true;
}

async function ddJackieEnsureTransportSelected() {
    const startedAt = Date.now();
    let attempts = 0;

    while (Date.now() - startedAt < 20000) {
        const recipient = ddJackieFindTransportRecipient();
        if (recipient?.input?.checked || ddJackieTransportIsSelected()) {
            await ddJackieDelay(700);
            if (ddJackieTransportIsSelected()) return true;
        }

        if (ddJackieCheckboxIsClickable(recipient?.input)) {
            attempts++;
            try { recipient.input.click(); } catch (e) {}
            await ddJackieDelay(450);

            if (!ddJackieTransportIsSelected() && attempts % 2 === 0) {
                const doc = recipient.input.ownerDocument;
                const label = Array.from(doc.querySelectorAll('label')).find(item =>
                    item.htmlFor === recipient.input.id ||
                    /transporter/i.test(String(item.textContent || ''))
                );
                try { label?.click(); } catch (e) {}
            }
        }
        await ddJackieDelay(650);
    }

    throw new Error('The Transport recipient could not be selected.');
}

function ddJackieSetSubjectValue(input, value) {
    const view = input.ownerDocument?.defaultView || window;
    const setter = Object.getOwnPropertyDescriptor(view.HTMLInputElement.prototype, 'value')?.set;
    if (setter) setter.call(input, value);
    else input.value = value;

    try {
        input.dispatchEvent(new view.InputEvent('input', { bubbles: true, inputType: 'insertText', data: '~*' }));
    } catch (e) {
        input.dispatchEvent(new view.Event('input', { bubbles: true }));
    }
    input.dispatchEvent(new view.Event('change', { bubbles: true }));
}

async function ddJackiePrefixSubject() {
    for (let attempt = 0; attempt < 3; attempt++) {
        const input = await ddJackieWaitFor(
            () => ddJackieFindSubjectInput(true), 15000, 250, 'The email subject field did not load.'
        );
        const current = String(input.value || '');
        if (current.startsWith('~*')) return;

        ddJackieSetSubjectValue(input, `~*${current}`);
        await ddJackieDelay(300);
        if (String(ddJackieFindSubjectInput(false)?.value || '').startsWith('~*')) return;
    }
    throw new Error('Dynamics did not keep the ~* subject prefix.');
}

async function runJackieWorkflow() {
    if (ddJackieWorkflowRunning) {
        showCenteredOverlayMessage('Blind Send is already working on this email.', false, 2500);
        return;
    }

    ddJackieWorkflowRunning = true;
    try {
        const emailButton = ddJackieFindVisibleElement(
            '[id^="gtt_referral\\|NoRelationship\\|Form\\|gtt\\.gtt_referral\\.EmailConfirmation\\.Command"][id*="-button"], ' +
            '[id*=".EmailConfirmation.Command"][id*="-button"], ' +
            'button[aria-label="Email Confirmation"], button[title="Email Confirmation"]'
        );
        if (!emailButton) throw new Error('Email Confirmation button not found. Open a referral and try again.');

        showCenteredOverlayMessage('Blind Send is opening Email Confirmation...', true, 2200);
        emailButton.click();

        await ddJackieWaitFor(
            () => ddJackieFindSubjectInput(true),
            60000,
            300,
            'Email Confirmation did not finish loading.'
        );

        const saveButton = await ddJackieWaitFor(
            ddJackieFindEmailSaveButton,
            15000,
            250,
            'The email loaded, but its Save button was not found.'
        );
        saveButton.click();
        await ddJackieDelay(1700);

        await ddJackieWaitFor(
            () => {
                const recipient = ddJackieFindTransportRecipient();
                return ddJackieCheckboxIsClickable(recipient?.input) ? recipient : null;
            },
            30000,
            250,
            'The Transport recipient did not become clickable after saving.'
        );

        await ddJackieEnsureTransportSelected();
        await ddJackiePrefixSubject();
        showCenteredOverlayMessage('Blind Send setup complete: Transport recipient and ~* subject are ready.', true, 4500);
    } catch (error) {
        console.error('Blind Send workflow failed:', error);
        showCenteredOverlayMessage(`Blind Send stopped: ${error?.message || error}`, false, 6500);
    } finally {
        ddJackieWorkflowRunning = false;
    }
}

function ddJackieAddAuthorizedButton(buttonContainer) {
    ddJackieCurrentUserIsAllowed().then(allowed => {
        if (!allowed) return;
        const currentContainer = buttonContainer?.isConnected
            ? buttonContainer
            : document.getElementById('custom-button-container');
        if (!currentContainer || document.getElementById('mtoy-jackie-button')) return;

        const button = createModernButton('Blind Send', '#ec4899', '#f472b6', runJackieWorkflow);
        button.id = 'mtoy-jackie-button';
        currentContainer.appendChild(button);
    });
}

/*
 * When a safeguard rule removes Client/Driver Confirmation, Dynamics opens a
 * second confirmation dialog asking whether to remove the Confirmation Form
 * content/attachment.  Once the user has acknowledged OUR safeguard alert, we
 * automatically accept that Dynamics dialog, remove every attachment from
 * the email, and SAVE that clean state before the user selects the correct form.
 * This cleanup runs only after a safeguard violation.
 */
let mtoySafeguardCleanupTimer = null;
let mtoySafeguardCleanupStartedAt = 0;
let mtoySafeguardCleanupLastActionAt = 0;
let mtoySafeguardCleanupNoAttachmentTicks = 0;
let mtoySafeguardCleanupSaveClicked = false;
let mtoySafeguardCleanupSaveClickedAt = 0;

function elementIsActuallyVisible(el) {
    if (!el || !el.isConnected) return false;
    try {
        const style = el.ownerDocument?.defaultView?.getComputedStyle(el);
        if (style && (style.display === 'none' || style.visibility === 'hidden' || style.opacity === '0')) {
            return false;
        }
        const rect = el.getBoundingClientRect?.();
        return !rect || rect.width > 0 || rect.height > 0;
    } catch (e) {
        return true;
    }
}

function getVisibleSafeguardCleanupDialogs() {
    const dialogs = [];

    for (const doc of getAllAccessibleDynamicsDocuments()) {
        let candidates = [];
        try {
            candidates = Array.from(doc.querySelectorAll(
                '[role="dialog"], [aria-modal="true"], .ms-Dialog-main, .fui-DialogSurface'
            ));
        } catch (e) {
            continue;
        }

        for (const dialog of candidates) {
            if (!elementIsActuallyVisible(dialog)) continue;

            const text = String(dialog.innerText || dialog.textContent || '')
                .replace(/\s+/g, ' ')
                .trim()
                .toLowerCase();

            // Exact native confirmation generated by unchecking Client/Driver Form.
            const isConfirmationFormRemoval =
                text.includes('you have selected to remove') &&
                text.includes('confirmation form') &&
                (text.includes('list of attachments') || text.includes('remove the content'));

            // Some Dynamics builds may show a separate attachment-delete confirm.
            const isAttachmentDeleteConfirm =
                text.includes('delete') &&
                text.includes('attachment');

            if (isConfirmationFormRemoval || isAttachmentDeleteConfirm) {
                dialogs.push(dialog);
            }
        }
    }

    return dialogs;
}

function clickSafeguardDialogAcceptButton() {
    const dialogs = getVisibleSafeguardCleanupDialogs();
    if (!dialogs.length) return false;

    for (const dialog of dialogs) {
        let buttons = [];
        try { buttons = Array.from(dialog.querySelectorAll('button')); } catch (e) {}

        const acceptButton = buttons.find(btn => {
            if (!elementIsActuallyVisible(btn)) return false;
            const txt = String(btn.innerText || btn.textContent || btn.getAttribute('aria-label') || '')
                .replace(/\s+/g, ' ')
                .trim()
                .toLowerCase();
            return txt === 'ok' || txt === 'yes' || txt === 'delete' || txt === 'remove';
        });

        if (acceptButton) {
            try {
                acceptButton.click();
                console.log('Email confirmation safeguard: accepted Dynamics removal dialog.');
                return true;
            } catch (e) {}
        }
    }

    return false;
}

function getEmailAttachmentContainers() {
    const containers = [];
    const seen = new Set();

    for (const doc of getAllAccessibleDynamicsDocuments()) {
        let found = [];
        try {
            found = Array.from(doc.querySelectorAll(
                '.attachmentsContainer .attachmentThumbnailIconContainer, .attachmentThumbnailIconContainer'
            ));
        } catch (e) {
            continue;
        }

        for (const el of found) {
            if (!seen.has(el)) {
                seen.add(el);
                containers.push(el);
            }
        }
    }

    return containers;
}

function getEmailAttachmentDeleteButtons() {
    const buttons = [];
    const seen = new Set();

    for (const container of getEmailAttachmentContainers()) {
        let btn = null;
        try { btn = container.querySelector('button[aria-label="Delete attachment"]'); } catch (e) {}
        if (btn && !seen.has(btn)) {
            seen.add(btn);
            buttons.push(btn);
        }
    }

    return buttons;
}

function clickNextEmailAttachmentDelete() {
    const buttons = getEmailAttachmentDeleteButtons();
    if (!buttons.length) return false;

    const btn = buttons[0];
    const container = btn.closest?.('.attachmentThumbnailIconContainer');

    try {
        // Fluent UI hides the thumbnail buttons until hover.  Programmatic click
        // normally works while hidden, but dispatch hover events too for builds
        // that require the button group to be activated first.
        if (container) {
            container.dispatchEvent(new MouseEvent('mouseover', { bubbles: true }));
            container.dispatchEvent(new MouseEvent('mouseenter', { bubbles: false }));
        }
    } catch (e) {}

    try {
        btn.click();
        console.log('Email confirmation safeguard: clicked Delete attachment.');
        return true;
    } catch (e) {
        console.warn('Email confirmation safeguard: could not click Delete attachment.', e);
        return false;
    }
}

function clickVisibleEmailSaveForSafeguard() {
    const candidates = [];

    for (const doc of getAllAccessibleDynamicsDocuments()) {
        let found = [];
        try {
            found = Array.from(doc.querySelectorAll(
                '[id*="SavePrimary"], ' +
                'button[data-id$=".Save"], ' +
                'button[aria-label="Save"]'
            ));
        } catch (e) {
            continue;
        }

        for (const btn of found) {
            if (!btn || !elementIsActuallyVisible(btn)) continue;
            if (btn.disabled || btn.getAttribute('aria-disabled') === 'true') continue;

            const label = String(
                btn.getAttribute('aria-label') ||
                btn.getAttribute('title') ||
                btn.innerText ||
                btn.textContent ||
                ''
            ).replace(/\s+/g, ' ').trim().toLowerCase();

            // Explicitly avoid Save & Close.  We only want to commit the cleanup
            // while leaving the email open so the user can choose the correct form.
            if (label.includes('save & close') || label.includes('save and close')) continue;
            candidates.push(btn);
        }
    }

    if (!candidates.length) return false;

    try {
        candidates[candidates.length - 1].click();
        console.log('Email confirmation safeguard: saved email to persist attachment removals.');
        return true;
    } catch (e) {
        console.warn('Email confirmation safeguard: could not click Save after cleanup.', e);
        return false;
    }
}

function stopSafeguardAttachmentCleanup() {
    if (mtoySafeguardCleanupTimer !== null) {
        clearInterval(mtoySafeguardCleanupTimer);
        mtoySafeguardCleanupTimer = null;
    }
}

function startSafeguardAttachmentCleanup() {
    // This cleanup is intentionally limited to the Confirmation Form workflow.
    // It NEVER blanket-deletes attachments. Manually added files must remain.
    // The continuous reconciler below removes only generated confirmation PDFs
    // whose type conflicts with the currently valid confirmation selection.
    mtoySafeguardCleanupStartedAt = Date.now();
    mtoySafeguardCleanupLastActionAt = Date.now();
    mtoySafeguardCleanupSaveClicked = false;
    mtoySafeguardCleanupSaveClickedAt = 0;

    if (mtoySafeguardCleanupTimer !== null) return;

    mtoySafeguardCleanupTimer = setInterval(() => {
        const now = Date.now();

        if (now - mtoySafeguardCleanupStartedAt > 15000) {
            stopSafeguardAttachmentCleanup();
            return;
        }

        // Accept only the native Dynamics dialog caused by removing one of the
        // generated Confirmation Forms (or a generated confirmation attachment).
        if (clickSafeguardDialogAcceptButton()) {
            mtoySafeguardCleanupLastActionAt = now;
            return;
        }

        // Continuously reconcile only Client/Driver/Interpreter confirmation PDFs.
        // Other/manual attachment names are ignored by the reconciler.
        try { reconcileConfirmationAttachments(); } catch (e) {}

        // Persist the checkbox/form removal after Dynamics has had time to settle.
        // If a stale confirmation PDF appears later, the reconciler will remove
        // that specific PDF and save again.
        if (
            !mtoySafeguardCleanupSaveClicked &&
            now - mtoySafeguardCleanupLastActionAt > 1400
        ) {
            if (clickVisibleEmailSaveForSafeguard()) {
                mtoySafeguardCleanupSaveClicked = true;
                mtoySafeguardCleanupSaveClickedAt = now;
            }
            return;
        }

        if (
            mtoySafeguardCleanupSaveClicked &&
            now - mtoySafeguardCleanupSaveClickedAt > 2200
        ) {
            stopSafeguardAttachmentCleanup();
        }
    }, 200);
}


/*
 * Keep the generated confirmation PDFs synchronized with the confirmation
 * checkboxes and the CURRENTLY selected referral recipients. Dynamics can create
 * a PDF several seconds after the checkbox/prompt flow finishes, so a one-time
 * cleanup is not enough. This reconciler runs continuously while the email is open:
 *
 *   Client checked       -> keep Client Confirmation (subject to existing rules)
 *   Driver checked       -> keep Driver Confirmation (subject to existing rules)
 *   Interpreter checked  -> keep Interpreter Confirmation; recipient eligibility
 *                           is validated event-by-event and again at Send
 *   unchecked/invalid    -> remove that generated confirmation PDF if it appears
 *
 * Other/non-confirmation attachments are never touched by this reconciler.
 */
let mtoyConfirmationAttachmentLastDeleteAt = 0;
let mtoyConfirmationAttachmentLastDeleteKey = '';
let mtoyConfirmationAttachmentNeedsSave = false;
let mtoyConfirmationAttachmentSaveAt = 0;

function getEmailAttachmentEntries() {
    const entries = [];

    for (const container of getEmailAttachmentContainers()) {
        let nameEl = null;
        let deleteButton = null;

        try {
            nameEl = container.querySelector('.attachmentName, [class*="attachmentName"]');
            deleteButton = container.querySelector('button[aria-label="Delete attachment"]');
        } catch (e) {}

        const name = String(
            nameEl?.innerText ||
            nameEl?.textContent ||
            container.getAttribute?.('aria-label') ||
            ''
        ).replace(/\s+/g, ' ').trim();

        if (!name) continue;

        entries.push({
            container,
            deleteButton,
            name,
            lowerName: name.toLowerCase()
        });
    }

    return entries;
}

function getRecipientSearchText(item) {
    return [
        item?.searchText,
        item?.text,
        item?.visibleText,
        item?.hiddenName,
        item?.rowText
    ]
        .filter(Boolean)
        .join(' ')
        .replace(/\s+/g, ' ')
        .trim();
}

// Interpreter matching is intentionally EVENT-DRIVEN instead of being invalidated
// continuously by the 250 ms poll. Dynamics can briefly rebuild the recipient web
// resource while a confirmation checkbox is changing; an "absence" check during that
// rebuild caused false Interpreter warnings. Client/Driver safeguards never rely on
// absence, so Interpreter now follows the same philosophy: validate when Interpreter
// is selected, when an Interpreter recipient changes, and again at Send.
function getCurrentInterpreterRecipient(selectedRecipients = null) {
    const recipients = selectedRecipients || getSelectedReferralRecipients();
    return recipients.find(item =>
        getRecipientSearchText(item).toLowerCase().includes('interpreter')
    ) || null;
}

function hasSelectedInterpreterRecipient() {
    return !!getCurrentInterpreterRecipient();
}

let mtoyInterpreterStableValidationToken = 0;

function validateInterpreterSelectionAfterRender(showAlert = true) {
    const token = ++mtoyInterpreterStableValidationToken;
    let attempts = 0;
    const maxAttempts = 12; // ~2.4 seconds, enough for the Dynamics web resource to rerender.

    function check() {
        if (token !== mtoyInterpreterStableValidationToken) return;

        const controls = findConfirmationControls();
        const interpreterCheck = controls.interpreterCheck;
        if (!interpreterCheck?.checked) return;

        const recipients = getSelectedReferralRecipients();
        const interpreterRecipient = getCurrentInterpreterRecipient(recipients);

        if (interpreterRecipient) {
            // Interpreter is valid. Now enforce one-of-three and the rest of the rules,
            // but do NOT run an absence-based Interpreter check again.
            validateEmailConfirmations(showAlert, false);
            return;
        }

        attempts++;
        if (attempts < maxAttempts) {
            setTimeout(check, 200);
            return;
        }

        // Recipient rows stayed stable long enough and no Interpreter recipient exists.
        validateEmailConfirmations(showAlert, true);
    }

    setTimeout(check, 175);
}

function isClientConfirmationAttachment(entry) {
    return entry.lowerName.includes('client confirmation');
}

function isDriverConfirmationAttachment(entry) {
    return entry.lowerName.includes('driver confirmation');
}

function isInterpreterConfirmationAttachment(entry) {
    return entry.lowerName.includes('interpreter confirmation');
}

function getMismatchedConfirmationAttachments() {
    const controls = findConfirmationControls();
    const clientChecked = !!controls.clientCheck?.checked;
    const driverChecked = !!controls.driverCheck?.checked;
    const interpreterChecked = !!controls.interpreterCheck?.checked;
    // There may be ZERO or ONE valid generated confirmation attachment.
    // Manual attachments with every other filename are always ignored.
    let allowedType = '';

    if (interpreterChecked) {
        allowedType = 'interpreter';
    } else if (clientChecked && !driverChecked && !interpreterChecked) {
        allowedType = 'client';
    } else if (driverChecked && !clientChecked && !interpreterChecked) {
        allowedType = 'driver';
    }

    return getEmailAttachmentEntries().filter(entry => {
        const isClient = isClientConfirmationAttachment(entry);
        const isDriver = isDriverConfirmationAttachment(entry);
        const isInterpreter = isInterpreterConfirmationAttachment(entry);

        if (!isClient && !isDriver && !isInterpreter) return false;

        if (!allowedType) return true;
        if (allowedType === 'client') return !isClient;
        if (allowedType === 'driver') return !isDriver;
        if (allowedType === 'interpreter') return !isInterpreter;

        return true;
    });
}

function clickSpecificAttachmentDelete(entry) {
    if (!entry?.deleteButton) return false;

    try {
        entry.container?.dispatchEvent(new MouseEvent('mouseover', { bubbles: true }));
        entry.container?.dispatchEvent(new MouseEvent('mouseenter', { bubbles: false }));
    } catch (e) {}

    try {
        entry.deleteButton.click();
        console.log('Email confirmation attachment reconciler: removing stale attachment:', entry.name);
        return true;
    } catch (e) {
        console.warn('Email confirmation attachment reconciler: failed to remove:', entry.name, e);
        return false;
    }
}

function reconcileConfirmationAttachments() {
    const now = Date.now();

    // If a delete click opened a native Dynamics confirmation, accept it. This
    // reuses the narrowly-scoped dialog detector already used by the safeguard.
    if (mtoyConfirmationAttachmentNeedsSave && clickSafeguardDialogAcceptButton()) {
        mtoyConfirmationAttachmentLastDeleteAt = now;
        return;
    }

    const mismatches = getMismatchedConfirmationAttachments();

    if (mismatches.length > 0) {
        const entry = mismatches[0];
        const key = `${entry.container?.id || ''}|${entry.name}`;

        // Give Dynamics time to rerender/remove the thumbnail after each click,
        // and avoid hammering the same hidden Delete button while a modal is open.
        if (
            key !== mtoyConfirmationAttachmentLastDeleteKey ||
            now - mtoyConfirmationAttachmentLastDeleteAt > 1500
        ) {
            if (clickSpecificAttachmentDelete(entry)) {
                mtoyConfirmationAttachmentLastDeleteKey = key;
                mtoyConfirmationAttachmentLastDeleteAt = now;
                mtoyConfirmationAttachmentNeedsSave = true;
            }
        }
        return;
    }

    // Once the wrong PDF is actually gone, persist that deletion. This is
    // especially important when Dynamics adds the stale PDF after the user's
    // earlier Save has already completed.
    if (
        mtoyConfirmationAttachmentNeedsSave &&
        now - mtoyConfirmationAttachmentLastDeleteAt > 1200 &&
        now - mtoyConfirmationAttachmentSaveAt > 2500
    ) {
        if (clickVisibleEmailSaveForSafeguard()) {
            mtoyConfirmationAttachmentSaveAt = now;
            mtoyConfirmationAttachmentNeedsSave = false;
            mtoyConfirmationAttachmentLastDeleteKey = '';
            console.log('Email confirmation attachment reconciler: saved removal of stale confirmation PDF.');
        }
    }
}

let mtoyEmailSafeguardHandling = false;

function validateEmailConfirmations(showAlert = true, enforceInterpreterRecipient = false) {
    if (mtoyEmailSafeguardHandling) return true;

    const controls = findConfirmationControls();
    const clientCheck = controls.clientCheck;
    const driverCheck = controls.driverCheck;
    const interpreterCheck = controls.interpreterCheck;

    if (!clientCheck && !driverCheck && !interpreterCheck) return true;

    const selectedRecipients = getSelectedReferralRecipients();
    const interpreterRecipient = getCurrentInterpreterRecipient(selectedRecipients);

    console.log(
        'Email confirmation safeguard - selected referral recipients:',
        selectedRecipients.map(item => ({
            id: item.input?.id || '',
            text: item.text,
            visibleText: item.visibleText,
            rowText: item.rowText,
            searchText: getRecipientSearchText(item),
            checked: !!item.input?.checked
        }))
    );

    // RULE 1: Only ONE generated confirmation may be selected.
    // When more than one is checked, prefer the confirmation that actually matches
    // the selected recipient type instead of blindly preferring Interpreter.
    //
    // Examples:
    //   Service Provider - Interpreter -> keep Interpreter
    //   Service Provider - Transporter -> keep Driver
    //   Adjuster / Customer Contact / Authorized -> keep Client
    // If more than one checked form is individually valid (for example a generic
    // recipient with both Client and Driver checked), clear the conflicting forms
    // rather than guessing.
    const checkedControls = [
        { key: 'client', label: 'Client Confirmation', cb: clientCheck },
        { key: 'driver', label: 'Driver Confirmation', cb: driverCheck },
        { key: 'interpreter', label: 'Interpreter Confirmation', cb: interpreterCheck }
    ].filter(item => item.cb?.checked);

    if (checkedControls.length > 1) {
        const recipientSearchTexts = selectedRecipients.map(item =>
            getRecipientSearchText(item).toLowerCase()
        );
        const hasTransportRecipient = recipientSearchTexts.some(text => text.includes('transport'));
        const hasClientTypeRecipient = recipientSearchTexts.some(text =>
            text.includes('adjuster') ||
            text.includes('customer contact') ||
            text.includes('authorized')
        );

        const isCompatible = (key) => {
            if (key === 'interpreter') return !!interpreterRecipient;
            if (key === 'driver') return !interpreterRecipient && !hasClientTypeRecipient;
            if (key === 'client') return !interpreterRecipient && !hasTransportRecipient;
            return false;
        };

        const compatibleChecked = checkedControls.filter(item => isCompatible(item.key));

        mtoyEmailSafeguardHandling = true;
        try {
            if (compatibleChecked.length === 1) {
                const keep = compatibleChecked[0];
                const removed = [];

                checkedControls.forEach(item => {
                    if (item !== keep) {
                        setConfirmationUnchecked(item.cb);
                        removed.push(item.label);
                    }
                });

                if (showAlert && removed.length) {
                    alert(
                        'EMAIL CONFIRMATION SAFEGUARD\n\n' +
                        'Only one confirmation form may be selected.\n\n' +
                        `${keep.label} has been kept because it matches the selected recipient; ` +
                        removed.join(' and ') +
                        (removed.length > 1 ? ' have' : ' has') +
                        ' been unchecked.'
                    );
                }
            } else {
                const removed = [];
                checkedControls.forEach(item => {
                    setConfirmationUnchecked(item.cb);
                    removed.push(item.label);
                });

                if (showAlert) {
                    alert(
                        'EMAIL CONFIRMATION SAFEGUARD\n\n' +
                        'Only one confirmation form may be selected.\n\n' +
                        (compatibleChecked.length === 0
                            ? 'None of the selected confirmation forms matches the selected recipient, so the conflicting confirmations have been unchecked.'
                            : 'More than one selected confirmation could apply, so the conflicting confirmations have been unchecked rather than guessing.')
                    );
                }
            }

            startSafeguardAttachmentCleanup();
        } finally {
            setTimeout(() => { mtoyEmailSafeguardHandling = false; }, 75);
        }
        return false;
    }

    // RULE 2: Only enforce the Interpreter-recipient requirement after a stable,
    // event-driven check (or at Send). The background poll deliberately skips this
    // absence test because Dynamics temporarily removes recipient rows while rerendering.
    if (enforceInterpreterRecipient && interpreterCheck?.checked && !interpreterRecipient) {
        mtoyEmailSafeguardHandling = true;
        try {
            setConfirmationUnchecked(interpreterCheck);

            if (showAlert) {
                alert(
                    'EMAIL CONFIRMATION SAFEGUARD\n\n' +
                    'Interpreter Confirmation can only be used when a selected recipient contains "Interpreter".\n\n' +
                    'Interpreter Confirmation has been unchecked.'
                );
            }

            startSafeguardAttachmentCleanup();
        } finally {
            setTimeout(() => { mtoyEmailSafeguardHandling = false; }, 75);
        }
        return false;
    }

    // RULE 3: If an Interpreter recipient is selected, Client and Driver
    // Confirmation are not allowed. This mirrors the existing positive-match
    // Transport/Adjuster safeguards: act only when the matching recipient is found.
    if (interpreterRecipient && (clientCheck?.checked || driverCheck?.checked)) {
        mtoyEmailSafeguardHandling = true;
        try {
            const removed = [];
            if (clientCheck?.checked) {
                setConfirmationUnchecked(clientCheck);
                removed.push('Client Confirmation');
            }
            if (driverCheck?.checked) {
                setConfirmationUnchecked(driverCheck);
                removed.push('Driver Confirmation');
            }

            if (showAlert && removed.length) {
                alert(
                    'EMAIL CONFIRMATION SAFEGUARD\n\n' +
                    `Selected recipient: ${interpreterRecipient.text}\n\n` +
                    'An Interpreter recipient can only use Interpreter Confirmation.\n\n' +
                    `${removed.join(' and ')} ${removed.length > 1 ? 'have' : 'has'} been unchecked.`
                );
            }

            startSafeguardAttachmentCleanup();
        } finally {
            setTimeout(() => { mtoyEmailSafeguardHandling = false; }, 75);
        }
        return false;
    }

    // RULE 4: Client Confirmation cannot be used when any selected referral
    // recipient contains "transport".
    if (clientCheck?.checked) {
        const offendingClientRecipient = selectedRecipients.find(item =>
            getRecipientSearchText(item).toLowerCase().includes('transport')
        );

        if (offendingClientRecipient) {
            mtoyEmailSafeguardHandling = true;
            try {
                setConfirmationUnchecked(clientCheck);

                if (showAlert) {
                    alert(
                        'EMAIL CONFIRMATION SAFEGUARD\n\n' +
                        `Selected recipient: ${offendingClientRecipient.text}\n\n` +
                        'Client Confirmation cannot be used for a recipient containing "Transport".\n\n' +
                        'Client Confirmation has been unchecked.'
                    );
                }

                startSafeguardAttachmentCleanup();
            } finally {
                setTimeout(() => { mtoyEmailSafeguardHandling = false; }, 75);
            }
            return false;
        }
    }

    // RULE 5: Driver Confirmation cannot be used for Adjuster,
    // Customer Contact, or Authorized recipient rows.
    if (driverCheck?.checked) {
        const blockedTerms = ['adjuster', 'customer contact', 'authorized'];
        let offendingDriverRecipient = null;
        let matchedTerm = '';

        for (const item of selectedRecipients) {
            const lower = getRecipientSearchText(item).toLowerCase();
            matchedTerm = blockedTerms.find(term => lower.includes(term)) || '';
            if (matchedTerm) {
                offendingDriverRecipient = item;
                break;
            }
        }

        if (offendingDriverRecipient) {
            mtoyEmailSafeguardHandling = true;
            try {
                setConfirmationUnchecked(driverCheck);

                if (showAlert) {
                    alert(
                        'EMAIL CONFIRMATION SAFEGUARD\n\n' +
                        `Selected recipient: ${offendingDriverRecipient.text}\n\n` +
                        `Driver Confirmation cannot be used when a selected recipient contains "${matchedTerm}".\n\n` +
                        'Driver Confirmation has been unchecked.'
                    );
                }

                startSafeguardAttachmentCleanup();
            } finally {
                setTimeout(() => { mtoyEmailSafeguardHandling = false; }, 75);
            }
            return false;
        }
    }

    return true;
}

function installEmailConfirmationListenersInDocument(doc) {
    if (!doc || doc.__mtoyConfirmationSafeguardInstalledV4224) return;
    doc.__mtoyConfirmationSafeguardInstalledV4224 = true;

    const validateAfterCheckboxChange = (event) => {
        const target = event.target;
        if (!target || target.type !== 'checkbox') return;

        setTimeout(() => {
            const id = target.id || '';
            let isReferralRecipient = false;
            try {
                isReferralRecipient = target.matches(
                    'input[type="checkbox"][onclick*="AddOrRemoveRecipient"]'
                );
            } catch (e) {}

            // Interpreter is validated only after Dynamics has had time to rebuild
            // the referral-recipient rows. This avoids the false "no Interpreter"
            // message that happened while the row was temporarily missing.
            if (id === 'cbInterpreterForm' && target.checked) {
                validateInterpreterSelectionAfterRender(true);
                return;
            }

            if (isReferralRecipient) {
                let rowHasInterpreter = false;
                try {
                    rowHasInterpreter = String(target.closest('tr')?.textContent || '')
                        .replace(/\s+/g, ' ')
                        .toLowerCase()
                        .includes('interpreter');
                } catch (e) {}

                if (rowHasInterpreter) {
                    // If Interpreter Confirmation is currently selected, re-check its
                    // eligibility only after the recipient list has stabilized.
                    const controls = findConfirmationControls();
                    if (controls.interpreterCheck?.checked) {
                        validateInterpreterSelectionAfterRender(true);
                        return;
                    }
                }
            }

            // Client/Driver and positive recipient matches can still be checked
            // immediately, exactly like the working Transport/Adjuster safeguards.
            validateEmailConfirmations(true, false);
        }, 150);
    };

    try {
        doc.addEventListener('click', validateAfterCheckboxChange, true);
        doc.addEventListener('change', validateAfterCheckboxChange, true);
    } catch (e) {}
}

function installEmailSendSafeguard() {
    if (window.mtoyEmailSendSafeguardInstalledV4224) return;
    window.mtoyEmailSendSafeguardInstalledV4224 = true;

    document.addEventListener('click', (event) => {
        const sendButton = event.target?.closest?.(
            'button[data-id="email|NoRelationship|Form|Mscrm.Form.email.Send"], ' +
            'button[id^="email|NoRelationship|Form|Mscrm.Form.email.Send"]'
        );

        if (!sendButton) return;

        if (!validateEmailConfirmations(true, true)) {
            event.preventDefault();
            event.stopPropagation();
            event.stopImmediatePropagation();
            console.warn('Email send stopped by confirmation safeguard.');
        }
    }, true);
}

function startEmailConfirmationSafeguardWatch() {
    if (window.mtoyEmailConfirmationSafeguardWatchStartedV4224) return;
    window.mtoyEmailConfirmationSafeguardWatchStartedV4224 = true;

    let lastFingerprint = '';

    function scan() {
        const docs = getAllAccessibleDynamicsDocuments();
        docs.forEach(installEmailConfirmationListenersInDocument);

        // Polling is intentional here. Dynamics can update checkbox state inside
        // a web resource without replacing nodes or generating a useful top-level
        // mutation.  The poll gives us a second independent safeguard path.
        const controls = findConfirmationControls();
        const recipients = getSelectedReferralRecipients();
        const fingerprint = [
            controls.clientCheck?.checked ? 'C1' : 'C0',
            controls.driverCheck?.checked ? 'D1' : 'D0',
            controls.interpreterCheck?.checked ? 'I1' : 'I0',
            recipients.map(r => `${r.input?.id || ''}:${r.input?.checked ? 1 : 0}:${getRecipientSearchText(r)}`).join('||')
        ].join('|');

        if (fingerprint !== lastFingerprint) {
            lastFingerprint = fingerprint;
            validateEmailConfirmations(true, false);
        }

        // Dynamics may generate confirmation PDFs several seconds after the
        // checkbox/prompt/save flow. Continuously remove any confirmation PDF
        // that conflicts with the currently selected form/recipient combination.
        reconcileConfirmationAttachments();
    }

    const observer = new MutationObserver(scan);
    observer.observe(document.documentElement, {
        childList: true,
        subtree: true
    });

    setInterval(scan, 250);
    scan();
}

installEmailSendSafeguard();
startEmailConfirmationSafeguardWatch();

function proceedWithRestOfFunction(claimant, claim, referralDate, selectedOption) {
    showProcessingMessage();

    waitForSavePrimary()
        .then((savePrimaryButton) => {
            setTimeout(() => {
            savePrimaryButton.click();
        }, 1500); // 👈 delay BEFORE clicking SavePrimary

            setTimeout(() => {
                var iframe = document.querySelector('#WebResource_AttachmentSelector');
                if (iframe) {
                    try {
                        var iframeDoc = iframe.contentDocument || iframe.contentWindow.document;
                        var checkBox = iframeDoc.querySelector('#cbClientForm');

                        // Install confirmation checkbox + recipient safeguards for this email.
                        installEmailConfirmationListenersInDocument(iframeDoc);

                        if (checkBox && !checkBox.checked) {
                            checkBox.click();
                        }

                        // Validate immediately after the script's automatic Client Confirmation selection.
                        setTimeout(() => validateEmailConfirmations(true), 25);

                        setTimeout(() => {
                            var templateButton = document.querySelector('[id*="Template"]');
                            if (templateButton) {
                                templateButton.click();

                                setTimeout(() => {
                                    selectCorrectRadioButton(selectedOption);
                                    //hideProcessingMessage();
                                }, 1800);
                            } else {
                                showMessage('Template button not found.', false);
                                hideProcessingMessage();
                            }
                        }, 1800);
                    } catch (e) {
                        console.error('Cannot access iframe content:', e);
                        hideProcessingMessage();
                    }
                } else {
                    console.error('No iframe found.');
                    hideProcessingMessage();
                }
            }, 3500);
        })
        .catch((error) => {
            showMessage(error, false);
            hideProcessingMessage();
        });
}

// Shows the animated processing message with a 5s auto-dismiss fail-safe
function showProcessingMessage() {
    // If an existing message is up, reset it (prevents stacking)
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

    // Create or reuse spinner CSS (avoid duplicate <style> nodes)
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

    // Optional: allow manual dismiss with ESC
    const escHandler = (e) => {
        if (e.key === 'Escape') hideProcessingMessage();
    };
    msg.dataset.escHandler = 'true';
    document.addEventListener('keydown', escHandler, { once: true });

    // Keep a reference so we can remove that listener even if auto-dismiss fires
    msg._escHandler = escHandler;

    document.body.appendChild(msg);

    // 🔒 Fail-safe: auto-hide after 14 seconds so users can finish manually
    processingTimeoutId = window.setTimeout(() => {
        hideProcessingMessage();
    }, 14000);
}

// Hides the processing message and clears the fail-safe timer
function hideProcessingMessage() {
    // Clear the one-shot timeout if it's pending
    if (processingTimeoutId !== null) {
        clearTimeout(processingTimeoutId);
        processingTimeoutId = null;
    }

    const msg = document.getElementById('processingMessage');
    if (msg) {
        // Remove the ESC handler if we attached one
        if (msg._escHandler) {
            document.removeEventListener('keydown', msg._escHandler);
        }
        msg.remove();
    }
}

// Function to detect when the page has fully loaded
function waitForPageToLoad() {
    return new Promise((resolve) => {
        const observer = new MutationObserver((mutations, observerInstance) => {
            if (document.readyState === 'complete') {
                observerInstance.disconnect();
                resolve();
            }
        });

        observer.observe(document.body, { childList: true, subtree: true });

        // Fallback timeout in case the observer misses the event
        setTimeout(resolve, 14000);
    });
}

function getVisibleElements(selector) {
    return Array.from(document.querySelectorAll(selector)).filter(el => {
        const style = window.getComputedStyle(el);
        return (
            el.offsetParent !== null &&
            style.display !== 'none' &&
            style.visibility !== 'hidden' &&
            style.opacity !== '0' &&
            !el.disabled &&
            el.getAttribute('aria-disabled') !== 'true'
        );
    });
}

function getLastVisibleElement(selector) {
    const els = getVisibleElements(selector);
    return els.length ? els[els.length - 1] : null;
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
                resolve(); // don't fail the whole flow if delete never appears
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

// Function to select the correct radio button
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
    } else if (selectedOption === "JBS Request for Higher Rates") {
        labelToFind = "JBS Request for Higher Rates (Default Rates)";
        showTemplateReminderPopup("Please send email to <b>AboveContractedRateRequest@sedgwick.com</b>");
    } else if (selectedOption === "Wait time request") {
        labelToFind = "Wait Time Request";
    } else if (selectedOption === "CareWorks Rate Request") {
        labelToFind = "CareWorks - Request for Higher Rates";
        showTemplateReminderPopup("Please send email to <b>AboveContractedRateRequest@sedgwick.com</b>");
    } else if (selectedOption === "Convergence Higher Rate Request") {
        labelToFind = "Convergence Higher Rate Request";
        showTemplateReminderPopup("Please always CC <b>kfimbres@convergencecare.com</b>");
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
                // keep the save -> signature order that actually worked
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

                            // non-staff: signature FIRST, then delete outbox, then save again
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

// Helper function to wait for an element to appear
function waitForElement(selector) {
    return new Promise(function(resolve, reject) {
        var interval = setInterval(function() {
            var element = document.querySelector(selector);
            if (element) {
                clearInterval(interval);
                resolve(element);
            }
        }, 500); // Check every 500ms
    });
}

// Function to extract the title part, prepend *, and copy it to the clipboard
function extractAndCopyTitle() {
    const titleElement = document.querySelector('[id^="formHeaderTitle"]');
    if (titleElement) {
        const title = titleElement.textContent.trim();
        console.log('Title Found:', title);

        // Check if title contains '4473' and show an alert
        if (title.startsWith("4473-")) {
            alert("Rates up to $3.75/mile are ok to apply and include in staffing email.\n\n" +
                  "Wait time ok if 25 miles or more each way and not surgery.\n\n" +
                  "If scheduled the day before appointment, ok to proceed with staffing above approved rates if needed and include in staffing email.");
        }


        // Check if title contains '5843-' and show a different alert
        if (title.startsWith("5843-")) {
            alert("Rate Increases do not need AUTH - provide rate increase in staffed email");
        }

        // Check if title contains '212-' and show a different alert
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

function createButtons() {
    const searchBox = document.querySelector('#searchBoxLiveRegion');

    if (searchBox) {
        // Prevent duplication by checking if container already exists
        if (document.getElementById('custom-button-container')) {
            console.log('Buttons already created.');
            return;
        }

        // Create container to hold buttons
        const buttonContainer = document.createElement('div');
        buttonContainer.id = 'custom-button-container';
        buttonContainer.style.display = 'inline-flex';
        buttonContainer.style.alignItems = 'center';
        buttonContainer.style.marginLeft = '10px';

        // Create and style buttons
        const createStyledButton = (text, color, clickHandler, tooltipText) => {
            const button = document.createElement('button');
            button.textContent = text;
            button.style.backgroundColor = color;
            button.style.color = '#fff';
            button.style.fontWeight = 'bold';
            button.style.letterSpacing = '1.2px';
            button.style.borderRadius = '15px';
            button.style.padding = '5px 10px';
            button.style.marginLeft = '10px';
            button.addEventListener('click', clickHandler);
            addTooltip(button, tooltipText);
            return button;
        };

        const button1 = createModernButton('Payer Emails', '#3b82f6', '#60a5fa', copyBoth);
        const button2 = createModernButton('Copy Name/SP Tab', '#3b82f6', '#60a5fa', copyClaimantName);
        const button3 = createModernButton('ClaimantID', '#3b82f6', '#60a5fa', extractAndCopyTitle);
        const button4 = createModernButton('Margin', '#22c55e', '#4ade80', showCalculatorBox);

        // Append buttons to container
        buttonContainer.appendChild(button3);
        buttonContainer.appendChild(button2);
        buttonContainer.appendChild(button1);
        buttonContainer.appendChild(button4);
        ddJackieAddAuthorizedButton(buttonContainer);

        // Insert button container into DOM
        searchBox.parentNode.insertBefore(buttonContainer, searchBox.nextSibling);

        console.log('Buttons created and inserted next to #searchBoxLiveRegion.');
    } else {
        console.log('Search box not found, retrying...');
    }
}

// Function to add a custom tooltip to the button
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

// Call the createButtons function to initialize the buttons
    createButtons();

    // ─── SPA navigation: re-inject buttons on every Dynamics navigation.
    // window 'load' only fires once on initial page load, never on SPA navigations.
    // observer.disconnect() after the first inject meant buttons disappeared after
    // navigating to a different record. This persistent observer re-injects whenever
    // #searchBoxLiveRegion reappears without a button container already present.
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
