// ==UserScript==
// @name         UIEnhancerforGOTANDTDynamics
// @namespace    https://github.com/mtoy30/GoTandT
// @version      1.1.9
// @updateURL   https://raw.githubusercontent.com/mtoy30/GoTandT/main/UIEnhancerforGOTANDTDynamics.user.js
// @downloadURL https://raw.githubusercontent.com/mtoy30/GoTandT/main/UIEnhancerforGOTANDTDynamics.user.js
// @description  Enhances UI with banner, row highlights, spacing, and styled notifications
// @author       Michael Toy
// @match        https://gotandt.crm.dynamics.com/*
// @match        https://gttqap2.crm.dynamics.com/*
// @grant        none
// ==/UserScript==

(function () {
    'use strict';

    const statusText = "Pending - RATE Authorization Requested";
    const headerSelector = '[id^="formHeaderTitle_"]';
    const buttonSelector = 'button[aria-label="Rate Approval Status"]';

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
        if (button.textContent.includes(statusText)) {
            insertBanner();
        } else {
            removeBanner();
        }
    }

    function observeRateStatusChanges() {
        const button = document.querySelector(buttonSelector);
        if (!button) return;

        const observer = new MutationObserver(() => {
            checkStatusAndInsertBanner();
        });

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

function addLegend() {
    document.querySelectorAll('[data-legend="true"]').forEach(el => el.remove());

    const targetSpans = document.querySelectorAll('span[id*="_text-value"]');
    targetSpans.forEach(span => {
        const label = span.textContent.toLowerCase();
        if (
            (label.includes("unassigned transportation") ||
             label.includes("same day confirmations") ||
             label.includes("same day (oncall)")) &&
            !label.includes("delete")
        ) {
            const type = (label.includes("-confirm") || label.includes("same day confirmations") || label.includes("same day (oncall)"))
                ? "-confirm"
                : "default";
            const legend = createLegendElement(type);
            span.parentElement?.appendChild(legend);
        }
    });

    const buttons = document.querySelectorAll('button[aria-label*="~Transport"], button[aria-label*="UBER"], button[aria-label*="Prev Vendor Search"], button[aria-label]');
    buttons.forEach(button => {
        const ariaLabel = button.getAttribute('aria-label')?.toLowerCase() || "";
        const labelText = button.querySelector('.ms-Button-label')?.textContent.toLowerCase() || "";

        if (ariaLabel.includes("delete") || labelText.includes("delete")) return;

        const isConfirm = labelText.includes("-confirm") ||
                          labelText.includes("same day confirmations") ||
                          labelText.includes("same day (oncall)");

        if (
            ariaLabel.includes("unassigned transport") ||
            ariaLabel.includes("~transport") ||
            ariaLabel.includes("uber") ||
            ariaLabel.includes("prev vendor search") ||
            isConfirm
        ) {
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

        if (header) {
            header.parentNode.insertBefore(banner, header.nextSibling);
        }
    }

function insertVipBannerIfNeeded() {
    const titleText = document.title;
    const header = document.querySelector(headerSelector);

    const vipIds = [
        "4474-65549",
        "4474-48338",
        "4474-48380",
        "202-46904",
        "202-50715",
        "4474-64737",
        "10837-61025",
        "4474-66551"
    ];

    if (!vipIds.some(id => titleText.includes(id)) || document.getElementById("vip-banner")) return;

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

    if (header) {
        header.parentNode.insertBefore(banner, header.nextSibling);
    }
}


    let timeout;
    const globalObserver = new MutationObserver(() => {
        clearTimeout(timeout);
        timeout = setTimeout(() => {
            checkStatusAndInsertBanner();
            highlightAllRowsGlobal();
            addLegend();
            adjustSpacing();
            styleNotificationWrapper();
            observeRateStatusChanges();
            insertJbaBannerIfNeeded();
            insertVipBannerIfNeeded();
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
            td.style.fontSize = "14px";
            td.style.fontWeight = "bold";

            const note = document.createElement("div");
            note.className = "monique-message";
            note.textContent = "Please combine staffing and/or auth requests into one email (include multiple dates into one email).";
            note.style.marginTop = "5px";
            note.style.color = "darkred";
            note.style.fontWeight = "bold";
            td.appendChild(note);
        } else if (retries > 0) {
            setTimeout(() => waitForMoniqueInIframe(retries - 3, delay), delay);
        }
    }

    const titleObserver = new MutationObserver(() => {
        if (document.title.includes("Email:")) {
            waitForMoniqueInIframe();
        }
    });

    titleObserver.observe(document.querySelector("title"), { childList: true });
})();
