// ==UserScript==
// @name         DD_Buttons_Admin
// @namespace    https://github.com/mtoy30/GoTandT
// @version      4.1.43
// @updateURL    https://raw.githubusercontent.com/mtoy30/GoTandT/main/DD_Buttons_Admin.user.js
// @downloadURL  https://raw.githubusercontent.com/mtoy30/GoTandT/main/DD_Buttons_Admin.user.js
// @description  Custom script for Dynamics 365 CRM page with multiple button functionalities
// @match        https://gotandt.crm.dynamics.com/*
// @match        https://gttqap2.crm.dynamics.com/*
// @author        Michael Toy
// @grant        GM_setClipboard
// @grant        GM_registerMenuCommand
// ==/UserScript==
//Moved to GitHub for 3.2.5+

(function() {
    'use strict';

    let processingTimeoutId = null;

    const style = document.createElement('style');
    style.innerHTML = `
       input[type="radio"] {
        width: 20px;
        height: 20px;
        accent-color: #3b82f6;
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
        box.style.minWidth = "500px";
        box.style.maxWidth = "500px";
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

        const lowMarginButton = createModernButton("Low Margin OK", "#3b82f6", "#60a5fa");
        lowMarginButton.onclick = () => {
            const textToCopy = "Management aware of Transportation low margin. Ok to staff";
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
        var startDateInput = document.querySelector('input[aria-label="Date of Start Date"]');
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
        "JBS Request for Higher Rates",
        "CareWorks Rate Request",
        "Staffed UBER Health",
        "Staffed Revised at Approved Rates"
    ];
} else if (use11525Rules) {
    exclusions = [
        "Standard Rate Request",
        "CareIQ Rate Request",
        "CareWorks Rate Request",
        "Homelink Rate Request",
        "Staffed UBER Health",
        "Staffed Revised at Approved Rates"
    ];
} else if (is4474) {
    exclusions = [
        "Standard Rate Request",
        "CareIQ Rate Request",
        "Homelink Rate Request",
        "JBS Request for Higher Rates"
    ];
} else if (headerTitle.startsWith("133-")) {
    exclusions = [
        "Standard Rate Request",
        "Homelink Rate Request",
        "JBS Request for Higher Rates",
        "CareWorks Rate Request",
        "Staffed UBER Health",
        "Staffed Revised at Approved Rates"
    ];
} else {
    exclusions = [
        "JBS Request for Higher Rates",
        "CareIQ Rate Request",
        "CareWorks Rate Request",
        "Homelink Rate Request",
        "Staffed UBER Health",
        "Staffed Revised at Approved Rates"
    ];
}

        const filteredOptions = fullOptions.filter(opt => !exclusions.includes(opt));

        filteredOptions.forEach(optionText => {
            const isStaff = /staff/i.test(optionText);
            const start = isStaff ? "#fde047" : "#3b82f6";
            const end   = isStaff ? "#facc15" : "#60a5fa";

            const button = createModernButton(optionText, start, end, () => {
                finalizeCopy(claimant, claim, referralDate, optionText);
                dropdownContainer.remove();
            });
            button.style.width = "100%";
            dropdownContainer.appendChild(button);
        });

        const closeButton = createModernButton("Close", "#7f1d1d", "#f87171", () => dropdownContainer.remove());
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
                const buttons = Array.from(document.querySelectorAll('button[aria-label="Insert Signature"]')).filter(btn => {
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
        } else if (selectedOption === "Homelink Rate Request") {
            labelToFind = "Homelink – Request for Higher Rates";
        } else if (selectedOption === "JBS Request for Higher Rates") {
            labelToFind = "JBS Request for Higher Rates (Default Rates)";
        } else if (selectedOption === "Wait time request") {
            labelToFind = "Wait Time Request";
        } else if (selectedOption === "CareWorks Rate Request") {
            labelToFind = "CareWorks - Request for Higher Rates";
        } else if (selectedOption === "Request Demographics") {
            labelToFind = "Request for Additional Information";
        } else if (selectedOption === "L-Orchid-CareWorks") {
            labelToFind = "L-Orchid-CareWorks";
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
            const button4 = createModernButton('Margin', '#22c55e', '#4ade80', showCalculatorBox);

            buttonContainer.appendChild(button3);
            buttonContainer.appendChild(button2);
            buttonContainer.appendChild(button1);
            buttonContainer.appendChild(button4);

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

    function waitForSearchBox() {
        const observer = new MutationObserver(() => {
            const searchBox = document.querySelector('#searchBoxLiveRegion');
            if (searchBox) {
                createButtons();
                observer.disconnect();
            }
        });

        observer.observe(document.body, {
            childList: true,
            subtree: true,
        });
    }

    window.addEventListener('load', () => {
        waitForSearchBox();
    });
})();
