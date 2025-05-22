// ==UserScript==
// @name         DD_Buttons_Admin_TEST
// @namespace    https://github.com/mtoy30/GoTandT
// @version      3.6.2
// @updateURL    https://raw.githubusercontent.com/mtoy30/GoTandT/main/DD_Buttons_Admin_TEST.user.js
// @downloadURL  https://raw.githubusercontent.com/mtoy30/GoTandT/main/DD_Buttons_Admin_TEST.user.js
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

// Function to show calculator UI
function showCalculatorBox() {
    // Check if the page title contains 'Referral: Information:'
    if (!document.title.includes("Referral: Information:")) {
        showMessage("Must be in a referral", false);
        return;
    }

    const tab8 = document.querySelector('[id^="tab8_"]');
    if (tab8) {
        tab8.click();
        console.log("Clicked tab8_ before showing calculator.");
        setTimeout(showCalculatorUI, 1000);
    } else {
        console.warn("tab8_ not found.");
        showCalculatorUI(); // fallback
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

    // Add fixed height and vertical scrollbar
    box.style.height = "800px"; // fixed height you want
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

    const inputLabel = document.createElement("label");
    inputLabel.innerText = "Enter Provider Rate:";
    inputLabel.style.display = "block";
    inputLabel.style.marginTop = "15px";
    inputLabel.style.fontWeight = "bold";

    const input = document.createElement("input");
    input.type = "number";
    input.style.width = "100%";
    input.style.marginTop = "10px";
    input.style.marginBottom = "15px";

    const result = document.createElement("div");
    result.style.marginTop = "10px";
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
    "Wheelchair Rental"
];

    const resetButton = document.createElement("button");
    resetButton.innerText = "Reset";
    resetButton.style.marginTop = "20px";
    resetButton.style.marginLeft = "10px";
    resetButton.style.padding = "8px 16px";
    resetButton.style.background = "#e74c3c";
    resetButton.style.color = "#fff";
    resetButton.style.border = "none";
    resetButton.style.borderRadius = "5px";
    resetButton.style.cursor = "pointer";
    resetButton.style.fontWeight = "bold";

    resetButton.onclick = () => {
        input.value = "";
        flatRadio.checked = true;
        mileRadio.checked = false;
        result.innerHTML = "";
        targetLabel.innerHTML = "";
        higherResult.innerHTML = "";

        Object.values(productInputs).forEach(field => {
            field.value = "";
        });
    };

    const lowMarginButton = document.createElement("button");
    lowMarginButton.innerText = "Low Margin OK";
    lowMarginButton.style.marginTop = "20px";
    lowMarginButton.style.marginLeft = "10px";
    lowMarginButton.style.padding = "8px 16px";
    lowMarginButton.style.background = "#3498db";
    lowMarginButton.style.color = "#fff";
    lowMarginButton.style.border = "none";
    lowMarginButton.style.borderRadius = "5px";
    lowMarginButton.style.cursor = "pointer";
    lowMarginButton.style.fontWeight = "bold";

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

        setTimeout(() => {
            copiedMsg.remove();
        }, 1000);
    });
};

    const waittimeButton = document.createElement("button");
    waittimeButton.innerText = "Wait Time Request";
    waittimeButton.style.marginTop = "20px";
    waittimeButton.style.marginLeft = "10px";
    waittimeButton.style.padding = "8px 16px";
    waittimeButton.style.background = "#3498db";
    waittimeButton.style.color = "#fff";
    waittimeButton.style.border = "none";
    waittimeButton.style.borderRadius = "5px";
    waittimeButton.style.cursor = "pointer";
    waittimeButton.style.fontWeight = "bold";

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

        setTimeout(() => {
            copiedMsg.remove();
        }, 1000);
    });
};

const boomerangButton = document.createElement("button");
    boomerangButton.innerText = "Boomerang";
    boomerangButton.style.marginTop = "20px";
    boomerangButton.style.marginLeft = "10px";
    boomerangButton.style.padding = "8px 16px";
    boomerangButton.style.background = "#3498db";
    boomerangButton.style.color = "#fff";
    boomerangButton.style.border = "none";
    boomerangButton.style.borderRadius = "5px";
    boomerangButton.style.cursor = "pointer";
    boomerangButton.style.fontWeight = "bold";

    boomerangButton.onclick = () => {
    const textToCopy = "Secure with Boomerang and leave in provider stage until rates approved";
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

        setTimeout(() => {
            copiedMsg.remove();
        }, 1000);
    });
};

const WaitStaffButton = document.createElement("button");
    WaitStaffButton.innerText = "Wait Time Staff";
    WaitStaffButton.style.marginTop = "20px";
    WaitStaffButton.style.marginLeft = "10px";
    WaitStaffButton.style.padding = "8px 16px";
    WaitStaffButton.style.background = "#3498db";
    WaitStaffButton.style.color = "#fff";
    WaitStaffButton.style.border = "none";
    WaitStaffButton.style.borderRadius = "5px";
    WaitStaffButton.style.cursor = "pointer";
    WaitStaffButton.style.fontWeight = "bold";

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

        setTimeout(() => {
            copiedMsg.remove();
        }, 1000);
    });
};

// Create the "Request Rates" button
const requestRatesButton = document.createElement("button");
requestRatesButton.innerText = "Request Rates";
requestRatesButton.style.marginTop = "20px";
requestRatesButton.style.marginLeft = "10px";
requestRatesButton.style.padding = "8px 16px";
requestRatesButton.style.background = "#2ecc71";
requestRatesButton.style.color = "#fff";
requestRatesButton.style.border = "none";
requestRatesButton.style.borderRadius = "5px";
requestRatesButton.style.cursor = "pointer";
requestRatesButton.style.fontWeight = "bold";

// Add the button to the page
document.body.appendChild(requestRatesButton);

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
        console.log("Matched Transport with quantity:", miles);
      }

      if (accountProductText.includes("Load Fee") && !isNaN(quantity)) {
        loadFeeQuantity = quantity;
        console.log("Matched Load Fee with quantity:", loadFeeQuantity);
      }
    }
  });

  // Now you can use miles and loadFeeQuantity accurately:
  let parts = [];

  Object.entries(productInputs).forEach(([label, input]) => {
  const value = input.value.trim();
  if (value !== "") {
    // Normalize labels
    let normalizedLabel = label === "Miscellaneous Dead Miles" ? "Dead Miles" :
                          label === "No Show" ? "No Show/Late Cancel" :
                          label;

    if (["Transport Ambulatory", "Transport Wheelchair", "Transport Stretcher, ALS & BLS"].includes(label)) {
      if (!isNaN(parseFloat(value))) {
        parts.push(`$${parseFloat(value).toFixed(2)}/mile x ${miles} miles`);
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
    } else if (normalizedLabel === "Dead Miles") {
      if (!isNaN(parseFloat(value))) {
        parts.push(`${parseFloat(value)} ${normalizedLabel}`);
      } else {
        parts.push(`${value} ${normalizedLabel}`);
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


  const finalText = "Request rates " + parts.join(", ");
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


    const calculateMargin = () => {
        const rateType = document.querySelector('input[name="rateType"]:checked').value;
        const inputValue = parseFloat(input.value);
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
        if (productsToTrack.includes(product) && product !== "Rush Fee") {
            if (!isNaN(totalValue)) {
                totalBilled += totalValue;
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

        ["Miscellaneous Dead Miles", "Tolls", "Other", "No Show", "Wait Time"].forEach(p => foundProducts.add(p));

const preferredOrder = [
  "Transport Ambulatory",
  "Transport Wheelchair",
  "Transport Stretcher, ALS & BLS",
  "Miscellaneous Dead Miles",
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

const label = document.createElement("label");
label.innerText = product ;
label.style.fontWeight = "bold";

labelRow.appendChild(label);

// Add Contract Rates button only for Wait Time and No Show
if (["Wait Time", "No Show"].includes(product)) {
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
    labelRow.appendChild(contractBtn);
}

const inputField = document.createElement("input");
inputField.type = "text";
inputField.style.width = "100%";

inputField.addEventListener("input", calculateMargin);

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

        let paidAmount = inputValue;
        if (rateType === "mile") {
            if (quantity === 0) {
                result.innerText = "Could not find transport quantity.";
                result.style.color = "black";
                higherResult.innerText = "";
                return;
            }
            paidAmount = inputValue * quantity;
        }

        const margin = 100 - ((paidAmount / totalBilled) * 100);
        let marginColor = "black";
        if (margin <= 24.99) marginColor = "red";
        else if (margin < 35) marginColor = "goldenrod";
        else marginColor = "green";

        let approvalNote = margin <= 24.99 ? `<br><span style="color: red; font-weight: bold;">Seek Management Approval</span>` : "";

        result.innerHTML = `
            <span>Total Billed: $${totalBilled.toFixed(2)}</span><br>
            ${rateType === "mile" ? `<span>Miles: ${quantity}, Total Paid: $${paidAmount.toFixed(2)}</span><br>` : ""}
            <span style="color: ${marginColor}; font-weight: bold;">Margin: ${margin.toFixed(2)}%</span>${approvalNote}
        `.trim();

        const target = totalBilled * (1 - 0.35);
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
    if (product === "Wait Time") return;

    // Prevent duplicates for Rush Fee, Tolls, etc.
    if (["Rush Fee", "Tolls", "Other", "Assistance Fee", "Passenger Fee", "Miscellaneous Dead Miles"].includes(product)) {
        if (alreadyCounted.has(product)) return; // Skip if already counted
        alreadyCounted.add(product);
    }

    const enteredValueRaw = productInputs[product]?.value;
    const enteredValue = parseFloat(enteredValueRaw);
    const qty = quantities[product] || 0;

    console.log(`Product: ${product}, Entered: ${enteredValueRaw}, Parsed: ${enteredValue}, Qty: ${qty}`);

    if (!isNaN(enteredValue)) {
        if (product === "Miscellaneous Dead Miles") {
            higherTotal += enteredValue * (activeTransportRate / 2);
        } else if (["Tolls", "Other", "Assistance Fee", "Passenger Fee", "Rush Fee"].includes(product)) {
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

        let higherApprovalNote = higherMargin <= 24.99 ? `<br><span style="color: red; font-weight: bold;">Seek Management Approval</span>` : "";

        higherResult.innerHTML = `
            <span>Total Using Higher Rates: $${higherTotal.toFixed(2)}</span><br>
            <span style="color: ${higherMarginColor}; font-weight: bold;">Margin: ${higherMargin.toFixed(2)}%</span>${higherApprovalNote}
        `.trim();
    };

    input.addEventListener("input", calculateMargin);
    flatRadio.addEventListener("change", () => {
        input.value = "";
        calculateMargin();
    });
    mileRadio.addEventListener("change", () => {
        input.value = "";
        calculateMargin();
    });

    box.appendChild(closeButton);
    box.appendChild(modeLabel);
    box.appendChild(flatRadio);
    box.appendChild(flatLabel);
    box.appendChild(mileRadio);
    box.appendChild(mileLabel);
    box.appendChild(inputLabel);
    box.appendChild(input);
    box.appendChild(result);
    box.appendChild(targetLabel);
    box.appendChild(higherHeader);
    box.appendChild(higherInputsWrapper);
    box.appendChild(higherResult);
    box.appendChild(lowMarginButton);
    box.appendChild(waittimeButton);
    box.appendChild(WaitStaffButton);
    box.appendChild(boomerangButton);
    box.appendChild(requestRatesButton);
    box.appendChild(resetButton);


    document.body.appendChild(box);
}

    // Add event listener for the calculator button
    const calculatorButton = document.querySelector('#yourCalculatorButtonSelector'); // Replace with actual button selector
    if (calculatorButton) {
        calculatorButton.addEventListener('click', showCalculatorBox);
    }

    // Function to copy claimant name
    function copyClaimantName() {
        var elementToCopy = document.querySelector('[id^="headerControlsList_"] > div:nth-child(3) > div[class^="pa-a pa-"].flexbox > a');
        if (elementToCopy) {
            var textToCopy = elementToCopy.textContent.trim();
            GM_setClipboard(textToCopy);
            showMessage(`Copied: "${textToCopy}" successfully.`);
            console.log('Copied to clipboard:', textToCopy);
            var tabs = document.querySelectorAll('[id^="tab6_"]');
            if (tabs.length) {
                tabs[0].click();
                console.log('Clicked on the first tab.');
                waitForButtonAndClick();
            } else {
                console.error('No tab6 elements found.');
            }
        } else {
            showMessage('Claimant Name not found. Please make sure you are in a referral.', false);
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
    var element1 = document.querySelector('[id^="headerControlsList_"] > div:nth-child(3) > div[class^="pa-a pa-"].flexbox > a');
    var element2 = document.querySelector('[id^="headerControlsList_"] > div:nth-child(5) > div[class^="pa-a pa-"].flexbox > a');
    var titleElement = document.querySelector('[id^="formHeaderTitle"]');

    if (element1 && element2) {
        var text1 = element1.textContent.trim();
        var text2 = element2.textContent.trim();

        // Check if title contains '4474' and show an alert
        if (titleElement && titleElement.textContent.includes("4474-")) {
            alert("Rate increase needs to go to the adjuster and authorized by aswell as AboveContractedRateRequest@sedgwick.com");
        }

        // Check if the claim number starts with "H" followed by numbers
        if (/^H\d+$/.test(text2)) {
            alert(`The claim number "${text2}" appears to be an HES claim. Please verify the payer is CareWorks and send related emails to HES@careworks.com`);
        }

        // Prompt user for referral date(s)
        var referralDate = prompt("Please enter the referral date(s):", "");

        if (referralDate === null) {
            // User clicked "Cancel", perform the default copy action
            var textToCopy = `Claimant: ${text1} - Claim: ${text2} - on DOS:`;
            GM_setClipboard(textToCopy);
            showMessage(`Copied: "${textToCopy}" successfully.`);
            return; // Stop further execution
        }

        if (!referralDate) {
            referralDate = "[No Date Provided]"; // Default text if nothing is entered
        }

        // Create the dropdown UI
        createDropdownMenu(text1, text2, referralDate);
    } else {
        showMessage('Claimant Name & Claim# not found. Please make sure you are in a referral.', false);
    }
}

// Function to create and show dropdown menu
function createDropdownMenu(claimant, claim, referralDate) {
    var existingDropdown = document.getElementById("customDropdownContainer");
    if (existingDropdown) existingDropdown.remove();

    var dropdownContainer = document.createElement("div");
    dropdownContainer.id = "customDropdownContainer";
    dropdownContainer.style.position = "fixed";
    dropdownContainer.style.top = "30%";
    dropdownContainer.style.left = "50%";
    dropdownContainer.style.transform = "translate(-50%, -50%)";
    dropdownContainer.style.background = "#FFF";
    dropdownContainer.style.padding = "15px";
    dropdownContainer.style.border = "1px solid #ccc";
    dropdownContainer.style.borderRadius = "8px";
    dropdownContainer.style.boxShadow = "0px 4px 6px rgba(0, 0, 0, 0.1)";
    dropdownContainer.style.zIndex = "10000";

    var label = document.createElement("label");
    label.innerText = "Select which template you would like to apply:";
    label.style.display = "block";
    label.style.marginBottom = "16px";
    label.style.color = "#000";
    label.style.letterSpacing = "1.5px";
    label.style.fontWeight = "bold";
    dropdownContainer.appendChild(label);

    var options = [
        { text: "Staffed Email", color: "black" },
        { text: "Standard Rate Request", color: "#FFFF99" }, // Light Yellow
        { text: "CareIQ Rate Request", color: "white" },
        { text: "Homelink Rate Request", color: "white" },
        { text: "JBS Staffed at Rates", color: "white" },
        { text: "Wait time request", color: "white" },
        { text: "Request Demographics", color: "white" },
        { text: "Other", color: "black" }
    ];

    options.forEach((option) => {
        var button = document.createElement("button");
        button.innerText = option.text;
        button.style.display = "block";
        button.style.width = "100%";
        button.style.fontWeight = "bold";
        button.style.padding = "8px";
        button.style.marginBottom = "15px";
        button.style.background = "#007bff";
        button.style.color = option.color;
        button.style.letterSpacing = "1.5px";
        button.style.border = "none";
        button.style.borderRadius = "15px";
        button.style.cursor = "pointer";

        button.onclick = function () {
            finalizeCopy(claimant, claim, referralDate, option.text);
            dropdownContainer.remove();
        };

        dropdownContainer.appendChild(button);
    });

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
                                }, 2000);
                            });
                        }
                    });

                    observer.observe(document.body, { childList: true, subtree: true });
                } else {
                    proceedWithRestOfFunction(claimant, claim, referralDate, selectedOption);
                }
            }, 1500);
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

function proceedWithRestOfFunction(claimant, claim, referralDate, selectedOption) {
    showProcessingMessage();

    waitForSavePrimary()
        .then((savePrimaryButton) => {
            savePrimaryButton.click();

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
                                    hideProcessingMessage();
                                }, 1500);
                            } else {
                                showMessage('Template button not found.', false);
                                hideProcessingMessage();
                            }
                        }, 1500);
                    } catch (e) {
                        console.error('Cannot access iframe content:', e);
                        hideProcessingMessage();
                    }
                } else {
                    console.error('No iframe found.');
                    hideProcessingMessage();
                }
            }, 2000);
        })
        .catch((error) => {
            showMessage(error, false);
            hideProcessingMessage();
        });
}

// Shows the animated processing message
function showProcessingMessage() {
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

    let msg = document.createElement('div');
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
    });

    const style = document.createElement('style');
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
    document.body.appendChild(msg);
}

// Hides the processing message
function hideProcessingMessage() {
    let msg = document.getElementById('processingMessage');
    if (msg) msg.remove();
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
        setTimeout(resolve, 5000);
    });
}

// Function to select the correct radio button
function selectCorrectRadioButton(selectedOption) {
    showProcessingMessage(); // Start the spinner

    var labelToFind = "";
    if (selectedOption === "Staffed Email") {
        labelToFind = "Staffed";
    } else if (selectedOption === "Standard Rate Request") {
        labelToFind = "Request for Higher Rates";
    } else if (selectedOption === "CareIQ Rate Request") {
        labelToFind = "CIQ Higher Rate Request";
    } else if (selectedOption === "Homelink Rate Request") {
        labelToFind = "Homelink – Request for Higher Rates";
    } else if (selectedOption === "JBS Staffed at Rates") {
        labelToFind = "JBS Staffed at Higher Rates (Default Rates)";
    } else if (selectedOption === "Wait time request") {
        labelToFind = "Wait Time Request";
    } else if (selectedOption === "Request Demographics") {
        labelToFind = "Request for Additional Information";
    } else if (selectedOption === "Other") {
        labelToFind = "";
    }

    if (labelToFind === "") {
        hideProcessingMessage(); // stop spinner if no action taken
        return;
    }

    setTimeout(function () {
        var labels = document.querySelectorAll('label');
        var correctLabel = Array.from(labels).find(label => {
            var titleSpan = label.querySelector('.titleText');
            return titleSpan && titleSpan.textContent.trim() === labelToFind;
        });

        if (correctLabel) {
            var associatedRadio = document.getElementById(correctLabel.getAttribute('for'));
            if (associatedRadio) {
                associatedRadio.click();

                setTimeout(function () {
                    var applyTemplateButton = document.querySelector('button[title="Apply template"]');
                    if (applyTemplateButton) {
                        applyTemplateButton.click();

                        setTimeout(function () {
                            var okButton = document.querySelector('button[title="OK"]');
                            if (okButton) {
                                okButton.click();

                                if (selectedOption !== "Staffed Email") {
                                    setTimeout(function () {
                                        var deleteButton = document.querySelector('button[aria-label="Delete Referral Outbox"]');
                                        if (deleteButton) {
                                            deleteButton.click();
                                        }
                                    }, 1500);
                                }

                                setTimeout(function () {
                                    var savePrimaryButton = document.querySelector('[id*="SavePrimary"]');
                                    if (savePrimaryButton) {
                                        savePrimaryButton.click();
                                    }

                                    hideProcessingMessage(); // ✅ Stop the spinner here

                                    // ✅ Now show final message
                                    setTimeout(function () {
                                        var messageDiv = document.createElement("div");
                                        messageDiv.innerText = "Actions completed. Please proceed with filling out the remainder of the information needed.";
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
                                            document.body.removeChild(messageDiv);
                                        }, 2000);
                                    }, 1500);
                                }, 1500);
                            }
                        }, 1500);
                    }
                }, 1500);
            }
        }
    }, 1000);
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
        if (title.includes("4473-")) {
            alert("Rates up to $3.75/mile are ok to apply and include in staffing email.\n\n" +
                  "Wait time ok if 25 miles or more each way and not surgery.\n\n" +
                  "If scheduled the day before appointment, ok to proceed with staffing above approved rates if needed and include in staffing email.");
        }


        // Check if title contains '5843-' and show a different alert
        if (title.includes("5843-")) {
            alert("Rate Increases do not need AUTH - provide rate increase in staffed email");
        }

        // Check if title contains '212-' and show a different alert
        if (title.includes("212-")) {
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
        // Create and style buttons
        var button1 = document.createElement('button');
        button1.textContent = 'Payer Emails';
        button1.style.backgroundColor = '#007BFF';
        button1.style.color = '#fff';
        button1.style.fontWeight = 'bold';
        button1.style.letterSpacing = "1.2px";
        button1.style.borderRadius = '15px';
        button1.style.padding = '5px 10px';
        button1.style.marginLeft = '10px';
        button1.addEventListener('click', copyBoth);
        addTooltip(button1, "Clicking Payer emails then cancel on date input will still copy name and claim");

        var button2 = document.createElement('button');
        button2.textContent = 'Copy Name/SP Tab';
        button2.style.backgroundColor = '#007BFF';
        button2.style.color = '#fff';
        button2.style.fontWeight = 'bold';
        button2.style.letterSpacing = "1.2px";
        button2.style.borderRadius = '15px';
        button2.style.padding = '5px 10px';
        button2.style.marginLeft = '10px';
        button2.addEventListener('click', copyClaimantName);
        addTooltip(button2, "This will copy name and take you to Service Provider Tab");

        var button3 = document.createElement('button');
        button3.textContent = 'ClaimantID';
        button3.style.backgroundColor = '#007BFF';
        button3.style.color = '#fff';
        button3.style.fontWeight = 'bold';
        button3.style.letterSpacing = "1.2px";
        button3.style.borderRadius = '15px';
        button3.style.padding = '5px 10px';
        button3.style.marginLeft = '10px';
        button3.addEventListener('click', extractAndCopyTitle);
        addTooltip(button3, "This will copy the claimant ID used for prev vendor search");

        var button4 = document.createElement('button');
        button4.textContent = 'Margin';
        button4.style.backgroundColor = 'green';
        button4.style.color = '#fff';
        button4.style.fontWeight = 'bold';
        button4.style.letterSpacing = "1.2px";
        button4.style.borderRadius = '15px';
        button4.style.padding = '5px 10px';
        button4.style.marginLeft = '10px';
        button4.addEventListener('click', showCalculatorBox);
        addTooltip(button4, "Calculate Margins");

        // Append buttons in order
        searchBox.parentNode.insertBefore(button3, searchBox.nextSibling);
        searchBox.parentNode.insertBefore(button2, button3.nextSibling);
        searchBox.parentNode.insertBefore(button1, button2.nextSibling);
        searchBox.parentNode.insertBefore(button4, button1.nextSibling);

        console.log('Buttons created and inserted in order next to #searchBoxLiveRegion.');
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


    // Function to continuously observe the DOM for the search box to appear
    function waitForSearchBox() {
        const observer = new MutationObserver(() => {
            const searchBox = document.querySelector('#searchBoxLiveRegion');
            if (searchBox) {
                createButtons();
                observer.disconnect(); // Stop observing once the buttons are created
            }
        });

        observer.observe(document.body, {
            childList: true,
            subtree: true,
        });
    }

    // Start observing the DOM for the search box and notification elements
    window.addEventListener('load', () => {
        waitForSearchBox();
    });
})();
