// ==UserScript==
// @name         DD_Buttons_Admin
// @namespace    https://github.com/mtoy30/GoTandT
// @version      3.2.7
// @updateURL    https://raw.githubusercontent.com/mtoy30/GoTandT/main/DD_Buttons_Admin.user.js
// @downloadURL  https://raw.githubusercontent.com/mtoy30/GoTandT/main/DD_Buttons_Admin.user.js
// @description  Custom script for Dynamics 365 CRM page with multiple button functionalities
// @match        https://gotandt.crm.dynamics.com/*
// @match        https://gttqap2.crm.dynamics.com/
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
            }, 1500);
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
        "Let me process that real quick...",
        "Hold on, good things take time..."
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


// Function to create and style the buttons
function createButtons() {
    const searchBox = document.querySelector('#searchBoxLiveRegion');
    if (searchBox) {
        // Payer Emails Both button
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
        searchBox.parentNode.insertBefore(button1, searchBox.nextSibling);

        // Add tooltip to the Payer Emails button
        addTooltip(button1, "Clicking Payer emails then cancel on date input will still copy name and claim");

        // Copy Name button
        var button2 = document.createElement('button');
        button2.textContent = 'Copy Name/SP Tab';
        button2.style.backgroundColor = '#007BFF';
        button2.style.color = '#fff';
        button2.style.fontWeight = 'bold';
        button2.style.letterSpacing = "1.2px"
        button2.style.borderRadius = '15px';
        button2.style.padding = '5px 10px';
        button2.style.marginLeft = '10px';
        button2.addEventListener('click', copyClaimantName);
        searchBox.parentNode.insertBefore(button2, searchBox.nextSibling);

        // Add tooltip to the Payer Emails button
        addTooltip(button2, "This will copy name and take you to Service Provider Tab");

        // Claimant ID button
        var button3 = document.createElement('button');
        button3.textContent = 'ClaimantID';
        button3.style.backgroundColor = '#007BFF';
        button3.style.color = '#fff';
        button3.style.fontWeight = 'bold';
        button3.style.letterSpacing = "1.2px"
        button3.style.borderRadius = '15px';
        button3.style.padding = '5px 10px';
        button3.style.marginLeft = '10px';
        button3.addEventListener('click', extractAndCopyTitle);
        searchBox.parentNode.insertBefore(button3, searchBox.nextSibling);

        // Add tooltip to the Payer Emails button
        addTooltip(button3, "This will copy the claimant ID used for prev vendor search");

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
