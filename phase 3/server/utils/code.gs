function doGet(e) {
    try {
        let page = e.parameter.view || "_employee-details";
        let html = HtmlService.createTemplateFromFile(page).evaluate();
        let htmlOutput = HtmlService.createHtmlOutput(html);
        htmlOutput.addMetaTag(
            "viewport",
            "width=device-width, initial-scale=1"
        );

        htmlOutput.setContent(
            htmlOutput.getContent().replace("{{SIDEBAR}}", getSidebar(page))
        );
        return htmlOutput;
    } catch (error) {
        Logger.log("Error in doGet: " + error.message);
        return HtmlService.createHtmlOutput(
            "An error occurred: " + error.message
        );
    }
}

//returns the URL of the Google Apps Script web app
function getScriptURL(qs = null) {
    var url = ScriptApp.getService().getUrl();
    if (qs) {
        if (qs.indexOf("?") === -1) {
            qs = "?" + qs;
        }
        url = url + qs;
    }
    return url;
}

//INCLUDE HTML PARTS, EG. JAVASCRIPT, CSS, OTHER HTML FILES
function include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getSidebar(activePage) {
    const views = {
        "employee-details": "Employee Details",
        "_sim-request-and-issuance": "SIM > Request and Issuance",
        "_sim-plans": "SIM > SIM Plans",
        "_sim-inventory": "SIM > SIM Inventory",
        "_returned-sim": "SIM > Returned SIM",
        "_mobile-request-and-issuance": "Mobile Phone > Request and Issuance",
        "_mobile-inventory": "Mobile Phone > Phone Inventory",
        "_mobile-returned": "Mobile Phone > Returned Phone",
        "rfp-group": "Bills and Payment > RFP Group",
        billing: "Bills and Payment > Billing",
        "rfp-summary": "Bills and Payment > RFP Summary",
        payment: "Bills and Payment > Payment",
        "excess-charges": "Excess Charges > Excess Charges",
        deduction: "Excess Charges > Deduction",
        "_ticket-management": "Ticket Management",
        "_audit-trail": "Audit Trail",
    };

    const generateUrl = (view) =>
        getScriptURL(view ? `view=${view}` : undefined);

    const generateListItem = (viewKey, label, isSubItem = false) => `
        <li>
            <a href="${generateUrl(viewKey)}"
                class="${
                    activePage === viewKey ? "bg-gray-100" : ""
                } flex items-center w-full p-2 text-gray-900 transition duration-75 rounded-lg ${
        isSubItem ? "pl-7" : ""
    } group hover:bg-gray-100">
                ${label}
            </a>
        </li>`;

    const generateSection = (title, items, sectionKey) => {
        const isOpen = items.some((item) => item.view === activePage);
        return `
            <li>
                <button type="button"
                    class="flex items-center w-full p-2 text-base text-gray-900 transition duration-75 rounded-lg group hover:bg-gray-100"
                    aria-controls="${sectionKey}" data-collapse-toggle="${sectionKey}" 
                    aria-expanded="${isOpen}">
                    <span class="flex-1 text-left rtl:text-right whitespace-wrap text-sm">${title}</span>
                    <svg class="w-3 h-3" aria-hidden="true" xmlns="http://www.w3.org/2000/svg" fill="none"
                        viewBox="0 0 10 6">
                        <path stroke="currentColor" stroke-linecap="round" stroke-linejoin="round" stroke-width="2"
                            d="m1 1 4 4 4-4" />
                    </svg>
                </button>
                <ul id="${sectionKey}" class="${
            isOpen ? "" : "hidden"
        } py-2 space-y-2">
                    ${items
                        .map((item) =>
                            generateListItem(item.view, item.label, true)
                        )
                        .join("")}
                </ul>
            </li>`;
    };

    const sidebarItems = [
        { view: "employee-details", label: "Employee Details" },
        {
            section: "SIM Card",
            key: "sim-card",
            items: [
                {
                    view: "_sim-request-and-issuance",
                    label: "Request and Issuance",
                },
                { view: "_sim-plans", label: "SIM Plans" },
                { view: "_sim-inventory", label: "SIM Inventory" },
                { view: "_returned-sim", label: "Returned SIM" },
            ],
        },
        {
            section: "Mobile Phone",
            key: "mobile-phone",
            items: [
                {
                    view: "_mobile-request-and-issuance",
                    label: "Request and Issuance",
                },
                { view: "_mobile-inventory", label: "Phone Inventory" },
                { view: "_mobile-returned", label: "Returned Phone" },
            ],
        },
        {
            section: "Bills and Payment",
            key: "bills-and-payment",
            items: [
                { view: "rfp-group", label: "RFP Group" },
                { view: "billing", label: "Billing" },
                { view: "rfp-summary", label: "RFP Summary" },
                { view: "payment", label: "Payment" },
            ],
        },
        {
            section: "Excess Charges and Deduction",
            key: "excess-charges-and-deduction",
            items: [
                {
                    view: "excess-charge-and-deduction-summary",
                    label: "Summary",
                },
                { view: "excess-charges", label: "Excess Charges" },
                { view: "deduction", label: "Deduction" },
            ],
        },
        { view: "_ticket-management", label: "Ticket Management" },
        { view: "_audit-trail", label: "Audit Trail" },
    ];

    const sidebarContent = sidebarItems
        .map((item) => {
            if (item.section) {
                return generateSection(item.section, item.items, item.key);
            }
            return generateListItem(item.view, item.label);
        })
        .join("");

    return `
        <aside id="logo-sidebar"
            class="fixed top-0 left-0 z-30 w-64 h-screen pt-20 transition-transform -translate-x-full bg-white border-r border-gray-200 sm:translate-x-0"
            aria-label="Sidebar">
            <div class="h-full px-3 pb-4 overflow-y-auto bg-white !text-sm">
                <ul class="space-y-2 !font-medium">
                    ${sidebarContent}
                </ul>
            </div>
        </aside>
    `;
}
