document.addEventListener("DOMContentLoaded", function () {
    const modal = document.getElementById("dynamic-modal");
    const modalTitle = document.getElementById("modal-title");
    const closeModalBtn = document.getElementById("modal-close-btn");

    const paginationInfo = document.getElementById("pagination-info");
    const prevPageBtn = document.getElementById("pagination-prev");
    const nextPageBtn = document.getElementById("pagination-next");
    const searchInput = document.getElementById("pagination-search");
    const exportCsvBtn = document.getElementById("export-csv");

    let currentTableKey = null;
    let tableInstance = null;
    let currentPage = 1;

    const sampleData = {
        "totalSimCards": [
            { SIM_CARD_ID: "101", MOBILE_NO: "09123456789", ACCOUNT_NO: "A12345", STATUS: "Active", PLAN_ID: "PL001", CATEGORY: "Postpaid" },
            { SIM_CARD_ID: "102", MOBILE_NO: "09987654321", ACCOUNT_NO: "A67890", STATUS: "Inactive", PLAN_ID: "PL002", CATEGORY: "Prepaid" }
        ],
        "activeUsers": [
            { USER_ID: "U001", NAME: "John Doe", EMAIL: "john@example.com", ROLE: "Admin", STATUS: "Active" },
            { USER_ID: "U002", NAME: "Jane Smith", EMAIL: "jane@example.com", ROLE: "User", STATUS: "Inactive" }
        ]
    };

    const tableConfigs = {
        "totalSimCards": {
            tableId: "#dynamic-table",
            colvisElement: "div#dynamic-table-colvis",
            csvButtonId: "button#export-csv",
            paginationLimitId: "select#rows-per-page",
            paginationNextId: "button#pagination-next",
            paginationPreviousId: "button#pagination-prev",
            paginationSearchId: "input#pagination-search",
            headers: ["SIM_CARD_ID", "MOBILE_NO", "ACCOUNT_NO", "STATUS", "PLAN_ID", "CATEGORY"],
            dataKey: "totalSimCards",
            filename: "total-sim-cards-report",
        },
        "activeUsers": {
            tableId: "#dynamic-table",
            colvisElement: "div#dynamic-table-colvis",
            csvButtonId: "button#export-csv",
            paginationLimitId: "select#rows-per-page",
            paginationNextId: "button#pagination-next",
            paginationPreviousId: "button#pagination-prev",
            paginationSearchId: "input#pagination-search",
            headers: ["USER_ID", "NAME", "EMAIL", "ROLE", "STATUS"],
            dataKey: "activeUsers",
            filename: "active-users-report",
        }
    };

    document.querySelectorAll(".modal-open-btn").forEach((button) => {
        button.addEventListener("click", function () {
            currentTableKey = this.getAttribute("data-table-key");
            modalTitle.textContent = this.getAttribute("data-modal-title");
            modal.classList.remove("hidden");

            currentPage = 1;
            loadData();
        });
    });

    closeModalBtn.addEventListener("click", function () {
        modal.classList.add("hidden");
    });

    function initializeTable(config, data) {
        return new TableJS(document.querySelector(config.tableId), {
            dataset: {
                init: false,
                collection: data,
                rendering: renderTable
            },
            colvis: {
                element: config.colvisElement,
                exclude: ['.checkbox-column', '.actions-column'],
                hide: [],
            },
            exportAs: [
                {
                    as: 'csv',
                    element: config.csvButtonId,
                    filename: config.filename,
                    exclude: ['.checkbox-column', '.actions-column'],
                }
            ],
            paginate: [
                {
                    to: 'api',
                    as: 'limit',
                    element: config.paginationLimitId,
                    output: updatePaginationInfo,
                },
                {
                    to: 'api',
                    as: 'next',
                    element: config.paginationNextId,
                    output: updatePaginationInfo,
                },
                {
                    to: 'api',
                    as: 'previous',
                    element: config.paginationPreviousId,
                    output: updatePaginationInfo,
                },
                {
                    to: 'local',
                    as: 'search',
                    element: config.paginationSearchId,
                    exclude: ['.checkbox-column', '.actions-column'],
                    output: updatePaginationInfo,
                }
            ]
        });
    }

    function loadData() {
        if (!tableConfigs[currentTableKey]) return;
        const config = tableConfigs[currentTableKey];
        const data = sampleData[config.dataKey] || [];

        tableInstance = initializeTable(config, data);
        tableInstance.dataset.collection = data;
        tableInstance.dataset.rendering();
    }

    function renderTable({ data, table }) {
        const thead = table.querySelector('thead');
        const tbody = table.querySelector('tbody');
        const config = tableConfigs[currentTableKey];

        thead.innerHTML = '';
        tbody.innerHTML = '';

        const headerRow = document.createElement('tr');
        config.headers.forEach(header => {
            const th = document.createElement('th');
            th.textContent = header.replace(/_/g, ' ');
            headerRow.appendChild(th);
        });
        thead.appendChild(headerRow);

        data.forEach(row => {
            const tr = document.createElement('tr');
            config.headers.forEach(header => {
                const td = document.createElement('td');
                td.textContent = row[header] || "No data";
                tr.appendChild(td);
            });
            tbody.appendChild(tr);
        });
    }

    function updatePaginationInfo({ current_page, total_page }) {
        paginationInfo.textContent = `Page ${current_page} of ${total_page}`;
    }
});
