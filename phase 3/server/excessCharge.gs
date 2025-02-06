const EXCESS_CHARGES_HISTORY = "EXCESS_CHARGES_HISTORY";
function addBillRecordToExcessCharges(data) {
    // data.id = await getLastId()
    // Adding record to Excess Charges sheet via key order
    addRecordToSheet(data, EXCESS_CHARGES, [], KEY_ORDER.EXCESS_CHARGES);

    // Appending the data using built-in function appendRow()
    // var sheet = accessSheet(EXCESS_CHARGES);
    // sheet.appendRow(data);
}

function processPayrollData() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const excessChargesSheet = ss.getSheetByName("EXCESS_CHARGES");
    const excessChargesHistorySheet = ss.getSheetByName(
        "EXCESS_CHARGES_HISTORY"
    );

    // Mocking the uploaded Payroll Data
    const payrollData = [
        [
            "TRANSACTION_ID",
            "EMPLOYEE_NO",
            "EMPLOYEE_NAME",
            "COMPANY",
            "DEDUCTION_AMOUNT",
            "DEDUCTION_DATE",
        ],
        ["TXN002", "EMP001", "John Doe", "Company A", 200.0, "2024-09-10"], // Payroll deduction exceeding available charge
    ];

    const headers = payrollData[0];
    const data = payrollData.slice(1);

    data.forEach((row) => {
        const transactionId = row[0];
        const employeeNo = row[1];
        const deductionAmount = parseFloat(row[4]) || 0;
        const deductionDate = row[5] || "";

        handlePayrollData(
            transactionId,
            employeeNo,
            deductionAmount,
            deductionDate,
            excessChargesSheet,
            excessChargesHistorySheet
        );
    });

    Logger.log("Processing complete.");
}

function handlePayrollData(
    transactionId,
    employeeNo,
    totalDeductionAmount,
    deductionDate,
    excessChargesSheet,
    excessChargesHistorySheet
) {
    Logger.log(
        `Processing payroll data for EMPLOYEE_NO: ${employeeNo}, TRANSACTION_ID: ${transactionId}`
    );

    const excessChargesData = excessChargesSheet.getDataRange().getValues();
    const headers = excessChargesData[0];
    const employeeExcessCharges = excessChargesData
        .slice(1)
        .filter((row) => row[headers.indexOf("EMPLOYEE_NO")] === employeeNo);

    let remainingDeduction = totalDeductionAmount;
    let newHistoryRecords = [];

    // Get the current maximum ECH_ID and increment for new records
    const existingHistoryData = excessChargesHistorySheet
        .getDataRange()
        .getValues();
    const historyHeaders = existingHistoryData[0];
    let lastECH_ID = 0;
    if (existingHistoryData.length > 1) {
        lastECH_ID = Math.max(
            ...existingHistoryData
                .slice(1)
                .map((row) => row[historyHeaders.indexOf("ECH_ID")])
        );
    }

    employeeExcessCharges.forEach((chargeRecord, index) => {
        if (remainingDeduction <= 0) return;

        const previousRemainingCharge =
            chargeRecord[headers.indexOf("REMAINING_EXCESS_CHARGE")];
        const deductionToApply = Math.min(
            previousRemainingCharge,
            remainingDeduction
        );

        // Update remaining charge in the current excess charge record
        excessChargesSheet
            .getRange(index + 2, headers.indexOf("REMAINING_EXCESS_CHARGE") + 1)
            .setValue(previousRemainingCharge - deductionToApply);
        excessChargesSheet
            .getRange(index + 2, headers.indexOf("LAST_UPDATE_DATE") + 1)
            .setValue(deductionDate);
        excessChargesSheet
            .getRange(index + 2, headers.indexOf("LAST_UPDATE_TYPE") + 1)
            .setValue("Payroll Deduction");

        remainingDeduction -= deductionToApply;

        // Append to excessChargesHistoryData
        const newECHRecord = [
            ++lastECH_ID, // Incremental ECH_ID
            chargeRecord[headers.indexOf("EC_ID")],
            transactionId,
            employeeNo,
            chargeRecord[headers.indexOf("EXCESS_CHARGE_DATE")],
            chargeRecord[headers.indexOf("EXCESS_CHARGE")],
            previousRemainingCharge -
                (previousRemainingCharge - deductionToApply),
            deductionDate,
            deductionToApply,
            previousRemainingCharge - deductionToApply,
            "Payroll Deduction",
            deductionDate,
            "Payroll",
        ];
        newHistoryRecords.push(newECHRecord);
    });

    // If there's remaining deduction, apply it as a negative value to the last record
    if (remainingDeduction > 0) {
        const lastRecord =
            employeeExcessCharges[employeeExcessCharges.length - 1];
        if (lastRecord) {
            const lastRecordIndex =
                headers.indexOf("REMAINING_EXCESS_CHARGE") + 1;
            const newRemainingCharge =
                lastRecord[lastRecordIndex - 1] - remainingDeduction;
            excessChargesSheet
                .getRange(
                    employeeExcessCharges.length + 1,
                    headers.indexOf("REMAINING_EXCESS_CHARGE") + 1
                )
                .setValue(newRemainingCharge);

            const newHistoryRecord = [
                ++lastECH_ID, // Incremental ECH_ID
                lastRecord[headers.indexOf("EC_ID")],
                transactionId,
                employeeNo,
                lastRecord[headers.indexOf("EXCESS_CHARGE_DATE")],
                lastRecord[headers.indexOf("EXCESS_CHARGE")],
                0,
                deductionDate,
                remainingDeduction,
                newRemainingCharge,
                "Payroll Deduction (Negative)",
                deductionDate,
                "Payroll",
            ];
            newHistoryRecords.push(newHistoryRecord);
        }
    }

    if (newHistoryRecords.length > 0) {
        excessChargesHistorySheet
            .getRange(
                excessChargesHistorySheet.getLastRow() + 1,
                1,
                newHistoryRecords.length,
                newHistoryRecords[0].length
            )
            .setValues(newHistoryRecords);
    }
}

// Helper function to format Excel date serial number to MM/DD/YYYY
function formatDateToMMDDYYYY(date) {
    const mm = String(date.getMonth() + 1).padStart(2, "0"); // Months are zero-based
    const dd = String(date.getDate()).padStart(2, "0");
    const yyyy = date.getFullYear();
    return mm + "/" + dd + "/" + yyyy;
}

// Function to get the next ECH_ID
function getNextECH_ID() {
    const excessChargesHistorySheet =
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
            "EXCESS_CHARGES_HISTORY"
        );
    const lastRow = excessChargesHistorySheet.getLastRow();
    if (lastRow <= 1) {
        return 1; // Start from 1 if there are no records yet
    } else {
        const lastECH_ID = excessChargesHistorySheet
            .getRange(lastRow, 1)
            .getValue();
        return parseInt(lastECH_ID) + 1;
    }
}

// Utility function to format date to MM/DD/YYYY
// function formatDateToMMDDYYYY(date) {
//   const mm = String(date.getMonth() + 1).padStart(2, '0'); // Month is zero-indexed
//   const dd = String(date.getDate()).padStart(2, '0');
//   const yyyy = date.getFullYear();
//   return mm + '/' + dd + '/' + yyyy;
// }

function applyWaiverLetter(uploadedData) {
    const excessChargesSheet =
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName("EXCESS_CHARGES");
    const excessChargesData = excessChargesSheet.getDataRange().getValues();

    const headerRow = excessChargesData[0];
    const ecIdIndex = headerRow.indexOf("EC_ID");
    const groupIdIndex = headerRow.indexOf("GROUP_ID");
    const remainingExcessChargeIndex = headerRow.indexOf(
        "REMAINING_EXCESS_CHARGE"
    );
    const excessChargeIndex = headerRow.indexOf("EXCESS_CHARGE");
    const previousExcessChargeIndex = headerRow.indexOf(
        "PREVIOUS_EXCESS_CHARGE"
    );
    const amountDeductedIndex = headerRow.indexOf("AMOUNT_DEDUCTED");
    const deductionDateIndex = headerRow.indexOf("DEDUCTION_DATE");
    const waivedAmount = headerRow.indexOf("WAIVED_AMOUNT");
    const waivedDateIndex = headerRow.indexOf("WAIVED_DATE");
    const referenceIndex = headerRow.indexOf("REFERENCE");
    const excessChargeDateIndex = headerRow.indexOf("EXCESS_CHARGE_DATE");

    const processedRecords = new Set();
    let updatedRecordsCount = 0;

    for (let i = 1; i < excessChargesData.length; i++) {
        if (excessChargesData[i][waivedDateIndex] !== "") {
            processedRecords.add(
                excessChargesData[i][groupIdIndex] +
                    "-" +
                    excessChargesData[i][waivedDateIndex]
            );
        }
    }

    for (let i = 0; i < uploadedData.length; i++) {
        const record = uploadedData[i];
        const groupId = record["GROUP ID"];
        const approvedBy = record["APPROVED BY"];
        const waivedDate = new Date(
            (record["APPROVAL DATE"] - 25569) * 86400 * 1000
        ); // Convert Excel date to JS date
        const waivedAmount = record["WAIVED_AMOUNT"];
        const waivedDateString = formatDateToMMDDYYYY(waivedDate);
        const excessChargeDate = new Date(
            (record["EXCESS CHARGE DATE"] - 25569) * 86400 * 1000
        );
        const excessChargeDateString = formatDateToMMDDYYYY(excessChargeDate);

        // Check if the record has already been processed
        if (processedRecords.has(groupId + "-" + waivedDateString)) {
            console.log(
                `Warning: Record for group ID ${groupId} with waived date ${waivedDateString} has already been processed.`
            );
            continue;
        }

        // Find the latest record for this group ID with matching billing date
        let latestRecordIndex = -1;
        for (let j = 1; j < excessChargesData.length; j++) {
            const billingDateString = formatDateToMMDDYYYY(
                new Date(excessChargesData[j][excessChargeDateIndex])
            );
            if (
                excessChargesData[j][groupIdIndex] == groupId &&
                excessChargesData[j][waivedDateIndex] === "" &&
                billingDateString === excessChargeDateString
            ) {
                if (
                    latestRecordIndex == -1 ||
                    excessChargesData[j][0] >
                        excessChargesData[latestRecordIndex][0]
                ) {
                    latestRecordIndex = j;
                }
            }
        }

        if (
            latestRecordIndex != -1 &&
            excessChargesData[latestRecordIndex][deductionDateIndex] === ""
        ) {
            // Update the sheet
            excessChargesSheet
                .getRange(latestRecordIndex + 1, remainingExcessChargeIndex + 1)
                .setValue(0);
            excessChargesSheet
                .getRange(latestRecordIndex + 1, excessChargeIndex + 1)
                .setValue(0);
            excessChargesSheet
                .getRange(latestRecordIndex + 1, previousExcessChargeIndex + 1)
                .setValue(0);
            excessChargesSheet
                .getRange(latestRecordIndex + 1, waivedDateIndex + 1)
                .setValue(waivedDate);
            excessChargesSheet
                .getRange(latestRecordIndex + 1, waivedAmountIndex + 1)
                .setValue(waivedAmount);
            // excessChargesSheet.getRange(latestRecordIndex + 1, referenceIndex + 1).setValue(approvedBy);
            excessChargesSheet
                .getRange(latestRecordIndex + 1, referenceIndex + 1)
                .setValue(formatDateToMMDDYYYY(new Date())); // Set the modification date

            console.log(
                `Updated excess charge for group ID ${groupId}: remainingExcessCharge=0, excessCharge=0, previousExcessCharge=0, waivedDate=${waivedDateString}`
            );
            // return "success";
            updatedRecordsCount++;
        } else {
            console.log(
                `Warning: No record found or already processed for group ID ${groupId}`
            );
            // return "error";
        }
    }

    return updatedRecordsCount;
}

// Calculates the Total Excess Charge, Total Amount Deducted, Total Remaining Excess Charge for each Employee
// function calculateSummaryData(dateFrom = null, dateTo = null) {
//   const excessChargesData = getMergedData(dateFrom, dateTo); // Fetch and filter based on dates if provided

//   const processedBillIds = new Set();

//   const summaryData = Object.values(excessChargesData.reduce((acc, record) => {
//     const {
//       EMPLOYEE_NO,
//       EMPLOYEE_NAME,
//       DEPARTMENT,
//       COMPANY,
//       EXCESS_CHARGE,
//       AMOUNT_DEDUCTED,
//       REMAINING_EXCESS_CHARGE,
//       BILL_ID
//     } = record;

//     if (!acc[EMPLOYEE_NO]) {
//       acc[EMPLOYEE_NO] = {
//         EMPLOYEE_NO,
//         EMPLOYEE_NAME,
//         DEPARTMENT,
//         COMPANY,
//         TOTAL_EXCESS_CHARGE: 0,
//         TOTAL_AMOUNT_DEDUCTED: 0,
//         TOTAL_REMAINING_EXCESS_CHARGE: 0
//       };
//     }

//     if (!processedBillIds.has(BILL_ID)) {
//       acc[EMPLOYEE_NO].TOTAL_EXCESS_CHARGE += Number(EXCESS_CHARGE) || 0;
//       processedBillIds.add(BILL_ID);
//     }

//     acc[EMPLOYEE_NO].TOTAL_AMOUNT_DEDUCTED += Number(AMOUNT_DEDUCTED) || 0;
//     acc[EMPLOYEE_NO].TOTAL_REMAINING_EXCESS_CHARGE = Number(REMAINING_EXCESS_CHARGE) || 0;

//     return acc;
//   }, {}));

//   return [excessChargesData, summaryData];
// }

// // Returns a merged data needed for Excess charge
// function getMergedData(dateFrom = null, dateTo = null) {
//   const billingData = getDataFromSheet(BILLING, SHEETS.BILLING.DATE_FIELDS);
//   const simRequestAndIssuanceData = getDataFromSheet(SIM_REQUEST_AND_ISSUANCE, SHEETS.SIM_REQUEST_AND_ISSUANCE.DATE_FIELDS);
//   const simInventoryData = getDataFromSheet(SIM_INVENTORY, SHEETS.SIM_INVENTORY.DATE_FIELDS);
//   const simPlanData = getDataFromSheet(SIM_PLANS, SHEETS.SIM_PLANS.DATE_FIELDS);
//   const employeeData = getDataFromSheet(EMPLOYEE_DETAILS, SHEETS.EMPLOYEE_DETAILS.DATE_FIELDS);
//   const excessChargesData = getDataFromSheet(EXCESS_CHARGES, SHEETS.EXCESS_CHARGES.DATE_FIELDS);

//   const fromDate = dateFrom ? new Date(dateFrom) : null;
//   const toDate = dateTo ? new Date(dateTo) : null;

//   const filteredExcessChargesData = excessChargesData.filter(excessCharge => {
//     const chargeDate = new Date(excessCharge.EXCESS_CHARGE_DATE);
//     let withinDateRange = true;

//     if (fromDate && toDate) {
//       withinDateRange = chargeDate >= fromDate && chargeDate <= toDate;
//     } else if (fromDate) {
//       withinDateRange = chargeDate >= fromDate;
//     } else if (toDate) {
//       withinDateRange = chargeDate <= toDate;
//     }

//     return withinDateRange;
//   });

//   return filteredExcessChargesData.map(excessCharge => {
//     const bill = billingData.find(b => b.BILL_ID == excessCharge.BILL_ID) || {};
//     const simRequest = simRequestAndIssuanceData.find(req => req.ISSUANCE_NO == bill.ISSUANCE_NO) || {};
//     const simInventory = simInventoryData.find(sim => sim.SIM_CARD_ID == bill.SIM_CARD_ID) || {};
//     const simPlan = simPlanData.find(plan => plan.PLAN_ID == simInventory.PLAN_ID) || {};
//     const employee = employeeData.find(emp => emp.GROUP_ID == excessCharge.GROUP_ID) || {};

//     return {
//       EC_ID: excessCharge.EC_ID,
//       BILL_ID: excessCharge.BILL_ID,
//       RFP_NO: bill.RFP_NO || '',
//       SIM_CARD_ID: bill.SIM_CARD_ID || '',
//       MOBILE_NO: simInventory.MOBILE_NO,
//       ACCOUNT_NO: simInventory.ACCOUNT_NO,
//       NETWORK_PROVIDER: simPlan.NETWORK_PROVIDER,
//       PROVIDER_COMPANY_NAME: simPlan.PROVIDER_COMPANY_NAME,
//       GROUP_ID: excessCharge.GROUP_ID || '',
//       EMPLOYEE_NO: excessCharge.EMPLOYEE_NO || '',
//       EMPLOYEE_NAME: employee.FULL_NAME || '',
//       DEPARTMENT: employee.DEPARTMENT || '',
//       COMPANY: employee.COMPANY_NAME || '',
//       EXCESS_CHARGE_DATE: excessCharge.EXCESS_CHARGE_DATE || '',
//       EXCESS_CHARGE: excessCharge.EXCESS_CHARGE || '',
//       AMOUNT_DEDUCTED: excessCharge.AMOUNT_DEDUCTED || '',
//       DEDUCTION_DATE: excessCharge.DEDUCTION_DATE || '',
//       PREVIOUS_EXCESS_CHARGE: excessCharge.PREVIOUS_EXCESS_CHARGE || '',
//       REMAINING_EXCESS_CHARGE: excessCharge.REMAINING_EXCESS_CHARGE || '',
//       WAIVED_DATE: excessCharge.WAIVED_DATE || '',
//       LAST_UPDATE_TYPE: excessCharge.LAST_UPDATE_TYPE || '',
//       LAST_UPDATE_DATE: excessCharge.LAST_UPDATE_DATE || '',
//     };
//   });
// }

function calculateSummaryData() {
    // Fetch the main records data
    const excessChargesData = getExcessChargesData();
    const deductionData = getDeductionData();

    // Fetch the summary data
    const excessChargesSummary = getExcessChargesSummary(
        excessChargesData,
        deductionData
    );

    console.log(excessChargesData);
    console.log(deductionData);
    console.log(excessChargesSummary);

    console.log(typeof excessChargesData);
    console.log(typeof deductionData);
    console.log(typeof excessChargesSummary);

    // Return both datasets as an array
    return [excessChargesData, excessChargesSummary, deductionData];
}

function getDeductionData() {
    try {
        const spaceRegex = /[\s]/g;
        const toMatrixHeaders = (key, headers) =>
            headers
                .map((item) => {
                    if (new RegExp(spaceRegex).test(item)) {
                        return `${key}."${item}"`;
                    }
                    return `${key}.${item}`;
                })
                .join(", ");
        const deductionSheet = new Utils.Sheet("DEDUCTION", {
            row: { start: 1 },
        });
        const deductionDataSet = deductionSheet.toObject();

        const query = `
        SELECT
          d.DEDUCTION_ID,
          d.EMPLOYEE_NO,
          d.EMPLOYEE_NAME,
          d.AMOUNT,
          d.DATE_UPLOADED,
          d.FILE_NAME
        FROM ? AS d
      `;

        const execution = Utils.sql(query, [deductionDataSet]);

        // Check if dataset is empty, return placeholder if needed
        if (!execution || execution.length === 0) {
            return JSON.stringify([
                {
                    DEDUCTION_ID: "",
                    EMPLOYEE_NO: "",
                    EMPLOYEE_NAME: "",
                    AMOUNT: "",
                    DATE_UPLOADED: "",
                    FILE_NAME: "",
                },
            ]);
        }

        execution.forEach((row) => {
            if (row.DATE_UPLOADED) {
                row.DATE_UPLOADED = Utilities.formatDate(
                    new Date(row.DATE_UPLOADED),
                    Session.getScriptTimeZone(),
                    "MM/dd/yyyy"
                );
            }
        });

        return JSON.stringify(execution);
    } catch (error) {
        return Utils.ErrorHandler(error, {
            arguments,
            value: [],
        });
    }
}

function getExcessChargesData() {
    try {
        const spaceRegex = /[\s]/g;
        const toMatrixHeaders = (key, headers) =>
            headers
                .map((item) => {
                    if (new RegExp(spaceRegex).test(item)) {
                        return `${key}."${item}"`;
                    }
                    return `${key}.${item}`;
                })
                .join(", ");

        const billingSheet = new Utils.Sheet("BILLING", { row: { start: 1 } });
        const excessChargesSheet = new Utils.Sheet("EXCESS_CHARGES", {
            row: { start: 1 },
        });
        const rfpSummarySheet = new Utils.Sheet("RFP_SUMMARY", {
            row: { start: 1 },
        });

        const billingDataSet = billingSheet.toObject();
        const excessChargesDataSet = excessChargesSheet.toObject();
        const rfpSummaryDataSet = rfpSummarySheet.toObject();

        const query = `
        SELECT
          ec.EC_ID,
          ec.BILL_ID,
          b.RFP_NO,
          b.BILL_PERIOD_FROM,
          b.BILL_PERIOD_TO,
          ec.EXCESS_CHARGE_DATE,
          ec.EXCESS_CHARGE,
          b.SIM_INFO,
          b.EMPLOYEE_INFO,
          r.RFP_DATE
        FROM ? AS ec
        LEFT JOIN ? AS b ON ec.BILL_ID = b.BILL_ID
        LEFT JOIN ? AS r ON b.RFP_NO = r.RFP_NO
      `;

        const execution = Utils.sql(query, [
            excessChargesDataSet,
            billingDataSet,
            rfpSummaryDataSet,
        ]);

        // Check if dataset is empty, return placeholder if needed
        if (!execution || execution.length === 0) {
            return JSON.stringify([
                {
                    EC_ID: "",
                    BILL_ID: "",
                    RFP_NO: "",
                    GROUP_ID: "",
                    EMPLOYEE_NO: "",
                    EMPLOYEE_NAME: "",
                    DEPARTMENT: "",
                    // EMPLOYEE_INF0: "",
                    SIM_INFO: "",
                    COMPANY: "",
                    MOBILE_NO: "",
                    ACCOUNT_NO: "",
                    NETWORK_PROVIDER: "",
                    EXCESS_CHARGE_DATE: "",
                    EXCESS_CHARGE: "",
                },
            ]);
        }

        execution.forEach((row) => {
            if (row.EXCESS_CHARGE_DATE) {
                row.EXCESS_CHARGE_DATE = Utilities.formatDate(
                    new Date(row.EXCESS_CHARGE_DATE),
                    Session.getScriptTimeZone(),
                    "MM/dd/yyyy"
                );
            }
            if (row.RFP_DATE) {
                row.RFP_DATE = Utilities.formatDate(
                    new Date(row.RFP_DATE),
                    Session.getScriptTimeZone(),
                    "MM/dd/yyyy"
                );
            }
            if (row.BILL_PERIOD_FROM) {
                row.BILL_PERIOD_FROM = Utilities.formatDate(
                    new Date(row.BILL_PERIOD_FROM),
                    Session.getScriptTimeZone(),
                    "MM/dd/yyyy"
                );
            }
            if (row.BILL_PERIOD_TO) {
                row.BILL_PERIOD_TO = Utilities.formatDate(
                    new Date(row.BILL_PERIOD_TO),
                    Session.getScriptTimeZone(),
                    "MM/dd/yyyy"
                );
            }
        });

        return JSON.stringify(execution);
    } catch (error) {
        return Utils.ErrorHandler(error, {
            arguments,
            value: [],
        });
    }
    // return excessChargesData;  // Example return
}

// function to get the excess charge summary
function getExcessChargesSummary(excessChargesData, deductionsData) {
    let summary = {};

    const parsedExcessChargesData = JSON.parse(excessChargesData);
    const parsedDeductionsData = JSON.parse(deductionsData);

    if (
        parsedExcessChargesData.length === 0 ||
        parsedExcessChargesData[0].EC_ID == ""
    ) {
        return [
            {
                GROUP_ID: "",
                EMPLOYEE_NAME: "",
                EMPLOYEE_NO: "",
                COMPANY: "",
                DEPARTMENT: "",
                TOTAL_EXCESS_CHARGE: 0,
                TOTAL_REMAINING_EXCESS_CHARGE: 0,
                TOTAL_AMOUNT_DEDUCTED: 0,
                LATEST_TRANSACTION_DATE: "",
            },
        ];
    }

    // Process the excess charges data
    parsedExcessChargesData.forEach((item) => {
        const employeeInfo = JSON.parse(item.EMPLOYEE_INFO);
        const employeeNo = employeeInfo.EMPLOYEE_NO;

        // Initialize the summary object for the employee if not already present
        if (!summary[employeeNo]) {
            summary[employeeNo] = {
                GROUP_ID: employeeInfo.GROUP_ID,
                EMPLOYEE_NAME: employeeInfo.FULL_NAME,
                EMPLOYEE_NO: employeeInfo.EMPLOYEE_NO,
                COMPANY: employeeInfo.COMPANY_NAME,
                DEPARTMENT: employeeInfo.DEPARTMENT,
                TOTAL_EXCESS_CHARGE: 0,
                TOTAL_REMAINING_EXCESS_CHARGE: 0,
                TOTAL_AMOUNT_DEDUCTED: 0,
                LATEST_TRANSACTION_DATE: item.EXCESS_CHARGE_DATE, // Set initial date as excess charge date
            };
        }

        // Accumulate the total excess charge
        summary[employeeNo].TOTAL_EXCESS_CHARGE += item.EXCESS_CHARGE;

        // Update the LATEST_TRANSACTION_DATE only if no deductions exist later
        if (
            !summary[employeeNo].LATEST_TRANSACTION_DATE ||
            new Date(item.EXCESS_CHARGE_DATE) >
                new Date(summary[employeeNo].LATEST_TRANSACTION_DATE)
        ) {
            summary[employeeNo].LATEST_TRANSACTION_DATE =
                item.EXCESS_CHARGE_DATE;
        }
    });

    // Process the deductions data
    parsedDeductionsData.forEach((deduction) => {
        const employeeNo = deduction.EMPLOYEE_NO;

        // Skip records with empty or invalid EMPLOYEE_NO
        if (!employeeNo || employeeNo.trim() === "") {
            return;
        }

        // If the employee exists in the summary, update the total amount deducted
        if (!summary[employeeNo]) {
            // Add employees only present in the deductions dataset
            summary[employeeNo] = {
                GROUP_ID: "",
                EMPLOYEE_NAME: deduction.EMPLOYEE_NAME,
                EMPLOYEE_NO: deduction.EMPLOYEE_NO,
                COMPANY: "",
                DEPARTMENT: "",
                TOTAL_EXCESS_CHARGE: 0,
                TOTAL_REMAINING_EXCESS_CHARGE: 0,
                TOTAL_AMOUNT_DEDUCTED: 0,
                LATEST_TRANSACTION_DATE: deduction.DATE_UPLOADED, // Set initial date as deduction file uploaded
            };
        }

        // Accumulate the total amount deducted
        summary[employeeNo].TOTAL_AMOUNT_DEDUCTED += Number(
            deduction.AMOUNT || 0
        );

        // Compare and update the LATEST_TRANSACTION_DATE if the deduction date is later
        if (
            !summary[employeeNo].LATEST_TRANSACTION_DATE ||
            new Date(deduction.DATE_UPLOADED) >
                new Date(summary[employeeNo].LATEST_TRANSACTION_DATE)
        ) {
            summary[employeeNo].LATEST_TRANSACTION_DATE =
                deduction.DATE_UPLOADED;
        }
    });

    // Calculate TOTAL_REMAINING_EXCESS_CHARGE for each employee
    Object.values(summary).forEach((employee) => {
        employee.TOTAL_REMAINING_EXCESS_CHARGE =
            employee.TOTAL_EXCESS_CHARGE - employee.TOTAL_AMOUNT_DEDUCTED;
    });

    // Convert the summary object into an array
    return Object.values(summary);
}

// printing function for excess charges

function getExcessChargesTemplateData(data) {
    console.log(data);
    getExcessChargePdfUrl(data);
}

function getExcessChargeHrPdfUrl(staticValue, data) {
    const sheetName = "EXCESS_CHARGE_HR";
    const titleText = "";

    // // console.log(data)
    console.log(data.length);

    const pageHeight = 50; // Rows per page (adjust based on your layout)
    const staticRows = 9; // Static rows (title, date, report details, headers)
    const grandTotalRow = 1; // Grand total row
    const signatoryRows = 9; // Prepared by, Checked by, Received by (and their positions)

    // // Calculate total rows used by static rows, data, and signatory section
    const usedRows = staticRows + data.length + grandTotalRow + signatoryRows;

    // // Calculate total pages and determine blank rows required
    const totalPages = Math.ceil(usedRows / pageHeight); // Total pages needed
    const lastPageStartRow = (totalPages - 1) * pageHeight + 1; // Start row of the last page
    const blankRows = Math.max(0, lastPageStartRow - usedRows); // Blank rows to push signatories to last page

    // Calculate the endRow dynamically
    const endRow = usedRows + blankRows;

    console.log(endRow);

    const pdfUrl = createPDF(
        SPREADSHEET_ID,
        sheetName,
        { startRow: 0, startCol: 0, endRow: endRow, endCol: 10 },
        () => {
            // Call your data population functions here
            // populateExcessChargeReportTemplate(staticValue, data); // First method
            populateExcessChargeReportTemplate(
                staticValue,
                data,
                sheetName,
                titleText
            );

            // OR
            // populateAnotherTemplate(data); // Second method
        },
        {
            paperSize: "A4",
            orientation: "landscape",
            scale: 2,
            gridlines: false,
            verticalAlignment: "TOP",
            pageNum: "CENTER",
            leftMargin: 0.25,
            rightMargin: 0.25,
        }
    );

    return pdfUrl;
}

function getExcessChargePayrollPdfUrl(staticValue, data) {
    const sheetName = "EXCESS_CHARGE_PAYROLL";
    const titleText = "PAYROLL COPY";

    // // console.log(data)
    console.log(data.length);

    const pageHeight = 50; // Rows per page (adjust based on your layout)
    const staticRows = 9; // Static rows (title, date, report details, headers)
    const grandTotalRow = 1; // Grand total row
    const signatoryRows = 9; // Prepared by, Checked by, Received by (and their positions)

    // // Calculate total rows used by static rows, data, and signatory section
    const usedRows = staticRows + data.length + grandTotalRow + signatoryRows;

    // // Calculate total pages and determine blank rows required
    const totalPages = Math.ceil(usedRows / pageHeight); // Total pages needed
    const lastPageStartRow = (totalPages - 1) * pageHeight + 1; // Start row of the last page
    const blankRows = Math.max(0, lastPageStartRow - usedRows); // Blank rows to push signatories to last page

    // Calculate the endRow dynamically
    const endRow = usedRows + blankRows;

    console.log(endRow);

    const pdfUrl = createPDF(
        SPREADSHEET_ID,
        sheetName,
        { startRow: 0, startCol: 0, endRow: endRow, endCol: 10 },
        () => {
            // Call your data population functions here
            // populateExcessChargeReportTemplate(staticValue, data); // First method
            populateExcessChargeReportTemplate(
                staticValue,
                data,
                sheetName,
                titleText
            );

            // OR
            // populateAnotherTemplate(data); // Second method
        },
        {
            paperSize: "A4",
            orientation: "landscape",
            scale: 2,
            gridlines: false,
            verticalAlignment: "TOP",
            pageNum: "CENTER",
            leftMargin: 0.25,
            rightMargin: 0.25,
        }
    );

    return pdfUrl;
}

function processAndWriteDeductionData(data, fileName) {
    // const data = [
    //   {
    //     "Formula": "Quantity",
    //     "__EMPTY": "Hours",
    //     "__EMPTY_1": "Constant",
    //     "__EMPTY_2": "Salary Rate Type"
    //   },
    //   {
    //     "__EMPTY_2": "Salary Type",
    //     "__EMPTY_3": "Rate Type",
    //     "__EMPTY_4": "Rate"
    //   },
    //   {
    //     "Line #": 1,
    //     "Employee Code": "000735",
    //     "Employee Name": "ABAD, ARIEL B.",
    //     "Pay Item Code": 5,
    //     "Pay Item Name": "PLDT/CEL",
    //     "Type": "Deduction",
    //     "__EMPTY": 0,
    //     "__EMPTY_1": 0,
    //     "__EMPTY_2": "",
    //     "__EMPTY_3": "",
    //     "Amount": 32.46
    //   },
    //   {
    //     "Line #": 2,
    //     "Employee Code": "001163",
    //     "Employee Name": "ABAD, FRANZ KEVIN R.",
    //     "Pay Item Code": 5,
    //     "Pay Item Name": "PLDT/CEL",
    //     "Type": "Deduction",
    //     "__EMPTY": 0,
    //     "__EMPTY_1": 0,
    //     "__EMPTY_2": "",
    //     "__EMPTY_3": "",
    //     "Amount": 32.46
    //   }
    // ];

    // Get the spreadsheet and the "Deduction" sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const deductionSheet = ss.getSheetByName("DEDUCTION");

    // Get the last row with data in column A of the "Deduction" sheet
    const lastRow = deductionSheet.getLastRow();

    // Get the last ID from the sheet (assuming it's in the first column)
    let lastId;
    if (lastRow > 0) {
        const lastIdRange = deductionSheet.getRange(lastRow, 1);
        lastId = lastIdRange.getValue();
        // Check if lastId is a number, otherwise set it to 0
        lastId = isNaN(lastId) ? 0 : lastId;
    } else {
        lastId = 0; // No records, start with ID 1
    }

    let processedData = [];

    // Process the data
    for (let i = 2; i < data.length; i++) {
        if (String(data[i]["Employee Code"] || "").trim() !== "") {
            processedData.push([
                lastId + i - 1, // Use lastId for the first ID and increment
                `'${String(data[i]["Employee Code"]).padStart(6, "0")}`,
                data[i]["Employee Name"],
                data[i]["Amount"],
                Utilities.formatDate(
                    new Date(),
                    Session.getScriptTimeZone(),
                    "MM/dd/yyyy"
                ),
                fileName,
            ]);
        }
    }

    // Write the collected data to the sheet in a single operation
    if (processedData.length > 0) {
        const startRow = lastRow + 1;
        const numRows = processedData.length;
        const startColumn = 1;
        deductionSheet
            .getRange(startRow, startColumn, numRows, processedData[0].length)
            .setValues(processedData);
        return processedData.length;
    }

    return 0;
}

// function to populate the excess charge report template
function populateExcessChargeReportTemplate(
    staticValue,
    data,
    sheetName,
    titleText
) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(sheetName);

    const reference = getReferenceData();
    const parsedReference = JSON.parse(reference);

    // Check if the sheet exists
    if (!sheet) {
        Logger.log(`Sheet with name "${sheetName}" does not exist.`);
        return;
    }

    // Clear the sheet before writing
    sheet.clear();

    // Common styles
    const headerFont = {
        fontSize: 12,
        fontFamily: "Poppins",
        fontWeight: "bold",
    };
    const regularFont = { fontSize: 10, fontFamily: "Poppins" };

    // Insert the sheet name and today's date in a single batch
    const today = new Date();
    const formattedDate = Utilities.formatDate(
        today,
        Session.getScriptTimeZone(),
        "MM/dd/yyyy"
    );
    const headerValues = [
        [titleText, "", "", "", "", "", "", "", "", `Date: ${formattedDate}`],
    ];
    sheet.getRange("A1:J1").setValues(headerValues);
    sheet
        .getRange("A1")
        .setFontSize(14)
        .setFontWeight("bold")
        .setHorizontalAlignment("LEFT");
    sheet.getRange("J1").setFontSize(14).setHorizontalAlignment("LEFT");

    // Insert report details in batch
    const reportDetails = [
        ["RFP NO", staticValue.RFP_NO],
        ["RFP DATE", staticValue.RFP_DATE],
        ["NETWORK PROVIDER", staticValue.NETWORK_PROVIDER],
    ];
    sheet.getRange("A3:B5").setValues(reportDetails);
    sheet
        .getRange("A3:B5")
        .setFontSize(10)
        .setFontFamily("Poppins")
        .setHorizontalAlignment("LEFT");
    sheet.getRange("A3:B5").setBorder(true, true, true, true, true, true);

    // Define headers for the main data table
    const headers = [
        "EMPLOYEE NO",
        "EMPLOYEE NAME",
        "DEPARTMENT",
        "COMPANY",
        "BILL PERIOD FROM",
        "BILL PERIOD TO",
        "SIM CARD ID",
        "MOBILE NO",
        "ACCOUNT NO",
        "EXCESS CHARGE",
    ];

    // Process the data with optimized processing
    const rows = data;

    // Insert headers and data
    const startRow = 10;
    sheet.getRange(startRow, 1, 1, headers.length).setValues([headers]);
    sheet
        .getRange(startRow + 1, 1, rows.length, headers.length)
        .setValues(rows);

    // Format headers in batch
    const headerRange = sheet.getRange(startRow, 1, 1, headers.length);
    headerRange
        .setFontWeight("bold")
        .setFontSize(12)
        .setHorizontalAlignment("CENTER")
        .setFontFamily("Poppins")
        .setBorder(true, true, true, true, true, true);

    // Format data in batch
    const dataRange = sheet.getRange(
        startRow + 1,
        1,
        rows.length,
        headers.length
    );
    dataRange
        .setFontSize(11)
        .setFontFamily("Poppins")
        .setHorizontalAlignment("CENTER")
        .setBorder(true, true, true, true, true, true);

    // Format excess charge column
    const excessChargeColumn = sheet.getRange(startRow + 1, 10, rows.length, 1);
    excessChargeColumn
        .setNumberFormat("#,##0.00")
        .setHorizontalAlignment("RIGHT");

    // Add grand total row
    const totalRowIndex = startRow + rows.length + 1;
    sheet
        .getRange(totalRowIndex, 1)
        .setValue("GRAND TOTAL")
        .setFontWeight("bold")
        .setFontSize(12)
        .setHorizontalAlignment("CENTER");
    sheet
        .getRange(totalRowIndex, 10)
        .setFormula(`=SUM(J${startRow + 1}:J${startRow + rows.length})`)
        .setNumberFormat("#,##0.00")
        .setFontWeight("bold")
        .setFontSize(12)
        .setHorizontalAlignment("RIGHT");

    // Apply borders to the grand total row
    sheet
        .getRange(totalRowIndex, 1, 1, 10)
        .setBorder(true, true, true, true, true, true);

    // Signatories section
    const pageHeight = 50;
    const totalUsedRows = totalRowIndex + 3;
    const totalPages = Math.ceil(totalUsedRows / pageHeight);
    const lastPageStartRow = (totalPages - 1) * pageHeight + 1;
    const blankRowsNeeded = Math.max(0, lastPageStartRow - totalUsedRows);

    if (blankRowsNeeded > 0) {
        sheet.insertRowsAfter(totalRowIndex, blankRowsNeeded);
    }

    const signatoryStartRow = totalRowIndex + blankRowsNeeded + 5;
    // const signatoryValues = [
    //   ['Prepared by:', 'Checked By:', 'Received By:'],
    //   ['', '', ''],
    //   ['Cottish Star Papa', 'Rienalyn Villanueva', 'Nida Fajiculay'],
    //   ['Admin and Benefits Specialist', 'Human Resources Supervisor', 'Payroll Specialist']
    // ];
    const signatoryValues = [
        ["Prepared by:", "Checked By:", "Received By:"],
        ["", "", ""],
        [
            parsedReference[3].NAME,
            parsedReference[4].NAME,
            parsedReference[5].NAME,
        ],
        [
            parsedReference[3].POSITION,
            parsedReference[4].POSITION,
            parsedReference[5].POSITION,
        ],
    ];
    sheet
        .getRange(signatoryStartRow, 1, 4, 3)
        .setFontWeight("bold")
        .setValues(signatoryValues);
    sheet.getRange(signatoryStartRow, 1, 1, 3).setFontSize(11);
    sheet.getRange(signatoryStartRow + 1, 1, 1, 3).setFontSize(11);

    Logger.log(`Report successfully populated in sheet: "${sheetName}"`);
}
