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

    console.log(excessChargesData);
    console.log(deductionData);

    // Fetch the summary data
    const excessChargesSummary = getExcessChargesSummary(
        excessChargesData,
        deductionData
    );

    console.log(excessChargesSummary);

    console.log(typeof excessChargesData);
    console.log(typeof excessChargesSummary);

    // Return both datasets as an array
    return [excessChargesData, excessChargesSummary, deductionData];
    // return excessChargesData;
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
          d.DATE_UPLOADED
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

        console.log(execution);
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

        // console.log(billingDataSet);
        // console.log(excessChargesDataSet);
        // console.log(rfpSummaryDataSet);

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

        // // Make this part dynamic
        //       const excessChargesHeaders = [
        //           'EC_ID',
        //           'BILL_ID',
        //           'RFP_NO',
        //           'GROUP_ID',
        //           'EMPLOYEE_NO',
        //           'EMPLOYEE_NAME',
        //           'DEPARTMENT',
        //           'COMPANY',
        //           'MOBILE_NO',
        //           'ACCOUNT_NO',
        //           'NETWORK_PROVIDER',
        //           'EXCESS_CHARGE_DATE',
        //           'EXCESS_CHARGE',
        //           'REMAINING_EXCESS_CHARGE',
        //           // 'LAST_UPDATE_TYPE',
        //           // 'LAST_UPDATE_DATE',
        //       ];

        //       const summaryHeaders = [
        //           'GROUP_ID',
        //           'EMPLOYEE_NO',
        //           'EMPLOYEE_NAME',
        //           'DEPARTMENT',
        //           'COMPANY',
        //           'TOTAL_EXCESS_CHARGE',
        //           'TOTAL_REMAINING_EXCESS_CHARGE',
        //           'TOTAL_AMOUNT_DEDUCTED',
        //           'AS_OF_DATE'
        //       ];
        // Format date fields (assuming EXCESS_CHARGE_DATE)
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
        });

        console.log(execution);
        return JSON.stringify(execution);
    } catch (error) {
        return Utils.ErrorHandler(error, {
            arguments,
            value: [],
        });
    }
    // return excessChargesData;  // Example return
}

// Function to calculate totals based on the data
// function getExcessChargesSummary(excessChargesData) {
//     let summary = {};

//     const parsedExcessChargesData = JSON.parse(excessChargesData);

//     if(parsedExcessChargesData[0].EC_ID == ""){

//       summary = [{
//         GROUP_ID: "",
//         EMPLOYEE_NAME: "",
//         EMPLOYEE_NO: "",
//         COMPANY: "",
//         DEPARTMENT: "",
//         TOTAL_EXCESS_CHARGE: 0,
//         TOTAL_REMAINING_EXCESS_CHARGE: 0,
//         TOTAL_AMOUNT_DEDUCTED: 0
//       }]

//       return Object.values(summary);
//     }

//     // Iterate through each entry in the data
//     parsedExcessChargesData.forEach(item => {
//         const employeeInfo = JSON.parse(item.EMPLOYEE_INFO);
//         const employeeNo = employeeInfo.EMPLOYEE_NO;

//         // If employee summary doesn't exist, initialize it
//         if (!summary[employeeNo]) {
//             summary[employeeNo] = {
//                 GROUP_ID: employeeInfo.GROUP_ID,
//                 EMPLOYEE_NAME: employeeInfo.FULL_NAME,
//                 EMPLOYEE_NO: employeeInfo.EMPLOYEE_NO,
//                 COMPANY: employeeInfo.COMPANY_NAME,
//                 DEPARTMENT: employeeInfo.DEPARTMENT,
//                 TOTAL_EXCESS_CHARGE: 0,
//                 TOTAL_REMAINING_EXCESS_CHARGE: 0,
//                 TOTAL_AMOUNT_DEDUCTED: 0  // Can be updated if needed
//             };
//         }

//         // Add the current item's excess charge and remaining excess charge to the total
//         summary[employeeNo].TOTAL_EXCESS_CHARGE += item.EXCESS_CHARGE;
//         summary[employeeNo].TOTAL_REMAINING_EXCESS_CHARGE += item.REMAINING_EXCESS_CHARGE;
//     });

//     // Convert the summary object into an array for front-end consumption
//     return Object.values(summary);
// }

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
            };
        }

        // Accumulate the total excess charge
        summary[employeeNo].TOTAL_EXCESS_CHARGE += item.EXCESS_CHARGE;
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
            };
        }

        // Accumulate the total amount deducted
        summary[employeeNo].TOTAL_AMOUNT_DEDUCTED += Number(
            deduction.AMOUNT || 0
        );
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

// function populateExcessChargeReportTemplate(data) {
//   const sheetName = 'EXCESS_CHARGES_TEMPLATE'; // Target sheet name
//   const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
//   const sheet = spreadsheet.getSheetByName(sheetName);

//   // Check if the sheet exists
//   if (!sheet) {
//     Logger.log(`Sheet with name "${sheetName}" does not exist.`);
//     return;
//   }

//   // Clear the sheet before writing
//   sheet.clear();

//   // Insert the sheet name and today's date
//   sheet.getRange('A1').setValue("HR COPY").setFontSize(14).setFontFamily('Poppins').setFontWeight('bold').setHorizontalAlignment('LEFT');
//   const today = new Date();
//   const formattedDate = Utilities.formatDate(today, Session.getScriptTimeZone(), 'MM/dd/yyyy');
//   sheet.getRange('J1').setValue(`Date: ${formattedDate}`).setFontSize(14).setFontFamily('Poppins').setHorizontalAlignment('LEFT');

//   // Insert report details
//   sheet.getRange('A3').setValue("RFP NO").setFontSize(10).setFontFamily('Poppins').setHorizontalAlignment('LEFT');
//   sheet.getRange('B3').setValue(data[0].RFP_NO).setFontSize(10).setFontFamily('Poppins').setHorizontalAlignment('LEFT');
//   sheet.getRange('A4').setValue("RFP DATE").setFontSize(10).setFontFamily('Poppins').setHorizontalAlignment('LEFT');
//   sheet.getRange('B4').setValue(data[0].RFP_DATE).setFontSize(10).setFontFamily('Poppins').setHorizontalAlignment('LEFT');
//   sheet.getRange('A5').setValue("NETWORK PROVIDER").setFontSize(10).setFontFamily('Poppins').setHorizontalAlignment('LEFT');
//   sheet.getRange('B5').setValue(data[0].NETWORK_PROVIDER).setFontSize(10).setFontFamily('Poppins').setHorizontalAlignment('LEFT');
//   sheet.getRange('A3:B5').setBorder(true, true, true, true, true, true);

//   // Define headers
//   const headers = [
//     'EMPLOYEE NO',
//     'EMPLOYEE NAME',
//     'DEPARTMENT',
//     'COMPANY',
//     'BILL PERIOD FROM',
//     'BILL PERIOD TO',
//     'SIM CARD ID',
//     'MOBILE NO',
//     'ACCOUNT NO',
//     'EXCESS CHARGE',
//   ];

//   // Map data to match headers
//   const rows = data.map((item) => [
//     `'${String(item.EMPLOYEE_NO).padStart(6, '0')}`,
//     item.EMPLOYEE_NAME,
//     item.DEPARTMENT,
//     item.COMPANY,
//     Utilities.formatDate(new Date(item.BILL_PERIOD_FROM), Session.getScriptTimeZone(), 'MM/dd/yyyy'),
//     Utilities.formatDate(new Date(item.BILL_PERIOD_TO), Session.getScriptTimeZone(), 'MM/dd/yyyy'),
//     JSON.parse(item.SIM_INFO).PLAN_ID,
//     item.MOBILE_NO,
//     item.ACCOUNT_NO,
//     item.EXCESS_CHARGE,
//   ]);

//   // Insert headers and data
//   const startRow = 10;
//   const headerRange = sheet.getRange(startRow, 1, 1, headers.length); // Header row
//   const dataRange = sheet.getRange(startRow + 1, 1, rows.length, headers.length); // Data rows

//   headerRange.setValues([headers]); // Insert headers
//   dataRange.setValues(rows); // Insert data

//   // Format headers
//   headerRange.setFontWeight('bold').setFontSize(12).setHorizontalAlignment('CENTER').setFontFamily('Poppins');
//   headerRange.setBorder(true, true, true, true, true, true); // Add borders to header row

//   // Format data
//   dataRange.setFontSize(11).setFontFamily('Poppins').setHorizontalAlignment('CENTER');
//   dataRange.setBorder(true, true, true, true, true, true); // Add borders to data rows

//   const excessChargeColumn = sheet.getRange(startRow + 1, 10, rows.length, 1); // EXCESS CHARGE column
//   excessChargeColumn.setNumberFormat('#,##0.00').setHorizontalAlignment('RIGHT');

//   // Add grand total row
//   const totalRowIndex = startRow + rows.length + 1;
//   sheet.getRange(totalRowIndex, 1).setValue('GRAND TOTAL').setFontWeight('bold').setFontSize(12).setHorizontalAlignment('CENTER');
//   sheet.getRange(totalRowIndex, 10)
//     .setFormula(`=SUM(J${startRow + 1}:J${startRow + rows.length})`)
//     .setNumberFormat('#,##0.00')
//     .setFontWeight('bold')
//     .setFontSize(12)
//     .setHorizontalAlignment('RIGHT');

//   // Apply borders to the entire row (from column A to column J)
//   sheet.getRange(totalRowIndex, 1, 1, 10) // Entire row range (from column A to J)
//     .setBorder(true, true, true, true, true, true); // Adds borders on all sides

//   // **Signatories Section**
//   const pageHeight = 50; // Rows per page
//   const totalUsedRows = totalRowIndex + 3; // Rows used by headers, data, and totals
//   const totalPages = Math.ceil(totalUsedRows / pageHeight); // Calculate total pages
//   const lastPageStartRow = (totalPages - 1) * pageHeight + 1; // Start row of the last page

//   // Add blank rows to push signatories to the last page
//   const blankRowsNeeded = Math.max(0, lastPageStartRow - totalUsedRows);
//   if (blankRowsNeeded > 0) {
//     sheet.insertRowsAfter(totalRowIndex, blankRowsNeeded);
//   }

//   // Add signatories at the bottom of the last page
//   const signatoryStartRow = totalRowIndex + blankRowsNeeded + 5;
//   sheet.getRange(signatoryStartRow, 1).setValue('Prepared by:').setFontWeight('bold').setFontSize(12);
//   sheet.getRange(signatoryStartRow, 2).setValue('Checked By:').setFontWeight('bold').setFontSize(12);
//   sheet.getRange(signatoryStartRow, 3).setValue('Received By:').setFontWeight('bold').setFontSize(12);

//   sheet.getRange(signatoryStartRow + 2, 1).setValue('Cottish Star Papa').setFontWeight('bold').setFontSize(12);
//   sheet.getRange(signatoryStartRow + 2, 2).setValue('Rienalyn Villanueva').setFontWeight('bold').setFontSize(12);
//   sheet.getRange(signatoryStartRow + 2, 3).setValue('Nida Fajiculay').setFontWeight('bold').setFontSize(12);

//   sheet.getRange(signatoryStartRow + 3, 1).setValue('Admin and Benefits Specialist').setFontSize(12);
//   sheet.getRange(signatoryStartRow + 3, 2).setValue('Human Resources Supervisor').setFontSize(12);
//   sheet.getRange(signatoryStartRow + 3, 3).setValue('Payroll Specialist').setFontSize(12);

//   Logger.log(`Report successfully populated in sheet: "${sheetName}"`);
// }

function getExcessChargeHrPdfUrl(staticValue, data) {
    const sheetName = "EXCESS_CHARGE_HR";
    const titleText = "HR COPY";

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
