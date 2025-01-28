// function addBilling(formData) {
//     formData.billID = generateID(BILLING, "BILL ID");
//     // formData.groupName = `${formData.company} - ${formData.networkProvider}`;
//     addRecordToSheet(formData, BILLING, [], [], KEY_ORDER.BILLING, "");
// }

function checkValueIfValid(value) {
    const sheet =
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName("HELPER");
    const range = sheet.getRange("A:A"); // Adjust the range as necessary
    const formula = `=IF(COUNTIF(BILLING!D:D, "${value}") > 0, TRUE, FALSE)`;

    const cell = sheet.getRange("A1"); // Use a temporary cell to store the formula
    cell.setFormula(formula);

    const result = cell.getValue();
    cell.clear(); // Clear the temporary cell

    console.log(result);

    return result;
}

// Function to append billing sheet data to Google Sheet
// function appendBillingSheetData(billingSheetData) {
//     var sheet = accessSheet('BILLING');
//     console.log(billingSheetData)
//     // Get the last row number to start appending new data
//     var lastRow = sheet.getLastRow();
//     // Append data
//     sheet.getRange(lastRow + 1, 1, billingSheetData.length, billingSheetData[0].length).setValues(billingSheetData);
// }

// async function editBillingRecord(formData) {
//     console.log(formData);

//     // Find and update record in the BILLING sheet
//     var row = findRowById(BILLING, formData.billId);
//     console.log(row);
//     if (row != -1) {
//         upsertRecord(BILLING, row, formData, formToSheetMap.BILLING);
//     } else {
//         Logger.log('Record not found in BILLING');
//     }

//     var excessCharge = parseFloat(formData.editExcessCharges); // Ensure excessCharge is a number

//     if (excessCharge > 0) {
//         // Get the EXCESS_CHARGES sheet
//         var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EXCESS_CHARGES);
//         var lastRow = sheet.getLastRow();

//         var excessRow = -1;

//         if (lastRow > 1) {  // Check if there are any data rows (excluding header)
//             // Get all the data from the EXCESS_CHARGES sheet (excluding headers)
//             var data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

//             // Search for the BILL_ID in the data (assuming BILL_ID is in column B, index 1)
//             for (var i = 0; i < data.length; i++) {
//                 if (data[i][1].toString().trim() === formData.billId.toString().trim()) {
//                     excessRow = i + 2; // Account for the header row
//                     break;
//                 }
//             }
//         }

//         var previousExcessCharge = await getPreviousExcessCharge(formData.employeeGroupId);
//         previousExcessCharge = parseFloat(previousExcessCharge) || 0; // Ensure it's a number
//         var remainingExcessCharge = excessCharge; // Adjust as necessary

//         var recordData = {
//             billId: formData.billId,
//             groupId: `${formData.employeeGroupId.toString().padStart(6, '0')}`,
//             employeeNo: `${formData.employeeNoId.toString().padStart(6, '0')}`,
//             simCardId: formData.simCardId,
//             excessChargeDate: formData.editBillPeriodFrom,
//             excessCharge: excessCharge,
//             remainingExcessCharge: remainingExcessCharge,
//         };

//         // If BILL_ID exists, update the record; otherwise, add a new one
//         if (excessRow != -1) {
//             console.log('Editing record in EXCESS_CHARGES');
//             upsertRecord(EXCESS_CHARGES, excessRow, recordData, formToSheetMap.EXCESS_CHARGES);
//         } else {
//             console.log('Adding new record to EXCESS_CHARGES');
//             var lastId = await getLastId(EXCESS_CHARGES, "A");
//             recordData.id = lastId;
//             await addRecordToSheet(recordData, EXCESS_CHARGES, [], KEY_ORDER.EXCESS_CHARGES);
//         }
//     }
// }

// function deleteBillingRecord(billId) {
//     // Delete the record from the BILLING sheet by BILL_ID
//     var result = deleteRecordByColumnValue(billId, 'BILL_ID', BILLING);

//     // Now check if the BILL_ID exists in the EXCESS_CHARGES sheet
//     var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EXCESS_CHARGES);
//     var lastRow = sheet.getLastRow();

//     if (lastRow > 1) {  // Ensure there's at least one data row
//         // Get all data from the EXCESS_CHARGES sheet (excluding headers)
//         var data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

//         // Search for the BILL_ID in the data (assuming BILL_ID is in column B, index 1)
//         for (var i = 0; i < data.length; i++) {
//             if (data[i][1].toString().trim() === billId.toString().trim()) {
//                 var rowToDelete = i + 2; // Account for the header row
//                 sheet.deleteRow(rowToDelete); // Delete the row
//                 Logger.log('Deleted record with BILL_ID from EXCESS_CHARGES');
//                 break;
//             }
//         }
//     } else {
//         Logger.log('No records found in EXCESS_CHARGES to delete.');
//     }

//     return result;
// }

// function deleteBillingRecord(billId) {
//     var result = deleteRecordByColumnValue(billId, 'BILL_ID', BILLING);
//     return result;
// }

async function getPreviousExcessCharge(employeeGroupId) {
    var excessChargesSheet =
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EXCESS_CHARGES);
    var data = excessChargesSheet.getDataRange().getValues();
    var latestRemainingExcessCharge = 0;
    var latestRecordIndex = -1;

    for (var i = 1; i < data.length; i++) {
        // start from 1 to skip header row
        if (data[i][2] == employeeGroupId) {
            // assuming GROUP_ID is in the 5th column
            // Check if this record is the latest one for this employeeGroupId
            if (
                latestRecordIndex == -1 ||
                data[i][0] > data[latestRecordIndex][0]
            ) {
                // assuming EC_ID is in the 1st column
                latestRecordIndex = i;
            }
        }
    }

    if (latestRecordIndex != -1) {
        latestRemainingExcessCharge =
            parseFloat(data[latestRecordIndex][6]) || 0; // assuming REMAINING_EXCESS_CHARGE is in the 7th column
    }

    console.log(latestRemainingExcessCharge);

    return latestRemainingExcessCharge;
}

function getBillingTableData(headers = []) {
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
        const simRequestSheet = new Utils.Sheet("SIM_REQUEST_AND_ISSUANCE", {
            row: { start: 1 },
        });
        const simInventorySheet = new Utils.Sheet("SIM_INVENTORY", {
            row: { start: 1 },
        });
        const employeeDetailsSheet = new Utils.Sheet("EMPLOYEE_DETAILS", {
            row: { start: 1 },
        });
        const rfpGroupSheet = new Utils.Sheet("RFP_GROUP", {
            row: { start: 1 },
        });
        const simPlanSheet = new Utils.Sheet("SIM_PLANS", {
            row: { start: 1 },
        });

        const billingDataSet = billingSheet.toObject();
        const simRequestDataset = simRequestSheet.toObject();
        const employeeDataset = employeeDetailsSheet.toObject();
        const simInventoryDataset = simInventorySheet.toObject();
        const rfpGroupDataset = rfpGroupSheet.toObject();
        const simPlanDataset = simPlanSheet.toObject();

        // Updated SQL query to use LEFT JOIN for SIM_REQUEST_AND_ISSUANCE
        const query = `
      SELECT 
        b.BILL_ID,
        b.RFP_NO,
        b.ISSUANCE_NO,
        b.SIM_CARD_ID,
        b.BILL_PERIOD_FROM,
        b.BILL_PERIOD_TO,
        b.MONTHLY_RECURRING_FEE,
        b.EXCESS_CHARGES,
        b.OTHER_CHARGES,
        b.PREVIOUS_BILL_AMOUNT,
        b.PREVIOUS_BILL_PAYMENT,
        b.CURRENT_CHARGE_AMOUNT,
        b.ADJ_AMOUNT,
        b.AMOUNT_DUE,
        b.RFP_AMOUNT,
        b.WITHHOLDING_TAX,
        b.AMOUNT_AFTER_TAX,
        b.CHARGE_TO_BDC,
        b.WITH_SOA,
        b.SIM_INFO,
        b.EMPLOYEE_INFO
      FROM ? AS b
    `;

        // Execute the query by passing datasets in the correct order without placeholders for tables
        const execution = Utils.sql(query, [
            billingDataSet,
            simInventoryDataset,
            simRequestDataset,
            rfpGroupDataset,
            employeeDataset,
            simPlanDataset,
        ]);

        // Check if dataset is empty, return placeholder if needed
        if (!execution || execution.length === 0) {
            return JSON.stringify([
                {
                    BILL_ID: "",
                    RFP_NO: "",
                    ISSUANCE_NO: "",
                    RFP_GROUP_NAME: "",
                    SIM_CARD_ID: "",
                    MOBILE_NO: "",
                    ACCOUNT_NO: "",
                    EMPLOYEE_NAME: "",
                    BILL_PERIOD_FROM: "",
                    BILL_PERIOD_TO: "",
                    MONTHLY_RECURRING_FEE: "",
                    EXCESS_CHARGES: "",
                    OTHER_CHARGES: "",
                    PREVIOUS_BILL_AMOUNT: "",
                    PREVIOUS_BILL_PAYMENT: "",
                    CURRENT_CHARGE_AMOUNT: "",
                    ADJ_AMOUNT: "",
                    AMOUNT_DUE: "",
                    RFP_AMOUNT: "",
                    WITHHOLDING_TAX: "",
                    AMOUNT_AFTER_TAX: "",
                    CHARGE_TO_BDC: "",
                    WITH_SOA: "",
                    NETWORK_PROVIDER: "",
                    CATEGORY: "",
                    PLAN_DETAILS: "",
                },
            ]);
        }

        console.log(execution);
        return JSON.stringify(execution);
    } catch (error) {
        return Utils.ErrorHandler(error, {
            arguments,
            value: [],
        });
    }
}

// Testing this function
// function getBillingWithIssuance(headers = []) {
//   try {
//     const billingSheet = new Utils.Sheet("BILLING", { row: { start: 1 } });
//     const simRequestSheet = new Utils.Sheet("SIM_REQUEST_AND_ISSUANCE", { row: { start: 1 } });

//     const billingDataSet = billingSheet.toObject();
//     const simRequestDataset = simRequestSheet.toObject();

//     // SQL query to select only records with a valid ISSUANCE_NO from Billing and fetch corresponding data from SIM_REQUEST_AND_ISSUANCE
//     const query = `
//       SELECT
//         b.BILL_ID,
//         b.RFP_NO,
//         b.ISSUANCE_NO,
//         sr.SIM_CARD_ID,
//         sr.EMPLOYEE_INFO,
//         sr.SIM_INFO

//       FROM ? AS b
//       LEFT JOIN ? AS sr ON b.ISSUANCE_NO = sr.ISSUANCE_NO
//       WHERE b.ISSUANCE_NO IS NOT NULL
//     `;

//     // Execute the query
//     const execution = Utils.sql(query, [billingDataSet, simRequestDataset]);

//     // If no results, return a placeholder
//     if (!execution || execution.length === 0) {
//       return JSON.stringify([{
//         BILL_ID: "",
//         RFP_NO: "",
//         ISSUANCE_NO: "",
//         SIM_CARD_ID: "",
//         EMPLOYEE_NAME: "",
//         MOBILE_NO: "",
//         ACCOUNT_NO: "",
//         PLAN_ID: ""
//       }]);
//     }

//     console.log(execution);
//     return JSON.stringify(execution);
//   } catch (error) {
//     return Utils.ErrorHandler(error, {
//       arguments,
//       value: [],
//     });
//   }
// }

function getBillingWithIssuance(headers = []) {
    try {
        const billingSheet = new Utils.Sheet("BILLING", { row: { start: 1 } });
        const simRequestSheet = new Utils.Sheet("SIM_REQUEST_AND_ISSUANCE", {
            row: { start: 1 },
        });

        const billingDataSet = billingSheet.toObject();
        const simRequestDataset = simRequestSheet.toObject();

        // SQL query to fetch Employee Info and Sim Info based on ISSUANCE_NO
        const query = `
      SELECT 
        b.BILL_ID,
        b.SIM_CARD_ID,
        b.ISSUANCE_NO,
        b.RFP_NO,
        b.BILL_PERIOD_FROM,
        b.BILL_PERIOD_TO,
        b.MONTHLY_RECURRING_FEE,
        b.EXCESS_CHARGES,
        b.OTHER_CHARGES,
        b.PREVIOUS_BILL_AMOUNT,
        b.PREVIOUS_BILL_PAYMENT,
        b.CURRENT_CHARGE_AMOUNT,
        b.ADJ_AMOUNT,
        b.AMOUNT_DUE,
        b.RFP_AMOUNT,
        b.WITHHOLDING_TAX,
        b.AMOUNT_AFTER_TAX,
        b.CHARGE_TO_BDC,
        b.WITH_SOA,
        -- Fetch EMPLOYEE_INFO from SIM_REQUEST_AND_ISSUANCE if ISSUANCE_NO exists
        CASE
          WHEN b.ISSUANCE_NO IS NOT NULL AND b.ISSUANCE_NO <> '' THEN sr.EMPLOYEE_INFO
          ELSE NULL
        END AS EMPLOYEE_INFO,
        -- Fetch SIM_INFO from SIM_REQUEST_AND_ISSUANCE if ISSUANCE_NO exists
        CASE
          WHEN b.ISSUANCE_NO IS NOT NULL AND b.ISSUANCE_NO <> '' THEN sr.SIM_INFO
          ELSE NULL
        END AS SIM_INFO

      FROM ? AS b
      -- Join with SIM_REQUEST_AND_ISSUANCE based on ISSUANCE_NO
      LEFT JOIN ? AS sr ON b.ISSUANCE_NO = sr.ISSUANCE_NO
    `;

        // Execute the query
        const execution = Utils.sql(query, [billingDataSet, simRequestDataset]);

        // If no results, return a placeholder
        if (!execution || execution.length === 0) {
            return JSON.stringify([
                {
                    BILL_ID: "",
                    SIM_CARD_ID: "",
                    ISSUANCE_NO: "",
                    RFP_NO: "",
                    SIM_INFO: null,
                    EMPLOYEE_INFO: null,
                },
            ]);
        }

        console.log(execution);
        return JSON.stringify(execution);
    } catch (error) {
        return Utils.ErrorHandler(error, {
            arguments,
            value: [],
        });
    }
}

function getRfpGroupData() {
    try {
        const simInventorySheet = new Utils.Sheet("SIM_INVENTORY", {
            row: { start: 1 },
        });
        const rfpGroupSheet = new Utils.Sheet("RFP_GROUP", {
            row: { start: 1 },
        });

        const simInventoryDataset = simInventorySheet.toObject();
        const rfpGroupDataset = rfpGroupSheet.toObject();

        // console.log("RFP Group Data:", rfpGroupDataset);
        // console.log("SIM Inventory Data:", simInventoryDataset);

        // SQL query to join the SIM Inventory and RFP Group by RFP_GROUP_ID and include group details
        const query = `
      SELECT rfp.RFP_GROUP_ID AS RFP_GROUP_ID, rfp.RFP_GROUP_NAME, rfp.RFP_COMPANY, rfp.NETWORK_PROVIDER, rfp.PAYABLE_TO,
             ARRAY({SIM_CARD_ID: si.SIM_CARD_ID, MOBILE_NO: si.MOBILE_NO, ACCOUNT_NO: si.ACCOUNT_NO}) AS SIM_CARDS
      FROM ? AS si
      JOIN ? AS rfp ON si.RFP_GROUP_ID = rfp.RFP_GROUP_ID
      GROUP BY rfp.RFP_GROUP_ID, rfp.RFP_GROUP_NAME, rfp.RFP_COMPANY, rfp.NETWORK_PROVIDER, rfp.PAYABLE_TO
    `;

        // Execute the query
        const execution = Utils.sql(query, [
            simInventoryDataset,
            rfpGroupDataset,
        ]);

        // Transform the result into the required format
        const dataset = {};
        execution.forEach((row) => {
            dataset[row.RFP_GROUP_ID] = {
                RFP_GROUP_NAME: row.RFP_GROUP_NAME,
                RFP_COMPANY: row.RFP_COMPANY,
                NETWORK_PROVIDER: row.NETWORK_PROVIDER,
                PAYABLE_TO: row.PAYABLE_TO,
                SIM_CARDS: row.SIM_CARDS,
            };
        });

        console.log(
            "Transformed Dataset with RFP Group Details:",
            JSON.stringify(dataset)
        );

        // Return the final dataset as JSON
        return JSON.stringify(dataset);
    } catch (error) {
        return Utils.ErrorHandler(error, {
            arguments,
            value: [],
        });
    }
}

// Sample query
function sampleQuery(headers = []) {
    try {
        const billingSheet = new Utils.Sheet("BILLING", { row: { start: 1 } });
        const simInventorySheet = new Utils.Sheet("SIM_INVENTORY", {
            row: { start: 1 },
        });
        const simRequestSheet = new Utils.Sheet("SIM_REQUEST_AND_ISSUANCE", {
            row: { start: 1 },
        });

        const billingDataset = billingSheet.toObject();
        const simInventoryDataset = simInventorySheet.toObject();
        const simRequestDataset = simRequestSheet.toObject();

        // SQL query to fetch SIM and Employee Info based on the issuance status
        const query = `
      SELECT DISTINCT
        b.BILL_ID,
        b.SIM_CARD_ID,
        b.ISSUANCE_NO,
        b.RFP_NO,
        b.BILL_PERIOD_FROM,
        b.BILL_PERIOD_TO,
        b.MONTHLY_RECURRING_FEE,
        b.EXCESS_CHARGES,
        b.OTHER_CHARGES,
        b.PREVIOUS_BILL_AMOUNT,
        b.PREVIOUS_BILL_PAYMENT,
        b.CURRENT_CHARGE_AMOUNT,
        b.ADJ_AMOUNT,
        b.AMOUNT_DUE,
        b.RFP_AMOUNT,
        b.WITHHOLDING_TAX,
        b.AMOUNT_AFTER_TAX,
        b.CHARGE_TO_BDC,
        b.WITH_SOA
        -- Fetch Employee Info if SIM is currently issued, otherwise set to NULL
        CASE
          WHEN sr.REQUEST_STATUS = 'Issued' THEN sr.EMPLOYEE_INFO
          ELSE NULL
        END AS EMPLOYEE_INFO,
        -- Fetch SIM Info from Sim Request and Issuance if currently issued, otherwise from Sim Inventory
        CASE
          WHEN sr.REQUEST_STATUS = 'Issued' THEN sr.SIM_INFO
          ELSE JSON_OBJECT(
            'SIM_CARD_ID', si.SIM_CARD_ID,
            'MOBILE_NO', si.MOBILE_NO,
            'PLAN_ID', si.PLAN_ID,
            'ACCOUNT_NO', si.ACCOUNT_NO
          )
        END AS SIM_INFO
      FROM ? AS b
      LEFT JOIN ? AS si ON b.SIM_CARD_ID = si.SIM_CARD_ID
      LEFT JOIN ? AS sr ON b.SIM_CARD_ID = sr.SIM_CARD_ID
      WHERE sr.REQUEST_STATUS IS NULL OR sr.REQUEST_STATUS = 'Issued'
    `;

        // Execute the query
        const execution = Utils.sql(query, [
            billingDataset,
            simInventoryDataset,
            simRequestDataset,
        ]);

        // If no results, return a placeholder
        if (!execution || execution.length === 0) {
            return JSON.stringify([
                {
                    BILL_ID: "",
                    SIM_CARD_ID: "",
                    ISSUANCE_NO: "",
                    RFP_NO: "",
                    MOBILE_NO: "",
                    PLAN_ID: "",
                    ACCOUNT_NO: "",
                    EMPLOYEE_INFO: null,
                    SIM_INFO: null,
                },
            ]);
        }

        console.log(execution);
        return JSON.stringify(execution);
    } catch (error) {
        return Utils.ErrorHandler(error, {
            arguments,
            value: [],
        });
    }
}

function getRfpSummaryNumbers() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("BILLING");

    if (!sheet) {
        throw new Error(`Sheet "BILLING" does not exist.`);
    }

    const range = sheet.getRange(`D1:D`); // Adjusted to start from row 2
    const values = range.getValues().flat().filter(String); // Flatten values and remove empty strings

    console.log("Fetched Values:", values); // Debug: log the values fetched from the sheet

    const uniqueIds = [...new Set(values)]; // Get unique IDs

    console.log("Unique IDs:", uniqueIds); // Debug: log unique IDs

    return uniqueIds;
}

// functions with audit trail
// Function to append billing sheet data to Google Sheet with audit trail logging
// function appendBillingSheetData(billingSheetData) {
//     var sheet = accessSheet('BILLING');
//     console.log(billingSheetData);

//     // Get the last row number to start appending new data
//     var lastRow = sheet.getLastRow();

//     // Append data
//     sheet.getRange(lastRow + 1, 1, billingSheetData.length, billingSheetData[0].length).setValues(billingSheetData);

//     // Log each appended row in the audit trail
//     billingSheetData.forEach((rowData) => {
//         // Map the row data to field names (assuming headers are in the first row of the sheet)
//         const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
//         const newValue = mapValuesToObject(headers, rowData);

//         // Log the action as "ADD" to the audit trail
//         logAuditTrail(
//             "ADD",
//             "BILLING",
//             rowData[0], // Assuming the first column is the unique ID
//             {},         // No old values for ADD actions
//             newValue,   // New values being added
//             [],         // No changed fields for ADD actions
//             "Billing record added"
//         );
//     });
// }

// BATCH_UPDATE
function appendBillingSheetData(billingSheetData) {
    const sheet = accessSheet("BILLING");
    console.log(billingSheetData);

    // Get the last row number to start appending new data
    const lastRow = sheet.getLastRow();

    // Append data
    sheet
        .getRange(
            lastRow + 1,
            1,
            billingSheetData.length,
            billingSheetData[0].length
        )
        .setValues(billingSheetData);

    // Create a batch log entry
    const headers = sheet
        .getRange(1, 1, 1, sheet.getLastColumn())
        .getValues()[0];
    const batchLog = billingSheetData.map((row) =>
        mapValuesToObject(headers, row)
    );

    // Log the batch addition to the audit trail
    logAuditTrail(
        "BATCH_ADD",
        "BILLING",
        null, // No single record ID for a batch
        {}, // No old values for batch add
        batchLog, // New values as an array of records
        [], // No changed fields
        `Added ${billingSheetData.length} billing records as a batch`
    );
}

async function editBillingRecord(formData) {
    console.log(formData);

    // Access the BILLING sheet and find the row for the given bill ID
    const row = findRowById(BILLING, formData.billId);
    if (row === -1) {
        Logger.log("Record not found in BILLING");
        return;
    }

    // Access the BILLING sheet and get the old values before updating
    const sheet = accessSheet(BILLING);
    const oldValues = sheet
        .getRange(row, 1, 1, sheet.getLastColumn())
        .getValues()[0];

    oldValues[4] = formatDate(oldValues[4]);
    oldValues[5] = formatDate(oldValues[5]);

    // Get the headers to map old and new values
    const headers = sheet
        .getRange(1, 1, 1, sheet.getLastColumn())
        .getValues()[0];

    // Format old values into an object for logging
    const oldRecord = mapValuesToObject(headers, oldValues);

    // Update the BILLING record
    upsertRecord(BILLING, row, formData, formToSheetMap.BILLING);

    // Get the new values from the updated record
    const newValues = sheet
        .getRange(row, 1, 1, sheet.getLastColumn())
        .getValues()[0];
    newValues[4] = formatDate(newValues[4]);
    newValues[5] = formatDate(newValues[5]);
    const newRecord = mapValuesToObject(headers, newValues);

    // Identify changed fields
    const changedFields = getChangedFields(oldRecord, newRecord);

    // Log the edit action in the audit trail
    logAuditTrail(
        "EDIT",
        "BILLING",
        formData.billId,
        oldRecord, // Old values
        newRecord, // New values
        changedFields, // Changed fields
        "Updated billing record"
    );

    console.log("Audit trail logged for billing record update.");

    // Process EXCESS_CHARGES if applicable
    const excessCharge = parseFloat(formData.editExcessCharges); // Ensure excessCharge is a number
    if (excessCharge > 0) {
        const sheetExcess =
            SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
                EXCESS_CHARGES
            );
        const lastRowExcess = sheetExcess.getLastRow();

        let excessRow = -1;

        if (lastRowExcess > 1) {
            // Check if there are any data rows (excluding header)
            const data = sheetExcess
                .getRange(2, 1, lastRowExcess - 1, sheetExcess.getLastColumn())
                .getValues();

            // Search for the BILL_ID in the data (assuming BILL_ID is in column B, index 1)
            for (let i = 0; i < data.length; i++) {
                if (
                    data[i][1].toString().trim() ===
                    formData.billId.toString().trim()
                ) {
                    excessRow = i + 2; // Account for the header row
                    break;
                }
            }
        }

        const previousExcessCharge =
            parseFloat(
                await getPreviousExcessCharge(formData.employeeGroupId)
            ) || 0;
        const remainingExcessCharge = excessCharge;

        const recordData = {
            billId: formData.billId,
            groupId: `${formData.employeeGroupId.toString().padStart(6, "0")}`,
            employeeNo: `${formData.employeeNoId.toString().padStart(6, "0")}`,
            simCardId: formData.simCardId,
            excessChargeDate: formData.editBillPeriodFrom,
            excessCharge: excessCharge,
            remainingExcessCharge: remainingExcessCharge,
        };

        // If BILL_ID exists, update the record; otherwise, add a new one
        if (excessRow !== -1) {
            console.log("Editing record in EXCESS_CHARGES");
            upsertRecord(
                EXCESS_CHARGES,
                excessRow,
                recordData,
                formToSheetMap.EXCESS_CHARGES
            );
        } else {
            console.log("Adding new record to EXCESS_CHARGES");
            const lastId = await getLastId(EXCESS_CHARGES, "A");
            recordData.id = lastId;
            await addRecordToSheet(
                recordData,
                EXCESS_CHARGES,
                [],
                KEY_ORDER.EXCESS_CHARGES
            );
        }
    }
}

function deleteBillingRecord(billId) {
    // Access the BILLING sheet
    const billingSheet = accessSheet(BILLING);

    // Find the row for the BILL_ID
    const row = findRowById(BILLING, billId);
    if (row === -1) {
        Logger.log(`Record with BILL_ID ${billId} not found in BILLING.`);
        return;
    }

    // Get the old values from the BILLING sheet before deletion
    const oldValues = billingSheet
        .getRange(row, 1, 1, billingSheet.getLastColumn())
        .getValues()[0];

    oldValues[4] = formatDate(oldValues[4]);
    oldValues[5] = formatDate(oldValues[5]);

    // Get the headers to map old values to field names
    const headers = billingSheet
        .getRange(1, 1, 1, billingSheet.getLastColumn())
        .getValues()[0];

    // Prepare the old values object
    const oldRecord = mapValuesToObject(headers, oldValues);

    // Delete the record from the BILLING sheet
    const result = deleteRecordByColumnValue(billId, "BILL_ID", BILLING);

    // Log the DELETE action to the audit trail
    logAuditTrail(
        "DELETE",
        "BILLING",
        billId,
        oldRecord, // Old values before deletion
        {}, // No new values for DELETE action
        [], // No changed fields for DELETE action
        "Billing record deleted"
    );

    // Access the EXCESS_CHARGES sheet to check for related records
    const excessSheet =
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EXCESS_CHARGES);
    const lastRow = excessSheet.getLastRow();

    if (lastRow > 1) {
        // Ensure there's at least one data row
        // Get all data from the EXCESS_CHARGES sheet (excluding headers)
        const data = excessSheet
            .getRange(2, 1, lastRow - 1, excessSheet.getLastColumn())
            .getValues();

        // Search for the BILL_ID in the data (assuming BILL_ID is in column B, index 1)
        for (let i = 0; i < data.length; i++) {
            if (data[i][1].toString().trim() === billId.toString().trim()) {
                const rowToDelete = i + 2; // Account for the header row

                // Get old values before deletion
                const excessHeaders = excessSheet
                    .getRange(1, 1, 1, excessSheet.getLastColumn())
                    .getValues()[0];
                const oldExcessValues = excessSheet
                    .getRange(rowToDelete, 1, 1, excessSheet.getLastColumn())
                    .getValues()[0];
                const oldExcessRecord = mapValuesToObject(
                    excessHeaders,
                    oldExcessValues
                );

                // Delete the row
                excessSheet.deleteRow(rowToDelete);

                // Log the DELETE action for the EXCESS_CHARGES record
                logAuditTrail(
                    "DELETE",
                    "EXCESS_CHARGES",
                    billId,
                    oldExcessRecord, // Old values before deletion
                    {}, // No new values for DELETE action
                    [], // No changed fields for DELETE action
                    "Excess charges record deleted"
                );

                Logger.log("Deleted record with BILL_ID from EXCESS_CHARGES");
                break;
            }
        }
    } else {
        Logger.log("No records found in EXCESS_CHARGES to delete.");
    }

    return result;
}
