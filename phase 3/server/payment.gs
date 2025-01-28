// function appendPaymentRecords(data) {
//     var sheet = accessSheet(PAYMENT);
//     // Get the last row number to start appending new data
//     var lastRow = sheet.getLastRow();
//     // Append data
//     sheet.getRange(lastRow + 1, 1, data.length, data[0].length).setValues(data);
// }

function batchPaymentPostingUpdate(generalData, tableData) {
    // const generalData = {
    //   "rfpNo": "HR-SIM-2024-002",
    //   "paymentPostedDate": "12/06/2024",
    //   "paymentBreakdownReceiptDate": "12/06/2024",
    //   "paymentReferenceDate": "12/06/2024",
    //   "paymentReferenceNo": "2345"
    // };

    // const dataset = [
    //     {
    //         "paymentId": "5",
    //         "billId": "5",
    //         "simCardId": "307",
    //         "mobileNo": "9209158078",
    //         "rfpAmount": "300.00",
    //         "status": "Posted"
    //     },
    //     {
    //         "paymentId": "6",
    //         "billId": "6",
    //         "simCardId": "310",
    //         "mobileNo": "9088686903",
    //         "rfpAmount": "300.00",
    //         "status": "Unposted"
    //     }
    // ];

    console.log("General Data:", generalData);
    console.log("Table Data:", tableData);

    try {
        const sheet = new Utils.Sheet(PAYMENT, {
            row: {
                start: 1,
            },
        });

        const query = sheet.experimentalQuery();
        const queried = query.findMany({ RFP_NO: generalData.rfpNo });
        const queriedData = queried.data;

        console.log(queried);

        // index: [3, 4, 5]; // +2

        const startRow = queried.index[0] + 2;
        const startCol = 1;
        const numRows = queried.length;
        const numCols = queriedData[0].length;

        const result = generateUpdatedData(queriedData, generalData, tableData);
        console.log(result);

        sheet.getRange(startRow, startCol, numRows, numCols).setValues(result);
    } catch (error) {
        return Utils.ErrorHandler(error, {
            arguments,
        });
    }
}

// function editPaymentRecord(formData) {
//     var row = findRowById(PAYMENT, formData.editPaymentId);
//     const status = formData.editPaymentPostedDate ? "Posted" : "Unposted";
//     formData.status = status;
//     console.log(formData);
//     console.log(row);
//     if (row != -1) {
//         upsertRecord(PAYMENT, row, formData, formToSheetMap.PAYMENT);
//     } else {
//         Logger.log('Record not found');
//     }
// }

function deletePaymentRecord(id) {
    var result = deleteRecordByColumnValue(id, "PAYMENT_ID", PAYMENT);
    return result;
}

function getPaymentData() {
    try {
        // Initialize sheets
        const paymentSheet = new Utils.Sheet("PAYMENT", { row: { start: 1 } });
        const billingSheet = new Utils.Sheet("BILLING", { row: { start: 1 } });
        const simInventorySheet = new Utils.Sheet("SIM_INVENTORY", {
            row: { start: 1 },
        }); // Add SIM Inventory sheet

        // Fetch data from sheets
        const paymentData = paymentSheet.toObject();
        const billingData = billingSheet.toObject();
        const simInventoryData = simInventorySheet.toObject(); // Fetch SIM Inventory data

        // Define an SQL query to join the data from PAYMENT, BILLING, and SIM_INVENTORY sheets
        const query = `
      SELECT 
        p.*, 
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
        si.MOBILE_NO  -- Include mobile number from SIM Inventory
      FROM ? AS p
      LEFT JOIN ? AS b ON p.BILL_ID = b.BILL_ID
      LEFT JOIN ? AS si ON b.SIM_CARD_ID = si.SIM_CARD_ID  -- Join SIM Inventory using SIM_CARD_ID
    `;

        // Execute the query using Utils.sql
        const dataset = Utils.sql(query, [
            paymentData,
            billingData,
            simInventoryData,
        ]);
        console.log(dataset);

        // Check if dataset is empty and return '' for each column if true
        if (!dataset || dataset.length === 0) {
            console.log("asdf");
            return JSON.stringify([
                {
                    BILL_ID: "",
                    SIM_CARD_ID: "",
                    ISSUANCE_NO: "",
                    RFP_NO: "",
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
                    MOBILE_NO: "",
                },
            ]);
        }

        // Format date fields (assuming BILL_PERIOD_FROM and BILL_PERIOD_TO are date fields)
        dataset.forEach((row) => {
            if (row.PAYMENT_REFERENCE_DATE) {
                row.PAYMENT_REFERENCE_DATE = Utilities.formatDate(
                    new Date(row.PAYMENT_REFERENCE_DATE),
                    Session.getScriptTimeZone(),
                    "MM/dd/yyyy"
                );
            }
            if (row.PAYMENT_POSTED_DATE) {
                row.PAYMENT_POSTED_DATE = Utilities.formatDate(
                    new Date(row.PAYMENT_POSTED_DATE),
                    Session.getScriptTimeZone(),
                    "MM/dd/yyyy"
                );
            }
            if (row.PAYMENT_BREAKDOWN_RECEIPT_DATE) {
                row.PAYMENT_BREAKDOWN_RECEIPT_DATE = Utilities.formatDate(
                    new Date(row.PAYMENT_BREAKDOWN_RECEIPT_DATE),
                    Session.getScriptTimeZone(),
                    "MM/dd/yyyy"
                );
            }
        });

        console.log("qwer");

        // Return the joined dataset with Mobile Numbers
        return JSON.stringify(dataset);
    } catch (error) {
        console.error("Error fetching data on server side:", error);
        throw new Error("Failed to fetch data for table rendering");
    }
}

const asdff = "HR-SIM-2024-001";
function getBillingDataByRfpNo(rfpNo) {
    try {
        const spaceRegex = /[\s]/g;

        // Initialize sheets
        const billingSheet = new Utils.Sheet("BILLING", { row: { start: 1 } });
        const simInventorySheet = new Utils.Sheet("SIM_INVENTORY", {
            row: { start: 1 },
        });
        const simRequestAndIssuanceSheet = new Utils.Sheet(
            "SIM_REQUEST_AND_ISSUANCE",
            { row: { start: 1 } }
        );
        const employeeDetailsSheet = new Utils.Sheet("EMPLOYEE_DETAILS", {
            row: { start: 1 },
        });
        const paymentSheet = new Utils.Sheet("PAYMENT", { row: { start: 1 } });

        // Fetch data from sheets
        const billingDataSet = billingSheet.toObject();
        const simInventoryDataSet = simInventorySheet.toObject();
        const simRequestAndIssuanceDataSet =
            simRequestAndIssuanceSheet.toObject();
        const employeeDetailsDataSet = employeeDetailsSheet.toObject();
        const paymentDataSet = paymentSheet.toObject(); // Fetch Payment data

        // Prepare list of RFP_NO from Payment sheet
        const paymentRfpNos = paymentDataSet.map((payment) => payment.RFP_NO);

        console.log(rfpNo);
        console.log(paymentRfpNos);

        // Check if the selected RFP_NO already exists in the Payment sheet
        if (paymentRfpNos.includes(rfpNo)) {
            // Return an empty result if the RFP_NO already exists
            return JSON.stringify([]);
        }

        // SQL query to get the data
        const query = `
      SELECT DISTINCT
        b.BILL_ID,
        b.SIM_CARD_ID,
        b.ISSUANCE_NO,
        b.RFP_NO,
        b.BILL_PERIOD_FROM,
        b.BILL_PERIOD_TO,
        b.WITH_SOA,
        b.MONTHLY_RECURRING_FEE,
        b.PREVIOUS_BILL_AMOUNT,
        b.PREVIOUS_BILL_PAYMENT,
        b.EXCESS_CHARGES,
        b.OTHER_CHARGES,
        b.CURRENT_CHARGE_AMOUNT,
        b.ADJ_AMOUNT,
        b.AMOUNT_DUE,
        b.RFP_AMOUNT,
        b.WITHHOLDING_TAX,
        b.AMOUNT_AFTER_TAX,
        b.CHARGE_TO_BDC,
        si.MOBILE_NO,
        COALESCE(e.EMPLOYEE_NAME, '') AS EMPLOYEE_NAME
      FROM ? AS b
      LEFT JOIN ? AS si ON b.SIM_CARD_ID = si.SIM_CARD_ID
      LEFT JOIN ? AS sra ON b.SIM_CARD_ID = si.SIM_CARD_ID AND b.ISSUANCE_NO = sra.ISSUANCE_NO
      LEFT JOIN ? AS e ON sra.EMPLOYEE_ID = e.EMPLOYEE_ID
      WHERE b.RFP_NO = ?
    `;

        // Execute the query
        const result = Utils.sql(query, [
            billingDataSet,
            simInventoryDataSet,
            simRequestAndIssuanceDataSet,
            employeeDetailsDataSet,
            rfpNo,
        ]);
        console.log("Query run", result);

        // Check if dataset is empty, return placeholder if needed
        if (!result || result.length === 0) {
            return JSON.stringify([]);
        }

        return JSON.stringify(result);
    } catch (error) {
        return Utils.ErrorHandler(error, {
            arguments,
            value: [],
        });
    }
}

// For loading billing data
// function getBillingDataExcludingPayments() {
//   try {
//     // Initialize sheets
//     const billingSheet = new Utils.Sheet("BILLING", { row: { start: 1 } });
//     const simInventorySheet = new Utils.Sheet("SIM_INVENTORY", { row: { start: 1 } });
//     const simRequestAndIssuanceSheet = new Utils.Sheet("SIM_REQUEST_AND_ISSUANCE", { row: { start: 1 } });
//     const employeeDetailsSheet = new Utils.Sheet("EMPLOYEE_DETAILS", { row: { start: 1 } });
//     const paymentSheet = new Utils.Sheet("PAYMENT", { row: { start: 1 } });
//     const rfpSummarySheet = new Utils.Sheet("RFP_SUMMARY", { row: { start: 1 } });

//     // Fetch data from sheets
//     const billingDataSet = billingSheet.toObject();
//     const simInventoryDataSet = simInventorySheet.toObject();
//     const simRequestAndIssuanceDataSet = simRequestAndIssuanceSheet.toObject();
//     const employeeDetailsDataSet = employeeDetailsSheet.toObject();
//     const paymentDataSet = paymentSheet.toObject();
//     const rfpSummaryDataSet = rfpSummarySheet.toObject();

//     // Prepare list of RFP_NO from Payment sheet
//     const paymentRfpNos = paymentDataSet.map(payment => payment.RFP_NO);

//     // SQL query to get the data excluding records where RFP_NO is in the Payment sheet
//     const query = `
//       SELECT DISTINCT
//         b.BILL_ID,
//         b.SIM_CARD_ID,
//         b.ISSUANCE_NO,
//         b.RFP_NO,
//         b.BILL_PERIOD_FROM,
//         b.BILL_PERIOD_TO,
//         b.WITH_SOA,
//         b.MONTHLY_RECURRING_FEE,
//         b.PREVIOUS_BILL_AMOUNT,
//         b.PREVIOUS_BILL_PAYMENT,
//         b.EXCESS_CHARGES,
//         b.OTHER_CHARGES,
//         b.CURRENT_CHARGE_AMOUNT,
//         b.ADJ_AMOUNT,
//         b.AMOUNT_DUE,
//         b.RFP_AMOUNT,
//         b.WITHHOLDING_TAX,
//         b.AMOUNT_AFTER_TAX,
//         b.CHARGE_TO_BDC,
//         si.MOBILE_NO,
//         COALESCE(e.EMPLOYEE_NAME, '') AS EMPLOYEE_NAME
//       FROM ? AS b
//       LEFT JOIN ? AS si ON b.SIM_CARD_ID = si.SIM_CARD_ID
//       LEFT JOIN ? AS sra ON b.SIM_CARD_ID = si.SIM_CARD_ID AND b.ISSUANCE_NO = sra.ISSUANCE_NO
//       LEFT JOIN ? AS e ON sra.EMPLOYEE_ID = e.EMPLOYEE_ID
//       WHERE b.RFP_NO NOT IN (${paymentRfpNos.map(() => '?').join(',')})
//     `;

//     // Execute the query
//     const result = Utils.sql(query, [billingDataSet, simInventoryDataSet, simRequestAndIssuanceDataSet, employeeDetailsDataSet, ...paymentRfpNos]);
//     console.log("Query run", result);

//     // Check if dataset is empty, return placeholder if needed
//     if (!result || result.length === 0) {
//       return JSON.stringify([]);
//     }

//     return JSON.stringify(result);
//   } catch (error) {
//     return Utils.ErrorHandler(error, {
//       arguments,
//       value: [],
//     });
//   }
// }

// Working function
function getBillingDataExcludingPayments() {
    try {
        // Initialize sheets
        const billingSheet = new Utils.Sheet("BILLING", { row: { start: 1 } });
        const simInventorySheet = new Utils.Sheet("SIM_INVENTORY", {
            row: { start: 1 },
        });
        const simRequestAndIssuanceSheet = new Utils.Sheet(
            "SIM_REQUEST_AND_ISSUANCE",
            { row: { start: 1 } }
        );
        const employeeDetailsSheet = new Utils.Sheet("EMPLOYEE_DETAILS", {
            row: { start: 1 },
        });
        const paymentSheet = new Utils.Sheet("PAYMENT", { row: { start: 1 } });
        const rfpSummarySheet = new Utils.Sheet("RFP_SUMMARY", {
            row: { start: 1 },
        });

        // Fetch data from sheets
        const billingDataSet = billingSheet.toObject();
        const simInventoryDataSet = simInventorySheet.toObject();
        const simRequestAndIssuanceDataSet =
            simRequestAndIssuanceSheet.toObject();
        const employeeDetailsDataSet = employeeDetailsSheet.toObject();
        const paymentDataSet = paymentSheet.toObject();
        const rfpSummaryDataSet = rfpSummarySheet.toObject();

        // Prepare list of RFP_NO from Payment sheet
        const paymentRfpNos = paymentDataSet.map((payment) => payment.RFP_NO);

        // Prepare list of RFP_NO from RFP_SUMMARY sheet
        const rfpSummaryNos = rfpSummaryDataSet.map((rfp) => rfp.RFP_NO);

        // Prepare list of RFP_NO from BILLING sheet
        const billingRfpNos = billingDataSet.map((billing) => billing.RFP_NO);

        // // Only include RFP_NO that exist in both BILLING and RFP_SUMMARY
        // const validRfpNos = billingRfpNos.filter(rfpNo => rfpSummaryNos.includes(rfpNo));

        // Only include RFP_NO that exist in RFP_SUMMARY but not in PAYMENT
        const validRfpNos = rfpSummaryNos.filter(
            (rfpNo) =>
                !paymentRfpNos.includes(rfpNo) && billingRfpNos.includes(rfpNo)
        );

        // SQL query to get the data excluding records where RFP_NO is not valid
        const query = `
      SELECT DISTINCT
        b.BILL_ID,
        b.SIM_CARD_ID,
        b.ISSUANCE_NO,
        b.RFP_NO,
        b.BILL_PERIOD_FROM,
        b.BILL_PERIOD_TO,
        b.WITH_SOA,
        b.MONTHLY_RECURRING_FEE,
        b.PREVIOUS_BILL_AMOUNT,
        b.PREVIOUS_BILL_PAYMENT,
        b.EXCESS_CHARGES,
        b.OTHER_CHARGES,
        b.CURRENT_CHARGE_AMOUNT,
        b.ADJ_AMOUNT,
        b.AMOUNT_DUE,
        b.RFP_AMOUNT,
        b.WITHHOLDING_TAX,
        b.AMOUNT_AFTER_TAX,
        b.CHARGE_TO_BDC,
        si.MOBILE_NO,
        COALESCE(e.EMPLOYEE_NAME, '') AS EMPLOYEE_NAME
      FROM ? AS b
      LEFT JOIN ? AS si ON b.SIM_CARD_ID = si.SIM_CARD_ID
      LEFT JOIN ? AS sra ON b.SIM_CARD_ID = si.SIM_CARD_ID AND b.ISSUANCE_NO = sra.ISSUANCE_NO
      LEFT JOIN ? AS e ON sra.EMPLOYEE_ID = e.EMPLOYEE_ID
      WHERE b.RFP_NO IN (${validRfpNos.map(() => "?").join(",")})
    `;

        // Execute the query
        const result = Utils.sql(query, [
            billingDataSet,
            simInventoryDataSet,
            simRequestAndIssuanceDataSet,
            employeeDetailsDataSet,
            ...validRfpNos,
        ]);
        console.log("Query run", result);
        console.log("Valid RFP No:", validRfpNos);

        // Check if dataset is empty, return placeholder if needed
        if (!result || result.length === 0) {
            return JSON.stringify([]);
        }

        return JSON.stringify(result);
    } catch (error) {
        return Utils.ErrorHandler(error, {
            arguments,
            value: [],
        });
    }
}

// Testing function
function getBillingDataWithUnpostedAccounts() {
    try {
        // Initialize sheets
        const billingSheet = new Utils.Sheet("BILLING", { row: { start: 1 } });
        const paymentSheet = new Utils.Sheet("PAYMENT", { row: { start: 1 } });
        const rfpSummarySheet = new Utils.Sheet("RFP_SUMMARY", {
            row: { start: 1 },
        });
        const simInventorySheet = new Utils.Sheet("SIM_INVENTORY", {
            row: { start: 1 },
        });
        const simRequestAndIssuanceSheet = new Utils.Sheet(
            "SIM_REQUEST_AND_ISSUANCE",
            { row: { start: 1 } }
        );
        const employeeDetailsSheet = new Utils.Sheet("EMPLOYEE_DETAILS", {
            row: { start: 1 },
        });

        // Fetch data from sheets
        const billingDataSet = billingSheet.toObject();
        const paymentDataSet = paymentSheet.toObject();
        const rfpSummaryDataSet = rfpSummarySheet.toObject();
        const simInventoryDataSet = simInventorySheet.toObject();
        const simRequestAndIssuanceDataSet =
            simRequestAndIssuanceSheet.toObject();
        const employeeDetailsDataSet = employeeDetailsSheet.toObject();

        // Get RFP_NO with unposted accounts from PAYMENT
        const rfpWithUnpostedAccounts = paymentDataSet
            .filter((payment) => payment.STATUS !== "Posted")
            .map((payment) => payment.RFP_NO);

        // Get all RFP_NO from RFP_SUMMARY
        const allRfpNosInSummary = rfpSummaryDataSet.map((rfp) => rfp.RFP_NO);

        // Get RFP_NO that exist in RFP_SUMMARY but not in PAYMENT (no posting generated)
        const rfpWithoutPayments = allRfpNosInSummary.filter(
            (rfpNo) =>
                !paymentDataSet.some((payment) => payment.RFP_NO === rfpNo)
        );

        // Combine RFP Nos with unposted accounts and those without any payments
        const validRfpNos = [
            ...new Set([...rfpWithUnpostedAccounts, ...rfpWithoutPayments]),
        ];

        // SQL query to retrieve billing data for valid RFP_NO entries with unposted accounts or without payments
        const query = `
      SELECT DISTINCT
        b.BILL_ID,
        b.SIM_CARD_ID,
        b.ISSUANCE_NO,
        b.RFP_NO,
        b.BILL_PERIOD_FROM,
        b.BILL_PERIOD_TO,
        b.WITH_SOA,
        b.MONTHLY_RECURRING_FEE,
        b.PREVIOUS_BILL_AMOUNT,
        b.PREVIOUS_BILL_PAYMENT,
        b.EXCESS_CHARGES,
        b.OTHER_CHARGES,
        b.CURRENT_CHARGE_AMOUNT,
        b.ADJ_AMOUNT,
        b.AMOUNT_DUE,
        b.RFP_AMOUNT,
        b.WITHHOLDING_TAX,
        b.AMOUNT_AFTER_TAX,
        b.CHARGE_TO_BDC,
        si.MOBILE_NO,
        COALESCE(e.EMPLOYEE_NAME, '') AS EMPLOYEE_NAME
      FROM ? AS b
      LEFT JOIN ? AS si ON b.SIM_CARD_ID = si.SIM_CARD_ID
      LEFT JOIN ? AS sra ON b.SIM_CARD_ID = si.SIM_CARD_ID AND b.ISSUANCE_NO = sra.ISSUANCE_NO
      LEFT JOIN ? AS e ON sra.EMPLOYEE_ID = e.EMPLOYEE_ID
      WHERE b.RFP_NO IN (${validRfpNos.map(() => "?").join(",")})
    `;

        // Execute the query
        const result = Utils.sql(query, [
            billingDataSet,
            simInventoryDataSet,
            simRequestAndIssuanceDataSet,
            employeeDetailsDataSet,
            ...validRfpNos,
        ]);

        // Check if dataset is empty, return placeholder if needed
        if (!result || result.length === 0) {
            return JSON.stringify([]);
        }

        console.log(JSON.stringify(result));

        return JSON.stringify(result);
    } catch (error) {
        return Utils.ErrorHandler(error, {
            arguments,
            value: [],
        });
    }
}

function getPaymentDataTestingFn() {
    try {
        // Initialize sheets
        const paymentSheet = new Utils.Sheet("PAYMENT", { row: { start: 1 } });
        const billingSheet = new Utils.Sheet("BILLING", { row: { start: 1 } });
        const simInventorySheet = new Utils.Sheet("SIM_INVENTORY", {
            row: { start: 1 },
        });
        const simRequestAndIssuanceSheet = new Utils.Sheet(
            "SIM_REQUEST_AND_ISSUANCE",
            { row: { start: 1 } }
        ); // Add SIM Request & Issuance sheet
        const employeeDetailsSheet = new Utils.Sheet("EMPLOYEE_DETAILS", {
            row: { start: 1 },
        }); // Add Employee Details sheet

        // Fetch data from sheets
        const paymentData = paymentSheet.toObject();
        const billingData = billingSheet.toObject();
        const simInventoryData = simInventorySheet.toObject();
        const simRequestAndIssuanceData = simRequestAndIssuanceSheet.toObject(); // Fetch SIM Request & Issuance data
        const employeeDetailsData = employeeDetailsSheet.toObject(); // Fetch Employee Details data

        // Define an SQL query to join data from PAYMENT, BILLING, SIM_INVENTORY, SIM_REQUEST_AND_ISSUANCE, and EMPLOYEE_DETAILS sheets
        const query = `
      SELECT 
        p.*, 
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
        si.MOBILE_NO,  -- Include mobile number from SIM Inventory
        COALESCE(e.EMPLOYEE_NAME, '') AS EMPLOYEE_NAME  -- Include employee name
      FROM ? AS p
      LEFT JOIN ? AS b ON p.BILL_ID = b.BILL_ID
      LEFT JOIN ? AS si ON b.SIM_CARD_ID = si.SIM_CARD_ID  -- Join SIM Inventory using SIM_CARD_ID
      LEFT JOIN ( 
        SELECT SIM_CARD_ID, GROUP_ID 
        FROM ? 
        WHERE ISSUANCE_DATE = (SELECT MAX(ISSUANCE_DATE) FROM ? WHERE SIM_CARD_ID = si.SIM_CARD_ID)
      ) AS sra ON b.SIM_CARD_ID = sra.SIM_CARD_ID  -- Get the latest issuance record for SIM_CARD_ID
      LEFT JOIN ? AS e ON sra.GROUP_ID = e.GROUP_ID  -- Join Employee Details to get EMPLOYEE_NAME
    `;

        // Execute the query using Utils.sql
        const dataset = Utils.sql(query, [
            paymentData,
            billingData,
            simInventoryData,
            simRequestAndIssuanceData,
            simRequestAndIssuanceData,
            employeeDetailsData,
        ]);
        console.log(dataset);

        if (!dataset || dataset.length === 0) {
            return JSON.stringify([
                {
                    PAYMENT_ID: "",
                    BILL_ID: "",
                    RFP_NO: "",
                    SIM_CARD_ID: "",
                    MOBILE_NO: "",
                    RFP_AMOUNT: "",
                    STATUS: "",
                    PAYMENT_REFERENCE_NO: "",
                    PAYMENT_REFERENCE_DATE: "",
                    PAYMENT_POSTED_DATE: "",
                    PAYMENT_BREAKDOWN_RECEIPT_DATE: "",
                },
            ]);
        }

        // Return the joined dataset with Mobile Numbers and Employee Names
        return JSON.stringify(dataset);
    } catch (error) {
        console.error("Error fetching data on server side:", error);
        throw new Error("Failed to fetch data for table rendering");
    }
}

function testPaymentGetSheet() {
    const a = getDataFromSheet("BILLING", []);
    console.log(a);
}

function generateUpdatedData(data, formData, tableData) {
    return data.map((row) => {
        const [
            paymentId,
            billId,
            rfpNo,
            createdDate,
            amount,
            postedDate,
            updatedDate,
            status,
        ] = row;

        // Find the matching record in the table data
        const tableRecord = tableData.find(
            (item) => item.paymentId == paymentId
        );

        // Retain already "Posted" rows as-is
        if (status === "Posted") {
            return row;
        }

        // Update only if tableData status is "Posted"
        if (
            status === "Unposted" &&
            tableRecord &&
            tableRecord.status === "Posted"
        ) {
            return [
                paymentId,
                billId,
                rfpNo,
                formData.paymentPostedDate || createdDate, // Update the posted date
                formData.paymentReferenceNo || amount, // Update the reference number
                formData.paymentBreakdownReceiptDate || postedDate, // Update the receipt date
                formData.paymentReferenceDate || updatedDate, // Update the reference date
                "Posted", // Change the status to "Posted"
            ];
        }

        // Retain rows that are "Unposted" in both tableData and data
        return row;
    });
}

// Functions with Audit Trail
function appendPaymentRecords(data) {
    var sheet = accessSheet(PAYMENT);

    // Get the last row number to start appending new data
    var lastRow = sheet.getLastRow();

    // Convert data to append as a 2D array (required for setValues)
    var dataToAppend = data;

    // Log the ADD action for each record in the batch
    dataToAppend.forEach((record) => {
        const newRecord = record; // This is the new record being added

        // Prepare the new record for audit trail
        const headers = sheet
            .getRange(1, 1, 1, sheet.getLastColumn())
            .getValues()[0];
        const newRecordObject = mapValuesToObject(headers, newRecord);

        // Log the ADD action
        logAuditTrail(
            "ADD",
            "PAYMENT",
            newRecord[0], // Assuming the first column in the record is the unique ID
            {}, // No old values for ADD action
            newRecordObject, // New values being added
            [], // No changed fields for ADD action
            "Payment record added"
        );
    });

    // Append data
    sheet
        .getRange(lastRow + 1, 1, dataToAppend.length, dataToAppend[0].length)
        .setValues(dataToAppend);
}

// single record edit
function editPaymentRecord(formData) {
    var row = findRowById(PAYMENT, formData.editPaymentId);
    if (row != -1) {
        // Fetch the old record from the sheet
        const sheet =
            SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PAYMENT);
        const oldRecord = sheet
            .getRange(row, 1, 1, sheet.getLastColumn())
            .getValues()[0]; // Assuming row contains all fields

        // Get column headers
        const headers = sheet
            .getRange(1, 1, 1, sheet.getLastColumn())
            .getValues()[0]; // First row is headers

        // Track changes: old and new values
        const changedFields = [];
        const oldValues = {};
        const newValues = {};

        // Compare each field and check for differences
        for (let i = 0; i < oldRecord.length; i++) {
            const columnName = headers[i]; // Get column name from headers
            if (oldRecord[i] !== formData[columnName]) {
                changedFields.push(columnName);
                oldValues[columnName] = oldRecord[i];
                newValues[columnName] = formData[columnName];
            }
        }

        // If there are changes, log them into the audit trail
        if (changedFields.length > 0) {
            const auditData = {
                action: "EDIT",
                entity: "PAYMENT",
                id: formData.editPaymentId,
                oldValues: oldValues,
                newValues: newValues,
                changedFields: changedFields,
                remarks: "Payment record updated",
            };
            logAuditTrail(auditData); // Log the changes
        }

        // Update the record
        formData.status = formData.editPaymentPostedDate
            ? "Posted"
            : "Unposted";
        upsertRecord(PAYMENT, row, formData, formToSheetMap.PAYMENT);
    } else {
        Logger.log("Record not found");
    }
}
