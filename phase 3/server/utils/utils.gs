const EMPLOYEE_DETAILS = "EMPLOYEE_DETAILS";
const SIM_INVENTORY = "SIM_INVENTORY";
const SIM_PLANS = "SIM_PLANS";
const SIM_REQUEST_AND_ISSUANCE = "SIM_REQUEST_AND_ISSUANCE";
const RFP_GROUP = "RFP_GROUP";
const BILLING = "BILLING";
const RFP_SUMMARY = "RFP_SUMMARY";
const PAYMENT = "PAYMENT";
const EXCESS_CHARGES = "EXCESS_CHARGES";

const SHEETS = {
    EMPLOYEE_DETAILS: {
        DATE_FIELDS: ["DATE_HIRED"],
    },
    SIM_INVENTORY: {
        DATE_FIELDS: [
            "RR_DATE",
            // 'PURCHASING_RS_DATE', // MOVED TO SIM REQUEST AND ISSUANCE
            "PLAN_EFFECTIVITY",
            "TERMINATED_DATE",
            "LOCK_IN_PERIOD_FROM",
            "LOCK_IN_PERIOD_TO",
        ],
    },
    SIM_PLANS: {
        DATE_FIELDS: [],
    },
    SIM_REQUEST_AND_ISSUANCE: {
        DATE_FIELDS: [
            "RECEIVED_AND_APPROVED_DATE",
            "PRE_ISSUANCE_DATE",
            "ISSUANCE_DATE",
            "PURCHASING_RS_DATE", // MOVED TO SIM REQUEST AND ISSUANCE
        ],
    },
    RFP_GROUP: {
        DATE_FIELDS: [],
    },
    BILLING: {
        DATE_FIELDS: ["BILL_PERIOD_FROM", "BILL_PERIOD_TO"],
    },
    RFP_SUMMARY: {
        DATE_FIELDS: [
            "RFP_DATE",
            "DATE_OF_PAYMENT",
            "DATE_RECEIVED_BY_ACCTG",
            "CV_DATE",
            "CHECK_DATE",
            "DEPOSIT_DATE",
        ],
    },
    PAYMENT: {
        DATE_FIELDS: [
            "OR_DATE",
            "PAYMENT_POSTED_DATE",
            "PAYMENT_BREAKDOWN_RECEIPT_DATE",
        ],
    },
    EXCESS_CHARGES: {
        DATE_FIELDS: [
            "DEDUCTION_DATE",
            "APPROVAL_DATE",
            "REFERENCE",
            "EXCESS_CHARGE_DATE",
        ],
    },
};

const KEY_ORDER = {
    RFP_GROUP: [
        "groupID",
        "groupName",
        "company",
        "networkProvider",
        "payableTo",
    ],
    RFP_SUMMARY: ["rfpNo", "rfpDate", "billPeriod", "rfpInfo"],
    EXCESS_CHARGES: [
        "id",
        "billId",
        "groupId",
        "employeeNo",
        "simCardId",
        "excessChargeDate",
        "excessCharge",
        "remainingExcessCharge",
    ],
};

const KEYS_TO_FORMAT_DATE = {
    RFP_SUMMARY: ["RFP_DATE"],
};

const DATE_FIELDS = {
    EMPLOYEE_DETAILS: ["DATE_HIRED"],
};

const formToSheetMap = {
    RFP_GROUP: {
        editGroupName: "RFP_GROUP_NAME",
        editNetworkProvider: "NETWORK_PROVIDER",
        editCompany: "RFP_COMPANY",
        editPayableTo: "PAYABLE_TO",
        groupId: "RFP_GROUP_ID",
    },
    BILLING: {
        editAdjAmount: "ADJ_AMOUNT",
        editExcessCharges: "EXCESS_CHARGES",
        editOtherCharges: "OTHER_CHARGES",
        editPreviousBillAmount: "PREVIOUS_BILL_AMOUNT",
        editPreviousBillPayment: "PREVIOUS_BILL_PAYMENT",
        editRfpAmount: "RFP_AMOUNT",
        editCurrentChargeAmount: "CURRENT_CHARGE_AMOUNT",
        editAmountDue: "AMOUNT_DUE",
        editWithholdingTax: "WITHHOLDING_TAX",
        editAmountAfterTax: "AMOUNT_AFTER_TAX",
        editChargeToBdc: "CHARGE_TO_BDC",
        editWithSoa: "WITH_SOA",
    },
    RFP_SUMMARY: {
        editRfpNo: "RFP_NO",
        editDateReceivedByAccounting: "DATE_RECEIVED_BY_ACCTG",
        editCheckNo: "CHECK_NO",
        editCheckDate: "CHECK_DATE",
        editCvNo: "CV_NO",
        editCvDate: "CV_DATE",
        editDepositDate: "DEPOSIT_DATE",
        editDateOfPayment: "DATE_OF_PAYMENT",
    },
    PAYMENT: {
        editBillId: "BILL_ID",
        editPaymentReferenceNo: "PAYMENT_REFERENCE_NO",
        editPaymentReferenceDate: "PAYMENT_REFERENCE_DATE",
        editPaymentPostedDate: "PAYMENT_POSTED_DATE",
        editPaymentBreakdownReceiptDate: "PAYMENT_BREAKDOWN_RECEIPT_DATE",
        status: "STATUS",
    },
    EXCESS_CHARGES: {
        id: "EC_ID",
        billId: "BILL_ID",
        groupId: "GROUP_ID",
        employeeNo: "EMPLOYEE_NO",
        simCardId: "SIM_CARD_ID",
        excessChargeDate: "EXCESS_CHARGE_DATE",
        excessCharge: "EXCESS_CHARGE",
        remainingExcessCharge: "REMAINING_EXCESS_CHARGE",
    },
    // TICKET_MANAGEMENT: {

    // }
};

// Function to Access Sheet via Sheet Name
function accessSheet(sheetName) {
    var id = SpreadsheetApp.getActiveSpreadsheet().getId();
    var sheet = SpreadsheetApp.openById(id).getSheetByName(sheetName);
    return sheet;
}

// Function to add record to a specific sheet
function addRecordToSheet(
    formData,
    sheetName,
    keysToFormatDate,
    keyOrder,
    successMessage = ""
) {
    try {
        // Access the sheet by name
        var sheet = accessSheet(sheetName);

        // Create an array to hold the values in the desired order
        var formDataArray = keyOrder.map(function (key) {
            var value = formData[key];

            // Format date fields if necessary
            if (keysToFormatDate.includes(key) && value) {
                value = formatDate(value);
            }
            return value;
        });

        // Append the array to the sheet
        sheet.appendRow(formDataArray);
        console.log(formDataArray);

        // return "Success";
        return { status: "SUCCESS", message: successMessage }; // Return a structured response
    } catch (error) {
        // return "An error occurred while adding new record: " + error.toString();
        return {
            status: "ERROR",
            message:
                "An error occurred while adding new record: " +
                error.toString(),
        };
    }
}

// Function to generate ID based on header
function generateID(sheetName, header) {
    var sheet = accessSheet(sheetName);
    var columnIndex = getColumnIndex(sheet, header);
    var result = accessSheet(sheetName)
        .getRange(sheet.getLastRow(), columnIndex)
        .getValue();
    return result === header ? 1 : `${parseInt(result) + 1}`;
}

// Function to get the column index based on header (Helper function of generateID function)
function getColumnIndex(sheet, header) {
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    return headers.indexOf(header) + 1; // Adding 1 because column index starts from 1
}

// Optimized Version of getDataFromSheet
function getDataFromSheet(sheetName, dateFields = []) {
    const sheet =
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!sheet) {
        console.error(`Sheet with name ${sheetName} not found.`);
        return [];
    }

    const rows = sheet.getDataRange().getValues();
    if (rows.length === 0) {
        return [];
    }

    const headers = rows[0];
    const dataRows = rows.slice(1);
    const headerKeys = headers.reduce(
        (acc, key) => ({ ...acc, [key]: "" }),
        {}
    );

    const objectsArray = dataRows.map((row) => {
        return headers.reduce((rowData, header, index) => {
            const cell = row[index];
            rowData[header] =
                dateFields.includes(header) && cell !== ""
                    ? formatDate(cell)
                    : cell;
            return rowData;
        }, {});
    });

    console.log(objectsArray.length > 0 ? objectsArray : [headerKeys]);
    return objectsArray.length > 0 ? objectsArray : [headerKeys];
}

// Function to format date as "mm-dd-yy"
// function formatDate(date) {
//   if (!date) return '';
//   var formattedDate = Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), 'MM-dd-yy');
//   return formattedDate;
// }

// Function to format date as "mm/dd/yyyy"
function formatDate(date) {
    if (!date) return "";
    var formattedDate = Utilities.formatDate(
        new Date(date),
        Session.getScriptTimeZone(),
        "MM/dd/yyyy"
    );
    return formattedDate;
}

// Function to delete a record by column value
function deleteRecordByColumnValue(identifierValue, columnName, sheetName) {
    var sheet = accessSheet(sheetName);

    if (!sheet) {
        return { status: "Error", message: "Sheet not found" };
    }

    var range = sheet.getDataRange();
    var values = range.getValues();
    if (values.length === 0) {
        return { status: "Error", message: "Sheet is empty" };
    }

    var headers = values[0];
    var columnIndex = headers.indexOf(columnName);

    if (columnIndex === -1) {
        return { status: "Error", message: "Column not found" };
    }

    for (var rowIndex = 1; rowIndex < values.length; rowIndex++) {
        if (values[rowIndex][columnIndex] == identifierValue) {
            sheet.deleteRow(rowIndex + 1); // Add 1 because row indices are 1-based in Google Sheets
            return { status: "Success" };
        }
    }

    return { status: "Error", message: "Record not found" };
}

function findRowById(sheetName, uniqueId, idPosition = 0) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    var data = sheet.getDataRange().getValues();

    // Assuming the unique ID is in the first column (Column A)
    for (var i = 1; i < data.length; i++) {
        if (data[i][idPosition] == uniqueId) {
            console.log(data);
            return i + 1; // Row number in Google Sheets is 1-based
        }
    }

    return -1; // Return -1 if the ID is not found
}

function upsertRecord(sheetName, row, formData, formToSheetMap) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]; // Get headers

    // Check if the row exists and is not empty
    var existingRowValues = sheet
        .getRange(row, 1, 1, sheet.getLastColumn())
        .getValues()[0];
    var isRowEmpty = existingRowValues.every(function (cell) {
        return cell == "";
    });

    // If the row is empty, append a new row
    if (isRowEmpty) {
        row = sheet.getLastRow() + 1;
        sheet.appendRow(new Array(headers.length).fill("")); // Ensure the new row has the same number of columns
    }

    var valuesToUpdate = sheet
        .getRange(row, 1, 1, sheet.getLastColumn())
        .getValues();
    var rowValues = valuesToUpdate[0];

    for (var formKey in formData) {
        var sheetHeader = formToSheetMap[formKey]; // Get the corresponding sheet header
        if (sheetHeader) {
            var colIndex = headers.indexOf(sheetHeader); // Find the column index for the sheet header (0-based)
            if (colIndex >= 0) {
                // Only update if the column exists
                var newValue = formData[formKey];
                rowValues[colIndex] = newValue; // Update the value in the row array
            }
        }
    }

    sheet.getRange(row, 1, 1, sheet.getLastColumn()).setValues([rowValues]); // Set all values at once
}

// Function to get the last id from a sheet by using google sheets formula
// function getLastId(sheetName) {
//     const HELPER = "HELPER";
//     var helperSheet = accessSheet(HELPER);
//     if (!helperSheet) {
//         helperSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(HELPER);
//     }

//     // Use a temporary cell to store the formula
//     const tempCell = helperSheet.getRange('A1');

//     // Formula to find the last ID, increment by 1, or set to 1 if no IDs are present
//     const formula = `=IF(ISNUMBER(INDEX(${sheetName}!A:A, MAX(IF(${sheetName}!A:A<>"", ROW(${sheetName}!A:A), 0)))), INDEX(${sheetName}!A:A, MAX(IF(${sheetName}!A:A<>"", ROW(${sheetName}!A:A), 0))) + 1, 1)`;

//     tempCell.setFormula(formula);

//     // Wait for the formula to calculate
//     SpreadsheetApp.flush();

//     const result = tempCell.getValue();
//     tempCell.clear(); // Clear the temporary cell

//     // Ensure the result is a number
//     return Number(result);
// }

// Function to get the last id from a sheet
function getLastId(sheetName, column) {
    var sheet = accessSheet(sheetName);
    if (!sheet) throw new Error("Sheet not found");

    var values = sheet
        .getRange(column + ":" + column)
        .getValues()
        .flat()
        .map((value) => {
            // Convert string to number if possible, otherwise return NaN
            const numValue = Number(value);
            return !isNaN(numValue) ? numValue : NaN;
        })
        .filter((value) => typeof value === "number" && !isNaN(value)); // Filter out NaN values

    // Return the next ID, or start with 1 if no valid IDs are found
    return values.length ? Math.max(...values) + 1 : 1;
}

const rfpSummaryDateFields = [
    "RFP_DATE",
    "DATE_OF_PAYMENT",
    "DATE_RECEIVED_BY_ACCTG",
    "CV_DATE",
    "CHECK_DATE",
    "DEPOSIT_DATE",
];
const billingDateFields = ["BILL_PERIOD_FROM", "BILL_PERIOD_TO"];

function testGetLastId() {
    var lastId = getLastId("PAYMENT", "A");
    console.log(lastId);
    // var data = getDataFromSheet(RFP_SUMMARY, [
    //               'RFP_DATE',
    //               'DATE_OF_PAYMENT',
    //               'DATE_RECEIVED_BY_ACCTG',
    //               'CV_DATE',
    //               'CHECK_DATE',
    //               'DEPOSIT_DATE'
    //           ]);
    // var data = getDataFromSheet(PAYMENT, []);
    // console.log(data);
}

function testfn() {
    const dataset = getDataFromSheet("BILLING", billingDateFields);
    const headers = [
        "BILL_ID",
        "RFP_GROUP_ID",
        "SIM_CARD_ID",
        "ISSUANCE_NO",
        "RFP_NO",
        "BILL_PERIOD_FROM",
        "BILL_PERIOD_TO",
        "MONTHLY_RECURRING_FEE",
        "EXCESS_CHARGES",
        "OTHER_CHARGES",
        "PREVIOUS_BILL_AMOUNT",
        "PREVIOUS_BILL_PAYMENT",
        "CURRENT_CHARGE_AMOUNT",
        "ADJ_AMOUNT",
        "AMOUNT_DUE",
        "RFP_AMOUNT",
        "WITHHOLDING_TAX",
        "AMOUNT_AFTER_TAX",
        "CHARGE_TO_BDC",
    ];
    const query = new Utils.Query(dataset, headers);
    console.log(query.one({ RFP_NO: { $eq: "HR-2024-001" } }));
}

function selectColumnsFromSheet(sheetName) {
    try {
        const sheet = new Utils.Sheet(sheetName, {
            row: {
                start: 1,
            },
        });
        // const values = sheet.getValuesByColumns(...headers); // sheet.getValues();
        const values = sheet.getValues(); // sheet.getValues();

        return JSON.stringify(values);
    } catch (error) {
        return Utils.ErrorHandler(error, {
            arguments,
            value: [],
        });
    }
}

function getRfpPaymentBreakdown(rfpNo) {
    try {
        const sheetName = "BILLING";

        const sheet = new Utils.Sheet(sheetName, {
            row: {
                start: 1,
            },
        });

        const exp = sheet.experimentalQuery();

        const queried = exp.findMany({ RFP_NO: { $eq: rfpNo } }).toObject();
        // console.log(JSON.stringify(queried));
        // return JSON.stringify(queried);
        return queried;
    } catch (error) {
        return Utils.ErrorHandler(error, {
            arguments,
        });
    }
}

function listExcessCharges() {
    const excessChargesSheet =
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName("EXCESS_CHARGES");

    // Fetch all data from the Excess Charges sheet
    const excessChargesData = excessChargesSheet.getDataRange().getValues();

    // Initialize a map to store previous excess charges by EMPLOYEE_ID
    const previousExcessChargesMap = new Map();

    // Initialize an array to store the results
    const results = [];

    // Iterate over Excess Charges data to find records with excess charges
    for (let i = 1; i < excessChargesData.length; i++) {
        const ecId = excessChargesData[i][0]; // Assuming EC_ID is in the first column
        const billId = excessChargesData[i][1]; // Assuming BILL_ID is in the second column
        const employeeId = excessChargesData[i][2]; // Assuming EMPLOYEE_ID is in the ninth column
        const excessCharge = excessChargesData[i][3]; // Assuming EXCESS_CHARGE is in the third column

        // Look up the previous excess charge using the map
        const previousExcessCharge =
            previousExcessChargesMap.get(employeeId) || 0;

        // Add the record to the results array
        results.push([
            ecId,
            billId,
            employeeId,
            excessCharge,
            previousExcessCharge,
        ]);

        // Update the map with the new excess charge
        previousExcessChargesMap.set(
            employeeId,
            previousExcessCharge + excessCharge
        );
    }

    // Log the results or write them to another sheet
    Logger.log(results);

    // Optionally, write to another sheet
    // const resultSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Results');
    // resultSheet.getRange(1, 1, results.length, results[0].length).setValues(results);
}

// Under code was testing
function deleteRfpGroup(query) {
    try {
        const sheet = new Utils.Sheet("RFP_GROUP", {
            row: {
                start: 1,
            },
        });

        return sheet.experimentalQuery().findOneAndDelete(query);
    } catch (error) {
        return Utils.ErrorHandler(error, {
            arguments,
        });
    }
}
