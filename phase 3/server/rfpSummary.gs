// const SPREADSHEET_ID = '1soomTiHZWeKUdo71IoEBEMgQrXb5xvz5ksfaKU0uMuE';
const SPREADSHEET_ID = "10rejSExQPa8GZcdyPEaALwbY6XrGsGYhYGoPL8ut_xI";

// function createRfpSummary(formData) {
//   var sheet = accessSheet(RFP_SUMMARY);
//     // Get the last row number to start appending new data
//     var lastRow = sheet.getLastRow();

//     // Convert formData into a 2D array (required for setValues)
//     var dataToAppend = [formData];

//     // Append data
//     sheet.getRange(lastRow + 1, 1, dataToAppend.length, dataToAppend[0].length).setValues(dataToAppend);
// }

// function editRfpSummary(formData) {
//     console.log(formData)
//     var row = findRowById(RFP_SUMMARY, formData.editRfpNo);
//     if (row != -1) {
//         upsertRecord(RFP_SUMMARY, row, formData, formToSheetMap.RFP_SUMMARY);
//     } else {
//         console.log("Editing Error");
//     }
// }

// function deleteRfpSummary(id) {
//     var result = deleteRecordByColumnValue(id, 'RFP_NO', RFP_SUMMARY);
//     return result;
// }

// Function to create PDF
// function createRfpPdf(data) {

//   populateRfpTemplate(data);

//   SpreadsheetApp.flush();

//     const fr = 0, fc = 0, lc = 19, lr = 100;

//     const url = "https://docs.google.com/spreadsheets/d/" + "111oaSBb2sQ9ZtMoIWWiy2BR7s1qGvCxiPYYB-3zN4QQ" + "/export" +
//         "?format=pdf&" +
//         "size=A4&" +  // Set paper size to A4
//         "portrait=true&" +  // Set orientation to portrait
//         "fitw=true&" +  // Disable fit to width
//         "scale=4&" +  // Scale to fit entire content on a single page
//         "gridlines=false&" +
//         "printtitle=false&" +
//         "top_margin=0.25&" +
//         "bottom_margin=0.25&" +
//         "left_margin=0.25&" +
//         "right_margin=0.25&" +
//         "horizontal_alignment=CENTER&" +
//         "vertical_alignment=MIDDLE&" +
//         "sheetnames=false&" +
//         "pagenum=UNDEFINED&" +
//         "attachment=true&" +
//         "gid=" + 103162632 + '&' +
//         "r1=" + fr + "&c1=" + fc + "&r2=" + lr + "&c2=" + lc;

//     const params = {
//         method: "GET",
//         headers: {
//             "authorization": "Bearer " + ScriptApp.getOAuthToken()
//         }
//     };

//     const blob = UrlFetchApp.fetch(url, params).getBlob();
//     const pdfUrl = "data:application/pdf;base64," + Utilities.base64Encode(blob.getBytes());

//     return pdfUrl;
// }

// function createPaymentBreakdownPdf(data) {

//   console.log(data);

//   populatePaymentBreakdown(data);

//   SpreadsheetApp.flush();

//     const fr = 0, fc = 0, lc = 13, lr = 100;

//     const url = "https://docs.google.com/spreadsheets/d/" + "111oaSBb2sQ9ZtMoIWWiy2BR7s1qGvCxiPYYB-3zN4QQ" + "/export" +
//         "?format=pdf&" +
//         "size=A4&" +  // Set paper size to A4
//         "portrait=true&" +  // Set orientation to portrait
//         "fitw=true&" +  // Disable fit to width
//         "scale=4&" +  // Scale to fit entire content on a single page
//         "gridlines=false&" +
//         "printtitle=false&" +
//         "top_margin=0.25&" +
//         "bottom_margin=0.25&" +
//         "left_margin=0.25&" +
//         "right_margin=0.25&" +
//         "horizontal_alignment=CENTER&" +
//         "vertical_alignment=MIDDLE&" +
//         "sheetnames=false&" +
//         "pagenum=UNDEFINED&" +
//         "attachment=true&" +
//         "gid=" + 802214753 + '&' +
//         "r1=" + fr + "&c1=" + fc + "&r2=" + lr + "&c2=" + lc;

//     const params = {
//         method: "GET",
//         headers: {
//             "authorization": "Bearer " + ScriptApp.getOAuthToken()
//         }
//     };

//     const blob = UrlFetchApp.fetch(url, params).getBlob();
//     const pdfUrl = "data:application/pdf;base64," + Utilities.base64Encode(blob.getBytes());

//     return pdfUrl;
// }

// Function to create PDF dynamically
// Working function
function createPDF(
    spreadsheetId,
    sheetName,
    range,
    populateDataFunc,
    options = {}
) {
    const {
        format = "pdf",
        paperSize = "A4",
        orientation = "landscape",
        scale = 4,
        topMargin = 0.25,
        bottomMargin = 0.25,
        leftMargin = 0.5,
        rightMargin = 0.5,
        gridlines = false,
        printTitle = false,
        horizontalAlignment = "CENTER",
        verticalAlignment = "MIDDLE",
        sheetNames = false,
        pageNum = "UNDEFINED",
    } = options;

    // Call the provided function to populate data
    if (typeof populateDataFunc === "function") {
        populateDataFunc(); // Call the function without arguments or modify it to accept parameters as needed
    } else {
        throw new Error("populateDataFunc must be a function");
    }

    SpreadsheetApp.flush();

    const fr = range.startRow || 0;
    const fc = range.startCol || 0;
    const lr = range.endRow || 100;
    const lc = range.endCol || 19;

    const url =
        `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export` +
        `?format=${format}` +
        `&size=${paperSize}` +
        `&portrait=${orientation === "portrait"}` +
        `&fitw=true` +
        `&scale=${scale}` +
        `&gridlines=${gridlines}` +
        `&printtitle=${printTitle}` +
        `&top_margin=${topMargin}` +
        `&bottom_margin=${bottomMargin}` +
        `&left_margin=${leftMargin}` +
        `&right_margin=${rightMargin}` +
        `&horizontal_alignment=${horizontalAlignment}` +
        `&vertical_alignment=${verticalAlignment}` +
        `&sheetnames=${sheetNames}` +
        `&pagenum=${pageNum}` +
        `&attachment=true` +
        `&gid=${getSheetGid(spreadsheetId, sheetName)}` +
        `&r1=${fr}&c1=${fc}&r2=${lr}&c2=${lc}`;

    const params = {
        method: "GET",
        headers: {
            authorization: "Bearer " + ScriptApp.getOAuthToken(),
        },
    };

    const blob = UrlFetchApp.fetch(url, params).getBlob();
    const pdfUrl =
        "data:application/pdf;base64," +
        Utilities.base64Encode(blob.getBytes());

    return pdfUrl;
}

// Helper function to get the GID of the specified sheet
function getSheetGid(spreadsheetId, sheetName) {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(sheetName);
    return sheet ? sheet.getSheetId() : null;
}

function getRfpPdfUrl(data) {
    console.log(data);
    const pdfUrl = createPDF(
        SPREADSHEET_ID,
        "RFP_TEMPLATE",
        { startRow: 0, startCol: 0, endRow: 100, endCol: 19 },
        () => {
            // Call your data population functions here
            populateRfpTemplate(data); // First method
            // OR
            // populateAnotherTemplate(data); // Second method
        },
        {
            paperSize: "A4",
            orientation: "portrait",
            scale: 4,
            gridlines: false,
            leftMargin: 0.1,
            rightMargin: 0.1,
        }
    );

    return pdfUrl;
}

function getPaymentBreakdownPdfUrl(data) {
    console.log(data);
    const pdfUrl = createPDF(
        SPREADSHEET_ID,
        "PAYMENT_BREAKDOWN",
        { startRow: 0, startCol: 0, endRow: data.length + 4, endCol: 13 },
        () => {
            // Call your data population functions here
            populatePaymentBreakdown(data); // First method
            // OR
            // populateAnotherTemplate(data); // Second method
        },
        {
            paperSize: "A4",
            orientation: "landscape",
            scale: 2,
            gridlines: true,
            verticalAlignment: "TOP",
            pageNum: "CENTER",
            leftMargin: 0.75,
            rightMargin: 0.25,
        }
    );

    return pdfUrl;
}

function formatDatezxcv(inputDate) {
    const date = new Date(inputDate);
    return date
        .toLocaleDateString("en-US", {
            month: "short",
            day: "2-digit",
            year: "numeric",
        })
        .replace(",", ".");
}

function populateRfpTemplate(data) {
    // { TOTAL_WITHHOLDING_TAX: 643.2000000000013,
    // DATE_RECEIVED_BY_ACCTG: '',
    // RFP_NO: 'HR-SIM-2024-001',
    // CHECK_NO: '',
    // BILL_PERIOD_TO: '10/31/2024',
    // BILL_PERIOD_FROM: '10/01/2024',
    // DATE_OF_PAYMENT: '',
    // PAYABLE_TO: 'Smart Communication. Inc.',
    // RFP_COMPANY: 'Borland Development Corporation',
    // TOTAL_AMOUNT_AFTER_TAX: 35356.79999999995,
    // RFP_INFO: '{"RFP_GROUP_ID":"1","RFP_GROUP_NAME":"BDC- Sun Fixed Load","RFP_COMPANY":"Borland Development Corporation","NETWORK_PROVIDER":"Sun","PAYABLE_TO":"Smart Communication. Inc."}',
    // TOTAL_RFP_AMOUNT: 36000,
    // CV_NO: '',
    // DEPOSIT_DATE: '',
    // CV_DATE: '',
    // CHECK_DATE: '',
    // RFP_DATE: '10/18/2024',
    // NETWORK_PROVIDER: 'Sun' }

    const reference = getReferenceData();
    const parsedReference = JSON.parse(reference);

    // console.log(reference);

    console.log(data);

    var sheet = accessSheet("RFP_TEMPLATE");

    // Parse the RFP_INFO string to extract RFP Company
    let rfpInfo = {};
    try {
        rfpInfo = JSON.parse(data.RFP_INFO);
    } catch (error) {
        console.error("Error parsing RFP_INFO:", error);
    }

    var billPeriod = `${formatDatezxcv(
        data.BILL_PERIOD_FROM
    )} - ${formatDatezxcv(data.BILL_PERIOD_TO)}`;
    console.log(billPeriod);
    var dateOfPayment = data.DATE_OF_PAYMENT
        ? formatDatezxcv(data.DATE_OF_PAYMENT)
        : "";

    const rfpCompany = rfpInfo.RFP_COMPANY.toUpperCase();

    sheet
        .getRange("Q7:S7")
        .setValue(
            formatDatezxcv(
                Utilities.formatDate(new Date(), "GMT+1", "MM/dd/yyyy")
            )
        );
    sheet.getRange("G2:M2").setValue(rfpCompany);
    sheet.getRange("D6:J6").setValue(data.NETWORK_PROVIDER);
    sheet.getRange("D7:J7").setValue(data.PAYABLE_TO);
    sheet.getRange("D9:F9").setValue(dateOfPayment);

    // RFP Template - Representative's Name
    sheet.getRange("M16:N16").setValue(parsedReference[0].NAME);
    // RFP Template - Requested by
    sheet.getRange("B35:G35").setValue(parsedReference[1].NAME);
    // RFP Template - Standard Approval
    sheet.getRange("B41:G41").setValue(parsedReference[2].NAME);

    sheet.getRange("Q6:S6").setValue(data.RFP_NO);
    sheet.getRange("D22:F22").setValue(billPeriod);
    sheet.getRange("S23").setValue(data.TOTAL_RFP_AMOUNT);
    sheet.getRange("S24").setValue(data.TOTAL_WITHHOLDING_TAX);
    sheet.getRange("S25").setValue(data.TOTAL_AMOUNT_AFTER_TAX);
}

function populatePaymentBreakdown(data) {
    const sheetName = "PAYMENT_BREAKDOWN"; // Change this to the name of your target sheet
    console.log(data);

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(sheetName);

    // Check if the sheet exists
    if (!sheet) {
        Logger.log(`Sheet with name "${sheetName}" does not exist.`);
        return;
    }

    // Clear the sheet before writing
    sheet.clear();

    // Insert the sheet name and today's date
    sheet.insertRowBefore(1); // Insert a row at the top
    sheet
        .getRange("A1")
        .setValue(sheetName)
        .setFontSize(14)
        .setFontFamily("Poppins")
        .setFontWeight("bold")
        .setHorizontalAlignment("LEFT");
    sheet.insertRowBefore(2); // Insert another row for the date
    const today = new Date();
    const formattedDate = Utilities.formatDate(
        today,
        Session.getScriptTimeZone(),
        "MM/dd/yyyy"
    );
    sheet
        .getRange("M1")
        .setValue(`Date: ${formattedDate}`)
        .setFontSize(14)
        .setFontFamily("Poppins")
        .setHorizontalAlignment("LEFT");

    // Define the header order with spaces
    const headers = [
        "RFP NO",
        "SIM CARD ID",
        "EMPLOYEE NAME",
        "ACCOUNTING_CODE",
        "ACCOUNT NO",
        "MOBILE NO",
        "BILL PERIOD FROM",
        "BILL PERIOD TO",
        "CHARGE TO BDC",
        "EXCESS CHARGE",
        "RFP AMOUNT",
        "WITHHOLDING TAX",
        "AMOUNT AFTER TAX",
    ];

    const rows = data.map((item) =>
        headers.map((header) => item[header.replace(/ /g, "_")])
    ); // Replace spaces with underscores in keys

    // Insert header and data in one go
    sheet
        .getRange(3, 1, rows.length + 1, headers.length)
        .setValues([headers, ...rows]);

    // Set the format of the SIM CARD ID column to text
    sheet.getRange(4, 2, rows.length, 1).setNumberFormat("@"); // 2nd column is SIM CARD ID

    // Apply formatting to the header row
    const headerRange = sheet.getRange(3, 1, 1, headers.length);
    headerRange
        .setFontWeight("bold")
        .setFontSize(12)
        .setHorizontalAlignment("center")
        .setFontFamily("Poppins");

    // Apply borders around the data range
    const dataRange = sheet.getRange(3, 1, rows.length + 1, headers.length);
    // dataRange.setBorder(true, true, true, true, true, true); // Uncomment if you want borders

    // Center align text and set font to Poppins for all rows
    const allDataRange = sheet.getRange(3, 1, rows.length + 2, headers.length); // +2 for headers and title/date
    allDataRange
        .setFontSize(12)
        .setFontFamily("Poppins")
        .setHorizontalAlignment("center");

    // Calculate totals for specified columns
    const lastRow = rows.length + 3; // Last row with data (including headers and the title/date)

    // Set TOTAL label
    sheet
        .getRange(lastRow + 1, 1)
        .setValue("TOTAL")
        .setFontWeight("bold");

    // Calculate sums
    sheet.getRange(lastRow + 1, 9).setFormula(`=SUM(I4:I${lastRow})`); // CHARGE TO BDC
    sheet.getRange(lastRow + 1, 11).setFormula(`=SUM(K4:K${lastRow})`); // RFP AMOUNT
    sheet.getRange(lastRow + 1, 12).setFormula(`=SUM(L4:L${lastRow})`); // WITHHOLDING TAX
    sheet.getRange(lastRow + 1, 13).setFormula(`=SUM(M4:M${lastRow})`); // AMOUNT AFTER TAX

    // Apply number formatting to specified columns
    const numberColumns = [9, 10, 11, 12, 13]; // Column indices for the specified headers
    numberColumns.forEach((col) => {
        sheet.getRange(4, col, rows.length, 1).setNumberFormat("#,##0.00"); // Format as number with commas and two decimal places
    });

    // Apply borders to the TOTAL row
    const totalRange = sheet.getRange(lastRow + 1, 1, 1, headers.length);
    totalRange.setFontWeight("bold").setFontSize(12); // Make TOTAL bold and larger

    // Center align text for total row
    totalRange.setHorizontalAlignment("center").setFontFamily("Poppins");
}

function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu("Custom Menu").addItem("Create PDF", "createPDF").addToUi();
}

// Function to format the date range
function formatDateRange(dateRange) {
    // Split the date range into two dates
    const [startDate, endDate] = dateRange.split(" - ");

    // Debugging step: Ensure the dates are split correctly
    console.log("Start Date:", startDate); // Should be '09/01/2024'
    console.log("End Date:", endDate); // Should be '09/30/2024'

    // Format both start and end dates
    const formattedStartDate = formatDateToStr(startDate);
    const formattedEndDate = formatDateToStr(endDate);

    // Return the correctly formatted date range
    return `${formattedStartDate} - ${formattedEndDate}`;
}

function formatDateToStr(dateStr) {
    const months = [
        "Jan.",
        "Feb.",
        "Mar.",
        "Apr.",
        "May.",
        "Jun.",
        "Jul.",
        "Aug.",
        "Sept.",
        "Oct.",
        "Nov.",
        "Dec.",
    ];
    const [month, day, year] = dateStr.split("/");
    const monthName = months[parseInt(month, 10) - 1];
    return `${monthName} ${day}, ${year}`;
}

function getRfpDetails(headers = []) {
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

        // Updated SQL query to get only the required fields
        const query = `
          SELECT 
          b.RFP_NO,
          b.BILL_PERIOD_FROM,
          b.BILL_PERIOD_TO,
          rg.RFP_GROUP_ID,
          rg.RFP_GROUP_NAME AS RFP_GROUP_NAME,
          rg.RFP_COMPANY AS RFP_COMPANY,
          rg.PAYABLE_TO,
          sp.NETWORK_PROVIDER,
          SUM(b.AMOUNT_AFTER_TAX) AS TOTAL_AMOUNT_AFTER_TAX,
          SUM(b.RFP_AMOUNT) AS TOTAL_RFP_AMOUNT,
          SUM(b.WITHHOLDING_TAX) AS TOTAL_WITHHOLDING_TAX

          FROM ? AS b
          LEFT JOIN ? AS si ON b.SIM_CARD_ID = si.SIM_CARD_ID
          LEFT JOIN ? AS sr ON b.ISSUANCE_NO = sr.ISSUANCE_NO
          LEFT JOIN ? AS rg ON si.RFP_GROUP_ID = rg.RFP_GROUP_ID
          LEFT JOIN ? AS ed ON sr.GROUP_ID = ed.GROUP_ID
          LEFT JOIN ? AS sp ON si.PLAN_ID = sp.PLAN_ID

          GROUP BY b.RFP_NO, b.BILL_PERIOD_FROM, b.BILL_PERIOD_TO, rg.RFP_GROUP_ID, rg.RFP_GROUP_NAME, rg.RFP_COMPANY, rg.PAYABLE_TO, sp.NETWORK_PROVIDER
      `;

        // Execute the query by passing datasets in the correct order
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
                    RFP_NO: "",
                    BILL_PERIOD: "",
                    RFP_GROUP_ID: "",
                    RFP_GROUP_NAME: "",
                    RFP_COMPANY: "",
                    NETWORK_PROVIDER: "",
                    PAYABLE_TO: "",
                    TOTAL_AMOUNT_AFTER_TAX: "0.00",
                    TOTAL_RFP_AMOUNT: "0.00",
                    TOTAL_WITHHOLDING_TAX: "0.00",
                },
            ]);
        }

        // Format date fields (assuming BILL_PERIOD_FROM and BILL_PERIOD_TO are date fields)
        execution.forEach((row) => {
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

        console.log(execution);
        return JSON.stringify(execution);
    } catch (error) {
        return Utils.ErrorHandler(error, {
            arguments,
            value: [],
        });
    }
}

function getRfpData() {
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

        // Initialize sheets
        const rfpSheet = new Utils.Sheet("RFP_SUMMARY", { row: { start: 1 } });
        const billingSheet = new Utils.Sheet("BILLING", { row: { start: 1 } });
        const simInventorySheet = new Utils.Sheet("SIM_INVENTORY", {
            row: { start: 1 },
        });
        const rfpGroupSheet = new Utils.Sheet("RFP_GROUP", {
            row: { start: 1 },
        });

        // Fetch data from sheets
        const rfpDataSet = rfpSheet.toObject();
        const billingDataSet = billingSheet.toObject();
        const simInventoryDataSet = simInventorySheet.toObject();
        const rfpGroupDataSet = rfpGroupSheet.toObject();

        // Log datasets to debug
        // console.log("RFP DataSet:", rfpDataSet);
        // console.log("Billing DataSet:", billingDataSet);

        // SQL query with required joins
        const query = `
      SELECT 
      r.RFP_NO,
      r.RFP_DATE,
      r.DATE_OF_PAYMENT,
      r.DATE_RECEIVED_BY_ACCTG,
      r.CV_NO,
      r.CV_DATE,
      r.CHECK_NO,
      r.CHECK_DATE,
      r.DEPOSIT_DATE,
      r.RFP_INFO,
      b.BILL_PERIOD_FROM,
      b.BILL_PERIOD_TO,
      COALESCE(SUM(b.AMOUNT_AFTER_TAX), 0) AS TOTAL_AMOUNT_AFTER_TAX,
      COALESCE(SUM(b.RFP_AMOUNT), 0) AS TOTAL_RFP_AMOUNT,
      COALESCE(SUM(b.WITHHOLDING_TAX), 0) AS TOTAL_WITHHOLDING_TAX
      FROM ? AS r
      LEFT JOIN ? AS b ON r.RFP_NO = b.RFP_NO
      LEFT JOIN ? AS si ON b.SIM_CARD_ID = si.SIM_CARD_ID
      LEFT JOIN ? AS rg ON si.RFP_GROUP_ID = rg.RFP_GROUP_ID
      GROUP BY r.RFP_NO, r.RFP_DATE, r.DATE_OF_PAYMENT, r.DATE_RECEIVED_BY_ACCTG, r.CV_NO, r.CV_DATE, r.CHECK_NO, r.CHECK_DATE, r.DEPOSIT_DATE, b.BILL_PERIOD_FROM, b.BILL_PERIOD_TO, RFP_INFO
    `;

        // Execute the query by passing datasets in the correct order
        const execution = Utils.sql(query, [
            rfpDataSet,
            billingDataSet,
            simInventoryDataSet,
            rfpGroupDataSet,
        ]);

        // Format date fields (assuming BILL_PERIOD_FROM and BILL_PERIOD_TO are date fields)
        // execution.forEach(row => {
        //   if (row.BILL_PERIOD_FROM) {
        //     row.BILL_PERIOD_FROM = Utilities.formatDate(new Date(row.BILL_PERIOD_FROM), Session.getScriptTimeZone(), "MM/dd/yyyy");
        //   }
        //   if (row.BILL_PERIOD_TO) {
        //     row.BILL_PERIOD_TO = Utilities.formatDate(new Date(row.BILL_PERIOD_TO), Session.getScriptTimeZone(), "MM/dd/yyyy");
        //   }
        //   if (row.RFP_DATE) {
        //     row.RFP_DATE = Utilities.formatDate(new Date(row.RFP_DATE), Session.getScriptTimeZone(), "MM/dd/yyyy");
        //   }
        //   if (row.DATE_OF_PAYMENT) {
        //     row.DATE_OF_PAYMENT = Utilities.formatDate(new Date(row.DATE_OF_PAYMENT), Session.getScriptTimeZone(), "MM/dd/yyyy");
        //   }
        //   if (row.DATE_RECEIVED_BY_ACCTG) {
        //     row.DATE_RECEIVED_BY_ACCTG = Utilities.formatDate(new Date(row.DATE_RECEIVED_BY_ACCTG), Session.getScriptTimeZone(), "MM/dd/yyyy");
        //   }
        //   if (row.CV_DATE) {
        //     row.CV_DATE = Utilities.formatDate(new Date(row.CV_DATE), Session.getScriptTimeZone(), "MM/dd/yyyy");
        //   }
        //   if (row.CHECK_DATE) {
        //     row.CHECK_DATE = Utilities.formatDate(new Date(row.CHECK_DATE), Session.getScriptTimeZone(), "MM/dd/yyyy");
        //   }
        //   if (row.DEPOSIT_DATE) {
        //     row.DEPOSIT_DATE = Utilities.formatDate(new Date(row.DEPOSIT_DATE), Session.getScriptTimeZone(), "MM/dd/yyyy");
        //   }
        // });

        // Check if dataset is empty, return placeholder if needed
        if (!execution || execution.length === 0) {
            return JSON.stringify([
                {
                    RFP_NO: "",
                    RFP_DATE: "",
                    BILL_PERIOD_FROM: "",
                    BILL_PERIOD_TO: "",
                    DATE_OF_PAYMENT: "",
                    DATE_RECEIVED_BY_ACCTG: "",
                    CV_NO: "",
                    CV_DATE: "",
                    CHECK_NO: "",
                    CHECK_DATE: "",
                    DEPOSIT_DATE: "",
                    TOTAL_AMOUNT_AFTER_TAX: "",
                    TOTAL_RFP_AMOUNT: "",
                    TOTAL_WITHHOLDING_TAX: "",
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

function getUniqueRfpNo() {
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

        // Initialize sheets
        const rfpSheet = new Utils.Sheet("RFP_SUMMARY", { row: { start: 1 } });
        const billingSheet = new Utils.Sheet("BILLING", { row: { start: 1 } });

        // Fetch data from sheets
        const rfpDataSet = rfpSheet.toObject();
        const billingDataSet = billingSheet.toObject();

        // SQL query to get unique RFP_NO from Billing not present in RFP_SUMMARY
        const query = `
      SELECT DISTINCT b.RFP_NO
      FROM ? AS b
      WHERE b.RFP_NO NOT IN (SELECT DISTINCT r.RFP_NO FROM ? AS r)
    `;

        // Execute the query by passing datasets in the correct order
        const execution = Utils.sql(query, [billingDataSet, rfpDataSet]);

        // Return the result as an array of RFP_NO values
        if (execution && execution.length > 0) {
            const s = execution.map((row) => row.RFP_NO);
            console.log(s);
            return execution.map((row) => row.RFP_NO);
        } else {
            return []; // Return an empty array if no data is found
        }
    } catch (error) {
        return Utils.ErrorHandler(error, {
            arguments,
            value: [],
        });
    }
}

function formatDateToString(date) {
    const monthNames = [
        "January",
        "February",
        "March",
        "April",
        "May",
        "June",
        "July",
        "August",
        "September",
        "October",
        "November",
        "December",
    ];

    // Get the formatted month name, day, and year
    const formattedDate = `${monthNames[date.getMonth()]} ${String(
        date.getDate()
    ).padStart(2, "0")}, ${date.getFullYear()}`;

    return formattedDate;
}

const asd = "HR-SIM-2024-001";
// working function to get sim card details, employee details
function getBillingSimAndEmployeeDetailsByRfpNo(rfpNo) {
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
        const rfpGroupSheet = new Utils.Sheet("RFP_GROUP", {
            row: { start: 1 },
        });
        const simPlanSheet = new Utils.Sheet("SIM_PLANS", {
            row: { start: 1 },
        });

        // Fetch data from sheets
        const billingData = billingSheet.toObject();
        const simInventoryData = simInventorySheet.toObject();
        const simRequestAndIssuanceData = simRequestAndIssuanceSheet.toObject();
        const employeeDetailsData = employeeDetailsSheet.toObject();
        const rfpGroupData = rfpGroupSheet.toObject();
        const simPlanData = simPlanSheet.toObject();

        // SQL-like query to select columns from BILLING, SIM_INVENTORY, and EMPLOYEE_DETAILS
        // const query = `
        //   SELECT
        //     b.RFP_NO,
        //     b.SIM_CARD_ID,             -- SIM Card ID from Billing
        //     b.BILL_PERIOD_FROM,        -- Bill Period From from Billing
        //     b.BILL_PERIOD_TO,          -- Bill Period To from Billing
        //     b.CHARGE_TO_BDC,           -- Charge to BDC from Billing
        //     b.EXCESS_CHARGES,          -- Excess Charges from Billing
        //     b.RFP_AMOUNT,              -- RFP Amount from Billing
        //     b.WITHHOLDING_TAX,         -- Withholding Tax from Billing
        //     b.AMOUNT_AFTER_TAX,        -- Amount After Tax from Billing
        //     si.ACCOUNT_NO,             -- Account No. from SIM Inventory
        //     si.MOBILE_NO,              -- Mobile No. from SIM Inventory
        //     e.FULL_NAME AS EMPLOYEE_NAME,                -- Employee Full Name from Employee Details
        //     e.ACCOUNTING_CODE AS CODES
        //   FROM ? AS b
        //   LEFT JOIN ? AS si ON b.SIM_CARD_ID = si.SIM_CARD_ID
        //   LEFT JOIN ? AS sra ON b.SIM_CARD_ID = si.SIM_CARD_ID
        //     AND b.ISSUANCE_NO = sra.ISSUANCE_NO
        //   LEFT JOIN ? AS e ON sra.GROUP_ID = e.GROUP_ID
        //   WHERE b.RFP_NO = ?
        // `;

        const query = `
      SELECT
        b.RFP_NO,
        b.SIM_CARD_ID,
        b.BILL_PERIOD_FROM,
        b.BILL_PERIOD_TO,
        b.CHARGE_TO_BDC,
        b.EXCESS_CHARGES,
        b.RFP_AMOUNT,
        b.WITHHOLDING_TAX,
        b.AMOUNT_AFTER_TAX,
        b.WITHHOLDING_TAX,
      CASE
        WHEN sr.REQUEST_STATUS = 'Issued' THEN sr.EMPLOYEE_INFO
        ELSE NULL
      END AS EMPLOYEE_INFO,
      CASE
        WHEN sr.REQUEST_STATUS = 'Issued' THEN sr.SIM_INFO
        ELSE NULL
      END AS SIM_INFO,
      sr.GROUP_ID,
      si.MOBILE_NO,
      si.ACCOUNT_NO,
      sp.NETWORK_PROVIDER,
      sp.CATEGORY,
      sp.PLAN_DETAILS
      FROM ? AS b
      LEFT JOIN ? AS si ON b.SIM_CARD_ID = si.SIM_CARD_ID
      LEFT JOIN ? AS sr ON b.ISSUANCE_NO = sr.ISSUANCE_NO AND sr.REQUEST_STATUS = 'Issued'
      LEFT JOIN ? AS rg ON si.RFP_GROUP_ID = rg.RFP_GROUP_ID
      LEFT JOIN ? AS ed ON sr.GROUP_ID = ed.GROUP_ID
      LEFT JOIN ? AS sp ON si.PLAN_ID = sp.PLAN_ID
      WHERE b.RFP_NO = ?
    `;

        // Execute the query
        const result = Utils.sql(query, [
            billingData,
            simInventoryData,
            simRequestAndIssuanceData,
            rfpGroupData,
            employeeDetailsData,
            simPlanData,
            rfpNo,
        ]);
        console.log(result);

        // Format date fields (assuming BILL_PERIOD_FROM and BILL_PERIOD_TO are date fields)
        result.forEach((row) => {
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

        // Return the result
        return JSON.stringify(result);
    } catch (error) {
        console.error(
            "Error fetching billing, SIM inventory, and employee details by RFP_NO:",
            error
        );
        throw new Error(
            "Failed to fetch billing, SIM inventory, and employee details for the selected RFP_NO"
        );
    }
}

// Audit Trail
function createRfpSummary(formData) {
    var sheet = accessSheet(RFP_SUMMARY);

    // Get the last row number to start appending new data
    var lastRow = sheet.getLastRow();

    // Convert formData into a 2D array (required for setValues)
    var dataToAppend = [formData];

    // Append data to the sheet
    sheet
        .getRange(lastRow + 1, 1, dataToAppend.length, dataToAppend[0].length)
        .setValues(dataToAppend);

    // Retrieve headers from the sheet to map formData fields to column names
    const headers = sheet
        .getRange(1, 1, 1, sheet.getLastColumn())
        .getValues()[0];

    // Map formData to an object for the audit trail
    const newRecord = mapValuesToObject(headers, formData);

    // Log the ADD action to the audit trail
    logAuditTrail(
        "ADD",
        "RFP_SUMMARY",
        formData[0], // Assuming the first field is a unique ID for this record
        {}, // No old values for ADD action
        newRecord, // New record data
        [], // No changed fields for ADD action
        "New RFP Summary created"
    );
}

function editRfpSummary(formData) {
    var row = findRowById(RFP_SUMMARY, formData.editRfpNo);

    if (row != -1) {
        var sheet = accessSheet(RFP_SUMMARY);

        // Get headers from the sheet
        var headers = sheet
            .getRange(1, 1, 1, sheet.getLastColumn())
            .getValues()[0];

        // Retrieve old values from the row
        var oldValues = sheet
            .getRange(row, 1, 1, sheet.getLastColumn())
            .getValues()[0];

        oldValues[1] = formatDate(oldValues[1]);
        oldValues[2] = formatDate(oldValues[2]);
        oldValues[3] = formatDate(oldValues[3]);
        oldValues[5] = formatDate(oldValues[5]);
        oldValues[7] = formatDate(oldValues[7]);
        oldValues[8] = formatDate(oldValues[8]);

        // Format the old values into an object for the audit trail
        var oldRecord = mapValuesToObject(headers, oldValues);

        // Update the record in the sheet
        upsertRecord(RFP_SUMMARY, row, formData, formToSheetMap.RFP_SUMMARY);

        // Retrieve new values after update
        var newValues = sheet
            .getRange(row, 1, 1, sheet.getLastColumn())
            .getValues()[0];

        newValues[1] = formatDate(newValues[1]);
        newValues[2] = formatDate(newValues[2]);
        newValues[3] = formatDate(newValues[3]);
        newValues[5] = formatDate(newValues[5]);
        newValues[7] = formatDate(newValues[7]);
        newValues[8] = formatDate(newValues[8]);

        var newRecord = mapValuesToObject(headers, newValues);

        // Identify changed fields
        var changedFields = [];
        for (var i = 0; i < headers.length; i++) {
            if (oldValues[i] !== newValues[i]) {
                changedFields.push(headers[i]);
            }
        }

        // Log the EDIT action to the audit trail
        logAuditTrail(
            "EDIT",
            "RFP_SUMMARY",
            formData.editRfpNo, // Unique ID
            oldRecord, // Old values
            newRecord, // New values
            changedFields, // Fields that were changed
            "RFP Summary updated"
        );
    } else {
        console.log("Editing Error");
    }
}

function deleteRfpSummary(id) {
    var sheet = accessSheet(RFP_SUMMARY);

    // Find the row to delete by ID
    var row = findRowById(RFP_SUMMARY, id);

    if (row === -1) {
        throw new Error(`Record with ID ${id} not found in RFP_SUMMARY.`);
    }

    // Retrieve headers from the sheet
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    // Retrieve old values before deletion
    var oldValues = sheet
        .getRange(row, 1, 1, sheet.getLastColumn())
        .getValues()[0];

    oldValues[1] = formatDate(oldValues[1]);
    oldValues[2] = formatDate(oldValues[2]);
    oldValues[3] = formatDate(oldValues[3]);
    oldValues[5] = formatDate(oldValues[5]);
    oldValues[7] = formatDate(oldValues[7]);
    oldValues[8] = formatDate(oldValues[8]);

    var oldRecord = mapValuesToObject(headers, oldValues);

    // Perform the deletion
    var result = deleteRecordByColumnValue(id, "RFP_NO", RFP_SUMMARY);

    // Log the DELETE action in the audit trail
    logAuditTrail(
        "DELETE",
        "RFP_SUMMARY",
        id, // Unique ID of the record
        oldRecord, // Old values before deletion
        {}, // No new values for DELETE action
        [], // No changed fields for DELETE action
        "RFP Summary deleted"
    );

    return result;
}
