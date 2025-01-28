function selectColumnsFromSheet(sheetName, headers) {
    try {
        const sheet = new Utils.Sheet(sheetName, {
            row: {
                start: 1,
            },
        });
        const values = sheet.getValuesByColumns(...headers);

        return JSON.stringify(values);
    } catch (error) {
        return Utils.ErrorHandler(error, {
            arguments,
            value: [],
        });
    }
}

function insertDataToSimPlans(rowData) {
    try {
        const sheet = new Utils.Sheet("SIM_PLANS", {
            row: {
                start: 1,
            },
        });

        const planIds = sheet.getValuesByColumn("PLAN_ID");

        const ids = planIds.map(([planId]) => Utils.toNumber(planId));
        const maxId = Math.max(...ids);
        const currentId = maxId + 1;

        const sheetHeaders = sheet.getHeaders();
        const template = sheet.getTemplate();
        const row = new Utils.Row(template, sheetHeaders);

        row.setMany({
            ["PLAN_ID"]: currentId,
            ...rowData,
        });

        const data = row.toArray();

        return sheet.insert(data);
    } catch (error) {
        return Utils.ErrorHandler(error, {
            arguments,
        });
    }
}

function editDataToSimPlans(query, rowData) {
    try {
        const sheet = new Utils.Sheet("SIM_PLANS", {
            row: {
                start: 1,
            },
        });

        return sheet.experimentalQuery().findOneAndUpdate(query, rowData);
    } catch (error) {
        return Utils.ErrorHandler(error, {
            arguments,
        });
    }
}

function editDataToEmployeeDetails(query, rowData) {
    try {
        const sheet = new Utils.Sheet("EMPLOYEE_DETAILS", {
            row: {
                start: 1,
            },
        });

        return sheet.experimentalQuery().findOneAndUpdate(query, rowData);
    } catch (error) {
        return Utils.ErrorHandler(error, {
            arguments,
        });
    }
}

function deleteDataToSimPlans(query) {
    try {
        const sheet = new Utils.Sheet("SIM_PLANS", {
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

function isGroupIdExisting(groupId) {
    try {
        const sheet = new Utils.Sheet("EMPLOYEE_DETAILS", {
            row: {
                start: 1,
            },
        });

        const { match } = sheet.findOne({
            GROUP_ID: {
                $exec: (value) => value == groupId,
            },
        });

        return match;
    } catch (error) {
        return Utils.ErrorHandler(error, {
            arguments,
        });
    }
}

function isEmployeeIdExisting(employeeId) {
    try {
        const sheet = new Utils.Sheet("EMPLOYEE_DETAILS", {
            row: {
                start: 1,
            },
        });

        const { match } = sheet.findOne({
            EMPLOYEE_ID: {
                $exec: (value) => value == employeeId,
            },
        });

        return match;
    } catch (error) {
        return Utils.ErrorHandler(error, {
            arguments,
        });
    }
}

function isImeiExisting(imei) {
    try {
        const sheet = new Utils.Sheet("PHONE_INVENTORY", {
            row: {
                start: 1,
            },
        });

        const { match } = sheet.findOne({
            IMEI: {
                $exec: (value) => value == imei,
            },
        });

        return match;
    } catch (error) {
        return Utils.ErrorHandler(error, {
            arguments,
        });
    }
}

function insertDataToEmployeeDetails(rowData) {
    try {
        const sheet = new Utils.Sheet("EMPLOYEE_DETAILS", {
            row: {
                start: 1,
            },
        });

        const groupIds = sheet.getValuesByColumn("GROUP_ID");

        const ids = groupIds.map(([groupId]) => Utils.toNumber(groupId));
        const maxId = Math.max(...ids);
        const currentId = maxId + 1;

        if (isGroupIdExisting(rowData.GROUP_ID)) {
            throw new Error(
                "GROUP_ID already exists. Cannot insert duplicate GROUP_ID."
            );
        }

        if (isEmployeeIdExisting(rowData.EMPLOYEE_ID)) {
            throw new Error(
                "EMPLOYEE_ID already exists. Cannot insert duplicate EMPLOYEE_ID."
            );
        }

        const sheetHeaders = sheet.getHeaders();
        const template = sheet.getTemplate();
        const row = new Utils.Row(template, sheetHeaders);

        row.setMany({
            ["GROUP_ID"]: currentId,
            ...rowData,
        });

        const data = row.toArray();

        return sheet.insert(data);
    } catch (error) {
        return Utils.ErrorHandler(error, {
            arguments,
        });
    }
}

function deleteDataToEmployeeDetails(query) {
    try {
        const sheet = new Utils.Sheet("EMPLOYEE_DETAILS", {
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

function deleteDataToSimInventory(query) {
    try {
        const sheet = new Utils.Sheet("SIM_INVENTORY", {
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

function deleteDataToMobileInventory(query) {
    try {
        const sheet = new Utils.Sheet("PHONE_INVENTORY", {
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

function insertDataToSimRequest(rowData) {
    try {
        const sheet = new Utils.Sheet("SIM_REQUEST_AND_ISSUANCE", {
            row: {
                start: 1,
            },
        });

        const simSheet = new Utils.Sheet("SIM_INVENTORY", {
            row: {
                start: 1,
            },
        });

        const columns = sheet.getValuesByColumns("REQUEST_NO", "ISSUANCE_NO");

        const ids = columns.reduce(
            (acc, column) => {
                if (!acc.request_no.has(column[0])) {
                    acc.request_no.add(Utils.toNumber(column[0]));
                }
                if (!acc.issuance_no.has(column[1])) {
                    acc.issuance_no.add(Utils.toNumber(column[1]));
                }
                return acc;
            },
            {
                request_no: new Set(),
                issuance_no: new Set(),
            }
        );

        const max = {
            request_no: (() => {
                const requestNumbers = Array.from(ids.request_no);
                const maxRequestNumber = Math.max(...requestNumbers);
                return maxRequestNumber + 1;
            })(),
            issuance_no: (() => {
                const issuanceNumbers = Array.from(ids.issuance_no);
                const maxIssuanceNumber = Math.max(...issuanceNumbers);
                return maxIssuanceNumber + 1;
            })(),
        };

        const sheetHeaders = sheet.getHeaders();
        const template = sheet.getTemplate();
        const row = new Utils.Row(template, sheetHeaders);

        row.setMany({
            REQUEST_NO: max.request_no,
            ...(rowData.SIM_CARD_ID ? { ISSUANCE_NO: max.issuance_no } : {}),
            ...rowData,
        });

        const data = row.toArray();

        const simCardId = rowData.SIM_CARD_ID;
        const simCardIdObject = { SIM_CARD_ID: simCardId };
        const statusObject = { STATUS: "Reserved" };

        const result = simSheet.findOneAndUpdate(simCardIdObject, statusObject);

        if (result) {
            console.info("Data was updated successfully", result);
        } else {
            console.info("Data was not updated", result);
        }

        return sheet.insert(data);
    } catch (error) {
        return Utils.ErrorHandler(error, {
            arguments,
        });
    }
}

function editDataToSimRequest(query, rowData) {
    try {
        const sheet = new Utils.Sheet("SIM_REQUEST_AND_ISSUANCE", {
            row: {
                start: 1,
            },
        });

        const simSheet = new Utils.Sheet("SIM_INVENTORY", {
            row: {
                start: 1,
            },
        });

        const issuanceNoIds = sheet.getValuesByColumn("ISSUANCE_NO");

        const ids = issuanceNoIds.map(([issuanceNoId]) =>
            Utils.toNumber(issuanceNoId)
        );
        const maxId = Math.max(...ids);
        const simCardId = rowData.SIM_CARD_ID;
        const issuanceNo = rowData.ISSUANCE_NO;

        const currentId = issuanceNo ? issuanceNo : maxId + 1;

        const preIssuanceDate = rowData.PRE_ISSUANCE_DATE;
        const issuanceDate = rowData.ISSUANCE_DATE;
        const cancellationDate = rowData.CANCELLATION_DATE;
        const simCardIdObject = { SIM_CARD_ID: simCardId };
        const statusObject = issuanceDate
            ? { STATUS: "Issued" }
            : cancellationDate
            ? { STATUS: "Available" }
            : simCardId
            ? { STATUS: "Reserved" }
            : {};

        simSheet
            .experimentalQuery()
            .findOneAndUpdate(simCardIdObject, statusObject);

        console.log(rowData);

        return sheet.findOneAndUpdate(query, {
            ...rowData,
            ...(preIssuanceDate ? { REQUEST_STATUS: "Pre-Issued" } : {}),
            ...(cancellationDate ? { REQUEST_STATUS: "Cancelled" } : {}),
            ...(issuanceDate ? { REQUEST_STATUS: "Issued" } : {}),
            ...(rowData.SIM_CARD_ID ? { ISSUANCE_NO: currentId } : {}),
        });
    } catch (error) {
        return Utils.ErrorHandler(error, {
            arguments,
        });
    }
}

function editDataToMobileRequest(query, rowData) {
    try {
        const sheet = new Utils.Sheet("PHONE_REQUEST_AND_ISSUANCE", {
            row: {
                start: 1,
            },
        });

        const phoneSheet = new Utils.Sheet("PHONE_INVENTORY", {
            row: {
                start: 1,
            },
        });

        const issuanceNoIds = sheet.getValuesByColumn("ISSUANCE_NO");

        const ids = issuanceNoIds.map(([issuanceNoId]) =>
            Utils.toNumber(issuanceNoId)
        );
        const maxId = Math.max(...ids);
        const mobileId = rowData.MOBILE_ID;
        const issuanceNo = rowData.ISSUANCE_NO;
        const purchasingRsNo = rowData.PURCHASING_RS_NO;

        const currentId = issuanceNo ? issuanceNo : maxId + 1;

        const preIssuanceDate = rowData.PRE_ISSUANCE_DATE;
        const issuanceDate = rowData.ISSUANCE_DATE;
        const cancellationDate = rowData.CANCELLATION_DATE;
        const mobileIdObject = { MOBILE_ID: mobileId };
        const statusObject = issuanceDate
            ? { STATUS: "Issued" }
            : cancellationDate
            ? { STATUS: "Available" }
            : mobileId
            ? { STATUS: "Reserved" }
            : {};

        phoneSheet
            .experimentalQuery()
            .findOneAndUpdate(mobileIdObject, statusObject);

        console.log(rowData);

        return sheet.findOneAndUpdate(query, {
            ...rowData,
            ...(purchasingRsNo ? { REQUEST_STATUS: "Received PR" } : {}),
            ...(preIssuanceDate ? { REQUEST_STATUS: "Pre-Issued" } : {}),
            ...(cancellationDate ? { REQUEST_STATUS: "Cancelled" } : {}),
            ...(issuanceDate ? { REQUEST_STATUS: "Issued" } : {}),
            ...(rowData.MOBILE_ID ? { ISSUANCE_NO: currentId } : {}),
        });
    } catch (error) {
        return Utils.ErrorHandler(error, {
            arguments,
        });
    }
}

function editDataToSimInventory(query, rowData) {
    try {
        const sheet = new Utils.Sheet("SIM_INVENTORY", {
            row: {
                start: 1,
            },
        });

        return sheet.findOneAndUpdate(query, rowData);
    } catch (error) {
        return Utils.ErrorHandler(error, {
            arguments,
        });
    }
}

function editDataToMobileInventory(query, rowData) {
    try {
        const sheet = new Utils.Sheet("PHONE_INVENTORY", {
            row: {
                start: 1,
            },
        });

        return sheet.findOneAndUpdate(query, rowData);
    } catch (error) {
        return Utils.ErrorHandler(error, {
            arguments,
        });
    }
}

function insertDataToSimInventory(rowData) {
    try {
        const sheet = new Utils.Sheet("SIM_INVENTORY", {
            row: {
                start: 1,
            },
        });

        const simCardIds = sheet.getValuesByColumn("SIM_CARD_ID");

        const ids = simCardIds.map(([simCardId]) => Utils.toNumber(simCardId));
        const maxId = Math.max(...ids);
        const currentId = maxId + 1;

        const sheetHeaders = sheet.getHeaders();
        const template = sheet.getTemplate();
        const row = new Utils.Row(template, sheetHeaders);

        row.setMany({
            ["SIM_CARD_ID"]: currentId,
            ...rowData,
        });

        const data = row.toArray();

        return sheet.insert(data);
    } catch (error) {
        return Utils.ErrorHandler(error, {
            arguments,
        });
    }
}

function insertDataToMobileInventory(rowData) {
    try {
        const sheet = new Utils.Sheet("PHONE_INVENTORY", {
            row: {
                start: 1,
            },
        });

        const mobileDeviceIds = sheet.getValuesByColumn("MOBILE_ID");

        const ids = mobileDeviceIds.map(([mobileId]) =>
            Utils.toNumber(mobileId)
        );
        const maxId = Math.max(...ids);
        const currentId = maxId + 1;

        const sheetHeaders = sheet.getHeaders();
        const template = sheet.getTemplate();
        const row = new Utils.Row(template, sheetHeaders);

        row.setMany({
            ["MOBILE_ID"]: currentId,
            ...rowData,
        });

        const data = row.toArray();

        return sheet.insert(data);
    } catch (error) {
        return Utils.ErrorHandler(error, {
            arguments,
        });
    }
}

function isPlanIdExits(planId) {
    try {
        const sheet = new Utils.Sheet("SIM_PLANS", {
            row: {
                start: 1,
            },
        });

        const { match } = sheet.findOne({
            PLAN_ID: {
                $exec: (value) => value == planId,
            },
        });

        return match;
    } catch (error) {
        return Utils.ErrorHandler(error, {
            arguments,
        });
    }
}

function insertUploadSimInventoryFile(params = {}) {
    try {
        const Id = {
            MobileNo: "MOBILE_NO",
            PlanId: "PLAN_ID",
        };

        const sheet = new Utils.Sheet("SIM_INVENTORY", {
            row: {
                start: 1,
            },
            constraints: {
                SIM_CARD_ID: {
                    primary: true,
                },
            },
        });

        const planSheet = new Utils.Sheet("SIM_PLANS", {
            row: {
                start: 1,
            },
        });

        const mobileNumbers = sheet.getValuesByColumn(Id.MobileNo);
        const planIds = planSheet.getValuesByColumn(Id.PlanId);

        const simProps = sheet.getProperties("headers", "values", "template");
        const ids = sheet.getValuesByColumn("SIM_CARD_ID");
        const insertedItems = new Set(
            mobileNumbers.map(([item]) => Utils.toNumber(item))
        );
        const planIdInsertedItems = new Set(
            planIds.map(([item]) => Utils.toNumber(item))
        );

        let skippedRows = [];

        const upsertData = sheet.upsertWith(function (
            action,
            row,
            index,
            array
        ) {
            try {
                const source = new Utils.Row(row, params.headers);

                const mobile = Utils.toNumber(source.get(Id.MobileNo));
                const planId = Utils.toNumber(source.get(Id.PlanId));

                if (!mobile) {
                    throw new Error(`Mobile number cannot be empty`);
                }

                if (!planId) {
                    throw new Error(`Mobile number cannot be empty`);
                }
                // If existed, throw an Error
                if (insertedItems.has(mobile)) {
                    // Add skipped row info to skippedRows array
                    skippedRows.push({
                        mobileNo: mobile,
                        rowIndex: index + 1, // Adjust for 1-based index
                    });
                    action.skip(row); // Mark row as skipped
                } else if (!planIdInsertedItems.has(planId)) {
                    // Add skipped row info to skippedRows array
                    skippedRows.push({
                        planId: planId,
                        rowIndex: index + 1, // Adjust for 1-based index
                    });
                    action.skip(row); // Mark row as skipped
                } else {
                    const destination = new Utils.Row(
                        simProps.template,
                        simProps.headers
                    );
                    const maxId = Math.max(...ids);
                    const currentId = maxId + 1;

                    destination.setMany({
                        SIM_CARD_ID: currentId,
                        STATUS: "Available",
                        ...source.toObject(),
                    });
                    const destinationValues = destination.toArray();
                    action.insert(destinationValues);

                    ids.push([currentId]);

                    // If not inserted, add it to the Set.
                    insertedItems.add(mobile);
                    planIdInsertedItems.add(planId);
                }
            } catch (error) {
                return Utils.ErrorHandler(error, {
                    value: {},
                });
            }
        },
        params.values);

        return { ...upsertData, skippedRows };
    } catch (error) {
        return Utils.ErrorHandler(error, {
            arguments,
        });
    }
}

function insertDataToReturnedSim(rowData) {
    try {
        const sheet = new Utils.Sheet("RETURNED_SIM", {
            row: {
                start: 1,
            },
        });

        const requestSheet = new Utils.Sheet("SIM_REQUEST_AND_ISSUANCE", {
            row: {
                start: 1,
            },
        });

        const simSheet = new Utils.Sheet("SIM_INVENTORY", {
            row: {
                start: 1,
            },
        });

        const returnNoIds = sheet.getValuesByColumn("RETURN_NO");

        const ids = returnNoIds.map(([returnNoId]) =>
            Utils.toNumber(returnNoId)
        );
        const maxId = Math.max(...ids);
        const currentId = maxId + 1;

        const sheetHeaders = sheet.getHeaders();
        const template = sheet.getTemplate();
        const row = new Utils.Row(template, sheetHeaders);

        row.setMany({
            ["RETURN_NO"]: currentId,
            ...rowData,
        });

        const data = row.toArray();

        const issuanceNo = rowData.ISSUANCE_NO;
        const issuanceNoObject = { ISSUANCE_NO: issuanceNo };
        const statusObject = { REQUEST_STATUS: "Returned" };

        requestSheet
            .experimentalQuery()
            .findOneAndUpdate(issuanceNoObject, statusObject);
        const findRequestDetails = requestSheet
            .experimentalQuery()
            .findOne(issuanceNoObject);
        const requestDetailsObject = findRequestDetails.toObject();
        const simCardId = requestDetailsObject.SIM_CARD_ID;
        const simCardIdObject = { SIM_CARD_ID: simCardId };
        const simStatusObject = { STATUS: "Available" };

        simSheet
            .experimentalQuery()
            .findOneAndUpdate(simCardIdObject, simStatusObject);

        return sheet.insert(data);
    } catch (error) {
        return Utils.ErrorHandler(error, {
            arguments,
        });
    }
}

function insertDataToReturnedMobile(rowData) {
    try {
        const sheet = new Utils.Sheet("RETURNED_PHONE", {
            row: {
                start: 1,
            },
        });

        const requestSheet = new Utils.Sheet("PHONE_REQUEST_AND_ISSUANCE", {
            row: {
                start: 1,
            },
        });

        const phoneInventorySheet = new Utils.Sheet("PHONE_INVENTORY", {
            row: {
                start: 1,
            },
        });

        const returnNoIds = sheet.getValuesByColumn("RETURN_NO");

        const ids = returnNoIds.map(([returnNoId]) =>
            Utils.toNumber(returnNoId)
        );
        const maxId = Math.max(...ids);
        const currentId = maxId + 1;

        const sheetHeaders = sheet.getHeaders();
        const template = sheet.getTemplate();
        const row = new Utils.Row(template, sheetHeaders);

        row.setMany({
            ["RETURN_NO"]: currentId,
            ...rowData,
        });

        const data = row.toArray();

        const issuanceNo = rowData.ISSUANCE_NO;
        const issuanceNoObject = { ISSUANCE_NO: issuanceNo };
        const statusObject = { REQUEST_STATUS: "Returned" };

        requestSheet
            .experimentalQuery()
            .findOneAndUpdate(issuanceNoObject, statusObject);
        const findRequestDetails = requestSheet
            .experimentalQuery()
            .findOne(issuanceNoObject);
        const requestDetailsObject = findRequestDetails.toObject();
        const mobileId = requestDetailsObject.MOBILE_ID;
        const mobileIdObject = { MOBILE_ID: mobileId };
        const mobileStatusObject = { STATUS: "Available" };

        phoneInventorySheet
            .experimentalQuery()
            .findOneAndUpdate(mobileIdObject, mobileStatusObject);

        return sheet.insert(data);
    } catch (error) {
        return Utils.ErrorHandler(error, {
            arguments,
        });
    }
}

function getGoogleSheetHeaders() {
    const sheet = SpreadsheetApp.openById(
        "1MLOQ4u7BDcU5_NBK91PFfUIPeiO1nyJy3unfiv2lz4I"
    ).getActiveSheet();
    return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

function submitDataToSheet(data) {
    const sheet = SpreadsheetApp.openById(
        "1MLOQ4u7BDcU5_NBK91PFfUIPeiO1nyJy3unfiv2lz4I"
    ).getActiveSheet();
    data.forEach((row) => {
        sheet.appendRow(Object.values(row));
    });
}

function getConsolidatedSimInventory(headers = ["*"]) {
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
        const simInventorySheet = new Utils.Sheet("SIM_INVENTORY", {
            row: {
                start: 1,
            },
        });
        const simPlansSheet = new Utils.Sheet("SIM_PLANS", {
            row: {
                start: 1,
            },
        });
        const requestSheet = new Utils.Sheet("SIM_REQUEST_AND_ISSUANCE", {
            row: {
                start: 1,
            },
        });
        const rfpSheet = new Utils.Sheet("RFP_GROUP", {
            row: {
                start: 1,
            },
        });
        const ticketSheet = new Utils.Sheet("TICKET_MANAGEMENT", {
            row: {
                start: 1,
            },
        });

        const simInventoryDataset = simInventorySheet.toObject();
        const simPlansDataset = simPlansSheet.toObject();
        const simRequestDataset = requestSheet.toObject();
        const rfpGroupDataset = rfpSheet.toObject();
        const ticketDataset = ticketSheet.toObject();

        const firstSet = Utils.sql(
            ` 
           SELECT SIM_CARD_ID, EMPLOYEE_INFO
            FROM ?
            WHERE REQUEST_STATUS = 'Received RS' OR REQUEST_STATUS = 'Issued' OR REQUEST_STATUS = 'Pre-Issued' OR REQUEST_STATUS = 'Received PR'
          `,
            [simRequestDataset]
        );

        const latestDates = Utils.sql(
            `
            SELECT SIM_CARD_ID, MAX(DATE_RECEIVED) AS DATE_RECEIVED
            FROM ?
            WHERE SIM_CARD_ID IS NOT NULL AND SIM_CARD_ID != ""
            GROUP BY SIM_CARD_ID
          `,
            [ticketDataset]
        );
        const secondSet = Utils.sql(
            `
            SELECT a.SIM_CARD_ID, a.STATUS as TICKET_STATUS, a.DATE_RECEIVED, a.TICKET_NO
            FROM ? AS a
            JOIN ? AS b
            ON a.SIM_CARD_ID = b.SIM_CARD_ID AND a.DATE_RECEIVED = b.DATE_RECEIVED
            ORDER BY a.DATE_RECEIVED DESC
          `,
            [ticketDataset, latestDates]
        );

        const execution = Utils.sql(
            ` 
          SELECT DISTINCT i.PLAN_ID, ${toMatrixHeaders(
              "i",
              headers
          )}, t.TICKET_STATUS, t.TICKET_NO, p.PLAN_DETAILS, p.CATEGORY, p.NETWORK_PROVIDER, p.PROVIDER_COMPANY_NAME, p.MONTHLY_RECURRING_FEE, s.EMPLOYEE_INFO, r.RFP_GROUP_NAME, r.RFP_COMPANY
          FROM ? AS i
          LEFT JOIN ? AS p ON i.PLAN_ID = p.PLAN_ID
          LEFT JOIN ? AS s ON i.SIM_CARD_ID = s.SIM_CARD_ID
          LEFT JOIN ? AS r ON i.RFP_GROUP_ID = r.RFP_GROUP_ID
          LEFT JOIN ? AS t ON i.SIM_CARD_ID = t.SIM_CARD_ID
        `,
            [
                simInventoryDataset,
                simPlansDataset,
                firstSet,
                rfpGroupDataset,
                secondSet,
            ]
        );

        const parsedResult = execution.map((row) => {
            // Parse EMPLOYEE_INFO if it exists
            if (row.EMPLOYEE_INFO) {
                try {
                    const employeeInfo = JSON.parse(row.EMPLOYEE_INFO);
                    // Add individual properties to the row
                    row.ACTIVE_USER = employeeInfo.FULL_NAME || "";
                } catch (error) {
                    console.error("Error parsing EMPLOYEE_INFO:", error);
                    row.ACTIVE_USER = "";
                }
            } else {
                // Handle the case where EMPLOYEE_INFO is empty
                row.ACTIVE_USER = "";
            }
            return row;
        });

        // Log the processed data and return as a JSON string
        console.log(JSON.stringify(secondSet));
        return JSON.stringify(parsedResult);
    } catch (error) {
        return Utils.ErrorHandler(error, {
            arguments,
            value: [],
        });
    }
}

function getConsolidatedMobilePhoneInventory(headers = ["*"]) {
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
        const mobileInventorySheet = new Utils.Sheet("PHONE_INVENTORY", {
            row: {
                start: 1,
            },
        });
        const requestSheet = new Utils.Sheet("PHONE_REQUEST_AND_ISSUANCE", {
            row: {
                start: 1,
            },
        });
        const ticketSheet = new Utils.Sheet("TICKET_MANAGEMENT", {
            row: {
                start: 1,
            },
        });

        const mobileInventoryDataset = mobileInventorySheet.toObject();
        const mobileRequestDataset = requestSheet.toObject();
        const ticketDataset = ticketSheet.toObject();

        const firstSet = Utils.sql(
            ` 
           SELECT MOBILE_ID, EMPLOYEE_INFO
            FROM ?
            WHERE REQUEST_STATUS = 'Received RS' OR REQUEST_STATUS = 'Issued' OR REQUEST_STATUS = 'Pre-Issued' OR REQUEST_STATUS = 'Received PR'
          `,
            [mobileRequestDataset]
        );

        const latestDates = Utils.sql(
            `
            SELECT MOBILE_ID, MAX(DATE_RECEIVED) AS DATE_RECEIVED
            FROM ?
            WHERE MOBILE_ID IS NOT NULL AND MOBILE_ID != ""
            GROUP BY MOBILE_ID
          `,
            [ticketDataset]
        );
        const secondSet = Utils.sql(
            `
            SELECT a.MOBILE_ID, a.STATUS as TICKET_STATUS, a.DATE_RECEIVED, a.TICKET_NO
            FROM ? AS a
            JOIN ? AS b
            ON a.MOBILE_ID = b.MOBILE_ID AND a.DATE_RECEIVED = b.DATE_RECEIVED
            ORDER BY a.DATE_RECEIVED DESC
          `,
            [ticketDataset, latestDates]
        );

        const execution = Utils.sql(
            ` 
          SELECT DISTINCT i.MOBILE_ID, ${toMatrixHeaders(
              "i",
              headers
          )}, t.TICKET_STATUS, t.TICKET_NO, s.EMPLOYEE_INFO
          FROM ? AS i
          LEFT JOIN ? AS s ON i.MOBILE_ID = s.MOBILE_ID
          LEFT JOIN ? AS t ON i.MOBILE_ID = t.MOBILE_ID
        `,
            [mobileInventoryDataset, firstSet, secondSet]
        );

        const parsedResult = execution.map((row) => {
            // Parse EMPLOYEE_INFO if it exists
            if (row.EMPLOYEE_INFO) {
                try {
                    const employeeInfo = JSON.parse(row.EMPLOYEE_INFO);
                    // Add individual properties to the row
                    row.ACTIVE_USER = employeeInfo.FULL_NAME || "";
                } catch (error) {
                    console.error("Error parsing EMPLOYEE_INFO:", error);
                    row.ACTIVE_USER = "";
                }
            } else {
                // Handle the case where EMPLOYEE_INFO is empty
                row.ACTIVE_USER = "";
            }
            return row;
        });

        // Log the processed data and return as a JSON string
        console.log(JSON.stringify(secondSet));
        return JSON.stringify(parsedResult);
    } catch (error) {
        return Utils.ErrorHandler(error, {
            arguments,
            value: [],
        });
    }
}

function getConsolidatedSimDetails(headers = []) {
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
        const simInventorySheet = new Utils.Sheet("SIM_INVENTORY", {
            row: {
                start: 1,
            },
        });
        const simPlansSheet = new Utils.Sheet("SIM_PLANS", {
            row: {
                start: 1,
            },
        });
        const issuanceSheet = new Utils.Sheet("SIM_REQUEST_AND_ISSUANCE", {
            row: {
                start: 1,
            },
        });
        const employeeSheet = new Utils.Sheet("EMPLOYEE_DETAILS", {
            row: {
                start: 1,
            },
        });
        const ticketSheet = new Utils.Sheet("TICKET_MANAGEMENT", {
            row: {
                start: 1,
            },
        });
        const simInventoryDataset = simInventorySheet.toObject();
        const simPlansDataset = simPlansSheet.toObject();
        const issuanceDataset = issuanceSheet.toObject();
        const employeeDataset = employeeSheet.toObject();
        const ticketDataset = ticketSheet.toObject();

        const firstSet = Utils.sql(
            ` 
           SELECT SIM_CARD_ID, OWNERSHIP, USAGE_TYPE, GROUP_ID
            FROM ?
            WHERE REQUEST_STATUS = 'Returned'
              AND ISSUANCE_NO IN (
                SELECT MAX(ISSUANCE_NO)
                FROM ?
                WHERE REQUEST_STATUS = 'Returned'
                GROUP BY SIM_CARD_ID
              )
          `,
            [issuanceDataset, issuanceDataset]
        );

        const latestDates = Utils.sql(
            `
            SELECT SIM_CARD_ID, MAX(DATE_RECEIVED) AS DATE_RECEIVED
            FROM ?
            WHERE SIM_CARD_ID IS NOT NULL AND SIM_CARD_ID != ""
            GROUP BY SIM_CARD_ID
          `,
            [ticketDataset]
        );
        const secondSet = Utils.sql(
            `
            SELECT a.SIM_CARD_ID, a.STATUS as TICKET_STATUS, a.DATE_RECEIVED, a.TICKET_NO
            FROM ? AS a
            JOIN ? AS b
            ON a.SIM_CARD_ID = b.SIM_CARD_ID AND a.DATE_RECEIVED = b.DATE_RECEIVED
            ORDER BY a.DATE_RECEIVED DESC
          `,
            [ticketDataset, latestDates]
        );

        const execution = Utils.sql(
            ` 
        SELECT
          ${toMatrixHeaders("i", headers)}, 
          p.PLAN_DETAILS, 
          p.CATEGORY,
          p.NETWORK_PROVIDER, 
          p.MONTHLY_RECURRING_FEE,
          s.OWNERSHIP,
          s.USAGE_TYPE,
          e.FULL_NAME, 
          t.TICKET_STATUS
        FROM ? AS i
        LEFT JOIN ? AS p
        ON i.PLAN_ID = p.PLAN_ID
        LEFT JOIN ? AS s 
        ON i.SIM_CARD_ID = s.SIM_CARD_ID 
        LEFT JOIN ? AS e 
        ON s.GROUP_ID = e.GROUP_ID
        LEFT JOIN ? AS t 
        ON i.SIM_CARD_ID = t.SIM_CARD_ID
        WHERE i.STATUS = 'Available' AND ( t.TICKET_STATUS = 'Resolved' OR t.TICKET_STATUS = '' OR t.TICKET_STATUS IS NULL )
        `,
            [
                simInventoryDataset,
                simPlansDataset,
                firstSet,
                employeeDataset,
                secondSet,
            ]
        );

        console.log(JSON.stringify(execution));
        return JSON.stringify(execution);
    } catch (error) {
        return Utils.ErrorHandler(error, {
            arguments,
            value: [],
        });
    }
}

function getConsolidatedReturnSim(headers = ["*"]) {
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
        const returnSimSheet = new Utils.Sheet("RETURNED_SIM", {
            row: {
                start: 1,
            },
        });
        const simRequestSheet = new Utils.Sheet("SIM_REQUEST_AND_ISSUANCE", {
            row: {
                start: 1,
            },
        });
        const returnSimDataset = returnSimSheet.toObject();
        const simRequestDataset = simRequestSheet.toObject();
        const execution = Utils.sql(
            ` 
          SELECT ${toMatrixHeaders(
              "r",
              headers
          )}, s.ISSUANCE_DATE, s.USAGE_TYPE, s.GROUP_ID, s.EMPLOYEE_INFO, s.SIM_INFO 
          FROM ? AS r
          LEFT JOIN ? AS s
          ON r.ISSUANCE_NO = s.ISSUANCE_NO
        `,
            [returnSimDataset, simRequestDataset]
        );
        console.log(JSON.stringify(execution));
        return JSON.stringify(execution);
    } catch (error) {
        return Utils.ErrorHandler(error, {
            arguments,
            value: [],
        });
    }
}

function getConsolidatedReturnMobile(headers = ["*"]) {
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
        const returnMobileSheet = new Utils.Sheet("RETURNED_PHONE", {
            row: {
                start: 1,
            },
        });
        const mobileRequestSheet = new Utils.Sheet(
            "PHONE_REQUEST_AND_ISSUANCE",
            {
                row: {
                    start: 1,
                },
            }
        );
        const returnMobileDataset = returnMobileSheet.toObject();
        const mobileRequestDataset = mobileRequestSheet.toObject();
        const execution = Utils.sql(
            ` 
          SELECT ${toMatrixHeaders(
              "r",
              headers
          )}, s.ISSUANCE_DATE, s.USAGE_TYPE, s.GROUP_ID, s.EMPLOYEE_INFO, s.EARPHONES, s.CHARGER, s.OTHERS, s.MOBILE_INFO, s.MOBILE_ID
          FROM ? AS r
          LEFT JOIN ? AS s
          ON r.ISSUANCE_NO = s.ISSUANCE_NO
        `,
            [returnMobileDataset, mobileRequestDataset]
        );
        console.log(JSON.stringify(execution));
        return JSON.stringify(execution);
    } catch (error) {
        return Utils.ErrorHandler(error, {
            arguments,
            value: [],
        });
    }
}

function getConsolidatedReturnRequestSim() {
    try {
        const simRequestSheet = new Utils.Sheet("SIM_REQUEST_AND_ISSUANCE", {
            row: {
                start: 1,
            },
        });
        const simRequestDataset = simRequestSheet.toObject();
        const execution = Utils.sql(
            ` 
          SELECT ISSUANCE_NO,ISSUANCE_DATE,SIM_CARD_ID,GROUP_ID,EMPLOYEE_INFO,USAGE_TYPE,SIM_INFO,REQUEST_STATUS
          FROM ? 
          where REQUEST_STATUS = 'Issued'
          ORDER BY ISSUANCE_NO
        `,
            [simRequestDataset]
        );
        console.log(JSON.stringify(execution));
        return JSON.stringify(execution);
    } catch (error) {
        return Utils.ErrorHandler(error, {
            arguments,
            value: [],
        });
    }
}

function getConsolidatedReturnRequestMobile() {
    try {
        const simRequestSheet = new Utils.Sheet("PHONE_REQUEST_AND_ISSUANCE", {
            row: {
                start: 1,
            },
        });
        const simRequestDataset = simRequestSheet.toObject();
        const execution = Utils.sql(
            `
          SELECT ISSUANCE_NO,ISSUANCE_DATE,MOBILE_ID,GROUP_ID,EMPLOYEE_INFO,MOBILE_INFO, EARPHONES, CHARGER, OTHERS ,REQUEST_STATUS
          FROM ? 
          where REQUEST_STATUS = 'Issued'
          ORDER BY ISSUANCE_NO
        `,
            [simRequestDataset]
        );
        console.log(JSON.stringify(execution));
        return JSON.stringify(execution);
    } catch (error) {
        return Utils.ErrorHandler(error, {
            arguments,
            value: [],
        });
    }
}

function getConsolidatedEmployeewithSim(headers = ["*"]) {
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
        const employeeSheet = new Utils.Sheet("EMPLOYEE_DETAILS", {
            row: {
                start: 1,
            },
        });
        const simRequestSheet = new Utils.Sheet("SIM_REQUEST_AND_ISSUANCE", {
            row: {
                start: 1,
            },
        });
        const inventorySheet = new Utils.Sheet("SIM_INVENTORY", {
            row: {
                start: 1,
            },
        });

        const simInventoryDataset = inventorySheet.toObject();
        const employeeDataset = employeeSheet.toObject();
        const simRequestDataset = simRequestSheet.toObject();
        const execution = Utils.sql(
            ` 
          SELECT ${toMatrixHeaders("e", headers)}, r.SIM_CARD_ID, i.MOBILE_NO
          FROM ? AS e
          LEFT JOIN ? AS r
          ON e.GROUP_ID = r.GROUP_ID
          LEFT JOIN ? AS i
          ON r.SIM_CARD_ID = i.SIM_CARD_ID
        `,
            [employeeDataset, simRequestDataset, simInventoryDataset]
        );
        console.log(JSON.stringify(execution));
        return JSON.stringify(execution);
    } catch (error) {
        return Utils.ErrorHandler(error, {
            arguments,
            value: [],
        });
    }
}

function insertDataToMobileRequest(rowData) {
    try {
        const sheet = new Utils.Sheet("PHONE_REQUEST_AND_ISSUANCE", {
            row: {
                start: 1,
            },
        });

        const phoneSheet = new Utils.Sheet("PHONE_INVENTORY", {
            row: {
                start: 1,
            },
        });

        const columns = sheet.getValuesByColumns("REQUEST_NO", "ISSUANCE_NO");

        const ids = columns.reduce(
            (acc, column) => {
                if (!acc.request_no.has(column[0])) {
                    acc.request_no.add(Utils.toNumber(column[0]));
                }
                if (!acc.issuance_no.has(column[1])) {
                    acc.issuance_no.add(Utils.toNumber(column[1]));
                }
                return acc;
            },
            {
                request_no: new Set(),
                issuance_no: new Set(),
            }
        );

        const max = {
            request_no: (() => {
                const requestNumbers = Array.from(ids.request_no);
                const maxRequestNumber = Math.max(...requestNumbers);
                return maxRequestNumber + 1;
            })(),
            issuance_no: (() => {
                const issuanceNumbers = Array.from(ids.issuance_no);
                const maxIssuanceNumber = Math.max(...issuanceNumbers);
                return maxIssuanceNumber + 1;
            })(),
        };

        const sheetHeaders = sheet.getHeaders();
        const template = sheet.getTemplate();
        const row = new Utils.Row(template, sheetHeaders);

        row.setMany({
            REQUEST_NO: max.request_no,
            ...(rowData.MOBILE_ID ? { ISSUANCE_NO: max.issuance_no } : {}),
            ...rowData,
        });

        const data = row.toArray();

        const mobileId = rowData.MOBILE_ID;
        const mobileIdObject = { MOBILE_ID: mobileId };
        const statusObject = { STATUS: "Reserved" };

        const result = phoneSheet.findOneAndUpdate(
            mobileIdObject,
            statusObject
        );

        if (result) {
            console.info("Data was updated successfully", result);
        } else {
            console.info("Data was not updated", result);
        }

        return sheet.insert(data);
    } catch (error) {
        return Utils.ErrorHandler(error, {
            arguments,
        });
    }
}

function getConsolidatedMobileInventory(headers = ["*"]) {
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
        const phoneInventorySheet = new Utils.Sheet("PHONE_INVENTORY", {
            row: {
                start: 1,
            },
        });
        const ticketSheet = new Utils.Sheet("TICKET_MANAGEMENT", {
            row: {
                start: 1,
            },
        });

        const phoneDataset = phoneInventorySheet.toObject();
        const ticketDataset = ticketSheet.toObject();

        const latestDates = Utils.sql(
            `
            SELECT MOBILE_ID, MAX(DATE_RECEIVED) AS DATE_RECEIVED
            FROM ?
            WHERE MOBILE_ID IS NOT NULL AND MOBILE_ID != ""
            GROUP BY MOBILE_ID
          `,
            [ticketDataset]
        );
        const secondSet = Utils.sql(
            `
            SELECT a.MOBILE_ID, a.STATUS as TICKET_STATUS, a.DATE_RECEIVED, a.TICKET_NO
            FROM ? AS a
            JOIN ? AS b
            ON a.MOBILE_ID = b.MOBILE_ID AND a.DATE_RECEIVED = b.DATE_RECEIVED
            ORDER BY a.DATE_RECEIVED DESC
          `,
            [ticketDataset, latestDates]
        );

        const execution = Utils.sql(
            ` 
          SELECT ${toMatrixHeaders("r", headers)}
          FROM ? AS r
          LEFT JOIN ? AS t
          ON r.MOBILE_ID = t.MOBILE_ID
          WHERE r.STATUS = 'Available' AND ( t.TICKET_STATUS = 'Resolved' OR t.TICKET_STATUS = '' OR t.TICKET_STATUS IS NULL )
        `,
            [phoneDataset, secondSet]
        );

        return JSON.stringify(execution);
    } catch (error) {
        return Utils.ErrorHandler(error, {
            arguments,
            value: [],
        });
    }
}
