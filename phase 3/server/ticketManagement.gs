function getTicketManagementData() {
    // const employeeDetailsSheet = new Utils.Sheet("EMPLOYEE_DETAILS", { row: { start: 1 } });
    // const simInventorySheet = new Utils.Sheet("SIM_INVENTORY", { row: { start: 1 } });
    // const mobileInventorySheet = new Utils.Sheet("MOBILE_INVENTORY", { row: { start: 1 } });

    // const employeeDetailsDataSet = employeeDetailsSheet.toObject();
    // const simInventoryDataSet = simInventorySheet.toObject();
    // const mobileInventoryDataSet = mobileInventorySheet.toObject();

    const ticketManagementSheet = new Utils.Sheet("TICKET_MANAGEMENT", {
        row: { start: 1 },
    });

    let ticketManagementDataSet = ticketManagementSheet.toObject();

    if (ticketManagementDataSet.length === 0) {
        ticketManagementDataSet = [
            {
                TICKET_ID: "",
                TICKET_NO: "",
                STATUS: "",
                DATE_RECEIVED: "",
                CREATED_AT: "",
                GROUP_ID: "",
                DEVICE_TYPE: "",
                SIM_CARD_ID: "",
                MOBILE_ID: "",
                DESCRIPTION: "",
                PRIORITY: "",
                ASSIGNED_PIC: "",
                PLAN_OF_ACTION: "",
                DATE_RESOLVED: "",
                RESOLUTION_DETAILS: "",
            },
        ];
    }

    console.log(ticketManagementDataSet.length);
    console.log(ticketManagementDataSet);
    return JSON.stringify(ticketManagementDataSet);
}

function getDevices() {
    const simInventorySheet = new Utils.Sheet("SIM_INVENTORY", {
        row: { start: 1 },
    });
    const simRequestSheet = new Utils.Sheet("SIM_REQUEST_AND_ISSUANCE", {
        row: { start: 1 },
    });
    const phoneInventorySheet = new Utils.Sheet("PHONE_INVENTORY", {
        row: { start: 1 },
    });
    const phoneRequestSheet = new Utils.Sheet("PHONE_REQUEST_AND_ISSUANCE", {
        row: { start: 1 },
    });

    // console.log(phoneInventorySheet)
    // console.log(phoneRequestSheet)

    const simInventoryDataset = simInventorySheet.toObject();
    const simRequestDataset = simRequestSheet.toObject();
    const phoneInventoryDataset = phoneInventorySheet.toObject();
    const phoneRequestDataset = phoneRequestSheet.toObject();

    // console.log(phoneInventoryDataset)
    // console.log(phoneRequestDataset)

    // Create a map of SIM_CARD_IDs from SIM_REQUEST_AND_ISSUANCE
    const simInventoryData = new Map(
        simRequestDataset.map((simRequest) => [
            simRequest.SIM_CARD_ID,
            simRequest,
        ])
    );

    // Create a map of MOBILE_IDs from PHONE_REQUEST_AND_ISSUANCE
    const phoneInventoryData = new Map(
        phoneRequestDataset.map((phoneRequest) => [
            phoneRequest.MOBILE_ID,
            phoneRequest,
        ])
    );

    // Process data
    const result = {
        SIM: simInventoryDataset.map((simInventory) => {
            const simRequest = simInventoryData.get(simInventory.SIM_CARD_ID);

            if (simRequest) {
                // SIM card has been issued
                const employeeInfo = JSON.parse(simRequest.EMPLOYEE_INFO);
                const simInfo = JSON.parse(simRequest.SIM_INFO);

                return {
                    SIM_CARD_ID: simInventory.SIM_CARD_ID,
                    MOBILE_NO: simInfo.MOBILE_NO,
                    NETWORK_PROVIDER: simInfo.NETWORK_PROVIDER,
                    PLAN_DETAILS: simInfo.PLAN_DETAILS,
                    STATUS: simInventory.STATUS,
                    EMPLOYEE_NAME: employeeInfo.FULL_NAME,
                    COMPANY_NAME: employeeInfo.COMPANY_NAME,
                    GROUP_ID: simRequest.GROUP_ID,
                    ACCOUNT_NO: simInfo.ACCOUNT_NO,
                };
            } else {
                // SIM card has not been issued
                return {
                    SIM_CARD_ID: simInventory.SIM_CARD_ID,
                    MOBILE_NO: simInventory.MOBILE_NO.toString(),
                    NETWORK_PROVIDER: null,
                    PLAN_DETAILS: null,
                    STATUS: simInventory.STATUS,
                    EMPLOYEE_NAME: null,
                    COMPANY_NAME: null,
                    GROUP_ID: null,
                    ACCOUNT_NO: simInventory.ACCOUNT_NO,
                };
            }
        }),
        PHONE: phoneInventoryDataset.map((phoneInventory) => {
            const phoneRequest = phoneInventoryData.get(
                phoneInventory.MOBILE_ID
            );

            if (phoneRequest) {
                // Phone has been issued
                const employeeInfo = JSON.parse(phoneRequest.EMPLOYEE_INFO);
                const mobileInfo = JSON.parse(phoneRequest.MOBILE_INFO);

                return {
                    MOBILE_ID: phoneInventory.MOBILE_ID,
                    IMEI: mobileInfo.IMEI,
                    BRAND: mobileInfo.BRAND,
                    MODEL: mobileInfo.MODEL,
                    RAM_ROM: phoneInventory.RAM_ROM,
                    CAMERA: phoneInventory.CAMERA,
                    COLOR: phoneInventory.COLOR,

                    STATUS: phoneInventory.STATUS,
                    EMPLOYEE_NAME: employeeInfo.FULL_NAME,
                    COMPANY_NAME: employeeInfo.COMPANY_NAME,
                    GROUP_ID: phoneRequest.GROUP_ID,
                };
            } else {
                // Phone has not been issued
                return {
                    MOBILE_ID: phoneInventory.MOBILE_ID,
                    IMEI: phoneInventory.IMEI,
                    BRAND: phoneInventory.BRAND,
                    MODEL: phoneInventory.MODEL,
                    RAM_ROM: phoneInventory.RAM_ROM,
                    CAMERA: phoneInventory.CAMERA,
                    COLOR: phoneInventory.COLOR,
                    STATUS: phoneInventory.STATUS,
                    EMPLOYEE_NAME: null,
                    COMPANY_NAME: null,
                    GROUP_ID: null,
                };
            }
        }),
    };

    console.log(result);

    return result;
}

function addTicketRecord(formData) {
    const sheet = accessSheet("TICKET_MANAGEMENT");
    const lastRow = sheet.getLastRow();
    // const lastId = sheet.getRange(lastRow, 1).getValue();

    console.log(formData);

    const lastId =
        lastRow > 1 ? sheet.getRange(lastRow, 1).getValue() : "TICKET_ID";

    const ticketId = lastId === "TICKET_ID" ? 1 : Number(lastId) + 1;
    const ticketNo = `TN-${ticketId}`;
    const dateCreated = new Date().toLocaleDateString("en-US");

    const deviceType = formData.deviceType;

    // Default values
    const status = "Pending";
    let data = [];

    if (deviceType == "SIM") {
        data = [
            ticketId,
            ticketNo,
            status,
            formData.dateReceived,
            dateCreated,
            formData.group_id,
            formData.deviceType,
            formData.sim_card_id,
            "",
            formData.ticketDescriptionId,
            formData.priority,
            formData.assignedPic,
            formData.planOfAction,
        ];
    } else if (deviceType == "PHONE") {
        data = [
            ticketId,
            ticketNo,
            status,
            formData.dateReceived,
            dateCreated,
            formData.group_id,
            formData.deviceType,
            "",
            formData.mobile_id,
            formData.ticketDescriptionId,
            formData.priority,
            formData.assignedPic,
            formData.planOfAction,
        ];
    } else {
        throw new Error("Invalid device type. Must be 'SIM' or 'PHONE'.");
    }

    const newValue = {
        TICKET_ID: ticketId,
        TICKET_NO: ticketNo,
        STATUS: status,
        DATE_RECEIVED: formData.dateReceived,
        CREATED_AT: dateCreated,
        GROUP_ID: formData.group_id,
        DEVICE_TYPE: formData.deviceType,
        SIM_CARD_ID: formData.sim_card_id || "",
        MOBILE_ID: formData.mobile_id || "",
        DESCRIPTION: formData.ticketDescriptionId,
        PRIORITY: formData.priority,
        ASSIGNED_PIC: formData.assignedPic,
        PLAN_OF_ACTION: formData.planOfAction,
    };

    sheet.appendRow(data);

    // Log the insertion to the audit trail
    logAuditTrail(
        "ADD",
        "TICKET_MANAGEMENT",
        ticketId,
        {}, // No old value for inserts
        newValue,
        // Object.keys(newValue), // Fields being added
        [],
        "Initial record creation" // Remarks
    );
}

// function editTicketRecord(formData) {
//   const sheet = accessSheet('TICKET_MANAGEMENT');
//   const id = formData.editTicketId;
//   const row = findRowById('TICKET_MANAGEMENT', id);

//   if (row === -1) {
//     throw new Error(`Record with ID ${id} not found.`);
//   }

//   // Get old values for the record
//   const oldValues = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

//   console.log(oldValues)

//   // Formats the date fields
//   oldValues[3] = formatDate(oldValues[3]);
//   oldValues[4] = formatDate(oldValues[4]);
//   oldValues[13] = formatDate(oldValues[13]);

//   // Determine new values based on device type
//   let newValues = [];
//   if(formData.device_type === "SIM"){
//     newValues = [
//       formData.editTicketId,
//       formData.editTicketNo,
//       "Resolved",
//       formData.editDateReceived,
//       formData.editCreatedAt,
//       formData.group_id,
//       formData.device_type,
//       formData.sim_card_id,
//       '',
//       formData.editTicketDescription,
//       formData.editPriority,
//       formData.editAssignedPic,
//       formData.editPlanOfAction,
//       formData.dateResolved,
//       formData.resolutionDetails

//     ];
//   } else if (formData.device_type === "PHONE"){
//     newValues = [
//       formData.editTicketId,
//       formData.editTicketNo,
//       "Resolved",
//       formData.editDateReceived,
//       formData.editCreatedAt,
//       formData.group_id,
//       formData.device_type,
//       '',
//       formData.mobile_id,
//       formData.editTicketDescription,
//       formData.editPriority,
//       formData.editAssignedPic,
//       formData.editPlanOfAction,
//       formData.dateResolved,
//       formData.resolutionDetails
//     ]
//   }

//   sheet.getRange(row, 1, 1, sheet.getLastColumn()).setValues([newValues]); // Set all values at once

//   // Identify changed fields
//   const changedFields = [];
//   const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]; // Get column headers
//   for (let i = 0; i < oldValues.length; i++) {
//     if (oldValues[i] != newValues[i]) {
//       changedFields.push(headers[i]);
//     }
//   }

//   // Log changes to the audit trail
//   logAuditTrail(
//     "EDIT",
//     "TICKET_MANAGEMENT",
//     id,
//     mapValuesToObject(headers, oldValues),
//     mapValuesToObject(headers, newValues),
//     changedFields,
//     "Ticket record updated"
//   );

// }

function deleteTicketRecord(id) {
    // Access the sheet where ticket records are stored
    const sheet = accessSheet("TICKET_MANAGEMENT");

    // Find the row for the ticket by its ID
    const row = findRowById("TICKET_MANAGEMENT", id);

    if (row === -1) {
        throw new Error(`Record with ID ${id} not found.`);
    }

    // Get the ticket's old values (before deletion)
    const oldValues = sheet
        .getRange(row, 1, 1, sheet.getLastColumn())
        .getValues()[0];

    // Formats the date fields
    oldValues[3] = formatDate(oldValues[3]);
    oldValues[4] = formatDate(oldValues[4]);
    oldValues[13] = formatDate(oldValues[13]);

    // Get column headers to map old values to field names
    const headers = sheet
        .getRange(1, 1, 1, sheet.getLastColumn())
        .getValues()[0];

    // Prepare the old values object
    const oldValuesObject = mapValuesToObject(headers, oldValues);

    // Log the delete action to the audit trail
    logAuditTrail(
        "DELETE",
        "TICKET_MANAGEMENT",
        id,
        oldValuesObject, // Old values before deletion
        {}, // No new values for DELETE action
        [], // No changed fields for DELETE action
        "Ticket record deleted"
    );

    var result = deleteRecordByColumnValue(
        id,
        "TICKET_ID",
        "TICKET_MANAGEMENT"
    );
    return result;
}

// can convert this function reusable to all to generate ID based on sheet name
function sequenceID() {
    const sheet = accessSheet("TICKET_MANAGEMENT");
    const lastRow = sheet.getLastRow();
    let lastId = sheet.getRange(lastRow, 1).getValue();
    let rowValues = [];

    const ticketId = lastId === "TICKET_ID" ? 1 : Number(lastId) + 1;
    const ticketNo = `TN-${ticketId}`;

    console.log("ticketId: ", ticketId);
    console.log("ticketNo: ", ticketNo);
}

function mapValuesToObject(headers, values) {
    const result = {};
    headers.forEach((header, index) => {
        result[header] = values[index];
    });
    return result;
}

function editTicketRecord(formData) {
    const sheet = accessSheet("TICKET_MANAGEMENT");
    const id = formData.editTicketId;
    const row = findRowById("TICKET_MANAGEMENT", id);

    if (row === -1) {
        throw new Error(`Record with ID ${id} not found.`);
    }

    // Get old values for the record
    const oldValues = sheet
        .getRange(row, 1, 1, sheet.getLastColumn())
        .getValues()[0];

    // Format the date fields
    oldValues[3] = formatDate(oldValues[3]);
    oldValues[4] = formatDate(oldValues[4]);
    oldValues[13] = formatDate(oldValues[13]);

    // Determine new values based on device type
    let newValues = [];
    if (formData.device_type === "SIM") {
        newValues = [
            formData.editTicketId,
            formData.editTicketNo,
            "Resolved",
            formData.editDateReceived,
            formData.editCreatedAt,
            formData.group_id,
            formData.device_type,
            formData.sim_card_id,
            "",
            formData.editTicketDescription,
            formData.editPriority,
            formData.editAssignedPic,
            formData.editPlanOfAction,
            formData.dateResolved,
            formData.resolutionDetails,
        ];
    } else if (formData.device_type === "PHONE") {
        newValues = [
            formData.editTicketId,
            formData.editTicketNo,
            "Resolved",
            formData.editDateReceived,
            formData.editCreatedAt,
            formData.group_id,
            formData.device_type,
            "",
            formData.mobile_id,
            formData.editTicketDescription,
            formData.editPriority,
            formData.editAssignedPic,
            formData.editPlanOfAction,
            formData.dateResolved,
            formData.resolutionDetails,
        ];
    }

    // Update the sheet with new values
    sheet.getRange(row, 1, 1, sheet.getLastColumn()).setValues([newValues]);

    // Identify changed fields and log only old and new values for them
    const headers = sheet
        .getRange(1, 1, 1, sheet.getLastColumn())
        .getValues()[0];
    const changedFields = [];
    const changedOldValues = {};
    const changedNewValues = {};

    for (let i = 0; i < oldValues.length; i++) {
        if (oldValues[i] != newValues[i]) {
            changedFields.push(headers[i]);
            changedOldValues[headers[i]] = oldValues[i];
            changedNewValues[headers[i]] = newValues[i];
        }
    }

    // Log changes to the audit trail
    logAuditTrail(
        "EDIT",
        "TICKET_MANAGEMENT",
        id,
        changedOldValues,
        changedNewValues,
        changedFields,
        "Ticket record updated"
    );
}
