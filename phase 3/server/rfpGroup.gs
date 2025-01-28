// function addRFPGroup(formData) {
//     formData.groupID = generateID(RFP_GROUP, "RFP_GROUP_ID");
//     formData.groupName = `${formData.company} - ${formData.networkProvider}`;
//     addRecordToSheet(formData, RFP_GROUP, [], KEY_ORDER.RFP_GROUP, "");
// }

// function deleteRFPGroup(groupID) {
//     var result = deleteRecordByColumnValue(groupID, 'RFP_GROUP_ID', RFP_GROUP);
//     return result;
// }

// function editRFPGroup(formData) {
//     console.log(formData);
//     var row = findRowById(RFP_GROUP, formData.groupId);
//     console.log(row);
//     if (row != -1) {
//         upsertRecord(RFP_GROUP, row, formData, formToSheetMap.RFP_GROUP);
//     } else {
//         Logger.log('Record not found');
//     }
// }

// With Audit Trail
function addRFPGroup(formData) {
    // Generate a new ID for the group
    formData.groupID = generateID(RFP_GROUP, "RFP_GROUP_ID");

    // Generate the group name
    formData.groupName = `${formData.company} - ${formData.networkProvider}`;

    // Log the addition of this new RFP group to the audit trail
    const actionType = "ADD"; // Action type for adding a new record
    const entityName = "RFP_GROUP";
    const entityId = formData.groupID;
    const oldValue = {}; // There is no old value since it's a new record
    const newValue = formData; // The new values of the record
    const changedFields = Object.keys(formData); // All fields have changed as it's a new record
    const remarks = "New RFP Group added";

    // Log the audit trail for this action
    logAuditTrail(
        actionType,
        entityName,
        entityId,
        oldValue,
        newValue,
        [],
        remarks
    );

    // Add the new record to the RFP_GROUP sheet
    addRecordToSheet(formData, RFP_GROUP, [], KEY_ORDER.RFP_GROUP, "");
}

function editRFPGroup(formData) {
    const sheet = accessSheet("RFP_GROUP");
    const id = formData.groupId;
    const row = findRowById("RFP_GROUP", id);

    if (row === -1) {
        throw new Error(`Record with ID ${id} not found.`);
    }

    // Get old values for the record
    const oldValues = sheet
        .getRange(row, 1, 1, sheet.getLastColumn())
        .getValues()[0];

    // Prepare new values
    const newValues = [
        formData.groupId, // Group ID
        formData.editGroupName, // Group Name
        formData.editCompany, // Company
        formData.editNetworkProvider, // Network Provider
        formData.editPayableTo || "", // Payable To (if present in the formData)
        // '',                                 // Status (if required, leave blank if not available)
        // ''                                  // Date Created or Last Updated (if required)
    ];

    // Update the sheet with the new values
    sheet.getRange(row, 1, 1, sheet.getLastColumn()).setValues([newValues]);

    // Identify changed fields
    const changedFields = [];
    const headers = sheet
        .getRange(1, 1, 1, sheet.getLastColumn())
        .getValues()[0]; // Get column headers
    for (let i = 0; i < oldValues.length; i++) {
        if (oldValues[i] !== newValues[i]) {
            changedFields.push(headers[i]);
        }
    }

    // Log changes to the audit trail
    logAuditTrail(
        "EDIT",
        "RFP_GROUP",
        id,
        mapValuesToObject(headers, oldValues),
        mapValuesToObject(headers, newValues),
        changedFields,
        "RFP Group record updated"
    );
}

function deleteRFPGroup(groupID) {
    // Access the sheet where RFP groups are stored
    const sheet = accessSheet("RFP_GROUP");

    // Find the row for the group by its ID
    const row = findRowById("RFP_GROUP", groupID);

    if (row === -1) {
        throw new Error(`Record with ID ${groupID} not found.`);
    }

    // Get the group's old values (before deletion)
    const oldValues = sheet
        .getRange(row, 1, 1, sheet.getLastColumn())
        .getValues()[0];

    // Get column headers to map old values to field names
    const headers = sheet
        .getRange(1, 1, 1, sheet.getLastColumn())
        .getValues()[0];

    // Prepare the old values object
    const oldValuesObject = mapValuesToObject(headers, oldValues);

    // Log the delete action to the audit trail
    logAuditTrail(
        "DELETE",
        "RFP_GROUP",
        groupID,
        oldValuesObject, // Old values before deletion
        {}, // No new values for DELETE action
        [], // No changed fields for DELETE action
        "RFP Group deleted"
    );

    // Delete the record
    const result = deleteRecordByColumnValue(
        groupID,
        "RFP_GROUP_ID",
        "RFP_GROUP"
    );
    return result;
}

/**
 * Retrieve a record from the sheet by its unique identifier.
 * @param {string} sheetName - The name of the sheet to search.
 * @param {string} id - The unique identifier to search for.
 * @returns {Object|null} - The record as an object or null if not found.
 */
function getRecordById(sheetName, id) {
    const sheet = accessSheet(sheetName); // Your function to access the sheet
    const data = sheet.getDataRange().getValues(); // Retrieve all data from the sheet
    const headers = data[0]; // Assuming the first row contains headers

    // Iterate over rows to find the matching ID
    for (let i = 1; i < data.length; i++) {
        if (data[i][0] == id) {
            // Assuming the ID is in the first column
            const record = {};
            headers.forEach((header, index) => {
                record[header] = data[i][index]; // Map headers to row values
            });
            return record;
        }
    }

    return null; // Return null if the ID is not found
}

/**
 * Compare old and new objects and return the fields that have changed.
 * @param {Object} oldValue - The original object.
 * @param {Object} newValue - The updated object.
 * @returns {Array<string>} - A list of changed field names.
 */
function getChangedFields(oldValue, newValue) {
    const changedFields = [];

    for (const key in newValue) {
        if (newValue[key] !== oldValue[key]) {
            changedFields.push(key);
        }
    }

    return changedFields;
}
