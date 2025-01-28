function logAuditTrail(
    actionType,
    entityName,
    entityId,
    oldValue = {},
    newValue = {},
    changedFields = [],
    remarks = ""
) {
    const sheet = accessSheet("AUDIT_TRAIL");
    const timestamp = new Date();

    const userEmail = Session.getActiveUser().getEmail() || "Unknown User";

    // Prepare the audit log entry
    const auditEntry = [
        sheet.getLastRow(), // Audit ID (incremental)
        entityName, // Entity Name
        entityId, // Entity ID
        actionType, // Action Type
        timestamp, // Timestamp
        JSON.stringify(oldValue), // Old Value (as JSON)
        JSON.stringify(newValue), // New Value (as JSON)
        JSON.stringify(changedFields), // Changed Fields
        userEmail, // User Email
        remarks, // remarks for Change
    ];

    // Append the entry to the Audit Trail sheet
    sheet.appendRow(auditEntry);
}

function getAuditTrailData() {
    const auditTrailSheet = new Utils.Sheet("AUDIT_TRAIL", {
        row: { start: 1 },
    });

    let auditTrailDataSet = auditTrailSheet.toObject();

    if (auditTrailDataSet.length === 0) {
        auditTrailDataSet = [
            {
                AUDIT_ID: "",
                SHEET_NAME: "",
                RECORD_ID: "",
                ACTION_TYPE: "",
                TIMESTAMP: "",
                OLD_VALUE: "",
                NEW_VALUE: "",
                USER: "",
                CHANGED_FIELDS: "",
                REMARKS: "",
            },
        ];
    }

    console.log(auditTrailDataSet.length);
    console.log(auditTrailDataSet);
    return JSON.stringify(auditTrailDataSet);
}
