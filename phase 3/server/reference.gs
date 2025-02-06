// function to get the reference data from Signatory sheet
function getReferenceData() {
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

        const referenceSheet = new Utils.Sheet("SIGNATORY", {
            row: { start: 1 },
        });

        const referenceDataSet = referenceSheet.toObject();

        const execution = Utils.sql(
            `SELECT 
                r.SIGNATORY_ID,
                r.NAME,
                r.POSITION,
                r.FIELD_NAME,
                r.DOCUMENT_NAME
                FROM ? r
            `,
            [
                referenceDataSet
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

// function to edit signatory record with audit trail
function editSignatory(formData) {
    const sheet = accessSheet("SIGNATORY");
    const id = formData.editSignatoryId;
    const row = findRowById("SIGNATORY", id);

    if (row === -1) {
        throw new Error(`Record with ID ${id} not found.`);
    }

    // Get old values for the record
    const oldValues = sheet
        .getRange(row, 1, 1, sheet.getLastColumn())
        .getValues()[0];

    // Define new values
    let newValues = [
        formData.editSignatoryId,
        formData.editName,
        formData.editPosition,
        formData.editFieldName,
        formData.editDocumentName
    ];

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
        "SIGNATORY",
        id,
        changedOldValues,
        changedNewValues,
        changedFields,
        "Signatory record updated"
    );
}