function getMonthlyPaymentData() {
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

        const billingDataSet = billingSheet.toObject();

        const query = `
        SELECT
          b.BILL_ID,
          b.SIM_CARD_ID,
          b.BILL_PERIOD_FROM,
          b.SIM_INFO,
          b.EMPLOYEE_INFO,
          b.RFP_AMOUNT
        FROM ? AS b
      `;

        const execution = Utils.sql(query, [billingDataSet]);

        // Check if dataset is empty, return placeholder if needed
        if (!execution || execution.length === 0) {
            return JSON.stringify([
                {
                    BILL_ID: "",
                    SIM_CARD_ID: "",
                    BILL_PERIOD_FROM: "",
                    SIM_INFO: "",
                    MOBILE_NO: "",
                    ACCOUNT_NO: "",
                    RFP_AMOUNT: "",
                },
            ]);
        }

        execution.forEach((row) => {
            if (row.BILL_PERIOD_FROM) {
                row.BILL_PERIOD_FROM = Utilities.formatDate(
                    new Date(row.BILL_PERIOD_FROM),
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
