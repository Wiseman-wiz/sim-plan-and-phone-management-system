function getDelayedPaymentsData() {
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
            row: { start: 1 },
        });
        const rfpSummarySheet = new Utils.Sheet("RFP_SUMMARY", {
            row: { start: 1 },
        });
        const billingSheet = new Utils.Sheet("BILLING", { row: { start: 1 } }); // Added Billing Sheet
        const paymentSheet = new Utils.Sheet("PAYMENT", { row: { start: 1 } });

        const simInventoryDataSet = simInventorySheet.toObject();
        const rfpSummaryDataSet = rfpSummarySheet.toObject();
        const billingDataSet = billingSheet.toObject();
        const paymentDataSet = paymentSheet.toObject();

        const execution = Utils.sql(
            `SELECT 
                b.BILL_ID,
                b.RFP_NO,
                b.SIM_CARD_ID,
                s.MOBILE_NO,
                s.ACCOUNT_NO,
                s.SIM_INFO,
                b.BILL_PERIOD_FROM,
                b.BILL_PERIOD_TO,
                s.DUE_DATE_DAY,
                r.RFP_DATE,
                r.DATE_RECEIVED_BY_ACCTG,
                r.CHECK_DATE,
                r.DATE_OF_PAYMENT,
                p.PAYMENT_POSTED_DATE
                FROM ? b
                JOIN ? s ON b.SIM_CARD_ID = s.SIM_CARD_ID
                LEFT JOIN ? r ON b.RFP_NO = r.RFP_NO
                LEFT JOIN ? p ON b.BILL_ID = p.BILL_ID
            `,
            [
                billingDataSet,
                simInventoryDataSet,
                rfpSummaryDataSet,
                paymentDataSet,
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
