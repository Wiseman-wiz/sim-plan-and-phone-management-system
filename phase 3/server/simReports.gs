function getReportsData() {
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
        const paymentSheet = new Utils.Sheet("PAYMENT", { row: { start: 1 } });

        const billingDataSet = billingSheet.toObject();
        const paymentDataSet = paymentSheet.toObject();

        const query = `
        SELECT
          b.BILL_ID,
          b.SIM_CARD_ID,
          b.BILL_PERIOD_FROM,
          b.SIM_INFO,
          p.BILL_ID,
          p.PAYMENT_REFERENCE_DATE,
          p.PAYMENT_REFERENCE_NO,
          b.RFP_AMOUNT,
          b.WITHHOLDING_TAX,
          b.AMOUNT_AFTER_TAX
        FROM ? AS b
        LEFT JOIN ? AS p ON b.BILL_ID = p.BILL_ID 
      `;

        const execution = Utils.sql(query, [billingDataSet, paymentDataSet]);

        // Check if dataset is empty, return placeholder if needed
        if (!execution || execution.length === 0) {
            return JSON.stringify([
                {
                    BILL_ID: "",
                    SIM_CARD_ID: "",
                    BILL_PERIOD_FROM: "",
                    SIM_INFO: "",
                    PAYMENT_REFERENCE_DATE: "",
                    PAYMENT_REFERENCE_NO: "",
                    RFP_AMOUNT: "",
                    WITHHOLDING_TAX: "",
                    AMOUNT_AFTER_TAX: "",
                    MOBILE_NO: "",
                    ACCOUNT_NO: "",
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
            } else if (row.PAYMENT_REFERENCE_DATE) {
                row.PAYMENT_REFERENCE_DATE = Utilities.formatDate(
                    new Date(row.PAYMENT_REFERENCE_DATE),
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

function getSimReportsData() {
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
        const simPlanSheet = new Utils.Sheet("SIM_PLANS", {
            row: { start: 1 },
        });

        const simInventoryDataSet = simInventorySheet.toObject();
        const simPlanDataSet = simPlanSheet.toObject();

        // NEED TO FIND A WAY TO INCLUDE HERE THE SIM CARDS WITH PHONE OR WITHOUT PHONE
        const query = `
        SELECT
          si.SIM_CARD_ID,
          si.MOBILE_NO,
          si.ACCOUNT_NO,
          si.STATUS,
          sp.PLAN_ID,
          sp.CATEGORY
        FROM ? AS si
        LEFT JOIN ? AS sp ON si.PLAN_ID = sp.PLAN_ID
      `;

        const execution = Utils.sql(query, [
            simInventoryDataSet,
            simPlanDataSet,
        ]);

        // Check if dataset is empty, return placeholder if needed
        if (!execution || execution.length === 0) {
            return JSON.stringify([
                {
                    SIM_CARD_ID: "",
                    MOBILE_NO: "",
                    ACCOUNT_NO: "",
                    STATUS: "",
                    PLAN_ID: "",
                    CATEGORY: "",
                },
            ]);
        }

        // execution.forEach(row => {
        //   if (row.BILL_PERIOD_FROM) {
        //     row.BILL_PERIOD_FROM = Utilities.formatDate(new Date(row.BILL_PERIOD_FROM), Session.getScriptTimeZone(), "MM/dd/yyyy");
        //   } else if (row.PAYMENT_REFERENCE_DATE) {
        //     row.PAYMENT_REFERENCE_DATE = Utilities.formatDate(new Date(row.PAYMENT_REFERENCE_DATE), Session.getScriptTimeZone(), "MM/dd/yyyy");
        //   }
        // });

        console.log(execution);
        return JSON.stringify(execution);
    } catch (error) {
        return Utils.ErrorHandler(error, {
            arguments,
            value: [],
        });
    }
}
