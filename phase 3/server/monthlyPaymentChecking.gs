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

        const result = transformData(execution);
        console.log(result);

        return result;
    } catch (error) {
        return Utils.ErrorHandler(error, {
            arguments,
            value: [],
        });
    }
}

function transformData(dataset) {
    const result = {};

    dataset.forEach(
        ({ SIM_CARD_ID, SIM_INFO, BILL_PERIOD_FROM, RFP_AMOUNT }) => {
            const { MOBILE_NO, ACCOUNT_NO } = JSON.parse(SIM_INFO);
            const dateObj = new Date(BILL_PERIOD_FROM);
            const billYear = dateObj.getFullYear(); // Extract Year (YYYY)
            const billMonth = `${billYear}-${String(
                dateObj.getMonth() + 1
            ).padStart(2, "0")}`; // Format YYYY-MM

            // Initialize SIM card entry if it doesn't exist
            if (!result[SIM_CARD_ID]) {
                result[SIM_CARD_ID] = {
                    SIM_CARD_ID,
                    MOBILE_NO,
                    ACCOUNT_NO,
                    YEARS: {}, // Store payments by year
                };
            }

            // Initialize the year with all 12 months set to 0 if not exists
            if (!result[SIM_CARD_ID].YEARS[billYear]) {
                result[SIM_CARD_ID].YEARS[billYear] = {};
                for (let month = 1; month <= 12; month++) {
                    const monthKey = `${billYear}-${String(month).padStart(
                        2,
                        "0"
                    )}`;
                    result[SIM_CARD_ID].YEARS[billYear][monthKey] = 0;
                }
            }

            // Store or update the monthly payment
            result[SIM_CARD_ID].YEARS[billYear][billMonth] += RFP_AMOUNT || 0;
        }
    );

    return Object.values(result);
}
