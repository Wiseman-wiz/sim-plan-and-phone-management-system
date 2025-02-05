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
        const ticketSheet = new Utils.Sheet("TICKET_MANAGEMENT", {
            row: { start: 1 },
        });
        const billingSheet = new Utils.Sheet("BILLING", { row: { start: 1 } }); // Added Billing Sheet

        const simInventoryDataSet = simInventorySheet.toObject();
        const simPlanDataSet = simPlanSheet.toObject();
        const ticketDataSet = ticketSheet.toObject();
        const billingDataSet = billingSheet.toObject();

        // Get the latest ticket for each SIM card
        const latestDates = Utils.sql(
            `SELECT SIM_CARD_ID, MAX(DATE_RECEIVED) AS DATE_RECEIVED
       FROM ?
       WHERE SIM_CARD_ID IS NOT NULL AND SIM_CARD_ID != ""
       GROUP BY SIM_CARD_ID`,
            [ticketDataSet]
        );

        // Fetch the most recent ticket details for each SIM
        const latestTicketDetails = Utils.sql(
            `SELECT a.SIM_CARD_ID, a.STATUS AS TICKET_STATUS, a.DATE_RECEIVED, a.TICKET_NO
       FROM ? AS a
       JOIN ? AS b
       ON a.SIM_CARD_ID = b.SIM_CARD_ID AND a.DATE_RECEIVED = b.DATE_RECEIVED
       ORDER BY a.DATE_RECEIVED DESC`,
            [ticketDataSet, latestDates]
        );

        // Count active ("Pending") tickets per SIM
        const activeTicketCount = Utils.sql(
            `SELECT SIM_CARD_ID, COUNT(TICKET_NO) AS ACTIVE_TICKET_COUNT
       FROM ?
       WHERE STATUS = 'Pending' AND SIM_CARD_ID IS NOT NULL AND SIM_CARD_ID != ""
       GROUP BY SIM_CARD_ID`,
            [ticketDataSet]
        );

        // Compute Billing Data (Total Charge, Total Payment, Outstanding Amount, Latest Bill Period)
        const billingDetails = Utils.sql(
            `SELECT SIM_CARD_ID, 
              SUM(CURRENT_CHARGE_AMOUNT) AS TOTAL_CHARGE, 
              SUM(RFP_AMOUNT) AS TOTAL_PAYMENT, 
              SUM(CURRENT_CHARGE_AMOUNT) - SUM(RFP_AMOUNT) AS OUTSTANDING_AMOUNT,
              MAX(BILL_PERIOD_FROM) AS BILL_PERIOD_FROM,
              MAX(BILL_PERIOD_TO) AS BILL_PERIOD_TO
       FROM ?
       WHERE SIM_CARD_ID IS NOT NULL AND SIM_CARD_ID != ""
       GROUP BY SIM_CARD_ID`,
            [billingDataSet]
        );

        // Final Consolidated Query
        const execution = Utils.sql(
            `SELECT si.SIM_CARD_ID, si.MOBILE_NO, si.ACCOUNT_NO, si.STATUS,
              sp.PLAN_ID, sp.CATEGORY, sp.NETWORK_PROVIDER, sp.PLAN_DETAILS,
              t.TICKET_STATUS, t.DATE_RECEIVED, t.TICKET_NO,
              COALESCE(a.ACTIVE_TICKET_COUNT, 0) AS ACTIVE_TICKET_COUNT,
              CASE WHEN a.ACTIVE_TICKET_COUNT > 0 THEN 'Yes' ELSE 'No' END AS WITH_TICKET,
              COALESCE(b.TOTAL_CHARGE, 0) AS TOTAL_CHARGE,
              COALESCE(b.TOTAL_PAYMENT, 0) AS TOTAL_PAYMENT,
              COALESCE(b.OUTSTANDING_AMOUNT, 0) AS OUTSTANDING_AMOUNT,
              COALESCE(b.BILL_PERIOD_FROM, 0) AS BILL_PERIOD_FROM,
              COALESCE(b.BILL_PERIOD_TO, 0) AS BILL_PERIOD_TO
       FROM ? AS si
       LEFT JOIN ? AS sp ON si.PLAN_ID = sp.PLAN_ID
       LEFT JOIN ? AS t ON si.SIM_CARD_ID = t.SIM_CARD_ID
       LEFT JOIN ? AS a ON si.SIM_CARD_ID = a.SIM_CARD_ID
       LEFT JOIN ? AS b ON si.SIM_CARD_ID = b.SIM_CARD_ID`,
            [
                simInventoryDataSet,
                simPlanDataSet,
                latestTicketDetails,
                activeTicketCount,
                billingDetails,
            ]
        );

        // Return default structure if empty
        if (!execution || execution.length === 0) {
            return JSON.stringify([
                {
                    SIM_CARD_ID: "",
                    MOBILE_NO: "",
                    ACCOUNT_NO: "",
                    STATUS: "",
                    PLAN_ID: "",
                    CATEGORY: "",
                    NETWORK_PROVIDER: "",
                    PLAN_DETAILS: "",
                    TICKET_STATUS: "",
                    DATE_RECEIVED: "",
                    TICKET_NO: "",
                    ACTIVE_TICKET_COUNT: 0,
                    WITH_TICKET: "No",
                    TOTAL_CHARGE: 0,
                    TOTAL_PAYMENT: 0,
                    OUTSTANDING_AMOUNT: 0,
                    BILL_PERIOD_FROM: "",
                    BILL_PERIOD_TO: "",
                },
            ]);
        }

        console.log(typeof execution);
        return JSON.stringify(execution);
    } catch (error) {
        return Utils.ErrorHandler(error, {
            arguments,
            value: [],
        });
    }
}
