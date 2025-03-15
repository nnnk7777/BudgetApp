/**
 * 受け取ったアクションに応じて処理を実行する
 * 
 * @param {Action} action 
 */
function calculateExpensesSummary(action) {
    const config = new Configure(
        "TEST_DATE_PLACEHOLDER"
            ? Env.STG
            : Env.PRD,
        "TARGET_EMAIL_ADDRESS"
    );
    const expenseSheet = new Sheet("🐖 家計簿");
    const calendar = new Calendar();
    const dateList = calendar.getDatesInWeek();
    config.adjustedBudget(dateList);

    const weeklyRecord = new WeeklyRecord(dateList, expenseSheet);
    weeklyRecord.getExpenseEntries();
    weeklyRecord.culculateTotalAmount();
    weeklyRecord.calculateExpenseUsage(config.adjustedBudget)
    weeklyRecord.calculateTop5Entries();


    const dailySummaryMessage = new Message(config.env);
    dailySummaryMessage.buildDaiilySummaryMessage(
        new Date(),
        dateList,
        weeklyRecord.totalAmount,
        weeklyRecord.budgetPercentage,
        weeklyRecord.culculateTotalAmount
    );
    const weeklySummaryMessage = new Message(config.env);
    if ((new Date()).getDay() === 0) {
        weeklySummaryMessage.buildWeeklySummaryMessage(
            weeklyRecord.dataEntries,
            dateList,
            weeklyRecord.totalAmount,
            weeklyRecord.budgetPercentage,
            weeklyRecord.budgetDifference,
            weeklyRecord.top5Entries
        );
    }

}