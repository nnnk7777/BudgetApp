// 月次サマリーのメイン処理
function calculateMonthlySummary(action) {
    var budgetPerWeek = 45000;
    var runtimeContext = getScriptRuntimeContext();
    var currentDate = runtimeContext.currentDate;
    var isStaging = runtimeContext.isStaging;
    var year = currentDate.getFullYear();
    var month = currentDate.getMonth(); // 0-based

    var startOfMonth = new Date(year, month, 1);
    var endOfMonth = new Date(year, month + 1, 0);
    var dateRangeStr = formatDate(startOfMonth) + "〜" + formatDate(endOfMonth);

    var expenseEntries = getMonthlyExpenseEntries(year, month);
    var incomeEntries = getMonthlyIncomeEntries(year, month);

    var totalExpenses = calculateTotalAmount(expenseEntries);
    var totalIncome = calculateTotalAmount(incomeEntries);

    var daysInMonth = endOfMonth.getDate();
    var adjustedBudget = Math.round((budgetPerWeek * daysInMonth / 7) / 100) * 100;
    var difference = totalExpenses - adjustedBudget;
    var percentage = adjustedBudget ? (totalExpenses / adjustedBudget) * 100 : 0;

    var categoryTotals = calculateCategoryTotals(expenseEntries);
    logMonthlySummaryDebug(month, expenseEntries, incomeEntries, categoryTotals);

    var geminiAnalysis = analyzeMonthlyWithGemini(expenseEntries, categoryTotals, totalExpenses, totalIncome, adjustedBudget, percentage, dateRangeStr);
    var body = buildMonthlySummaryMessage(
        dateRangeStr,
        totalIncome,
        totalExpenses,
        adjustedBudget,
        difference,
        percentage,
        expenseEntries,
        incomeEntries,
        categoryTotals,
        geminiAnalysis
    );

    return sendMonthlySummaryResult(action, currentDate, isStaging, body);
}
