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

    var specialExpenseReview = reviewSpecialExpensesWithAI(expenseEntries);
    var budgetTargetExpenseEntries = getBudgetTargetEntries(expenseEntries, specialExpenseReview);
    var actualTotalExpenses = calculateTotalAmount(expenseEntries);
    var totalExpenses = calculateTotalAmount(budgetTargetExpenseEntries);
    var totalIncome = calculateTotalAmount(incomeEntries);

    var daysInMonth = endOfMonth.getDate();
    var adjustedBudget = Math.round((budgetPerWeek * daysInMonth / 7) / 100) * 100;
    var difference = totalExpenses - adjustedBudget;
    var percentage = adjustedBudget ? (totalExpenses / adjustedBudget) * 100 : 0;

    var categoryTotals = calculateCategoryTotals(budgetTargetExpenseEntries);
    logMonthlySummaryDebug(month, budgetTargetExpenseEntries, incomeEntries, categoryTotals);

    var aiAnalysis = analyzeMonthlyWithAI(budgetTargetExpenseEntries, categoryTotals, totalExpenses, totalIncome, adjustedBudget, percentage, dateRangeStr);
    var body = buildMonthlySummaryMessage(
        dateRangeStr,
        totalIncome,
        totalExpenses,
        actualTotalExpenses,
        adjustedBudget,
        difference,
        percentage,
        expenseEntries,
        incomeEntries,
        categoryTotals,
        aiAnalysis,
        specialExpenseReview
    );

    return sendMonthlySummaryResult(action, currentDate, isStaging, body);
}
