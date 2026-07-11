function analyzeMonthlyWithAI(expenseEntries, categoryTotals, totalExpenses, totalIncome, adjustedBudget, percentage, dateRangeStr) {
    var prompt = buildMonthlySummaryPrompt(expenseEntries, categoryTotals, totalExpenses, totalIncome, adjustedBudget, percentage, dateRangeStr);

    return generatePreferredAiText(prompt, {
        temperature: 0.4,
        maxOutputTokens: 1000,
        thinkingConfig: {
            thinkingBudget: 400
        }
    }, {
        logContext: "monthly_summary"
    });
}
