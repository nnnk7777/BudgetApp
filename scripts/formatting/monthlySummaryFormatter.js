function logMonthlySummaryDebug(month, expenseEntries, incomeEntries, categoryTotals) {
    Logger.log("【MonthlySummary Debug】Target sheet: 🐖 家計簿");
    Logger.log("【MonthlySummary Debug】Expenses range: rows 35-185, cols " + getColumnsForMonth(month).dateCol + "-" + (getColumnsForMonth(month).dateCol + 3));
    Logger.log("【MonthlySummary Debug】Income range: rows 22-33, cols " + getColumnsForMonth(month).dateCol + "-" + (getColumnsForMonth(month).dateCol + 3));
    Logger.log("【MonthlySummary Debug】Expenses (sample):");
    expenseEntries.slice(0, 5).forEach(function (entry, idx) {
        Logger.log("  Expense[" + idx + "] " + formatDate(entry.date) + " " + (entry.category || "未分類") + " " + entry.name + " " + entry.amount);
    });
    Logger.log("【MonthlySummary Debug】Incomes (sample):");
    incomeEntries.slice(0, 3).forEach(function (entry, idx) {
        Logger.log("  Income[" + idx + "] " + formatDate(entry.date) + " " + entry.name + " " + entry.amount);
    });
    Logger.log("【MonthlySummary Debug】Category totals (top3): " + Object.keys(categoryTotals).sort(function (a, b) {
        return categoryTotals[b] - categoryTotals[a];
    }).slice(0, 3).map(function (category) {
        return category + "=" + categoryTotals[category];
    }).join(", "));
}

function buildMonthlySummaryMessage(dateRangeStr, totalIncome, totalExpenses, adjustedBudget, difference, percentage, expenseEntries, incomeEntries, categoryTotals, geminiAnalysis) {
    var sortedCategories = Object.keys(categoryTotals).sort(function (a, b) {
        return categoryTotals[b] - categoryTotals[a];
    });
    var body = "";

    body += "◆ " + dateRangeStr + " の月次サマリー\n\n";
    body += "収入合計: " + totalIncome + "円\n";
    body += "支出合計: " + totalExpenses + "円\n";
    body += "月間予算(週予算換算): " + adjustedBudget + "円\n";
    body += "予算差分: " + (difference >= 0 ? "+" : "-") + Math.abs(difference) + "円\n";
    body += "予算消化率: " + percentage.toFixed(2) + "%\n\n";
    body += "◆ カテゴリ別支出\n";

    sortedCategories.forEach(function (category) {
        body += "・" + category + ": " + categoryTotals[category] + "円\n";
    });
    body += "\n";
    body += "◆ 支出一覧\n";
    expenseEntries.forEach(function (entry) {
        body += "・" + formatDate(entry.date) + " - " + entry.category + " - " + entry.name + ": " + entry.amount + "円\n";
    });
    body += "\n◆ 収入一覧\n";
    incomeEntries.forEach(function (entry) {
        body += "・" + formatDate(entry.date) + " - " + entry.name + ": " + entry.amount + "円\n";
    });
    body += "\n";

    if (geminiAnalysis) {
        body += "◆ Gemini分析\n" + geminiAnalysis + "\n";
    } else {
        body += "◆ Gemini分析\n(Geminiからの回答を取得できませんでした。ログを確認してください)\n";
    }

    return body;
}

function sendMonthlySummaryResult(action, currentDate, isStaging, body) {
    switch (action) {
        case 'mail':
            var emailAddress = getTargetEmailAddress();
            var subjectPrefix = isStaging ? "<test>" : "";
            var subject = subjectPrefix + "家計簿月次レポート（" + (currentDate.getMonth() + 1) + "月）";
            MailApp.sendEmail(emailAddress, subject, body);
            return "Successfully sent monthly summary mail";
        case 'text':
            return body;
        default:
            throw new Error('actionが定義されていません');
    }
}
