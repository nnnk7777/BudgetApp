function handleWeeklySummaryResult(dateRangeStr, totalAmount, dataEntries, difference, percentage, adjustedBudget, isStaging, action, currentDate) {
    var upcomingPlannedExpenses = getPlannedExpensesForCurrentWeek(currentDate);
    var upcomingExpenseLines = formatUpcomingPlannedExpenseLines(upcomingPlannedExpenses);
    var plannedExpenseTotal = calculatePlannedExpenseTotal(upcomingPlannedExpenses);
    var differenceSign = difference >= 0 ? "+" : "-";
    var differenceAbs = Math.abs(difference);
    var percentageStr = percentage.toFixed(2);
    var projectedPercentage = adjustedBudget
        ? (((totalAmount + plannedExpenseTotal) / adjustedBudget) * 100).toFixed(2)
        : "0.00";
    var categoryRankingLines = getCategoryRankingLines(dataEntries);
    var top5Entries = getTopExpenseEntries(dataEntries, 5);
    var body = "";

    body += "◆ " + dateRangeStr + " の週次サマリー\n\n";
    body += "++++++++++ 💸 予算サマリー 💸 ++++++++++\n";
    body += dateRangeStr + " の実支出: " + totalAmount + " 円\n";
    body += "今後の予定金額: " + plannedExpenseTotal + " 円\n";
    body += "支出＋予定の合計見込み: " + (totalAmount + plannedExpenseTotal) + " 円\n";
    body += "\n";
    body += "予算に対して\n";
    body += "・実支出: " + percentageStr + "%\n";
    body += "・合計見込み: " + projectedPercentage + "%\n";
    body += "（設定予算：" + adjustedBudget + "円）\n";
    body += "++++++++++++++++++++++++++++++++++++\n";
    body += "* 予算差分：" + differenceSign + differenceAbs + "円\n\n";
    body += "◆ カテゴリ別支出ランキング\n";
    if (categoryRankingLines.length) {
        categoryRankingLines.forEach(function (line) {
            body += line + "\n";
        });
    } else {
        body += "・なし\n";
    }
    body += "\n";
    body += "◆ 支出TOP5\n";
    top5Entries.forEach(function (entry) {
        body += "・" + formatDate(entry.date) + " - " + entry.name + ": " + entry.amount + "円\n";
    });
    body += "\n";
    body += "◆ 支出一覧\n";
    dataEntries.forEach(function (entry) {
        body += "・" + formatDate(entry.date) + " - " + entry.name + ": " + entry.amount + "円\n";
    });
    body += "\n";
    body += "◆ 直近の予定支出\n";
    if (upcomingExpenseLines.length) {
        upcomingExpenseLines.forEach(function (line) {
            body += line + "\n";
        });
    } else {
        body += "・なし\n";
    }
    body += "\n";

    var geminiAnalysis = analyzeExpensesWithGemini(dataEntries, totalAmount, adjustedBudget, percentage, currentDate);
    if (geminiAnalysis) {
        body += "◆ Gemini分析\n" + geminiAnalysis + "\n";
    } else {
        body += "◆ Gemini分析\n(Geminiからの回答を取得できませんでした。ログを確認してください)\n";
    }

    if (!isStaging) {
        upsertWeeklyBudgetCarryoverMemo(currentDate, difference, adjustedBudget, totalAmount, dateRangeStr);
    } else {
        Logger.log("staging実行のため前週予算差分メモの保存をスキップしました");
    }

    switch (action) {
        case 'mail':
            var emailAddress = getTargetEmailAddress();
            var subject = (isStaging ? "<test>" : "") + "家計簿週次レポート" + "（" + dateRangeStr + "）";
            MailApp.sendEmail(emailAddress, subject, body);
            return "Successfully sent mail";
        case 'text':
            return body;
        default:
            throw new Error('actionが定義されていません');
    }
}

function handleDailySummaryResult(currentDate, datesInWeek, adjustedBudget, isStaging, action) {
    var upcomingPlannedExpenses = getPlannedExpensesForCurrentWeek(currentDate);
    var upcomingExpenseLines = formatUpcomingPlannedExpenseLines(upcomingPlannedExpenses);
    var plannedExpenseTotal = calculatePlannedExpenseTotal(upcomingPlannedExpenses);
    var datesUpToToday = datesInWeek.filter(function (date) {
        return date <= currentDate;
    });
    var dataEntries = getExpenseEntriesForDates(datesUpToToday).reverse().map(function (entry) {
        if (entry.name.length >= 16) {
            entry.name = entry.name.substring(0, 14) + "...";
        }
        return entry;
    });
    var totalAmount = calculateTotalAmount(dataEntries);
    var percentage = (totalAmount / adjustedBudget) * 100;
    var projectedPercentage = adjustedBudget
        ? (((totalAmount + plannedExpenseTotal) / adjustedBudget) * 100).toFixed(2)
        : "0.00";
    var categoryRankingLines = getCategoryRankingLines(dataEntries);
    var uncategorizedCount = countUncategorizedEntries(dataEntries);
    var geminiAnalysis = analyzeExpensesWithGemini(dataEntries, totalAmount, adjustedBudget, percentage, currentDate);
    var subject = (isStaging ? "<test>" : "") + "家計簿日次レポート（" + formatDate(currentDate) + "）";
    var body = "++++++++++ 💸 予算サマリー 💸 ++++++++++\n";

    body += formatDate(datesInWeek[0]) + " から " + formatDate(currentDate) + " までの実支出: " + totalAmount + " 円\n";
    body += "今後の予定金額: " + plannedExpenseTotal + " 円\n";
    body += "支出＋予定の合計見込み: " + (totalAmount + plannedExpenseTotal) + " 円\n";
    body += "\n";
    body += "予算に対して\n";
    body += "・実支出: " + percentage.toFixed(2) + "%\n";
    body += "・合計見込み: " + projectedPercentage + "%\n";
    body += "（設定予算：" + adjustedBudget + "円）\n";
    if (uncategorizedCount > 0) {
        body += "・未分類の支出: " + uncategorizedCount + "件\n";
    }
    body += "++++++++++++++++++++++++++++++++++++\n\n";
    body += "◆ カテゴリ別支出ランキング\n";
    if (categoryRankingLines.length) {
        categoryRankingLines.forEach(function (line) {
            body += line + "\n";
        });
    } else {
        body += "・なし\n";
    }
    body += "\n";
    body += "詳細:\n";
    dataEntries.forEach(function (entry) {
        body += "・" + formatDate(entry.date) + " - " + entry.name + ": " + entry.amount + "円\n";
    });
    body += "\n";
    body += "◆ 直近の予定支出\n";
    if (upcomingExpenseLines.length) {
        upcomingExpenseLines.forEach(function (line) {
            body += line + "\n";
        });
    } else {
        body += "・なし\n";
    }
    body += "\n";

    if (geminiAnalysis) {
        body += "◆ Gemini分析\n" + geminiAnalysis + "\n";
    } else {
        body += "◆ Gemini分析\n(Geminiからの回答を取得できませんでした。ログを確認してください)\n";
    }

    switch (action) {
        case 'mail':
            var emailAddress = getTargetEmailAddress();
            MailApp.sendEmail(emailAddress, subject, body);
            return "Successfully sent mail";
        case 'text':
            return body;
        default:
            throw new Error('actionが定義されていません');
    }
}
