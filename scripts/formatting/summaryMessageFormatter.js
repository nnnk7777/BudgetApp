function handleWeeklySummaryResult(dateRangeStr, totalAmount, dataEntries, difference, percentage, adjustedBudget, isStaging, action, currentDate) {
    var nextWeekPlannedExpenses = getPlannedExpensesForNextWeek(currentDate);
    var nextWeekExpenseLines = formatUpcomingPlannedExpenseLines(nextWeekPlannedExpenses);
    var nextWeekPlannedExpenseTotal = calculatePlannedExpenseTotal(nextWeekPlannedExpenses);
    var weeklyBudgetCarryoverMemo = getWeeklyBudgetCarryoverMemoForWeek(currentDate);
    var weeklyAnalysisMode = getWeeklyAnalysisMode(currentDate);
    var differenceSign = difference >= 0 ? "+" : "-";
    var differenceAbs = Math.abs(difference);
    var percentageStr = percentage.toFixed(2);
    var categoryRankingLines = getCategoryRankingLines(dataEntries);
    var top5Entries = getTopExpenseEntries(dataEntries, 5);
    var body = "";

    body += "◆ " + dateRangeStr + " の週次サマリー\n\n";
    body += "++++++++++ 💸 予算サマリー 💸 ++++++++++\n";
    body += dateRangeStr + " の実支出: " + totalAmount + " 円\n";
    body += "\n";
    body += "予算に対して\n";
    body += "・実支出: " + percentageStr + "%\n";
    body += "（設定予算：" + adjustedBudget + "円）\n";
    body += "・分析モード: " + formatWeeklyAnalysisModeForMessage(weeklyAnalysisMode) + "\n";
    body += "++++++++++++++++++++++++++++++++++++\n";
    body += "* 予算差分：" + differenceSign + differenceAbs + "円\n\n";
    body += "◆ 前週からの持ち越し\n";
    body += formatWeeklyBudgetCarryoverSummaryForMessage(weeklyBudgetCarryoverMemo) + "\n\n";
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
    body += "◆ 来週の支出予定\n";
    if (nextWeekExpenseLines.length) {
        body += "合計見込み: " + nextWeekPlannedExpenseTotal + "円\n";
        nextWeekExpenseLines.forEach(function (line) {
            body += line + "\n";
        });
    } else {
        body += "・なし\n";
    }
    body += "\n";

    var geminiAnalysis = analyzeExpensesWithGemini(dataEntries, totalAmount, adjustedBudget, percentage, currentDate, weeklyAnalysisMode, {
        plannedExpenses: nextWeekPlannedExpenses,
        plannedExpenseLabel: "来週の予定メモ"
    });
    if (geminiAnalysis) {
        body += "◆ Gemini分析\n" + geminiAnalysis + "\n";
    } else {
        body += "◆ Gemini分析\n(Geminiからの回答を取得できませんでした。ログを確認してください)\n";
    }

    if (action === 'mail') {
        upsertWeeklyBudgetCarryoverMemo(currentDate, difference, adjustedBudget, totalAmount, dateRangeStr);
    } else if (action !== 'mail') {
        Logger.log("mail送信以外のため前週予算差分メモの保存をスキップしました");
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
    var weeklyBudgetCarryoverMemo = getWeeklyBudgetCarryoverMemoForWeek(currentDate);
    var weeklyAnalysisMode = getWeeklyAnalysisMode(currentDate);
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
    var geminiAnalysis = analyzeExpensesWithGemini(dataEntries, totalAmount, adjustedBudget, percentage, currentDate, weeklyAnalysisMode);
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
    body += "・分析モード: " + formatWeeklyAnalysisModeForMessage(weeklyAnalysisMode) + "\n";
    if (uncategorizedCount > 0) {
        body += "・未分類の支出: " + uncategorizedCount + "件\n";
    }
    body += "++++++++++++++++++++++++++++++++++++\n\n";
    body += "◆ 前週からの持ち越し\n";
    body += formatWeeklyBudgetCarryoverSummaryForMessage(weeklyBudgetCarryoverMemo) + "\n\n";
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

function formatWeeklyBudgetCarryoverSummaryForMessage(memo) {
    var overrunRatio;

    if (!memo) {
        return "・前週差分メモはありません。今週の予算と支出状況を中心に見ます。";
    }

    if (memo.difference <= 0) {
        return (
            "・前週（" +
            (memo.dateRangeStr || "対象週不明") +
            "）は予算内で、" +
            Math.abs(memo.difference) +
            "円の余裕がありました。ただし今週は今週の支出状況を優先して判断します。"
        );
    }

    overrunRatio = memo.adjustedBudget ? ((memo.difference / memo.adjustedBudget) * 100).toFixed(1) : "0.0";
    return (
        "・前週（" +
        (memo.dateRangeStr || "対象週不明") +
        "）は予算を" +
        memo.difference +
        "円超過していました（週予算比 " +
        overrunRatio +
        "% 超）。今週はこの反動も意識して、裁量支出をやや引き締め気味に見ます。"
    );
}

function formatWeeklyAnalysisModeForMessage(modeResult) {
    if (!modeResult || modeResult.mode !== WEEKLY_ANALYSIS_MODE_FRUGAL) {
        return "通常";
    }

    return modeResult.label + "（やや厳しめに評価）";
}
