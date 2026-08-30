function handleWeeklySummaryResult(dateRangeStr, totalAmount, dataEntries, difference, percentage, adjustedBudget, isStaging, action, currentDate) {
    var nextWeekCalendarMemos = getCalendarMemoEntriesInRange(getNextWeekStartDate(currentDate), getWeekAfterNextStartDate(currentDate));
    var nextWeekPlannedExpenses = filterPlannedExpenseMemos(nextWeekCalendarMemos);
    var nextWeekContextualMemos = filterContextualCalendarMemos(nextWeekCalendarMemos);
    var nextWeekExpenseLines = formatUpcomingPlannedExpenseLines(nextWeekPlannedExpenses);
    var nextWeekContextualMemoLines = formatUpcomingPlannedExpenseLines(nextWeekContextualMemos);
    var nextWeekPlannedExpenseTotal = calculatePlannedExpenseTotal(nextWeekPlannedExpenses);
    var saleIncomeEntries = getSaleIncomeEntriesForDates(getDatesInWeek(currentDate));
    var saleIncomeTotal = calculateTotalAmount(saleIncomeEntries);
    var weeklyBudgetCarryoverMemo = getWeeklyBudgetCarryoverMemoForWeek(currentDate);
    var weeklyAnalysisMode = getWeeklyAnalysisMode(currentDate);
    var differenceSign;
    var differenceAbs;
    var percentageStr;
    var actualWeeklyTotalAmount = totalAmount;
    var monthlyBudget = calculateMonthlyBudgetForDate(currentDate);
    var monthEntries = getCurrentMonthExpenseEntries(currentDate);
    var specialExpenseReview = reviewSpecialExpensesWithAI(monthEntries);
    var budgetTargetEntries = getBudgetTargetEntries(dataEntries, specialExpenseReview);
    var monthlyBudgetTargetEntries = getBudgetTargetEntries(monthEntries, specialExpenseReview);
    var approvedSpecialExpenseTotal = calculateApprovedSpecialExpenseTotal(specialExpenseReview);
    var approvedWeeklySpecialExpenseTotal = calculateApprovedSpecialExpenseTotalForEntries(dataEntries, specialExpenseReview);
    var categoryRankingLines;
    var top5Entries;
    var monthlyTotalAmount;
    var actualMonthlyTotalAmount;
    var body = "";

    totalAmount = calculateTotalAmount(budgetTargetEntries);
    difference = totalAmount - adjustedBudget;
    percentage = adjustedBudget ? (totalAmount / adjustedBudget) * 100 : 0;
    differenceSign = difference >= 0 ? "+" : "-";
    differenceAbs = Math.abs(difference);
    percentageStr = percentage.toFixed(2);
    monthlyTotalAmount = calculateTotalAmount(monthlyBudgetTargetEntries);
    actualMonthlyTotalAmount = calculateTotalAmount(monthEntries);
    categoryRankingLines = getCategoryRankingLines(budgetTargetEntries);
    top5Entries = getTopExpenseEntries(budgetTargetEntries, 5);

    body += "◆ " + dateRangeStr + " の週次サマリー\n\n";
    body += "+++ 💸 予算サマリー 💸 +++\n";
    body += dateRangeStr + " の予算対象支出: " + totalAmount + " 円\n";
    body += dateRangeStr + " の実支出合計: " + actualWeeklyTotalAmount + " 円\n";
    body += "今月の承認済み特別費: " + approvedSpecialExpenseTotal + " 円\n";
    body += "\n";
    body += "予算に対して\n";
    body += "・予算対象支出: " + percentageStr + "%\n";
    body += "（設定予算：" + adjustedBudget + "円）\n";
    body += "・分析モード: " + formatWeeklyAnalysisModeForMessage(weeklyAnalysisMode) + "\n";
    body += "++++++++++++++++++++\n";
    body += "* 予算差分：" + differenceSign + differenceAbs + "円\n\n";
    body += buildSpecialExpenseReviewSection(specialExpenseReview, "今月");
    body += "\n";
    body += buildSaleIncomeSummarySection(saleIncomeEntries, saleIncomeTotal, actualWeeklyTotalAmount);
    body += "\n";
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
    body += "◆ 来週のユーザー補足メモ\n";
    body += nextWeekContextualMemoLines.length ? nextWeekContextualMemoLines.join("\n") + "\n" : "・なし\n";
    body += "\n";

    var aiAnalysis = analyzeExpensesWithAI(budgetTargetEntries, totalAmount, adjustedBudget, percentage, currentDate, weeklyAnalysisMode, {
        plannedExpenses: nextWeekPlannedExpenses,
        plannedExpenseLabel: "来週の予定支出",
        contextualMemos: nextWeekContextualMemos,
        contextualMemoLabel: "来週のユーザー補足メモ",
        saleIncomeEntries: saleIncomeEntries
    });
    body += buildAiSummarySection("◆ AI分析", aiAnalysis);

    if (action === 'mail') {
        upsertWeeklyBudgetCarryoverMemo(currentDate, difference, adjustedBudget, totalAmount, dateRangeStr);
    } else if (action !== 'mail') {
        Logger.log("mail送信以外のため前週予算差分メモの保存をスキップしました");
    }

    switch (action) {
        case 'mail':
            var emailAddress = getTargetEmailAddress();
            var subject = (isStaging ? "<test>" : "") + "家計簿週次レポート" + "（" + dateRangeStr + "）";
            var weeklySummaryBudgetCharts = createWeeklySummaryBudgetCharts(actualMonthlyTotalAmount, monthlyBudget, currentDate, actualWeeklyTotalAmount, adjustedBudget, approvedSpecialExpenseTotal, approvedWeeklySpecialExpenseTotal);
            MailApp.sendEmail({
                to: emailAddress,
                subject: subject,
                body: body,
                htmlBody: buildWeeklySummaryHtmlBody(body, actualMonthlyTotalAmount, monthlyBudget, currentDate, actualWeeklyTotalAmount, adjustedBudget, approvedSpecialExpenseTotal, approvedWeeklySpecialExpenseTotal),
                inlineImages: {
                    monthlyBudgetChart: weeklySummaryBudgetCharts.monthlyBudgetChart,
                    weeklyBudgetChart: weeklySummaryBudgetCharts.weeklyBudgetChart
                }
            });
            return "Successfully sent mail";
        case 'text':
            return body;
        default:
            throw new Error('actionが定義されていません');
    }
}

function getCurrentMonthExpenseEntries(currentDate) {
    var monthEntries = getMonthlyExpenseEntries(currentDate.getFullYear(), currentDate.getMonth());
    var endOfCurrentDate = new Date(currentDate.getFullYear(), currentDate.getMonth(), currentDate.getDate(), 23, 59, 59, 999);

    return monthEntries.filter(function (entry) {
        return entry.date <= endOfCurrentDate;
    });
}

function getCurrentMonthExpenseTotal(currentDate) {
    return calculateTotalAmount(getCurrentMonthExpenseEntries(currentDate));
}

function handleDailySummaryResult(currentDate, datesInWeek, adjustedBudget, isStaging, action) {
    var upcomingCalendarMemos = getCalendarMemoEntriesInRange(currentDate, getCurrentWeekEndExclusive(currentDate));
    var upcomingPlannedExpenses = filterPlannedExpenseMemos(upcomingCalendarMemos);
    var upcomingContextualMemos = filterContextualCalendarMemos(upcomingCalendarMemos);
    var upcomingExpenseLines = formatUpcomingPlannedExpenseLines(upcomingPlannedExpenses);
    var upcomingContextualMemoLines = formatUpcomingPlannedExpenseLines(upcomingContextualMemos);
    var plannedExpenseTotal = calculatePlannedExpenseTotal(upcomingPlannedExpenses);
    var weeklyBudgetCarryoverMemo = getWeeklyBudgetCarryoverMemoForWeek(currentDate);
    var weeklyAnalysisMode = getWeeklyAnalysisMode(currentDate);
    var datesUpToToday = datesInWeek.filter(function (date) {
        return date <= currentDate;
    });
    var rawDataEntries = getExpenseEntriesForDates(datesUpToToday).reverse();
    var saleIncomeEntries = getSaleIncomeEntriesForDates(datesUpToToday);
    var saleIncomeTotal = calculateTotalAmount(saleIncomeEntries);
    var specialExpenseReview = reviewSpecialExpensesWithAI(rawDataEntries);
    var budgetTargetEntries = getBudgetTargetEntries(rawDataEntries, specialExpenseReview);
    var actualTotalAmount = calculateTotalAmount(rawDataEntries);
    var approvedSpecialExpenseTotal = calculateApprovedSpecialExpenseTotal(specialExpenseReview);
    var dataEntries = rawDataEntries.map(function (entry) {
        var displayEntry = {
            date: entry.date,
            category: entry.category,
            name: entry.name,
            amount: entry.amount
        };
        if (displayEntry.name.length >= 16) {
            displayEntry.name = displayEntry.name.substring(0, 14) + "...";
        }
        return displayEntry;
    });
    var totalAmount = calculateTotalAmount(budgetTargetEntries);
    var percentage = (totalAmount / adjustedBudget) * 100;
    var projectedPercentage = adjustedBudget
        ? (((totalAmount + plannedExpenseTotal) / adjustedBudget) * 100).toFixed(2)
        : "0.00";
    var categoryRankingLines = getCategoryRankingLines(budgetTargetEntries);
    var uncategorizedCount = countUncategorizedEntries(budgetTargetEntries);
    var aiAnalysis = analyzeExpensesWithAI(budgetTargetEntries, totalAmount, adjustedBudget, percentage, currentDate, weeklyAnalysisMode, {
        plannedExpenses: upcomingPlannedExpenses,
        plannedExpenseLabel: "今後の予定支出",
        contextualMemos: upcomingContextualMemos,
        contextualMemoLabel: "今後のユーザー補足メモ",
        saleIncomeEntries: saleIncomeEntries
    });
    var subject = (isStaging ? "<test>" : "") + "家計簿日次レポート（" + formatDate(currentDate) + "）";
    var body = "+++ 💸 予算サマリー 💸 +++\n";

    body += formatDate(datesInWeek[0]) + " から " + formatDate(currentDate) + " までの予算対象支出: " + totalAmount + " 円\n";
    body += "実支出合計: " + actualTotalAmount + " 円\n";
    body += "承認済み特別費: " + approvedSpecialExpenseTotal + " 円\n";
    body += "今後の予定金額: " + plannedExpenseTotal + " 円\n";
    body += "支出＋予定の合計見込み: " + (totalAmount + plannedExpenseTotal) + " 円\n";
    body += "\n";
    body += "◆ 直近のユーザー補足メモ\n";
    body += upcomingContextualMemoLines.length ? upcomingContextualMemoLines.join("\n") + "\n" : "・なし\n";
    body += "\n";
    body += "予算に対して\n";
    body += "・予算対象支出: " + percentage.toFixed(2) + "%\n";
    body += "・合計見込み: " + projectedPercentage + "%\n";
    body += "（設定予算：" + adjustedBudget + "円）\n";
    body += "・分析モード: " + formatWeeklyAnalysisModeForMessage(weeklyAnalysisMode) + "\n";
    if (uncategorizedCount > 0) {
        body += "・未分類の支出: " + uncategorizedCount + "件\n";
    }
    body += "++++++++++++++++++++\n\n";
    body += buildSpecialExpenseReviewSection(specialExpenseReview, "今週");
    body += "\n";
    body += buildSaleIncomeSummarySection(saleIncomeEntries, saleIncomeTotal, actualTotalAmount);
    body += "\n";
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

    body += buildAiSummarySection("◆ AI分析", aiAnalysis);

    switch (action) {
        case 'mail':
            var emailAddress = getTargetEmailAddress();
            var dailyBudgetChart = createDailyBudgetChartBlob(actualTotalAmount, adjustedBudget, approvedSpecialExpenseTotal);
            MailApp.sendEmail({
                to: emailAddress,
                subject: subject,
                body: body,
                htmlBody: buildDailySummaryHtmlBody(body, actualTotalAmount, adjustedBudget, approvedSpecialExpenseTotal),
                inlineImages: {
                    dailyBudgetChart: dailyBudgetChart
                }
            });
            return "Successfully sent mail";
        case 'text':
            return body;
        default:
            throw new Error('actionが定義されていません');
    }
}

function buildSaleIncomeSummarySection(saleIncomeEntries, saleIncomeTotal, actualTotalAmount) {
    if (!saleIncomeEntries.length) {
        return "";
    }

    var lines = [
        "◆ 売却収入（分析の補足）",
        "・売却収入: " + saleIncomeTotal + "円",
        "・売却を差し引いた実質支出: " + (actualTotalAmount - saleIncomeTotal) + "円"
    ];

    saleIncomeEntries.forEach(function (entry) {
        lines.push("・" + formatDate(entry.date) + " - " + entry.name + ": " + entry.amount + "円");
    });
    lines.push("※予算対象支出・支出割合は、売却収入を差し引かずに計算しています。");

    return lines.join("\n") + "\n";
}

function filterPlannedExpenseMemos(calendarMemos) {
    return calendarMemos.filter(function (entry) {
        return entry.intent === "planned_expense";
    });
}

function filterContextualCalendarMemos(calendarMemos) {
    return calendarMemos.filter(function (entry) {
        return entry.intent === "contextual_note" || entry.intent === "reservation_info";
    });
}

function getCurrentWeekEndExclusive(currentDate) {
    var range = getWeekRange(currentDate);
    var endDateExclusive = new Date(range.endDate);
    endDateExclusive.setDate(endDateExclusive.getDate() + 1);
    return endDateExclusive;
}

function getNextWeekStartDate(currentDate) {
    var range = getWeekRange(currentDate);
    var startDate = new Date(range.endDate);
    startDate.setDate(startDate.getDate() + 1);
    return startDate;
}

function getWeekAfterNextStartDate(currentDate) {
    var startDate = getNextWeekStartDate(currentDate);
    var endDateExclusive = new Date(startDate);
    endDateExclusive.setDate(endDateExclusive.getDate() + 7);
    return endDateExclusive;
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
