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
    var categorySpendingIncreaseSummary;
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
    categorySpendingIncreaseSummary = getCategorySpendingIncreaseSummary(budgetTargetEntries, getDatesInWeek(currentDate));
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
    body += "◆ 普段より増えたカテゴリ\n";
    body += formatCategorySpendingIncreaseSummary(categorySpendingIncreaseSummary) + "\n\n";
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
    var categorySpendingIncreaseSummary = getCategorySpendingIncreaseSummary(budgetTargetEntries, datesUpToToday);
    var uncategorizedCount = countUncategorizedEntries(budgetTargetEntries);
    var aiAnalysis = analyzeExpensesWithAI(budgetTargetEntries, totalAmount, adjustedBudget, percentage, currentDate, weeklyAnalysisMode, {
        plannedExpenses: upcomingPlannedExpenses,
        plannedExpenseLabel: "今後の予定支出",
        contextualMemos: upcomingContextualMemos,
        contextualMemoLabel: "今後のユーザー補足メモ",
        saleIncomeEntries: saleIncomeEntries
    });
    var subject = (isStaging ? "<test>" : "") + "家計簿日次レポート（" + formatDate(currentDate) + "）";
    var body = buildDailySummaryDecisionSection({
        totalAmount: totalAmount,
        actualTotalAmount: actualTotalAmount,
        plannedExpenseTotal: plannedExpenseTotal,
        adjustedBudget: adjustedBudget,
        percentage: percentage,
        projectedPercentage: projectedPercentage,
        weeklyAnalysisMode: weeklyAnalysisMode,
        weeklyBudgetCarryoverMemo: weeklyBudgetCarryoverMemo,
        categorySpendingIncreaseSummary: categorySpendingIncreaseSummary,
        uncategorizedCount: uncategorizedCount,
        approvedSpecialExpenseTotal: approvedSpecialExpenseTotal
    });

    body += "\n\n" + buildAiSummarySection("◆ AI分析", aiAnalysis);
    body += "\n\n" + buildDailySummaryDetailsSection({
        upcomingContextualMemoLines: upcomingContextualMemoLines,
        specialExpenseReview: specialExpenseReview,
        saleIncomeEntries: saleIncomeEntries,
        saleIncomeTotal: saleIncomeTotal,
        actualTotalAmount: actualTotalAmount,
        weeklyBudgetCarryoverMemo: weeklyBudgetCarryoverMemo,
        categoryRankingLines: categoryRankingLines,
        dataEntries: dataEntries,
        upcomingExpenseLines: upcomingExpenseLines
    });

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

function buildDailySummaryDecisionSection(options) {
    var projectedTotalAmount = options.totalAmount + options.plannedExpenseTotal;
    var signals = [];
    var lines = [
        "◆ 今日の判断",
        "・予定込み着地: " + projectedTotalAmount + "円 / " + options.adjustedBudget + "円（" + options.projectedPercentage + "%）",
        "・今週の実績: " + options.totalAmount + "円（" + options.percentage.toFixed(2) + "%）",
        "・分析モード: " + formatWeeklyAnalysisModeForMessage(options.weeklyAnalysisMode)
    ];

    if (projectedTotalAmount > options.adjustedBudget) {
        signals.push("・予定込みで予算を" + (projectedTotalAmount - options.adjustedBudget) + "円超過する見込みです。");
    }
    if (options.weeklyBudgetCarryoverMemo && options.weeklyBudgetCarryoverMemo.difference > 0) {
        signals.push(formatWeeklyBudgetCarryoverSummaryForMessage(options.weeklyBudgetCarryoverMemo));
    }
    if (options.uncategorizedCount > 0) {
        signals.push("・未分類の支出: " + options.uncategorizedCount + "件");
    }
    if (options.approvedSpecialExpenseTotal > 0) {
        signals.push("・今週の承認済み特別費: " + options.approvedSpecialExpenseTotal + "円");
    }
    (options.categorySpendingIncreaseSummary.increases || []).forEach(function (entry) {
        signals.push("・" + entry.category + ": 同曜日時点の平均より +" + entry.difference + "円");
    });

    if (!signals.length) {
        lines.push("・注意: 現時点で特に確認が必要な項目はありません。");
        return lines.join("\n");
    }

    lines.push("注意:");
    return lines.concat(signals.slice(0, 3)).join("\n");
}

function buildDailySummaryDetailsSection(options) {
    var lines = ["◆ 詳細・確認用"];

    if (options.upcomingContextualMemoLines.length) {
        lines.push("◆ 直近のユーザー補足メモ");
        lines = lines.concat(options.upcomingContextualMemoLines);
    }
    if (options.specialExpenseReview.hasCandidates) {
        lines.push(buildSpecialExpenseReviewSection(options.specialExpenseReview, "今週").trim());
    }
    if (options.saleIncomeEntries.length) {
        lines.push(buildSaleIncomeSummarySection(options.saleIncomeEntries, options.saleIncomeTotal, options.actualTotalAmount).trim());
    }
    if (options.weeklyBudgetCarryoverMemo && options.weeklyBudgetCarryoverMemo.difference <= 0) {
        lines.push("◆ 前週からの持ち越し");
        lines.push(formatWeeklyBudgetCarryoverSummaryForMessage(options.weeklyBudgetCarryoverMemo));
    }
    if (options.categoryRankingLines.length) {
        lines.push("◆ カテゴリ別支出ランキング");
        lines = lines.concat(options.categoryRankingLines);
    }
    if (options.dataEntries.length) {
        lines.push("◆ 支出詳細");
        options.dataEntries.forEach(function (entry) {
            lines.push("・" + formatDate(entry.date) + " - " + entry.name + ": " + entry.amount + "円");
        });
    }
    if (options.upcomingExpenseLines.length) {
        lines.push("◆ 直近の予定支出");
        lines = lines.concat(options.upcomingExpenseLines);
    }

    return lines.join("\n");
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
