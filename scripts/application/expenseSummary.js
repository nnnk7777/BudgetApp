// 行いたい操作を引数actionで受け取る
function calculateExpensesSummary(action) {
    return calculateAutomaticSummaries(action);
}

function calculateAutomaticSummaries(action) {
    var runtimeContext = getScriptRuntimeContext();
    var currentDate = runtimeContext.currentDate;
    var summaryUnits = ['daily'];
    var results = [];

    if (isSunday(currentDate)) {
        summaryUnits.push('weekly');
    }

    if (isLastDayOfMonth(currentDate)) {
        summaryUnits.push('monthly');
    }

    summaryUnits.forEach(function (unit) {
        results.push(runSummaryByUnit(unit, action));
    });

    return formatSummaryExecutionResult(summaryUnits, action, results);
}

function calculateDailySummary(action) {
    return runSummaryByUnit('daily', action);
}

function calculateWeeklySummary(action) {
    return runSummaryByUnit('weekly', action);
}

function calculateMonthlySummaryByUnit(action) {
    return runSummaryByUnit('monthly', action);
}

function runSummaryByUnit(unit, action) {
    switch (unit) {
        case 'daily':
            return buildDailySummaryResult(action);
        case 'weekly':
            return buildWeeklySummaryResult(action);
        case 'monthly':
            return calculateMonthlySummary(action);
        default:
            throw new Error('summary unitが定義されていません');
    }
}

function buildDailySummaryResult(action) {
    var budgetPerWeek = 40000;
    var runtimeContext = getScriptRuntimeContext();
    var currentDate = runtimeContext.currentDate;
    var isStaging = runtimeContext.isStaging;
    var datesInWeek = getDatesInWeek(currentDate);
    var adjustedBudget = Math.round((budgetPerWeek * datesInWeek.length / 7) / 100) * 100;

    return handleDailySummaryResult(currentDate, datesInWeek, adjustedBudget, isStaging, action);
}

function buildWeeklySummaryResult(action) {
    var budgetPerWeek = 40000;
    var runtimeContext = getScriptRuntimeContext();
    var currentDate = runtimeContext.currentDate;
    var isStaging = runtimeContext.isStaging;
    var datesInWeek = getDatesInWeek(currentDate);
    var startOfWeek = datesInWeek[0];
    var endOfWeek = datesInWeek[datesInWeek.length - 1];
    var dateRangeStr = formatDate(startOfWeek) + "〜" + formatDate(endOfWeek);
    var dataEntries = getExpenseEntriesForDates(datesInWeek);
    var totalAmount = calculateTotalAmount(dataEntries);
    var adjustedBudget = Math.round((budgetPerWeek * datesInWeek.length / 7) / 100) * 100;
    var difference = totalAmount - adjustedBudget;
    var percentage = adjustedBudget ? (totalAmount / adjustedBudget) * 100 : 0;

    Logger.log("データエントリ一覧:");
    dataEntries.forEach(function (entry) {
        Logger.log("日付: " + formatDate(entry.date) + ", 名称: " + entry.name + ", 金額: " + entry.amount);
    });
    Logger.log(dateRangeStr + " の合計金額: " + totalAmount + "円");
    if (difference > 0) {
        Logger.log("予算を " + difference + " 円上回りました。");
    } else {
        Logger.log("予算を " + Math.abs(difference) + " 円下回りました。");
    }
    Logger.log("予算の " + percentage.toFixed(2) + "% を使用しました。");

    return handleWeeklySummaryResult(
        dateRangeStr,
        totalAmount,
        dataEntries,
        difference,
        percentage,
        adjustedBudget,
        isStaging,
        action,
        currentDate
    );
}

function formatSummaryExecutionResult(summaryUnits, action, results) {
    if (action === 'text') {
        return results.join("\n\n==============================\n\n");
    }

    return "Successfully sent " + summaryUnits.join(', ') + " summary mail";
}
