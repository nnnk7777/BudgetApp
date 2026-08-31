var CATEGORY_SPENDING_BASELINE_WEEKS = 4;
var CATEGORY_SPENDING_BASELINE_MIN_WEEKS = 3;
var CATEGORY_SPENDING_INCREASE_MIN_YEN = 1500;
var CATEGORY_SPENDING_INCREASE_MIN_RATIO = 0.3;

function getCategorySpendingIncreaseSummary(currentEntries, comparisonDates) {
    var currentYear;
    var currentCategoryTotals;
    var historicalCategoryTotals = {};
    var completedWeeks = 0;
    var weeksAgo;

    if (!comparisonDates.length) {
        return createCategorySpendingIncreaseSummary([], 0, {});
    }

    currentYear = comparisonDates[0].getFullYear();
    currentCategoryTotals = calculateCategoryTotals(currentEntries);

    for (weeksAgo = 1; weeksAgo <= CATEGORY_SPENDING_BASELINE_WEEKS; weeksAgo++) {
        var historicalDates = getDatesShiftedByWeeks(comparisonDates, weeksAgo);
        var historicalEntries;
        var historicalCategoryTotalsForWeek;

        if (!historicalDates.length || !areDatesInYear(historicalDates, currentYear)) {
            continue;
        }

        historicalEntries = getBudgetTargetEntriesWithCachedSpecialExpenseDecisions(
            getExpenseEntriesForDates(historicalDates)
        );
        historicalCategoryTotalsForWeek = calculateCategoryTotals(historicalEntries);
        addCategoryTotals(historicalCategoryTotals, historicalCategoryTotalsForWeek);
        completedWeeks++;
    }

    return createCategorySpendingIncreaseSummary(
        getCategorySpendingIncreases(currentCategoryTotals, historicalCategoryTotals, completedWeeks),
        completedWeeks,
        currentCategoryTotals
    );
}

function getDatesShiftedByWeeks(dates, weeksAgo) {
    return dates.map(function (date) {
        var shiftedDate = new Date(date.getFullYear(), date.getMonth(), date.getDate());
        shiftedDate.setDate(shiftedDate.getDate() - weeksAgo * 7);
        return shiftedDate;
    });
}

function areDatesInYear(dates, year) {
    return dates.every(function (date) {
        return date.getFullYear() === year;
    });
}

function addCategoryTotals(targetTotals, sourceTotals) {
    Object.keys(sourceTotals).forEach(function (category) {
        targetTotals[category] = (targetTotals[category] || 0) + sourceTotals[category];
    });
}

function getCategorySpendingIncreases(currentCategoryTotals, historicalCategoryTotals, completedWeeks) {
    if (completedWeeks < CATEGORY_SPENDING_BASELINE_MIN_WEEKS) {
        return [];
    }

    return Object.keys(currentCategoryTotals)
        .map(function (category) {
            var currentAmount = currentCategoryTotals[category];
            var averageAmount = Math.round((historicalCategoryTotals[category] || 0) / completedWeeks);
            var difference = currentAmount - averageAmount;
            var increaseRatio = averageAmount ? difference / averageAmount : null;

            return {
                category: category,
                currentAmount: currentAmount,
                averageAmount: averageAmount,
                difference: difference,
                increaseRatio: increaseRatio
            };
        })
        .filter(function (entry) {
            return (
                entry.difference >= CATEGORY_SPENDING_INCREASE_MIN_YEN &&
                (entry.increaseRatio === null || entry.increaseRatio >= CATEGORY_SPENDING_INCREASE_MIN_RATIO)
            );
        })
        .sort(function (left, right) {
            return right.difference - left.difference;
        })
        .slice(0, 3);
}

function createCategorySpendingIncreaseSummary(increases, completedWeeks, currentCategoryTotals) {
    return {
        increases: increases,
        completedWeeks: completedWeeks,
        currentCategoryTotals: currentCategoryTotals
    };
}

function formatCategorySpendingIncreaseSummary(summary) {
    if (summary.completedWeeks < CATEGORY_SPENDING_BASELINE_MIN_WEEKS) {
        return "・比較データ不足（同年内の過去" + CATEGORY_SPENDING_BASELINE_MIN_WEEKS + "週分がそろうまで表示しません）";
    }

    if (!summary.increases.length) {
        return "・普段より大きく増えたカテゴリはありません。";
    }

    return summary.increases
        .map(function (entry) {
            return "・" + entry.category + ": 平均 " + entry.averageAmount + "円 / +" + entry.difference + "円";
        })
        .join("\n");
}
