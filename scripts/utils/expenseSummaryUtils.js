function calculateTotalAmount(dataEntries) {
    var total = 0;
    dataEntries.forEach(function (entry) {
        var amount = parseFloat(entry.amount);
        if (!isNaN(amount)) {
            total += amount;
        }
    });
    return total;
}

function calculateCategoryTotals(dataEntries) {
    var totals = {};
    dataEntries.forEach(function (entry) {
        var key = entry.category || "未分類";
        var amount = parseFloat(entry.amount);
        if (isNaN(amount)) {
            return;
        }
        if (!totals[key]) {
            totals[key] = 0;
        }
        totals[key] += amount;
    });
    return totals;
}

function getCategoryRankingLines(dataEntries) {
    var categoryTotals = calculateCategoryTotals(dataEntries);
    return Object.keys(categoryTotals)
        .sort(function (a, b) {
            return categoryTotals[b] - categoryTotals[a];
        })
        .map(function (category, index) {
            return "・" + (index + 1) + "位 " + category + ": " + categoryTotals[category] + "円";
        });
}

function countUncategorizedEntries(dataEntries) {
    return dataEntries.filter(function (entry) {
        return !entry.category || String(entry.category).trim() === "";
    }).length;
}

function calculatePlannedExpenseTotal(plannedExpenses) {
    var total = 0;

    plannedExpenses.forEach(function (entry) {
        var memo = entry.memo || "";
        var matches = memo.match(/([0-9,]+)\s*円/g) || [];

        matches.forEach(function (match) {
            var amount = parseInt(match.replace(/[^\d]/g, ""), 10);
            if (!isNaN(amount)) {
                total += amount;
            }
        });
    });

    return total;
}

function getTopExpenseEntries(dataEntries, limit) {
    var topEntries = dataEntries.slice();
    topEntries.sort(function (a, b) {
        return parseFloat(b.amount) - parseFloat(a.amount);
    });
    return topEntries.slice(0, limit);
}
