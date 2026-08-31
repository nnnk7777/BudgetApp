const assert = require('node:assert/strict');
const fs = require('node:fs');
const test = require('node:test');
const vm = require('node:vm');

function loadCategorySpendingBaseline(getEntries) {
    const context = {
        Date,
        Math,
        Object,
        calculateCategoryTotals: (entries) => entries.reduce((totals, entry) => {
            totals[entry.category] = (totals[entry.category] || 0) + entry.amount;
            return totals;
        }, {}),
        getExpenseEntriesForDates: getEntries,
        getBudgetTargetEntriesWithCachedSpecialExpenseDecisions: (entries) => entries.filter((entry) => !entry.approvedSpecial)
    };
    vm.runInNewContext(fs.readFileSync('scripts/utils/categorySpendingBaseline.js', 'utf8'), context);
    return context;
}

function entry(category, amount, approvedSpecial) {
    return { category, amount, approvedSpecial: Boolean(approvedSpecial) };
}

test('同じ週進捗の過去4週平均と比べ、大きく増えたカテゴリだけを返す', () => {
    const context = loadCategorySpendingBaseline((dates) => {
        const weekStartDay = dates[0].getDate();
        const entriesByWeekStartDay = {
            10: [entry('外食', 2000), entry('日用品', 500)],
            3: [entry('外食', 3000), entry('日用品', 500)],
            27: [entry('外食', 4000), entry('日用品', 500)],
            20: [entry('外食', 3000), entry('日用品', 500)]
        };
        return entriesByWeekStartDay[weekStartDay] || [];
    });
    const summary = context.getCategorySpendingIncreaseSummary(
        [entry('外食', 8000), entry('日用品', 800)],
        [new Date(2026, 7, 17), new Date(2026, 7, 18), new Date(2026, 7, 19)]
    );

    assert.equal(summary.completedWeeks, 4);
    assert.equal(JSON.stringify(summary.increases), JSON.stringify([
        { category: '外食', currentAmount: 8000, averageAmount: 3000, difference: 5000, increaseRatio: 5 / 3 }
    ]));
});

test('年またぎで同年内の比較週が3週未満なら比較データ不足にする', () => {
    const context = loadCategorySpendingBaseline(() => [entry('外食', 1000)]);
    const summary = context.getCategorySpendingIncreaseSummary(
        [entry('外食', 5000)],
        [new Date(2026, 0, 22), new Date(2026, 0, 23), new Date(2026, 0, 24)]
    );

    assert.equal(summary.completedWeeks, 3);
    assert.equal(summary.increases.length, 1);

    const insufficientSummary = context.getCategorySpendingIncreaseSummary(
        [entry('外食', 5000)],
        [new Date(2026, 0, 15), new Date(2026, 0, 16), new Date(2026, 0, 17)]
    );

    assert.equal(insufficientSummary.completedWeeks, 2);
    assert.match(context.formatCategorySpendingIncreaseSummary(insufficientSummary), /比較データ不足/);
});

test('履歴の承認済み特別費は平均から除外し、未承認分は予算対象に残す', () => {
    const context = loadCategorySpendingBaseline(() => [
        entry('特別費', 20000, true),
        entry('特別費', 1000, false),
        entry('外食', 1000)
    ]);
    const summary = context.getCategorySpendingIncreaseSummary(
        [entry('特別費', 4000), entry('外食', 1000)],
        [new Date(2026, 7, 17)]
    );

    assert.equal(summary.completedWeeks, 4);
    assert.equal(summary.increases[0].category, '特別費');
    assert.equal(summary.increases[0].averageAmount, 1000);
});
