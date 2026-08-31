const assert = require('node:assert/strict');
const fs = require('node:fs');
const test = require('node:test');
const vm = require('node:vm');

function loadDailySummaryMessageLayout() {
    const context = {
        formatWeeklyAnalysisModeForMessage: (mode) => mode.label,
        formatWeeklyBudgetCarryoverSummaryForMessage: (memo) => '・前週は予算を' + memo.difference + '円超過していました。',
        buildSpecialExpenseReviewSection: () => '◆ 今週の特別費チェック\n・確認あり\n',
        buildSaleIncomeSummarySection: () => '◆ 売却収入（分析の補足）\n・売却収入: 3000円\n',
        formatDate: (date) => (date.getMonth() + 1) + '/' + date.getDate(),
        WEEKLY_ANALYSIS_MODE_FRUGAL: 'frugal',
        Date,
        Object
    };
    vm.runInNewContext(fs.readFileSync('scripts/formatting/summaryMessageFormatter.js', 'utf8'), context);
    return context;
}

test('今日の判断は注意事項を最大3件に絞り、予定込み着地を先頭に置く', () => {
    const formatter = loadDailySummaryMessageLayout();
    const section = formatter.buildDailySummaryDecisionSection({
        totalAmount: 18000,
        actualTotalAmount: 19000,
        plannedExpenseTotal: 4000,
        adjustedBudget: 20000,
        percentage: 90,
        projectedPercentage: '110.00',
        weeklyAnalysisMode: { label: '節制モード' },
        weeklyBudgetCarryoverMemo: { difference: 3000 },
        categorySpendingIncreaseSummary: {
            increases: [
                { category: '外食', difference: 2000 },
                { category: '日用品', difference: 1800 }
            ]
        },
        uncategorizedCount: 1,
        approvedSpecialExpenseTotal: 5000
    });

    assert.match(section, /^◆ 今日の判断\n・予定込み着地: 22000円 \/ 20000円（110.00%）/);
    assert.equal((section.match(/^・/gm) || []).length, 6);
    assert.doesNotMatch(section, /外食/);
});

test('確認事項がない日は落ち着いている旨を表示する', () => {
    const formatter = loadDailySummaryMessageLayout();
    const section = formatter.buildDailySummaryDecisionSection({
        totalAmount: 5000,
        actualTotalAmount: 5000,
        plannedExpenseTotal: 1000,
        adjustedBudget: 20000,
        percentage: 25,
        projectedPercentage: '30.00',
        weeklyAnalysisMode: { label: '通常' },
        weeklyBudgetCarryoverMemo: null,
        categorySpendingIncreaseSummary: { increases: [] },
        uncategorizedCount: 0,
        approvedSpecialExpenseTotal: 0
    });

    assert.match(section, /現時点で特に確認が必要な項目はありません/);
    assert.doesNotMatch(section, /注意:\n/);
});

test('詳細・確認用は内容のないセクションを表示しない', () => {
    const formatter = loadDailySummaryMessageLayout();
    const section = formatter.buildDailySummaryDetailsSection({
        upcomingContextualMemoLines: [],
        specialExpenseReview: { hasCandidates: false },
        saleIncomeEntries: [],
        saleIncomeTotal: 0,
        actualTotalAmount: 0,
        weeklyBudgetCarryoverMemo: null,
        categoryRankingLines: ['・1位 食費: 3000円'],
        dataEntries: [{ date: new Date(2026, 8, 1), name: 'スーパー', amount: 3000 }],
        upcomingExpenseLines: []
    });

    assert.match(section, /◆ 詳細・確認用\n◆ カテゴリ別支出ランキング/);
    assert.match(section, /◆ 支出詳細\n・9\/1 - スーパー: 3000円/);
    assert.doesNotMatch(section, /なし|特別費|予定支出|補足メモ/);
});
