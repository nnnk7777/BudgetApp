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

test('今日の判断は裁量枠を先頭に置き、注意事項を最大3件に絞る', () => {
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

    assert.match(section, /^◆ 今日の判断\n・裁量枠: 0円（予定込みで2000円超過見込み）\n・予定込み着地: 22000円 \/ 20000円（110.00%）/);
    assert.equal((section.match(/^・/gm) || []).length, 7);
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
    assert.match(section, /・裁量枠: 14000円/);
    assert.doesNotMatch(section, /注意:\n/);
});

test('来週の見込みは週予算から金額入り予定支出だけを差し引く', () => {
    const formatter = loadDailySummaryMessageLayout();
    const section = formatter.buildNextWeekBudgetOutlookSection(20000, 11834);

    assert.equal(section, [
        '◆ 来週の見込み',
        '・週予算: 20000円',
        '・予定支出: 11834円',
        '・来週の裁量見込み: 8166円'
    ].join('\n'));
});

test('来週の予定支出が予算を超える場合は裁量見込みを0円にする', () => {
    const formatter = loadDailySummaryMessageLayout();
    const section = formatter.buildNextWeekBudgetOutlookSection(20000, 22500);

    assert.match(section, /・来週の裁量見込み: 0円（予定込みで2500円超過見込み）/);
});

test('既存AI分析の行動ルールを独立表示し、AI分析本文から重複行を除く', () => {
    const formatter = loadDailySummaryMessageLayout();
    const analysis = {
        text: '📊 現状\n・外食が増えています。\n行動ルール: 予定外の外食は1回までにする。',
        provider: 'openai'
    };
    const section = formatter.buildWeeklyActionRuleSection(analysis);
    const cleaned = formatter.removeExplicitActionRuleFromAiAnalysis(analysis);

    assert.equal(section, '◆ 今週の一つの行動ルール\n・予定外の外食は1回までにする。');
    assert.doesNotMatch(cleaned.text, /行動ルール/);
    assert.match(cleaned.text, /外食が増えています/);
});

test('固定形式でないAI応答でも次の行動セクションからルールを抽出する', () => {
    const formatter = loadDailySummaryMessageLayout();
    const actionRule = formatter.extractActionRuleFromAiAnalysis({
        text: '📊 現状\n・予算内です。\n✅ 次の行動\n・コンビニ利用は週1回までにする。'
    });

    assert.equal(actionRule, 'コンビニ利用は週1回までにする。');
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
