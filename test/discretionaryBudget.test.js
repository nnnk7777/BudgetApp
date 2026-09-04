const assert = require('node:assert/strict');
const fs = require('node:fs');
const test = require('node:test');
const vm = require('node:vm');

function loadExpenseSummaryUtils() {
    const context = { Date, Math, Object, String, parseFloat, parseInt, isNaN };
    vm.runInNewContext(fs.readFileSync('scripts/utils/summaryDateUtils.js', 'utf8'), context);
    vm.runInNewContext(fs.readFileSync('scripts/utils/expenseSummaryUtils.js', 'utf8'), context);
    return context;
}

function loadExpenseSummaryAi(overrides = {}) {
    const calls = [];
    const context = {
        Date,
        Math,
        Object,
        String,
        WEEKLY_ANALYSIS_MODE_FRUGAL: 'frugal',
        Logger: { log: () => {} },
        getUpcomingPlannedExpenses: () => [],
        calculateTotalAmount: () => 0,
        getCategoryRankingLines: () => [],
        getWeeklyBudgetCarryoverMemoForWeek: () => null,
        buildWeeklyBudgetCarryoverGuidanceForPrompt: () => '',
        buildWeeklyAnalysisModeGuidanceForPrompt: () => '',
        formatDate: (date) => (date.getMonth() + 1) + '/' + date.getDate(),
        getWeekRange: (date) => ({ startDate: date, endDate: date }),
        buildAnalysisDateContext: () => '分析日テスト',
        formatWeeklyAnalysisModeForPrompt: () => '通常',
        formatWeeklyBudgetCarryoverMemoForPrompt: () => 'なし',
        generatePreferredAiText: (...args) => {
            calls.push(args);
            return { text: 'result', provider: 'openai' };
        },
        ...overrides
    };
    vm.runInNewContext(fs.readFileSync('scripts/domain/ai/expenseSummaryPrompts.js', 'utf8'), context);
    vm.runInNewContext(fs.readFileSync('scripts/infrastructure/ai/expenseSummaryAi.js', 'utf8'), context);
    return { context, calls };
}

test('金額不明の予定は予定支出合計から差し引かない', () => {
    const utils = loadExpenseSummaryUtils();
    const total = utils.calculatePlannedExpenseTotal([
        { title: '友人と食事', memo: '金額未定' },
        { title: '買い物', memo: '2,000円' }
    ]);

    assert.equal(total, 2000);
});

test('週次分析は既存の1回のAI呼び出しに行動ルール指定を含める', () => {
    const { context, calls } = loadExpenseSummaryAi();

    context.analyzeExpensesWithAI([], 10000, 20000, 50, new Date(2026, 8, 6), { mode: 'normal' }, {
        plannedExpenses: [],
        includeActionRule: true
    });

    assert.equal(calls.length, 1);
    assert.match(calls[0][0], /最後の行は必ず「行動ルール: 」で始め/);
});

test('日次分析には週次専用の行動ルール指定を追加しない', () => {
    const { context, calls } = loadExpenseSummaryAi();

    context.analyzeExpensesWithAI([], 10000, 20000, 50, new Date(2026, 8, 3), { mode: 'normal' }, {
        plannedExpenses: []
    });

    assert.equal(calls.length, 1);
    assert.doesNotMatch(calls[0][0], /最後の行は必ず「行動ルール: 」で始め/);
});
