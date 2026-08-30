const assert = require('node:assert/strict');
const fs = require('node:fs');
const test = require('node:test');
const vm = require('node:vm');

function loadBudgetConfig() {
    const context = {};
    vm.runInNewContext(
        fs.readFileSync('scripts/config/budget.js', 'utf8'),
        context
    );
    return context;
}

test('週予算の設定値から日次・週次・月次の予算を計算する', () => {
    const budget = loadBudgetConfig();

    assert.equal(budget.getWeeklyBudget(), 20000);
    assert.equal(budget.calculateBudgetForDays(7), 20000);
    assert.equal(budget.calculateMonthlyBudgetForDate(new Date(2026, 8, 1)), 85700);
});

test('週予算を変更すると月次予算も同じ設定値で再計算される', () => {
    const budget = loadBudgetConfig();

    budget.WEEKLY_BUDGET = 40000;

    assert.equal(budget.calculateBudgetForDays(7), 40000);
    assert.equal(budget.calculateMonthlyBudgetForDate(new Date(2026, 8, 1)), 171400);
});
