const assert = require('node:assert/strict');
const fs = require('node:fs');
const test = require('node:test');
const vm = require('node:vm');

function createEvent(title, date, options = {}) {
    const start = new Date(date.getFullYear(), date.getMonth(), date.getDate(), options.hour || 0);
    const end = new Date(start);
    end.setHours(start.getHours() + (options.durationHours || 24));

    return {
        getTitle: () => title,
        getDescription: () => options.description || '',
        getStartTime: () => start,
        getEndTime: () => end
    };
}

function loadCommuteExpense() {
    const context = {
        Date,
        Math,
        String,
        parseFloat,
        isNaN,
        Logger: { log: () => {} },
        calculateDateDistanceInDays: (left, right) => {
            const leftDay = new Date(left.getFullYear(), left.getMonth(), left.getDate());
            const rightDay = new Date(right.getFullYear(), right.getMonth(), right.getDate());
            return Math.round((leftDay - rightDay) / 86400000);
        }
    };
    vm.runInNewContext(fs.readFileSync('scripts/config/commute.js', 'utf8'), context);
    vm.runInNewContext(fs.readFileSync('scripts/domain/commuteExpense.js', 'utf8'), context);
    return context;
}

test('月水木の平日だけ通勤費1278円を予定支出にする', () => {
    const context = loadCommuteExpense();
    const expenses = context.buildAutomaticCommutePlannedExpenses(
        new Date(2026, 8, 7),
        new Date(2026, 8, 14),
        [],
        []
    );

    assert.equal(JSON.stringify(expenses.map((entry) => entry.date.getDay())), JSON.stringify([1, 3, 4]));
    assert.equal(expenses.every((entry) => entry.memo === '通勤費 1,278円'), true);
    assert.equal(expenses.every((entry) => entry.intent === 'planned_expense'), true);
});

test('祝日と全休の日は除外し、休日出勤があれば全休を上書きする', () => {
    const context = loadCommuteExpense();
    const monday = new Date(2026, 8, 7);
    const wednesday = new Date(2026, 8, 9);
    const thursday = new Date(2026, 8, 10);
    const expenses = context.buildAutomaticCommutePlannedExpenses(
        monday,
        new Date(2026, 8, 14),
        [createEvent('全休', wednesday), createEvent('休日出勤', wednesday), createEvent('有休', thursday)],
        [createEvent('成人の日', monday)]
    );

    assert.equal(JSON.stringify(expenses.map((entry) => entry.date.getDate())), JSON.stringify([9]));
});

test('祝日カレンダー上の行事でも国民の祝日でなければ除外しない', () => {
    const context = loadCommuteExpense();
    const monday = new Date(2026, 1, 2);
    const expenses = context.buildAutomaticCommutePlannedExpenses(
        monday,
        new Date(2026, 1, 3),
        [],
        [createEvent('節分', monday)]
    );

    assert.equal(expenses.length, 1);
});

test('午前休・午後休・半休も通勤なしとして扱う', () => {
    const context = loadCommuteExpense();
    const monday = new Date(2026, 8, 7);
    const expenses = context.buildAutomaticCommutePlannedExpenses(
        monday,
        new Date(2026, 8, 14),
        [
            createEvent('午前有休', monday),
            createEvent('午後休', new Date(2026, 8, 9)),
            createEvent('半休', new Date(2026, 8, 10))
        ],
        []
    );

    assert.equal(expenses.length, 0);
});

test('記録済み交通費を通勤費見込みから差し引き、全額記録済みなら除外する', () => {
    const context = loadCommuteExpense();
    const date = new Date(2026, 8, 7);
    const plannedExpense = {
        title: '通勤費（自動見込み）',
        date,
        memo: '通勤費 1,278円',
        intent: 'planned_expense',
        source: 'automatic_commute'
    };
    const partial = context.adjustAutomaticCommuteExpenseForRecordedEntries(plannedExpense, [
        { date, category: '交通費', name: '通勤（往路）', amount: 639 }
    ]);
    const recorded = context.adjustAutomaticCommuteExpenseForRecordedEntries(plannedExpense, [
        { date, category: '交通費', name: '通勤（往路）', amount: 639 },
        { date, category: '交通費', name: '通勤（復路）', amount: 639 }
    ]);

    assert.equal(partial.memo, '通勤費 残り見込み 639円');
    assert.equal(recorded, null);
});

test('交通費カテゴリでも名称に通勤を含まない移動は通勤費から差し引かない', () => {
    const context = loadCommuteExpense();
    const date = new Date(2026, 8, 7);
    const plannedExpense = {
        title: '通勤費（自動見込み）',
        date,
        memo: '通勤費 1,278円',
        intent: 'planned_expense',
        source: 'automatic_commute'
    };
    const adjusted = context.adjustAutomaticCommuteExpenseForRecordedEntries(plannedExpense, [
        { date, category: '交通費', name: '吉祥寺→渋谷', amount: 420 },
        { date, category: '日用品', name: '通勤用バッグ', amount: 5000 }
    ]);

    assert.equal(adjusted.memo, '通勤費 1,278円');
});
