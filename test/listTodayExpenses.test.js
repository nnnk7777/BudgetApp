const assert = require('node:assert/strict');
const fs = require('node:fs');
const test = require('node:test');
const vm = require('node:vm');

function loadListTodayExpenses(options) {
    const context = {
        JSON,
        getScriptRuntimeContext: () => ({ currentDate: options.currentDate }),
        getExpenseEntriesForDates: options.getExpenseEntriesForDates
    };
    vm.runInNewContext(
        fs.readFileSync('scripts/application/listTodayExpenses.js', 'utf8'),
        context
    );
    return context;
}

test('実行時日付に紐づく支出の確認に必要な項目だけを返す', () => {
    const currentDate = new Date(2026, 8, 3);
    let requestedDates;
    const context = loadListTodayExpenses({
        currentDate,
        getExpenseEntriesForDates: (dates) => {
            requestedDates = dates;
            return [
                { date: currentDate, category: '交通費', name: '電車', amount: 639, row: 40 },
                { date: currentDate, category: '外食', name: 'ランチ', amount: 1200, row: 41 }
            ];
        }
    });

    const result = JSON.parse(context.listTodayExpenses());

    assert.equal(requestedDates.length, 1);
    assert.equal(requestedDates[0], currentDate);
    assert.deepEqual(result, {
        ok: true,
        items: [
            { category: '交通費', name: '電車', amount: 639 },
            { category: '外食', name: 'ランチ', amount: 1200 }
        ]
    });
});

test('今日の支出がない場合は空配列を返す', () => {
    const context = loadListTodayExpenses({
        currentDate: new Date(2026, 8, 3),
        getExpenseEntriesForDates: () => []
    });

    assert.deepEqual(JSON.parse(context.listTodayExpenses()), {
        ok: true,
        items: []
    });
});
