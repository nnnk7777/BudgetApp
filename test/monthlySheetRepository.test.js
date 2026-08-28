const assert = require('node:assert/strict');
const fs = require('node:fs');
const test = require('node:test');
const vm = require('node:vm');

function loadMonthlySheetRepository(rows) {
    const context = {
        getExpenseSheet: () => ({
            getRange: (startRow, dateCol, rowCount, columnCount) => {
                assert.equal(startRow, 20);
                assert.equal(rowCount, 14);
                assert.equal(columnCount, 4);
                return { getValues: () => rows };
            }
        }),
        getColumnsForMonth: () => ({ dateCol: 31 }),
        parseDate: (value, year) => {
            const match = String(value).match(/^(\d{1,2})\/(\d{1,2})$/);
            return match
                ? new Date(year, Number(match[1]) - 1, Number(match[2]))
                : new Date(NaN);
        },
        Date,
        Object,
        isNaN
    };
    vm.runInNewContext(
        fs.readFileSync('scripts/infrastructure/gas/monthlySheetRepository.js', 'utf8'),
        context
    );
    return context;
}

test('月次収入は予定・空行をまたいでも収入欄の金額を集計対象にする', () => {
    const context = loadMonthlySheetRepository([
        ['07/04', '売却', 'HARDOFF', 4300],
        ['', '', '', ''],
        ['', '', '', ''],
        ['予定', '給料', '給料', 365219],
        ...Array.from({ length: 10 }, () => ['', '', '', ''])
    ]);
    const entries = context.getMonthlyIncomeEntries(2026, 6);

    const actual = entries.map(({ date, dateLabel, name, amount }) => ({
            month: date.getMonth() + 1,
            day: date.getDate(),
            dateLabel,
            name,
            amount
        }));

    assert.equal(
        JSON.stringify(actual),
        JSON.stringify([
            { month: 7, day: 4, dateLabel: '', name: 'HARDOFF', amount: 4300 },
            { month: 7, day: 1, dateLabel: '予定', name: '給料', amount: 365219 }
        ])
    );
});

test('週内の売却収入は実日付の売却だけを取得する', () => {
    const context = loadMonthlySheetRepository([
        ['07/04', '売却', 'HARDOFF', 4300],
        ['07/04', '給料', '給料', 250000],
        ['予定', '売却', 'フリマ出品予定', 12000],
        ['', '売却', '日付なし', 3000],
        ['07/11', '売却', '翌週の売却', 5000],
        ...Array.from({ length: 9 }, () => ['', '', '', ''])
    ]);

    const entries = context.getSaleIncomeEntriesForDates([
        new Date(2026, 6, 3),
        new Date(2026, 6, 4),
        new Date(2026, 6, 5)
    ]);

    assert.equal(
        JSON.stringify(entries.map(({ category, name, amount }) => ({ category, name, amount }))),
        JSON.stringify([{ category: '売却', name: 'HARDOFF', amount: 4300 }])
    );
});
