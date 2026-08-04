const assert = require('node:assert/strict');
const fs = require('node:fs');
const test = require('node:test');
const vm = require('node:vm');

function loadMonthlyIncomeEntries(rows) {
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
    return context.getMonthlyIncomeEntries(2026, 6);
}

test('月次収入は予定・空行をまたいでも収入欄の金額を集計対象にする', () => {
    const entries = loadMonthlyIncomeEntries([
        ['07/04', '売却', 'HARDOFF', 4300],
        ['', '', '', ''],
        ['', '', '', ''],
        ['予定', '給料', '給料', 365219],
        ...Array.from({ length: 10 }, () => ['', '', '', ''])
    ]);

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
