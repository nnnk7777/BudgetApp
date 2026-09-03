const assert = require('node:assert/strict');
const fs = require('node:fs');
const test = require('node:test');
const vm = require('node:vm');

test('GASの手動実行から当日支出一覧を取得できる', () => {
    let callCount = 0;
    const context = {
        listTodayExpenses: () => {
            callCount++;
            return '{"ok":true,"items":[]}';
        }
    };
    vm.runInNewContext(
        fs.readFileSync('scripts/entrypoints/0_manualEntryPoints.js', 'utf8'),
        context
    );

    const result = context.listTodayExpensesManual();

    assert.equal(result, '{"ok":true,"items":[]}');
    assert.equal(callCount, 1);
});
