const assert = require('node:assert/strict');
const fs = require('node:fs');
const test = require('node:test');
const vm = require('node:vm');

test('GASの手動実行から当日支出一覧を取得してログ出力できる', () => {
    let callCount = 0;
    const logs = [];
    const context = {
        Logger: {
            log: (message) => logs.push(message)
        },
        listTodayExpenses: () => {
            callCount++;
            return '{"ok":true,"items":[{"category":"交通費","name":"電車","amount":639}]}';
        }
    };
    vm.runInNewContext(
        fs.readFileSync('scripts/entrypoints/0_manualEntryPoints.js', 'utf8'),
        context
    );

    const result = context.listTodayExpensesManual();

    assert.equal(result, '{"ok":true,"items":[{"category":"交通費","name":"電車","amount":639}]}');
    assert.equal(callCount, 1);
    assert.deepEqual(logs, [result]);
});
