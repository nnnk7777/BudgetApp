const assert = require('node:assert/strict');
const fs = require('node:fs');
const test = require('node:test');
const vm = require('node:vm');

function loadApiCommon(overrides) {
    const context = {
        Error,
        JSON,
        Object,
        ...overrides
    };
    vm.runInNewContext(
        fs.readFileSync('scripts/entrypoints/apiCommon.js', 'utf8'),
        context
    );
    return context;
}

test('list_today_expenses actionは当日支出一覧の取得処理だけを呼び出す', () => {
    let callCount = 0;
    const context = loadApiCommon({
        listTodayExpenses: () => {
            callCount++;
            return '{"ok":true,"items":[]}';
        }
    });

    const result = context.dispatchApiAction({ action: 'list_today_expenses' });

    assert.equal(result, '{"ok":true,"items":[]}');
    assert.equal(callCount, 1);
});
