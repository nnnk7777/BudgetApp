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

test('GASの手動実行で設定した項目を共通の収入登録処理に渡す', () => {
    const item = { title: '機材の売却', amount: 12000, category: '売却' };
    const calls = [];
    const logs = [];
    let itemJson = null;
    const context = {
        PropertiesService: {
            getScriptProperties: () => ({
                getProperty(key) {
                    assert.equal(key, 'MANUAL_INCOME_ITEM');
                    return itemJson;
                }
            })
        },
        Logger: { log: (message) => logs.push(message) },
        addIncomeRecord: (input) => {
            calls.push(JSON.stringify(input));
            return { ok: true };
        }
    };
    vm.runInNewContext(
        fs.readFileSync('scripts/entrypoints/0_manualEntryPoints.js', 'utf8'),
        context
    );

    assert.throws(() => context.addIncomeRecordManual(), /MANUAL_INCOME_ITEM/);
    assert.deepEqual(calls, []);
    assert.deepEqual(logs, []);

    itemJson = JSON.stringify(item);
    assert.deepEqual(context.addIncomeRecordManual(), { ok: true });
    assert.deepEqual(calls, [itemJson]);
    assert.deepEqual(logs, ['{"ok":true}']);
});
