const assert = require('node:assert/strict');
const fs = require('node:fs');
const test = require('node:test');
const vm = require('node:vm');

const income = { title: '機材の売却', amount: 12000, category: '売却' };
const saleDate = new Date(2026, 11, 1);

function incomeApi({ full = false, hash = 'test-hash', lockAvailable = true } = {}) {
    const rows = Array.from({ length: 14 }, () => full ? ['予定', '', '', ''] : ['', '', '', '']);
    const formulas = Array.from({ length: 14 }, () => ['', '', '', '']);
    const events = [];
    const listCriteria = 'VALUE_IN_LIST';
    const context = vm.createContext({
        PropertiesService: { getScriptProperties: () => ({ getProperty: () => hash }) },
        ContentService: {
            MimeType: { JSON: 'json' },
            createTextOutput: (body) => ({ body, setMimeType: () => {} })
        },
        Logger: { log: () => {} },
        getScriptRuntimeContext: () => ({ currentDate: new Date('2026-11-30T15:00:00Z') }),
        Utilities: {
            formatDate(date, timeZone, format) {
                assert.equal(format, 'MM/dd');
                return new Intl.DateTimeFormat('en-US', {
                    timeZone, month: '2-digit', day: '2-digit'
                }).format(date);
            }
        },
        LockService: {
            getScriptLock: () => ({
                tryLock: () => { events.push('lock'); return lockAvailable; },
                releaseLock: () => events.push('release')
            })
        },
        SpreadsheetApp: {
            DataValidationCriteria: { VALUE_IN_LIST: listCriteria },
            flush: () => events.push('flush'),
            getActiveSpreadsheet: () => ({
                getSheetByName: () => ({
                    getRange(startRow, startColumn, rowCount, columnCount) {
                        assert.deepEqual([startRow, startColumn, rowCount, columnCount], [20, 51, 14, 4]);
                        return {
                            getValues: () => rows,
                            getFormulas: () => formulas,
                            offset(index, column, height, width) {
                                assert.deepEqual([column, height, width], [0, 1, 4]);
                                return {
                                    getCell(row, col) {
                                        assert.deepEqual([row, col], [1, 2]);
                                        return {
                                            getDataValidation: () => ({
                                                getCriteriaType: () => listCriteria,
                                                getCriteriaValues: () => [['売却', '給料'], true]
                                            })
                                        };
                                    },
                                    setValues([values]) {
                                        events.push('write');
                                        assert.ok(values.slice(0, 3).every(value => value.startsWith("'")));
                                        rows[index] = Array.from(values, value =>
                                            typeof value === 'string' ? value.slice(1) : value
                                        );
                                    }
                                };
                            }
                        };
                    }
                })
            })
        },
        addExpenseRecord: (title, amount, category) => ({ title, amount, category })
    });

    for (const file of [
        'scripts/entrypoints/apiCommon.js',
        'scripts/entrypoints/handleApi.js',
        'scripts/application/addIncomeRecord.js',
        'scripts/infrastructure/gas/expenseSheetRepository.js',
        'scripts/infrastructure/gas/monthlySheetRepository.js',
        'scripts/utils/summaryDateUtils.js'
    ]) {
        vm.runInContext(fs.readFileSync(file, 'utf8'), context, { filename: file });
    }

    return {
        rows,
        formulas,
        events,
        context,
        post(data) {
            return JSON.parse(context.doPost({
                postData: { contents: JSON.stringify({ hash: 'test-hash', action: 'add_income', ...data }) }
            }).body);
        }
    };
}

test('日本時間の月境界で収入を登録し、既存の月次・売却収入集計で各行を読み取れる', () => {
    const { post, rows, formulas, events, context } = incomeApi();
    rows[0][0] = '予定';
    formulas[1][2] = '=IF(FALSE, 1, "")';

    assert.deepEqual(post({ item: income }), { ok: true });
    assert.deepEqual(post({ item: { ...income, title: '=ABC', amount: 3000 } }), { ok: true });
    assert.deepEqual(rows.slice(0, 4), [
        ['予定', '', '', ''],
        ['', '', '', ''],
        ['12/01', '売却', '機材の売却', 12000],
        ['12/01', '売却', '=ABC', 3000]
    ]);
    assert.deepEqual(events, ['lock', 'write', 'flush', 'release', 'lock', 'write', 'flush', 'release']);
    const monthlyEntries = context.getMonthlyIncomeEntries(2026, 11);
    assert.equal(monthlyEntries.reduce((sum, entry) => sum + entry.amount, 0), 15000);
    assert.equal(context.getSaleIncomeEntriesForDates([saleDate]).length, 2);
});

test('文字列で渡された円単位の金額を数値として保存する', () => {
    const { post, rows } = incomeApi();
    for (const [index, amount] of ['12000', ' １２，０００ ', '¥12,000', '12,000円'].entries()) {
        assert.deepEqual(post({ item: { ...income, amount } }), { ok: true });
        assert.equal(rows[index][3], 12000);
    }
});

test('認証エラーと入力不備はシートへの書き込み前にエラーを返す', () => {
    const { post, events } = incomeApi();
    assert.match(post({ hash: 'wrong', item: income }).error, /hash/);
    assert.match(incomeApi({ hash: null }).post({ hash: null, item: income }).error, /hash/);
    assert.match(post({}).error, /item/);
    assert.match(post({ item: { ...income, title: ' ' } }).error, /title/);
    assert.match(post({ item: { ...income, category: '' } }).error, /カテゴリ/);
    for (const amount of [
        0, -1, 1.5, '0', '1.5', '12,00', '1 200', '1200abc', '1e3',
        null, true, [1200], Number.MAX_SAFE_INTEGER + 1, '9007199254740992'
    ]) {
        assert.match(post({ item: { ...income, amount } }).error, /amount/);
    }
    assert.deepEqual(events, []);
    assert.deepEqual(post({ item: { ...income, category: '交通費' } }), {
        ok: false,
        error: '収入カテゴリが不正です'
    });
    assert.deepEqual(events, ['lock', 'release']);
});

test('満欄やロックを取得できない場合は収入を上書きしない', () => {
    const full = incomeApi({ full: true });
    assert.match(full.post({ item: income }).error, /空きがありません/);
    assert.deepEqual(full.events, ['lock', 'release']);
    const busy = incomeApi({ lockAvailable: false });
    assert.match(busy.post({ item: income }).error, /使用中/);
    assert.deepEqual(busy.events, ['lock']);
});

test('既存の add は引き続き支出登録に振り分けられる', () => {
    const { post, events } = incomeApi();
    assert.deepEqual(post({ action: 'add', item: { title: '電車', price: 300, category: '交通費' } }), {
        title: '電車', amount: 300, category: '交通費'
    });
    assert.deepEqual(events, []);
});
