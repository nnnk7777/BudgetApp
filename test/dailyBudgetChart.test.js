const assert = require('node:assert/strict');
const fs = require('node:fs');
const test = require('node:test');
const vm = require('node:vm');

function loadChartFunctions() {
    const context = {};
    vm.runInNewContext(
        fs.readFileSync('scripts/formatting/dailyBudgetChart.js', 'utf8'),
        context
    );
    return context;
}

test('日次予算グラフは支出割合に応じた色を返す', () => {
    const chart = loadChartFunctions();

    assert.equal(chart.getDailyBudgetChartTheme(27000, 40000).color, '#188038');
    assert.equal(chart.getDailyBudgetChartTheme(27000, 40000).label, '予算内');
    assert.equal(chart.getDailyBudgetChartTheme(28000, 40000).color, '#f9ab00');
    assert.equal(chart.getDailyBudgetChartTheme(28000, 40000).label, '注意');
    assert.equal(chart.getDailyBudgetChartTheme(40001, 40000).color, '#d93025');
    assert.equal(chart.getDailyBudgetChartTheme(40001, 40000).label, '超過');
});

test('日次予算グラフのタイトルは超過額を明示する', () => {
    const chart = loadChartFunctions();

    assert.equal(
        chart.buildDailyBudgetChartTitle(43400, 40000),
        '今週の実支出 43400円 / 40000円（108.5%）  3400円超過'
    );
});

test('日次予算グラフは予算を100%として支出と残りを分ける', () => {
    const chart = loadChartFunctions();

    assert.equal(chart.getDailyBudgetChartAmounts(9400, 40000).spent, 9400);
    assert.equal(chart.getDailyBudgetChartAmounts(9400, 40000).remaining, 30600);
    assert.equal(chart.getDailyBudgetChartAmounts(43400, 40000).spent, 40000);
    assert.equal(chart.getDailyBudgetChartAmounts(43400, 40000).remaining, 0);
});

test('日次予算グラフは0%から100%までを25%刻みで表示する', () => {
    const chart = loadChartFunctions();
    const ticks = chart.getDailyBudgetChartTicks(40000);

    assert.equal(JSON.stringify(ticks), JSON.stringify([
        { v: 0, f: '0%' },
        { v: 10000, f: '25%' },
        { v: 20000, f: '50%' },
        { v: 30000, f: '75%' },
        { v: 40000, f: '100%' }
    ]));
});

test('HTMLメール本文はグラフを先頭に置き、テキスト本文をエスケープする', () => {
    const chart = loadChartFunctions();
    const html = chart.buildDailySummaryHtmlBody('支出: <1000>円 & 確認', 1000, 40000);

    assert.match(html, /^<img src="cid:dailyBudgetChart"/);
    assert.match(html, /支出: &lt;1000&gt;円 &amp; 確認/);
});
