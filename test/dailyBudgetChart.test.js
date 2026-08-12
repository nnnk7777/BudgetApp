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

    assert.equal(chart.getDailyBudgetChartTheme(15600, 40000).color, '#188038');
    assert.equal(chart.getDailyBudgetChartTheme(15600, 40000).label, '予算内');
    assert.equal(chart.getDailyBudgetChartTheme(16000, 40000).color, '#f9ab00');
    assert.equal(chart.getDailyBudgetChartTheme(24000, 40000).color, '#f4511e');
    assert.equal(chart.getDailyBudgetChartTheme(32000, 40000).color, '#e53935');
    assert.equal(chart.getDailyBudgetChartTheme(36000, 40000).color, '#c62828');
    assert.equal(chart.getDailyBudgetChartTheme(40001, 40000).color, '#b42318');
    assert.equal(chart.getDailyBudgetChartTheme(40001, 40000).overBudgetColor, '#8c1d18');
    assert.equal(chart.getDailyBudgetChartTheme(40001, 40000).label, '超過');
});

test('日次予算グラフのタイトルは超過額を明示する', () => {
    const chart = loadChartFunctions();

    assert.equal(
        chart.buildDailyBudgetChartTitle(43400, 40000),
        '【緊急】週予算の1.1倍（3400円超過）'
    );
});

test('週次サマリ用の月次・週次グラフはそれぞれの予算を基準にする', () => {
    const chart = loadChartFunctions();

    assert.equal(
        chart.buildWeeklySummaryMonthlyChartTitle(170000, 160000, null),
        '【緊急】月次予算の1.1倍（10000円超過）'
    );
    assert.equal(
        chart.buildWeeklySummaryWeeklyChartTitle(43400, 40000),
        '【緊急】週予算の1.1倍（3400円超過）'
    );
});

test('週次サマリ用の月次グラフは分析時点の月内ペースを表示する', () => {
    const chart = loadChartFunctions();
    const currentDate = new Date(2026, 1, 15);

    assert.equal(chart.getMonthlyBudgetPacePercentage(currentDate), 15 / 28 * 100);
    assert.equal(
        chart.buildWeeklySummaryMonthlyChartTitle(80000, 160000, currentDate),
        '今月の実支出 80000円 / 160000円（50.0%）　｜　2/15時点の目安 53.6%'
    );
});

test('日次予算グラフは予算内・超過・残りを分ける', () => {
    const chart = loadChartFunctions();

    assert.equal(chart.getDailyBudgetChartAmounts(9400, 40000).withinBudget, 9400);
    assert.equal(chart.getDailyBudgetChartAmounts(9400, 40000).overBudget, 0);
    assert.equal(chart.getDailyBudgetChartAmounts(9400, 40000).criticalOverBudget, 0);
    assert.equal(chart.getDailyBudgetChartAmounts(9400, 40000).remaining, 30600);
    assert.equal(chart.getDailyBudgetChartAmounts(43400, 40000).withinBudget, 40000);
    assert.equal(chart.getDailyBudgetChartAmounts(43400, 40000).overBudget, 3400);
    assert.equal(chart.getDailyBudgetChartAmounts(43400, 40000).criticalOverBudget, 0);
    assert.equal(chart.getDailyBudgetChartAmounts(43400, 40000).remaining, 0);
    assert.equal(chart.getDailyBudgetChartAmounts(159400, 40000).overBudget, 20000);
    assert.equal(chart.getDailyBudgetChartAmounts(159400, 40000).criticalOverBudget, 99400);
});

test('日次予算グラフは超過時に25%刻みで表示範囲を広げる', () => {
    const chart = loadChartFunctions();
    const ticks = chart.getDailyBudgetChartTicks(40000, chart.getDailyBudgetChartScalePercentage(64400, 40000));

    assert.equal(chart.getDailyBudgetChartScalePercentage(40000, 40000), 100);
    assert.equal(chart.getDailyBudgetChartScalePercentage(40100, 40000), 125);
    assert.equal(chart.getDailyBudgetChartScalePercentage(64400, 40000), 175);
    assert.equal(JSON.stringify(ticks), JSON.stringify([
        { v: 0, f: '0%' },
        { v: 10000, f: '25%' },
        { v: 20000, f: '50%' },
        { v: 30000, f: '75%' },
        { v: 40000, f: '100%' },
        { v: 50000, f: '125%' },
        { v: 60000, f: '150%' },
        { v: 70000, f: '175%' }
    ]));
});

test('日次予算グラフは大幅超過時に目盛りの間隔を広げる', () => {
    const chart = loadChartFunctions();
    const ticks = chart.getDailyBudgetChartTicks(40000, 400);

    assert.equal(chart.getDailyBudgetChartTickStep(200), 25);
    assert.equal(chart.getDailyBudgetChartTickStep(300), 50);
    assert.equal(chart.getDailyBudgetChartTickStep(400), 100);
    assert.equal(JSON.stringify(ticks), JSON.stringify([
        { v: 0, f: '0%' },
        { v: 40000, f: '100% 予算上限' },
        { v: 80000, f: '200%' },
        { v: 120000, f: '300%' },
        { v: 160000, f: '400%' }
    ]));
});

test('週次サマリ用グラフは共通の割合を25%ごとに表示する', () => {
    const chart = loadChartFunctions();
    const ticks = chart.getMonthlyBudgetChartTicks(160000, 100);

    assert.equal(JSON.stringify(ticks), JSON.stringify([
        { v: 0, f: '0%' },
        { v: 40000, f: '25%' },
        { v: 80000, f: '50%' },
        { v: 120000, f: '75%' },
        { v: 160000, f: '100% 予算上限' }
    ]));
});

test('週次サマリ用グラフは月次・週次で100%の位置を共通にする', () => {
    const chart = loadChartFunctions();

    assert.equal(
        chart.getWeeklySummaryBudgetChartScalePercentage(80000, 160000, 159400, 40000),
        400
    );
});

test('HTMLメール本文はグラフを先頭に置き、テキスト本文をエスケープする', () => {
    const chart = loadChartFunctions();
    const html = chart.buildDailySummaryHtmlBody('支出: <1000>円 & 確認', 1000, 40000);

    assert.match(html, /^<img src="cid:dailyBudgetChart"/);
    assert.match(html, /支出: &lt;1000&gt;円 &amp; 確認/);
});

test('週次HTMLメール本文は月次・週次のグラフを順に表示する', () => {
    const chart = loadChartFunctions();
    const html = chart.buildWeeklySummaryHtmlBody('週次本文', 80000, 160000, new Date(2026, 1, 15), 40000, 40000);

    assert.match(html, /^<img src="cid:monthlyBudgetChart"/);
    assert.match(html, /<img src="cid:weeklyBudgetChart"/);
    assert.match(html, /monthlyBudgetChart"[^>]+margin:0 0 4px/);
    assert.match(html, /週次本文/);
});
