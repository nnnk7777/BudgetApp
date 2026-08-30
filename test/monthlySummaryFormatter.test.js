const assert = require('node:assert/strict');
const fs = require('node:fs');
const test = require('node:test');
const vm = require('node:vm');

function loadMonthlySummaryFormatter(context) {
    vm.runInNewContext(
        fs.readFileSync('scripts/formatting/monthlySummaryFormatter.js', 'utf8'),
        context
    );
    return context;
}

test('月次メールは予算グラフを先頭に埋め込んで送信する', () => {
    let sentMessage;
    const chartBlob = { name: 'monthly-summary-budget-chart.png' };
    const formatter = loadMonthlySummaryFormatter({
        getTargetEmailAddress: () => 'user@example.com',
        createMonthlySummaryBudgetChartBlob: (total, budget, special) => {
            assert.equal(total, 193000);
            assert.equal(budget, 193000);
            assert.equal(special, 100000);
            return chartBlob;
        },
        buildMonthlySummaryHtmlBody: (body, total, budget, special) => {
            assert.equal(body, '月次本文');
            assert.equal(total, 193000);
            assert.equal(budget, 193000);
            assert.equal(special, 100000);
            return '<img src="cid:monthlySummaryBudgetChart">月次本文';
        },
        MailApp: {
            sendEmail: (message) => {
                sentMessage = message;
            }
        }
    });

    const result = formatter.sendMonthlySummaryResult(
        'mail',
        new Date(2026, 8, 30),
        false,
        '月次本文',
        193000,
        193000,
        100000
    );

    assert.equal(result, 'Successfully sent monthly summary mail');
    assert.equal(sentMessage.to, 'user@example.com');
    assert.equal(sentMessage.subject, '家計簿月次レポート（9月）');
    assert.equal(sentMessage.body, '月次本文');
    assert.equal(sentMessage.htmlBody, '<img src="cid:monthlySummaryBudgetChart">月次本文');
    assert.equal(sentMessage.inlineImages.monthlySummaryBudgetChart, chartBlob);
});
