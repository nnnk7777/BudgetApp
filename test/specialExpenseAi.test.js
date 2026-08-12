const assert = require('node:assert/strict');
const fs = require('node:fs');
const test = require('node:test');
const vm = require('node:vm');

function loadSpecialExpenseAi(context) {
    vm.runInNewContext(
        fs.readFileSync('scripts/infrastructure/ai/specialExpenseAi.js', 'utf8'),
        context
    );
    return context;
}

function makeEntry(day, category, name, amount) {
    return { date: new Date(2026, 1, day), category, name, amount };
}

test('特別費AI判定は承認した行だけを予算対象外にする', () => {
    const context = loadSpecialExpenseAi({
        calculateTotalAmount: (entries) => entries.reduce((sum, entry) => sum + entry.amount, 0)
    });
    const travel = makeEntry(10, '特別費', '福岡旅行の宿泊費', 60000);
    const gadget = makeEntry(11, '特別費', 'イヤホン', 20000);
    const food = makeEntry(12, 'スーパー・食品など', 'スーパー', 3000);
    const review = {
        approvedEntries: [{ entry: travel, approved: true, reason: '旅行費' }],
        rejectedEntries: [{ entry: gadget, approved: false, reason: '趣味の購入' }],
        hasCandidates: true
    };

    const budgetTargetEntries = context.getBudgetTargetEntries([travel, gadget, food], review);

    assert.equal(
        JSON.stringify(budgetTargetEntries.map((entry) => entry.name)),
        JSON.stringify(['イヤホン', 'スーパー'])
    );
    assert.equal(context.calculateApprovedSpecialExpenseTotal(review), 60000);
});

test('特別費AI判定は不完全な応答を予算対象に残す', () => {
    const context = loadSpecialExpenseAi({});
    const travel = makeEntry(10, '特別費', '福岡旅行の宿泊費', 60000);
    const unknown = makeEntry(11, '特別費', 'Amazon', 20000);

    const decisions = context.parseSpecialExpenseReviewResponse('0|approved|旅行費', [travel, unknown]);

    assert.equal(decisions[0].approved, true);
    assert.equal(decisions[1].approved, false);
    assert.equal(decisions[1].reason, 'AI判定を取得できなかったため');
});

test('特別費AI判定の文面は支出の必要性ではなく分類だけを審査する', () => {
    const context = loadSpecialExpenseAi({
        buildSpecialExpenseReviewPrompt: () => 'prompt',
        generatePreferredAiText: () => ({ text: '0|rejected|高額だけでは不十分' })
    });
    const gadget = makeEntry(11, '特別費', 'イヤホン', 20000);

    const review = context.reviewSpecialExpensesWithAI([gadget]);

    assert.equal(review.approvedEntries.length, 0);
    assert.equal(review.rejectedEntries.length, 1);
    assert.equal(review.rejectedEntries[0].reason, '高額だけでは不十分');
});

test('確定した特別費AI判定は保存し、同じ行を再判定しない', () => {
    let generated = 0;
    let savedCache;
    const properties = {
        getProperty: () => null,
        setProperty: (_key, value) => {
            savedCache = JSON.parse(value);
        }
    };
    const context = loadSpecialExpenseAi({
        PropertiesService: { getScriptProperties: () => properties },
        buildSpecialExpenseReviewPrompt: () => 'prompt',
        generatePreferredAiText: () => {
            generated += 1;
            return { text: '0|approved|旅行費' };
        }
    });
    const travel = makeEntry(10, '特別費', '福岡旅行の宿泊費', 60000);

    const firstReview = context.reviewSpecialExpensesWithAI([travel]);
    properties.getProperty = () => JSON.stringify(savedCache);
    const secondReview = context.reviewSpecialExpensesWithAI([travel]);

    assert.equal(firstReview.approvedEntries.length, 1);
    assert.equal(secondReview.approvedEntries.length, 1);
    assert.equal(generated, 1);
});
