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

test('ふるさと納税だけならAIとキャッシュ保存なしで予算対象外にし、実支出には残す', () => {
    const context = loadSpecialExpenseAi({
        calculateTotalAmount: (entries) => entries.reduce((sum, entry) => sum + entry.amount, 0),
        generatePreferredAiText: () => assert.fail('ふるさと納税はAI判定しない'),
        PropertiesService: {
            getScriptProperties: () => ({
                getProperty: () => null,
                setProperty: () => assert.fail('ふるさと納税は判定キャッシュを保存しない')
            })
        }
    });
    const donation = makeEntry(10, 'ふるさと納税', '北海道○○市', 50000);
    const food = makeEntry(11, 'スーパー・食品など', 'スーパー', 3000);
    const entries = [donation, food];
    const review = context.reviewSpecialExpensesWithAI(entries);

    assert.equal(review.hasCandidates, true);
    assert.equal(review.aiResult, null);
    assert.equal(review.approvedEntries[0].entry, donation);
    assert.equal(review.rejectedEntries.length, 0);
    assert.equal(context.calculateTotalAmount(entries), 53000);
    assert.equal(context.calculateTotalAmount(context.getBudgetTargetEntries(entries, review)), 3000);
    assert.equal(context.calculateApprovedSpecialExpenseTotal(review), 50000);
    assert.match(context.buildSpecialExpenseReviewSection(review, '今週'), /予算対象外: 北海道○○市 50000円（判定理由: ふるさと納税のため）/);
});

test('ふるさと納税と特別費が混在してもAIには特別費だけを渡し、対象期間の除外額を合算する', () => {
    const prompts = [];
    const context = loadSpecialExpenseAi({
        calculateTotalAmount: (entries) => entries.reduce((sum, entry) => sum + entry.amount, 0),
        buildSpecialExpenseReviewPrompt: (entries) => {
            prompts.push(entries.map((entry) => entry.name));
            return 'prompt';
        },
        generatePreferredAiText: () => ({ text: '0|approved|旅行費\n1|rejected|通常支出' })
    });
    const donation = makeEntry(10, 'ふるさと納税', 'ふるさと納税', 50000);
    const travel = makeEntry(11, '特別費', '旅行宿泊費', 20000);
    const gadget = makeEntry(12, '特別費', 'イヤホン', 10000);
    const nextDonation = makeEntry(17, 'ふるさと納税', '翌週の寄附', 30000);
    const review = context.reviewSpecialExpensesWithAI([donation, travel, gadget, nextDonation]);

    assert.equal(JSON.stringify(prompts), JSON.stringify([['旅行宿泊費', 'イヤホン']]));
    assert.equal(review.approvedEntries.length, 3);
    assert.equal(review.rejectedEntries[0].entry, gadget);
    assert.equal(context.calculateApprovedSpecialExpenseTotal(review), 100000);
    assert.equal(context.calculateApprovedSpecialExpenseTotalForEntries([donation, travel, gadget], review), 70000);
    assert.equal(context.calculateTotalAmount(context.getBudgetTargetEntries([donation, travel, gadget], review)), 10000);
});

test('履歴のふるさと納税は未キャッシュでも古い却下キャッシュがあっても除外する', () => {
    const context = loadSpecialExpenseAi({
        PropertiesService: {
            getScriptProperties: () => ({
                getProperty: () => JSON.stringify({
                    '2026-2-10|ふるさと納税|寄附|50000': { approved: false, reason: '古い判定' }
                })
            })
        }
    });
    const donation = makeEntry(10, 'ふるさと納税', '寄附', 50000);
    const uncachedDonation = makeEntry(11, ' ふるさと納税 ', '寄附', 30000);
    const special = makeEntry(12, '特別費', '未判定', 20000);
    const food = makeEntry(13, 'スーパー・食品など', 'スーパー', 3000);
    const entries = context.getBudgetTargetEntriesWithCachedSpecialExpenseDecisions([donation, uncachedDonation, special, food]);

    assert.equal(JSON.stringify(entries.map((entry) => entry.name)), JSON.stringify(['未判定', 'スーパー']));
    const review = context.reviewSpecialExpensesWithAI([donation, uncachedDonation]);
    assert.equal(review.approvedEntries.length, 2);
    assert.equal(review.aiResult, null);
});

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

test('保存済みの承認済み特別費だけを履歴比較の予算対象外にする', () => {
    const context = loadSpecialExpenseAi({
        PropertiesService: {
            getScriptProperties: () => ({
                getProperty: () => JSON.stringify({
                    '2026-2-10|特別費|福岡旅行の宿泊費|60000': { approved: true, reason: '旅行費' },
                    '2026-2-11|特別費|イヤホン|20000': { approved: false, reason: '通常支出' }
                })
            })
        }
    });
    const travel = makeEntry(10, '特別費', '福岡旅行の宿泊費', 60000);
    const gadget = makeEntry(11, '特別費', 'イヤホン', 20000);

    const entries = context.getBudgetTargetEntriesWithCachedSpecialExpenseDecisions([travel, gadget]);

    assert.equal(JSON.stringify(entries.map((entry) => entry.name)), JSON.stringify(['イヤホン']));
});

test('特別費AI判定は不完全な応答を予算対象に残す', () => {
    const context = loadSpecialExpenseAi({});
    const travel = makeEntry(10, '特別費', '福岡旅行の宿泊費', 60000);
    const unknown = makeEntry(11, '特別費', 'Amazon', 20000);

    const decisions = context.parseSpecialExpenseReviewResponse('0|approved|旅行費', [travel, unknown]);

    assert.equal(decisions[0].approved, true);
    assert.equal(decisions[1].approved, false);
    assert.equal(decisions[1].reason, 'AIの判定・理由を取得できなかったため');
});

test('特別費AI判定はOpenAIの不完全な応答を未判定の行だけリトライする', () => {
    const prompts = [];
    const context = loadSpecialExpenseAi({
        buildSpecialExpenseReviewPrompt: (entries) => {
            prompts.push(entries.map((entry) => entry.name));
            return 'prompt';
        },
        generatePreferredAiText: (_prompt, _config, options) => {
            if (options.logContext === 'special_expense_review_openai_attempt_1') {
                return { text: '0|approved|旅行の宿泊費' };
            }
            return { text: '0|approved|旅行の航空券' };
        },
        Logger: { log: () => {} }
    });
    const hotel = makeEntry(10, '特別費', '福岡旅行ホテル代', 34900);
    const flight = makeEntry(11, '特別費', '旅行帰りチケット：福岡→羽田', 24870);

    const review = context.reviewSpecialExpensesWithAI([hotel, flight]);

    assert.equal(JSON.stringify(prompts), JSON.stringify([['福岡旅行ホテル代', '旅行帰りチケット：福岡→羽田'], ['旅行帰りチケット：福岡→羽田']]));
    assert.equal(review.approvedEntries.length, 2);
    assert.equal(review.rejectedEntries.length, 0);
});

test('特別費AI判定はOpenAIの再試行後にGeminiへフォールバックする', () => {
    const optionsHistory = [];
    const context = loadSpecialExpenseAi({
        buildSpecialExpenseReviewPrompt: () => 'prompt',
        generatePreferredAiText: (_prompt, _config, options) => {
            optionsHistory.push(options);
            if (options.skipOpenAi) {
                return { text: '0|approved|旅行費', provider: 'gemini', model: 'gemini', usedFallback: true, fallbackReason: options.fallbackReason };
            }
            return { text: null, provider: null, model: null, usedFallback: false };
        },
        Logger: { log: () => {} }
    });
    const hotel = makeEntry(10, '特別費', '福岡旅行ホテル代', 34900);

    const review = context.reviewSpecialExpensesWithAI([hotel]);

    assert.equal(optionsHistory.length, 3);
    assert.equal(optionsHistory[0].skipGemini, true);
    assert.equal(optionsHistory[1].skipGemini, true);
    assert.equal(optionsHistory[2].skipOpenAi, true);
    assert.equal(optionsHistory[2].fallbackReason, 'openai_response_incomplete');
    assert.equal(review.approvedEntries.length, 1);
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

test('期間内の承認済み特別費だけを合計できる', () => {
    const context = loadSpecialExpenseAi({
        calculateTotalAmount: (entries) => entries.reduce((sum, entry) => sum + entry.amount, 0)
    });
    const travel = makeEntry(10, '特別費', '福岡旅行の宿泊費', 50000);
    const otherWeekTravel = makeEntry(17, '特別費', '次週の旅行費', 30000);
    const review = {
        approvedEntries: [
            { entry: travel, approved: true, reason: '旅行費' },
            { entry: otherWeekTravel, approved: true, reason: '旅行費' }
        ],
        rejectedEntries: [],
        hasCandidates: true
    };

    assert.equal(
        context.calculateApprovedSpecialExpenseTotalForEntries([travel], review),
        50000
    );
});

test('理由の欠落・空白・理由なし・20文字超過は承認も却下も確定しない', () => {
    const context = loadSpecialExpenseAi({});
    const entry = makeEntry(10, '特別費', '機器の買い替え', 10549);
    for (const status of ['approved', 'rejected']) {
        for (const reason of ['', '   ', '理由なし', 'あ'.repeat(21)]) {
            const decision = context.parseSpecialExpenseReviewResponse(`0|${status}|${reason}`, [entry])[0];
            assert.equal(decision.hasAiDecision, false);
            assert.equal(decision.approved, false);
        }
    }
    assert.equal(context.parseSpecialExpenseReviewResponse(`0|approved|${'あ'.repeat(20)}`, [entry])[0].hasAiDecision, true);
});

test('理由が空の承認は再試行し、取得した短い理由を保存してメールに表示する', () => {
    let calls = 0;
    let saved;
    const entry = makeEntry(10, '特別費', 'JCOM解約に伴い返却する機器の代わりを購入', 10549);
    const context = loadSpecialExpenseAi({
        buildSpecialExpenseReviewPrompt: () => 'prompt',
        generatePreferredAiText: () => ({ text: ++calls === 1 ? '0|approved|' : '0|approved|返却に伴う必須機器の購入' }),
        Logger: { log: () => {} },
        PropertiesService: { getScriptProperties: () => ({
            getProperty: () => null,
            setProperty: (_key, value) => { saved = JSON.parse(value); }
        }) }
    });
    const review = context.reviewSpecialExpensesWithAI([entry]);
    assert.equal(calls, 2);
    assert.equal(review.approvedEntries.length, 1);
    assert.equal(saved[context.getSpecialExpenseEntryKey(entry)].reason, '返却に伴う必須機器の購入');
    assert.match(context.buildSpecialExpenseReviewSection(review, '今週'), /（判定理由: 返却に伴う必須機器の購入）/);
});

test('理由なしのキャッシュを再判定し、全試行で理由が得られなければ予算に残して古い承認を削除する', () => {
    for (const reason of ['', '理由なし']) {
        let calls = 0;
        const entry = makeEntry(10, '特別費', '機器の買い替え', 10549);
        const key = '2026-2-10|特別費|機器の買い替え|10549';
        let cache = { [key]: { approved: true, reason } };
        const context = loadSpecialExpenseAi({
            calculateTotalAmount: entries => entries.reduce((sum, item) => sum + item.amount, 0),
            buildSpecialExpenseReviewPrompt: () => 'prompt',
            generatePreferredAiText: () => { calls += 1; return { text: '0|approved|' }; },
            Logger: { log: () => {} },
            PropertiesService: { getScriptProperties: () => ({
                getProperty: () => JSON.stringify(cache),
                setProperty: (_key, value) => { cache = JSON.parse(value); }
            }) }
        });
        assert.equal(context.getBudgetTargetEntriesWithCachedSpecialExpenseDecisions([entry]).length, 1);
        const review = context.reviewSpecialExpensesWithAI([entry]);
        assert.equal(calls, 3);
        assert.equal(review.approvedEntries.length, 0);
        assert.equal(context.calculateTotalAmount(context.getBudgetTargetEntries([entry], review)), 10549);
        assert.equal(cache[key], undefined);
        assert.match(context.buildSpecialExpenseReviewSection(review, '今週'), /予算対象に含める: .*判定理由: AIの判定・理由を取得できなかったため/);
    }
});

test('保存済みの理由なしは短い理由を取得して更新し、次回から再利用する', () => {
    let calls = 0;
    const entry = makeEntry(10, '特別費', '機器の買い替え', 10549);
    const key = '2026-2-10|特別費|機器の買い替え|10549';
    let cache = { [key]: { approved: true, reason: '理由なし' } };
    const context = loadSpecialExpenseAi({
        buildSpecialExpenseReviewPrompt: () => 'prompt',
        generatePreferredAiText: () => { calls += 1; return { text: '0|approved|故障による必須機器の交換' }; },
        PropertiesService: { getScriptProperties: () => ({
            getProperty: () => JSON.stringify(cache),
            setProperty: (_key, value) => { cache = JSON.parse(value); }
        }) }
    });
    context.reviewSpecialExpensesWithAI([entry]);
    assert.equal(cache[key].reason, '故障による必須機器の交換');
    context.reviewSpecialExpensesWithAI([entry]);
    assert.equal(calls, 1);
    assert.equal(context.getBudgetTargetEntriesWithCachedSpecialExpenseDecisions([entry]).length, 0);
});
