const assert = require('node:assert/strict');
const fs = require('node:fs');
const test = require('node:test');
const vm = require('node:vm');

function loadAiClient(context) {
    vm.runInNewContext(
        fs.readFileSync('scripts/infrastructure/ai/aiClient.js', 'utf8'),
        context
    );
    return context;
}

test('AIクライアントは特別費判定の再試行中にGeminiを呼ばない', () => {
    let geminiCalled = false;
    const client = loadAiClient({
        getOpenAiApiKey: () => 'openai-key',
        getGeminiApiKey: () => 'gemini-key',
        generateOpenAiText: () => null,
        generateGeminiText: () => {
            geminiCalled = true;
            return 'gemini result';
        },
        Logger: { log: () => {} }
    });

    const result = client.generatePreferredAiText('prompt', {}, { skipGemini: true });

    assert.equal(result.text, null);
    assert.equal(geminiCalled, false);
});

test('AIクライアントはOpenAIを省略してGeminiへフォールバックできる', () => {
    let openAiCalled = false;
    const client = loadAiClient({
        getOpenAiApiKey: () => 'openai-key',
        getGeminiApiKey: () => 'gemini-key',
        generateOpenAiText: () => {
            openAiCalled = true;
            return { text: 'openai result', model: 'gpt-test' };
        },
        generateGeminiText: () => 'gemini result',
        getGeminiModelLabel: () => 'gemini-test',
        Logger: { log: () => {} }
    });

    const result = client.generatePreferredAiText('prompt', {}, {
        skipOpenAi: true,
        fallbackReason: 'openai_response_incomplete'
    });

    assert.equal(openAiCalled, false);
    assert.equal(result.text, 'gemini result');
    assert.equal(result.provider, 'gemini');
    assert.equal(result.fallbackReason, 'openai_response_incomplete');
});
