const assert = require('node:assert/strict');
const fs = require('node:fs');
const test = require('node:test');
const vm = require('node:vm');

function loadOpenAiClient(context) {
    vm.runInNewContext(
        fs.readFileSync('scripts/infrastructure/openai/openaiClient.js', 'utf8'),
        context
    );
    return context;
}

test('OpenAIの既定モデルはGPT-5.6 Lunaを使う', () => {
    const client = loadOpenAiClient({});

    assert.equal(client.getOpenAiModel(), 'gpt-5.6-luna');
});

test('OPENAI_MODELが設定されている場合は既定値より優先する', () => {
    const client = loadOpenAiClient({
        PropertiesService: {
            getScriptProperties: () => ({
                getProperty: () => 'custom-model'
            })
        }
    });

    assert.equal(client.getOpenAiModel(), 'custom-model');
});
