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

test('GPT-5.6 Lunaへのリクエストではtemperatureを送らない', () => {
    let requestPayload;
    const client = loadOpenAiClient({
        UrlFetchApp: {
            fetch: (_url, options) => {
                requestPayload = JSON.parse(options.payload);
                return {
                    getResponseCode: () => 200,
                    getContentText: () => JSON.stringify({ output_text: '分析結果' })
                };
            }
        },
        Logger: { log: () => {} }
    });

    const result = client.generateOpenAiText('test-key', 'テスト', { temperature: 0.4, maxOutputTokens: 100 });

    assert.equal(result.text, '分析結果');
    assert.equal(requestPayload.model, 'gpt-5.6-luna');
    assert.equal(requestPayload.temperature, undefined);
    assert.equal(requestPayload.max_output_tokens, 100);
});

test('GPT-5.6以外の指定モデルではtemperatureを維持する', () => {
    let requestPayload;
    const client = loadOpenAiClient({
        PropertiesService: {
            getScriptProperties: () => ({
                getProperty: () => 'custom-model'
            })
        },
        UrlFetchApp: {
            fetch: (_url, options) => {
                requestPayload = JSON.parse(options.payload);
                return {
                    getResponseCode: () => 200,
                    getContentText: () => JSON.stringify({ output_text: '分析結果' })
                };
            }
        },
        Logger: { log: () => {} }
    });

    client.generateOpenAiText('test-key', 'テスト', { temperature: 0.4 });

    assert.equal(requestPayload.temperature, 0.4);
});
