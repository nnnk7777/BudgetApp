function getOpenAiApiKey(silent) {
    if (typeof PropertiesService === 'undefined') {
        return null;
    }

    var apiKey = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
    if (!apiKey && !silent) {
        Logger.log("OpenAI API key is not set in script properties.");
    }
    return apiKey;
}

function getOpenAiModel() {
    if (typeof PropertiesService === 'undefined') {
        return "gpt-5.4-mini";
    }

    return PropertiesService.getScriptProperties().getProperty("OPENAI_MODEL") || "gpt-5.4-mini";
}

function generateOpenAiText(apiKey, prompt, generationConfig) {
    var model = getOpenAiModel();
    var payload = {
        model: model,
        input: prompt
    };
    var response;
    var parsed;
    var text;

    generationConfig = generationConfig || {};
    if (generationConfig.temperature !== undefined) {
        payload.temperature = generationConfig.temperature;
    }
    if (generationConfig.maxOutputTokens !== undefined) {
        payload.max_output_tokens = generationConfig.maxOutputTokens;
    }

    try {
        response = UrlFetchApp.fetch("https://api.openai.com/v1/responses", {
            method: "post",
            contentType: "application/json",
            headers: {
                Authorization: "Bearer " + apiKey
            },
            payload: JSON.stringify(payload),
            muteHttpExceptions: true
        });

        if (response.getResponseCode() !== 200) {
            Logger.log("OpenAI request failed (model=" + model + "): " + response.getResponseCode() + " " + response.getContentText());
            return null;
        }

        parsed = JSON.parse(response.getContentText());
        text = extractOpenAiOutputText(parsed);
        if (!text) {
            Logger.log("OpenAI response did not include output text: " + response.getContentText());
            return null;
        }

        return {
            text: text,
            provider: "openai",
            model: model
        };
    } catch (error) {
        Logger.log("OpenAI request error (model=" + model + "): " + error);
        return null;
    }
}

function extractOpenAiOutputText(responseJson) {
    if (!responseJson) {
        return "";
    }

    if (typeof responseJson.output_text === "string" && responseJson.output_text.trim() !== "") {
        return responseJson.output_text.trim();
    }

    if (!responseJson.output || !responseJson.output.length) {
        return "";
    }

    return responseJson.output.map(function (item) {
        if (!item || !item.content || !item.content.length) {
            return "";
        }

        return item.content.map(function (contentPart) {
            if (!contentPart) {
                return "";
            }

            if (typeof contentPart.text === "string") {
                return contentPart.text;
            }

            if (contentPart.type === "output_text" && typeof contentPart.text === "string") {
                return contentPart.text;
            }

            return "";
        }).join("");
    }).join("").trim();
}

function generatePreferredAiText(prompt, generationConfig, options) {
    var openAiApiKey = getOpenAiApiKey(true);
    var geminiApiKey = getGeminiApiKey(true);
    var openAiResult;
    var geminiText;
    var fallbackReason;

    options = options || {};
    if (openAiApiKey) {
        openAiResult = generateOpenAiText(openAiApiKey, prompt, generationConfig);
        if (openAiResult && openAiResult.text) {
            return {
                text: openAiResult.text,
                provider: "openai",
                model: openAiResult.model,
                usedFallback: false
            };
        }

        fallbackReason = "openai_unavailable";
    } else {
        fallbackReason = "openai_api_key_missing";
    }

    if (geminiApiKey) {
        geminiText = generateGeminiText(geminiApiKey, prompt, generationConfig);
        if (geminiText) {
            return {
                text: geminiText,
                provider: "gemini",
                model: getGeminiModelLabel(),
                usedFallback: true,
                fallbackReason: fallbackReason
            };
        }
    }

    return {
        text: null,
        provider: null,
        model: null,
        usedFallback: false,
        fallbackReason: fallbackReason || "no_provider_available"
    };
}

function buildAiSummarySection(title, analysisResult) {
    var body = title + "\n";

    if (analysisResult && analysisResult.text) {
        if (analysisResult.provider === "gemini") {
            body += "（OpenAIを利用できなかったため、Geminiの結果を使用しています）\n";
        }
        body += analysisResult.text + "\n";
        return body;
    }

    return body + "(AIからの回答を取得できませんでした。ログを確認してください)\n";
}
