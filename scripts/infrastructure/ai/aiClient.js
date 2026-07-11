function generatePreferredAiText(prompt, generationConfig, options) {
    var openAiApiKey = getOpenAiApiKey(true);
    var geminiApiKey = getGeminiApiKey(true);
    var openAiResult;
    var geminiText;
    var fallbackReason;
    var logContext;

    options = options || {};
    logContext = options.logContext || "default";
    if (openAiApiKey) {
        openAiResult = generateOpenAiText(openAiApiKey, prompt, generationConfig);
        if (openAiResult && openAiResult.text) {
            Logger.log("AI利用: context=" + logContext + " provider=openai model=" + openAiResult.model);
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
            Logger.log("AI利用: context=" + logContext + " provider=gemini model=" + getGeminiModelLabel() + " fallbackReason=" + fallbackReason);
            return {
                text: geminiText,
                provider: "gemini",
                model: getGeminiModelLabel(),
                usedFallback: true,
                fallbackReason: fallbackReason
            };
        }
    }

    Logger.log("AI利用失敗: context=" + logContext + " fallbackReason=" + (fallbackReason || "no_provider_available"));
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
