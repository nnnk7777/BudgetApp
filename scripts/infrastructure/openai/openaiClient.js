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
