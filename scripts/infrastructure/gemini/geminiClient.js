function getGeminiApiKey(silent) {
    if (typeof PropertiesService === 'undefined') {
        return null;
    }

    var apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
    if (!apiKey && !silent) {
        Logger.log("Gemini API key is not set in script properties.");
    }
    return apiKey;
}

function getGeminiModelsToTry(apiKey) {
    var modelFromProp = PropertiesService.getScriptProperties().getProperty("GEMINI_MODEL");
    if (modelFromProp && modelFromProp.indexOf("models/") === 0) {
        modelFromProp = modelFromProp.replace(/^models\//, "");
    }

    var modelsToTry = [];
    if (modelFromProp) {
        modelsToTry.push(modelFromProp);
    } else {
        modelsToTry = modelsToTry.concat(fetchGenerativeModels(apiKey));
    }

    if (modelsToTry.length === 0) {
        modelsToTry = [
            "gemini-2.5-flash",
            "gemini-2.5-pro",
            "gemini-2.5-pro-preview-06-05",
            "gemini-2.0-flash",
            "gemini-1.5-flash"
        ];
    }

    return modelsToTry;
}

function getGeminiModelLabel() {
    var modelFromProp;

    if (typeof PropertiesService === 'undefined') {
        return "gemini";
    }

    modelFromProp = PropertiesService.getScriptProperties().getProperty("GEMINI_MODEL");
    if (modelFromProp && modelFromProp.indexOf("models/") === 0) {
        modelFromProp = modelFromProp.replace(/^models\//, "");
    }

    return modelFromProp || "gemini";
}

function generateGeminiText(apiKey, prompt, generationConfig) {
    var modelsToTry = getGeminiModelsToTry(apiKey);
    var apiVersions = ["v1beta", "v1"];
    var payload = {
        contents: [{
            parts: [{
                text: prompt
            }]
        }],
        generationConfig: generationConfig
    };

    for (var i = 0; i < apiVersions.length; i++) {
        var version = apiVersions[i];
        for (var j = 0; j < modelsToTry.length; j++) {
            var model = modelsToTry[j];
            var url = "https://generativelanguage.googleapis.com/" + version + "/models/" + model + ":generateContent?key=" + apiKey;

            try {
                var response = UrlFetchApp.fetch(url, {
                    method: "post",
                    contentType: "application/json",
                    payload: JSON.stringify(payload),
                    muteHttpExceptions: true
                });

                if (response.getResponseCode() !== 200) {
                    Logger.log("Gemini request failed (version=" + version + ", model=" + model + "): " + response.getResponseCode() + " " + response.getContentText());
                    continue;
                }

                var result = JSON.parse(response.getContentText());
                if (result && result.candidates && result.candidates.length > 0) {
                    var parts = result.candidates[0].content && result.candidates[0].content.parts;
                    if (parts && parts.length > 0) {
                        return parts.map(function (part) {
                            return part.text || "";
                        }).join("").trim();
                    }
                }
            } catch (error) {
                Logger.log("Gemini request error (version=" + version + ", model=" + model + "): " + error);
            }
        }
    }

    return null;
}

function fetchGenerativeModels(apiKey) {
    try {
        var listUrl = "https://generativelanguage.googleapis.com/v1beta/models?key=" + apiKey;
        var listResponse = UrlFetchApp.fetch(listUrl, { method: "get", muteHttpExceptions: true });
        if (listResponse.getResponseCode() !== 200) {
            Logger.log("Failed to list models: status=" + listResponse.getResponseCode() + " body=" + listResponse.getContentText());
            return [];
        }
        var parsed = JSON.parse(listResponse.getContentText());
        if (!parsed.models || !parsed.models.length) {
            Logger.log("No models found in listModels response.");
            return [];
        }
        var candidates = parsed.models.filter(function (model) {
            return model.supportedGenerationMethods && model.supportedGenerationMethods.indexOf("generateContent") !== -1;
        }).map(function (model) {
            return model.name.replace(/^models\//, "");
        });
        Logger.log("generateContent-capable models (auto-detected): " + candidates.join(", "));
        return candidates;
    } catch (error) {
        Logger.log("Error while listing models: " + error);
        return [];
    }
}
