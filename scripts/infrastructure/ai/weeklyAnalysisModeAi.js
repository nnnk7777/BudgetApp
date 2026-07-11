function detectWeeklyAnalysisModeWithAI(title, description) {
    var result;
    var normalized;
    var prompt;

    prompt = buildWeeklyAnalysisModePrompt(title, description);

    result = generatePreferredAiText(prompt, {
        temperature: 0,
        maxOutputTokens: 10
    }, {
        logContext: "weekly_analysis_mode_detection"
    });
    if (!result.text) {
        return null;
    }

    normalized = String(result.text).trim().toUpperCase();
    Logger.log("分析モードAI判定: title=" + title + " provider=" + result.provider + " response=" + normalized);
    return normalized.indexOf("YES") === 0 ? WEEKLY_ANALYSIS_MODE_FRUGAL : null;
}
