// 月次サマリーを生成するトリガー用
function handleCalculateMonthlySummaryTrigger() {
    calculateMonthlySummary('mail');
}

// 月次サマリーのメイン処理
function calculateMonthlySummary(action) {
    var budgetPerWeek = 45000;
    var runtimeContext = getScriptRuntimeContext();
    var currentDate = runtimeContext.currentDate;
    var isStaging = runtimeContext.isStaging;
    var year = currentDate.getFullYear();
    var month = currentDate.getMonth(); // 0-based

    var startOfMonth = new Date(year, month, 1);
    var endOfMonth = new Date(year, month + 1, 0);
    var dateRangeStr = formatDate(startOfMonth) + "〜" + formatDate(endOfMonth);

    var expenseEntries = getMonthlyExpenseEntries(year, month);
    var incomeEntries = getMonthlyIncomeEntries(year, month);

    var totalExpenses = calculateTotalAmount(expenseEntries);
    var totalIncome = calculateTotalAmount(incomeEntries);

    var daysInMonth = endOfMonth.getDate();
    var adjustedBudget = Math.round((budgetPerWeek * daysInMonth / 7) / 100) * 100;
    var difference = totalExpenses - adjustedBudget;
    var percentage = adjustedBudget ? (totalExpenses / adjustedBudget) * 100 : 0;

    var categoryTotals = calculateCategoryTotals(expenseEntries);
    logMonthlySummaryDebug(month, expenseEntries, incomeEntries, categoryTotals);

    var geminiAnalysis = analyzeMonthlyWithGemini(expenseEntries, categoryTotals, totalExpenses, totalIncome, adjustedBudget, percentage, dateRangeStr);
    var body = buildMonthlySummaryMessage(
        dateRangeStr,
        totalIncome,
        totalExpenses,
        adjustedBudget,
        difference,
        percentage,
        expenseEntries,
        incomeEntries,
        categoryTotals,
        geminiAnalysis
    );

    return sendMonthlySummaryResult(action, currentDate, isStaging, body);
}

// 月次サマリーを Gemini で分析する
function analyzeMonthlyWithGemini(expenseEntries, categoryTotals, totalExpenses, totalIncome, adjustedBudget, percentage, dateRangeStr) {
    if (typeof PropertiesService === 'undefined') {
        return null;
    }

    var apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
    if (!apiKey) {
        Logger.log("Gemini API key is not set in script properties.");
        return null;
    }

    var modelFromProp = PropertiesService.getScriptProperties().getProperty("GEMINI_MODEL");
    if (modelFromProp && modelFromProp.indexOf("models/") === 0) {
        modelFromProp = modelFromProp.replace(/^models\//, "");
    }
    var modelsToTry = [];
    if (modelFromProp) {
        modelsToTry.push(modelFromProp);
    }
    if (typeof fetchGenerativeModels === 'function') {
        var autoModels = fetchGenerativeModels(apiKey);
        if (autoModels && autoModels.length) {
            modelsToTry = modelsToTry.concat(autoModels);
        }
    }
    if (modelsToTry.length === 0) {
        modelsToTry = [
            "gemini-2.5-flash",
            "gemini-2.5-pro",
            "gemini-2.0-flash",
            "gemini-1.5-flash"
        ];
    }
    var apiVersions = ["v1beta", "v1"];

    var expenseLines = expenseEntries.map(function (entry) {
        return formatDate(entry.date) + " [" + (entry.category || "未分類") + "] " + entry.name + " " + entry.amount + "円";
    });
    var categoryLines = Object.keys(categoryTotals).map(function (key) {
        return key + ": " + categoryTotals[key] + "円";
    });

    var prompt = [
        "あなたはプロの家計管理アドバイザーです。挨拶や自己紹介は禁止です。金銭感覚の改善を目的としたコーチとして、冷静な分析と、時には優しく、時には厳しく指導してください。1週間分の支出について、予算を超えないようアドバイスをください。カジュアルな敬語て対応してください。",
        "レシートは保管していませんが、代わりに全ての支出・収入をスプレッドシートに記録しています。単なる分析にとどまらず、「行動に落とし込める改善提案」を重視してください。感情的にならず、客観的かつ現実的な判断で、飴と鞭を使い分けてください。",
        "通勤時（勤務地：永田町）の通勤定期はありませんが、給与で補填されます。食事はスーパーでまとめ買いした上でほぼ自炊しており、外食は友人と会う時が多いです。",
        "対象期間: " + dateRangeStr,
        "月間予算(週予算換算): " + adjustedBudget + "円 / これまでの支出: " + totalExpenses + "円 (" + percentage.toFixed(1) + "%) / 収入: " + totalIncome + "円",
        "カテゴリ別支出: " + categoryLines.join(", "),
        "支出一覧:",
        expenseLines.join("\n"),
        "日本語で、1) 今月の傾向 2) 予算に対する評価 3) カテゴリ別の改善提案 を箇条書き最大3つずつ、プレーンテキストと絵文字のみで800文字以内にまとめてください。"
    ].join("\n");

    var payload = {
        contents: [{
            parts: [{
                text: prompt
            }]
        }],
        generationConfig: {
            temperature: 0.4,
            maxOutputTokens: 1000,
            thinkingConfig: {
                thinkingBudget: 400
            }
        }
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
