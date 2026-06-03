function analyzeMonthlyWithGemini(expenseEntries, categoryTotals, totalExpenses, totalIncome, adjustedBudget, percentage, dateRangeStr) {
    var apiKey = getGeminiApiKey();
    if (!apiKey) {
        return null;
    }

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

    return generateGeminiText(apiKey, prompt, {
        temperature: 0.4,
        maxOutputTokens: 1000,
        thinkingConfig: {
            thinkingBudget: 400
        }
    });
}
