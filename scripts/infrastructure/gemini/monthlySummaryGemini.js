function analyzeMonthlyWithAI(expenseEntries, categoryTotals, totalExpenses, totalIncome, adjustedBudget, percentage, dateRangeStr) {
    var categoryLines = Object.keys(categoryTotals).map(function (key) {
        return key + ": " + categoryTotals[key] + "円";
    });

    var prompt = [
        "あなたはプロの家計管理アドバイザーです。挨拶や自己紹介は禁止です。金銭感覚の改善を目的としたコーチとして、冷静な分析と、時には優しく、時には厳しく指導してください。今回は月次レポートです。1か月分の支出と収入について、予算を超えないようアドバイスをください。カジュアルな敬語で対応してください。",
        "レシートは保管していませんが、代わりに全ての支出・収入をスプレッドシートに記録しています。単なる分析にとどまらず、「行動に落とし込める改善提案」を重視してください。感情的にならず、客観的かつ現実的な判断で、飴と鞭を使い分けてください。",
        "通勤時（勤務地：永田町）の通勤定期はありませんが、給与で補填されます。食事はスーパーでまとめ買いした上でほぼ自炊しており、外食は友人と会う時が多いです。",
        "これは月次分析なので、今週・来週・1週間・週末・今後7日間といった週次前提の表現や助言は禁止です。対象月全体だけを根拠にしてください。",
        "『これからの1週間をどう過ごすか』のような文言は禁止です。対象期間は月単位であり、月全体の振り返りと次の月に向けた改善だけを書いてください。",
        "対象期間: " + dateRangeStr,
        "月間予算(週予算換算): " + adjustedBudget + "円 / これまでの支出: " + totalExpenses + "円 (" + percentage.toFixed(1) + "%) / 収入: " + totalIncome + "円",
        "カテゴリ別支出: " + categoryLines.join(", "),
        "支出一覧:",
        expenseEntries.map(function (entry) {
            return formatDate(entry.date) + " [" + (entry.category || "未分類") + "] " + entry.name + " " + entry.amount + "円";
        }).join("\n"),
        "日本語で、1) 今月の傾向 2) 予算に対する評価 3) 来月に向けた改善提案 を書いてください。",
        "各項目は最大3つの箇条書きにし、全体で800文字以内にしてください。",
        "Markdown見出し、#、###、太字、表、コードブロックは使わず、プレーンテキストと絵文字のみで出力してください。"
    ].join("\n");

    return generatePreferredAiText(prompt, {
        temperature: 0.4,
        maxOutputTokens: 1000,
        thinkingConfig: {
            thinkingBudget: 400
        }
    }, {
        logContext: "monthly_summary"
    });
}
