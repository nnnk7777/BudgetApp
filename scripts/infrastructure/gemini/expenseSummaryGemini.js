function analyzeExpensesWithGemini(dataEntries, totalAmount, adjustedBudget, percentage, baseDate) {
    var apiKey = getGeminiApiKey();
    if (!apiKey) {
        return null;
    }
    var expenseLines = dataEntries.map(function (entry) {
        return formatDate(entry.date) + " [" + (entry.category || "未分類") + "] " + entry.name + " " + entry.amount + "円";
    });
    var upcomingPlannedExpenses = getUpcomingPlannedExpenses(baseDate);
    var upcomingExpenseLines = upcomingPlannedExpenses.map(function (entry) {
        return formatDate(entry.date) + " [" + entry.title + "] " + entry.memo;
    });
    var categoryRankingLines = getCategoryRankingLines(dataEntries);
    var weeklyBudgetCarryoverMemo = getWeeklyBudgetCarryoverMemoForWeek(baseDate);
    var weeklyBudgetCarryoverGuidance = buildWeeklyBudgetCarryoverGuidanceForPrompt(weeklyBudgetCarryoverMemo);

    var prompt = [
        "あなたはプロの家計管理アドバイザーです。挨拶や自己紹介は禁止です。金銭感覚の改善を目的としたコーチとして、冷静な分析と、時には優しく、時には厳しく指導してください。1週間分の支出について、予算を超えないようアドバイスをください。カジュアルな敬語て対応してください。",
        "レシートは保管していませんが、代わりに全ての支出・収入をスプレッドシートに記録しています。単なる分析にとどまらず、「行動に落とし込める改善提案」を重視してください。感情的にならず、客観的かつ現実的な判断で、飴と鞭を使い分けてください。",
        "通勤時（勤務地：永田町）の通勤定期はありませんが、給与で補填されます。食事はスーパーでまとめ買いした上でほぼ自炊しており、外食は友人と会う時が多いです。",
        "Googleカレンダーの今後の予定メモに書かれた支出予定も考慮して助言してください。近いうちに大きな支出予定があるなら、今週の節約を強めに促してください。",
        "前週の予算差分メモがあれば補助情報として参照してください。特に前週が大きく超過していた場合は、その影響を締めの一言だけで済ませず、今週の傾向分析と次に意識すべき点の両方に反映してください。",
        "前週超過がある場合は、裁量支出や先送りできる支出への姿勢を普段より一段引き締めて提案してください。ただし今週の週予算、実支出、予定支出を優先し、前週差分だけで過度に断定したり、今週の予算を実質的に減額したような言い方はしないでください。",
        "前週が予算内に収まっていた場合でも、安心しすぎる助言にはせず、今週の数値を主軸に冷静に判断してください。",
        "1週間は月曜始まり・日曜終わりで考えて、今日までの傾向を数個と、次に意識すべき点を数個、箇条書きでまとめてください。箇条書きは最大5個までにし、それぞれに補足を付けてください。Markdown記法は使わず、プレーンな文字と絵文字のみで出力し、全体で500文字以内に収めてください。",
        "週予算: " + adjustedBudget + "円 / これまでの支出: " + totalAmount + "円 (" + percentage.toFixed(1) + "%)",
        "カテゴリ別支出ランキング: " + (categoryRankingLines.length ? categoryRankingLines.join(" / ") : "なし"),
        "前週の予算差分メモ: " + formatWeeklyBudgetCarryoverMemoForPrompt(weeklyBudgetCarryoverMemo),
        "前週差分の反映方針: " + weeklyBudgetCarryoverGuidance,
        "今後の予定メモ: " + (upcomingExpenseLines.length ? upcomingExpenseLines.join(" / ") : "なし"),
        "以下は支出一覧(日付、カテゴリ、名称、金額)です。名称だけで内容が不明瞭な場合はカテゴリから内容を推定してください:",
        expenseLines.join("\n")
    ].join("\n");

    Logger.log("Geminiプロンプト：");
    Logger.log(prompt);

    return generateGeminiText(apiKey, prompt, {
        temperature: 0.4,
        maxOutputTokens: 1000,
        thinkingConfig: {
            thinkingBudget: 400
        }
    });
}

function getUpcomingPlannedExpenses(baseDate) {
    var startDate = new Date(baseDate.getFullYear(), baseDate.getMonth(), baseDate.getDate());
    var lookaheadDays = getUpcomingExpenseLookaheadDays();
    var endDate = new Date(startDate);
    endDate.setDate(endDate.getDate() + lookaheadDays);
    return getPlannedExpensesInRange(startDate, endDate);
}

function getPlannedExpensesForCurrentWeek(baseDate) {
    var range = getWeekRange(baseDate);
    var startDate = new Date(baseDate.getFullYear(), baseDate.getMonth(), baseDate.getDate());
    var endDateExclusive = new Date(range.endDate);
    endDateExclusive.setDate(endDateExclusive.getDate() + 1);
    return getPlannedExpensesInRange(startDate, endDateExclusive);
}

function getPlannedExpensesInRange(startDate, endDate) {
    var events = getCalendarEventsInRange(startDate, endDate);
    var plannedExpenses = [];
    Logger.log(
        "予定支出検索: start=" +
            formatDate(startDate) +
            " end=" +
            formatDate(endDate) +
            " events=" +
            events.length
    );

    events.forEach(function (event) {
        var title = event.getTitle() || "予定";
        var description = sanitizePlannedExpenseMemo(event.getDescription() || "");
        if (isWeeklyBudgetCarryoverMemo(title, description)) {
            return;
        }
        if (!hasPlannedExpenseMemo(description, title)) {
            return;
        }

        plannedExpenses.push({
            title: title,
            date: event.getStartTime(),
            memo: description
        });
    });

    plannedExpenses = cleanPlannedExpenseMemosWithGemini(plannedExpenses);

    plannedExpenses.sort(function (a, b) {
        return a.date.getTime() - b.date.getTime();
    });

    Logger.log("予定支出取得件数: " + plannedExpenses.length);

    return plannedExpenses;
}

function formatUpcomingPlannedExpenseLines(plannedExpenses) {
    return plannedExpenses.map(function (entry) {
        return "・" + formatDate(entry.date) + " - " + entry.title + ": " + entry.memo;
    });
}

function formatWeeklyBudgetCarryoverMemoForPrompt(memo) {
    if (!memo) {
        return "なし";
    }

    return [
        memo.dateRangeStr || "対象週不明",
        "予算差分 " + (memo.difference >= 0 ? "+" : "-") + Math.abs(memo.difference) + "円",
        "実支出 " + memo.totalAmount + "円",
        "週予算 " + memo.adjustedBudget + "円"
    ].join(" / ");
}

function buildWeeklyBudgetCarryoverGuidanceForPrompt(memo) {
    var overrunRatio;

    if (!memo) {
        return "前週差分なし。今週の実績と予定支出だけで判断する。";
    }

    if (memo.difference <= 0) {
        return "前週は予算内。評価としては軽く触れる程度に留め、今週を緩めすぎない。";
    }

    overrunRatio = memo.adjustedBudget ? (memo.difference / memo.adjustedBudget) : 0;
    if (overrunRatio >= 0.25) {
        return "前週は大きな予算超過。今週は全体的に一段厳しめのトーンで、節約余地や回避策を複数箇所で具体的に指摘する。ただし今週の数値が良好なら前週だけで悲観しすぎない。";
    }

    if (overrunRatio >= 0.1) {
        return "前週はやや大きめの予算超過。今週は通常より引き締め寄りに評価し、無理なく削れる支出を明確に示す。ただし今週の実績が落ち着いていれば過度には引きずらない。";
    }

    return "前週は軽度の予算超過。分析のどこかで一度は触れつつ、今週の数値を主軸にバランスよく助言する。";
}

function cleanPlannedExpenseMemosWithGemini(plannedExpenses) {
    if (!plannedExpenses.length) {
        return plannedExpenses;
    }

    var apiKey = getGeminiApiKey();
    if (!apiKey) {
        return plannedExpenses;
    }

    var prompt = [
        "以下はGoogleカレンダー予定のタイトルとメモです。",
        "目的は、支出予定として意味のある情報だけを残し、Googleの自動追記やURLや案内文など無関係な文を除去することです。",
        "各要素について cleanedMemo を返してください。",
        "ルール:",
        "- 支出予定として意味がある部分だけ残す",
        "- URL、Googleの案内文、自動生成の説明、予約メール由来の定型文は削除する",
        "- 金額がなくても、予定に関係する買い物や外食のメモなら残してよい",
        "- 予定に関係する情報が何も残らないなら cleanedMemo を空文字にする",
        "- 出力は各行 `index|cleanedMemo` のみ",
        "- 説明文、Markdown、コードブロック、jsonという語は出力しない",
        "- 例: `0|予定：服4000円`",
        JSON.stringify(plannedExpenses.map(function (entry, index) {
            return {
                index: index,
                title: entry.title,
                memo: entry.memo
            };
        }))
    ].join("\n");

    var responseText = generateGeminiText(apiKey, prompt, {
        temperature: 0,
        maxOutputTokens: 800
    });
    if (!responseText) {
        return plannedExpenses;
    }

    var cleanedByIndex = {};
    var parsedLines = parsePipeResponse(responseText);
    if (!parsedLines.length) {
        Logger.log("予定メモ整形の行解析に失敗: " + responseText);
        return plannedExpenses;
    }

    parsedLines.forEach(function (item) {
        if (typeof item.index === 'number' && typeof item.cleanedMemo === 'string') {
            cleanedByIndex[item.index] = item.cleanedMemo.trim();
        }
    });

    var cleanedExpenses = plannedExpenses.map(function (entry, index) {
        var cleanedMemo = cleanedByIndex.hasOwnProperty(index)
            ? cleanedByIndex[index]
            : entry.memo;

        return {
            title: entry.title,
            date: entry.date,
            memo: cleanedMemo
        };
    }).filter(function (entry) {
        return entry.memo !== "";
    });

    Logger.log("予定メモ整形後件数: " + cleanedExpenses.length);
    return cleanedExpenses;
}

function hasPlannedExpenseMemo(text, title) {
    var combined = [title || "", text || ""].join(" ").trim();
    if (!combined) {
        return false;
    }

    if (/予定/.test(combined)) {
        return true;
    }

    if (/[¥￥]/.test(combined) || /\d[\d,]*\s*円/.test(combined)) {
        return true;
    }

    return /(交通費|食費|外食|買い物|ランチ|ディナー|飲み|会費|チケット|ホテル|宿|タクシー|電車|バス|飛行機)/.test(combined);
}

function sanitizePlannedExpenseMemo(text) {
    return text
        .replace(/\r\n/g, "\n")
        .replace(/\r/g, "\n")
        .split("\n")
        .map(function (line) {
            return line.trim();
        })
        .filter(function (line) {
            return line !== "";
        })
        .join(" / ");
}
