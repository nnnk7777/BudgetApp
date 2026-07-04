function analyzeExpensesWithGemini(dataEntries, totalAmount, adjustedBudget, percentage, baseDate, weeklyAnalysisMode, options) {
    var apiKey = getGeminiApiKey();
    var analysisOptions = options || {};
    var plannedExpenses = analysisOptions.plannedExpenses || getUpcomingPlannedExpenses(baseDate);
    var plannedExpenseLabel = analysisOptions.plannedExpenseLabel || "今後の予定メモ";
    if (!apiKey) {
        return null;
    }
    var expenseLines = dataEntries.map(function (entry) {
        return formatDate(entry.date) + " [" + (entry.category || "未分類") + "] " + entry.name + " " + entry.amount + "円";
    });
    var upcomingExpenseLines = plannedExpenses.map(function (entry) {
        return formatDate(entry.date) + " [" + entry.title + "] " + entry.memo;
    });
    var categoryRankingLines = getCategoryRankingLines(dataEntries);
    var weeklyBudgetCarryoverMemo = getWeeklyBudgetCarryoverMemoForWeek(baseDate);
    var weeklyBudgetCarryoverGuidance = buildWeeklyBudgetCarryoverGuidanceForPrompt(weeklyBudgetCarryoverMemo);
    var weeklyAnalysisModeGuidance = buildWeeklyAnalysisModeGuidanceForPrompt(weeklyAnalysisMode);
    var prompt = buildExpenseSummaryPrompt(
        dataEntries,
        totalAmount,
        adjustedBudget,
        percentage,
        baseDate,
        categoryRankingLines,
        weeklyAnalysisMode,
        weeklyAnalysisModeGuidance,
        weeklyBudgetCarryoverMemo,
        weeklyBudgetCarryoverGuidance,
        plannedExpenseLabel,
        upcomingExpenseLines
    );

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

function buildExpenseSummaryPrompt(
    dataEntries,
    totalAmount,
    adjustedBudget,
    percentage,
    baseDate,
    categoryRankingLines,
    weeklyAnalysisMode,
    weeklyAnalysisModeGuidance,
    weeklyBudgetCarryoverMemo,
    weeklyBudgetCarryoverGuidance,
    plannedExpenseLabel,
    upcomingExpenseLines
) {
    var expenseLines = dataEntries.map(function (entry) {
        return formatDate(entry.date) + " [" + (entry.category || "未分類") + "] " + entry.name + " " + entry.amount + "円";
    });
    var prompt = [
        "あなたはプロの家計管理アドバイザーです。挨拶や自己紹介は禁止です。金銭感覚の改善を目的としたコーチとして、冷静な分析と、時には優しく、時には厳しく指導してください。1週間分の支出について、予算を超えないようアドバイスをください。カジュアルな敬語て対応してください。",
        "レシートは保管していませんが、代わりに全ての支出・収入をスプレッドシートに記録しています。単なる分析にとどまらず、「行動に落とし込める改善提案」を重視してください。感情的にならず、客観的かつ現実的な判断で、飴と鞭を使い分けてください。",
        "通勤時（勤務地：永田町）の通勤定期はありませんが、給与で補填されます。食事はスーパーでまとめ買いした上でほぼ自炊しており、外食は友人と会う時が多いです。",
        "Googleカレンダーの今後の予定メモに書かれた支出予定も考慮して助言してください。近いうちに大きな支出予定があるなら、今週の節約を強めに促してください。",
        "前週の予算差分メモがあれば補助情報として参照してください。特に前週が大きく超過していた場合は、その影響を締めの一言だけで済ませず、今週の傾向分析と次に意識すべき点の両方に反映してください。",
        "前週超過がある場合は、裁量支出や先送りできる支出への姿勢を普段より一段引き締めて提案してください。ただし今週の週予算、実支出、予定支出を優先し、前週差分だけで過度に断定したり、今週の予算を実質的に減額したような言い方はしないでください。",
        "前週が予算内に収まっていた場合でも、安心しすぎる助言にはせず、今週の数値を主軸に冷静に判断してください。",
        buildAnalysisDateContext(baseDate),
        "与えられた分析基準日・対象期間・支出一覧だけを根拠に判断してください。曜日や週の進捗を勝手に推測しないでください。",
        "特に、基準日が火曜日なのに『今日が日曜日』と書いたり、対象週が月をまたぐのに月末で週が終わったかのように扱うのは禁止です。",
        "1週間は月曜始まり・日曜終わりで考えて、今日までの傾向を数個と、次に意識すべき点を数個、箇条書きでまとめてください。箇条書きは最大5個までにし、それぞれに補足を付けてください。Markdown記法は使わず、プレーンな文字と絵文字のみで出力し、全体で500文字以内に収めてください。",
        "週予算: " + adjustedBudget + "円 / これまでの支出: " + totalAmount + "円 (" + percentage.toFixed(1) + "%)",
        "カテゴリ別支出ランキング: " + (categoryRankingLines.length ? categoryRankingLines.join(" / ") : "なし"),
        "今週の分析モード: " + formatWeeklyAnalysisModeForPrompt(weeklyAnalysisMode),
        "分析モードの反映方針: " + weeklyAnalysisModeGuidance,
        "前週の予算差分メモ: " + formatWeeklyBudgetCarryoverMemoForPrompt(weeklyBudgetCarryoverMemo),
        "前週差分の反映方針: " + weeklyBudgetCarryoverGuidance,
        plannedExpenseLabel + ": " + (upcomingExpenseLines.length ? upcomingExpenseLines.join(" / ") : "なし"),
        "以下は支出一覧(日付、カテゴリ、名称、金額)です。名称だけで内容が不明瞭な場合はカテゴリから内容を推定してください:",
        expenseLines.join("\n")
    ].join("\n");

    return prompt;
}

function formatWeeklyAnalysisModeForPrompt(modeResult) {
    if (!modeResult || modeResult.mode !== WEEKLY_ANALYSIS_MODE_FRUGAL) {
        return "通常モード";
    }

    return modeResult.label;
}

function buildWeeklyAnalysisModeGuidanceForPrompt(modeResult) {
    if (!modeResult || modeResult.mode !== WEEKLY_ANALYSIS_MODE_FRUGAL) {
        return "通常モード。必要以上に厳しくしすぎず、数値と予定支出に基づいて冷静に評価する。";
    }

    return "節制モード。通常より明確に厳しめに評価し、裁量支出、先送りできる支出、習慣化すると危険な出費には甘くしない。数値がまだ良くても油断を促す言い方は避け、今のうちに削れる支出や抑えるべき行動を具体的に指摘する。ただし感情的な説教にはせず、改善行動が明確になる実務的で手厳しい助言を優先する。";
}

function buildAnalysisDateContext(baseDate) {
    var weekRange = getWeekRange(baseDate);
    var weekday = getJapaneseWeekday(baseDate.getDay());
    var isWeekClosed = baseDate.getDay() === 0;
    var weekStatus = isWeekClosed
        ? "今週は日曜まで終了した確定値として扱う。"
        : "今週はまだ途中であり、" + weekday + "時点の途中経過として扱う。";

    return [
        "分析基準日: " + formatDate(baseDate) + " (" + weekday + ")",
        "分析対象の週: " + formatDate(weekRange.startDate) + "〜" + formatDate(weekRange.endDate) + " の1週間",
        "週の扱い: 1週間は月曜始まり・日曜終わり。月をまたいでも同じ週として扱う。年をまたぐ場合だけ、その年に含まれる日付までを対象にする。",
        "分析時点: " + weekStatus
    ].join(" / ");
}

function getJapaneseWeekday(dayIndex) {
    return ["日曜日", "月曜日", "火曜日", "水曜日", "木曜日", "金曜日", "土曜日"][dayIndex] || "";
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

function getPlannedExpensesForNextWeek(baseDate) {
    var range = getWeekRange(baseDate);
    var startDate = new Date(range.endDate);
    var endDateExclusive = new Date(range.endDate);
    startDate.setDate(startDate.getDate() + 1);
    endDateExclusive.setDate(endDateExclusive.getDate() + 8);
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
        if (isWeeklyAnalysisModeEvent(title, description)) {
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
    plannedExpenses = filterRecordedPlannedExpenses(plannedExpenses);

    plannedExpenses.sort(function (a, b) {
        return a.date.getTime() - b.date.getTime();
    });

    Logger.log("予定支出取得件数: " + plannedExpenses.length);

    return plannedExpenses;
}

function filterRecordedPlannedExpenses(plannedExpenses) {
    if (!plannedExpenses.length) {
        return plannedExpenses;
    }

    var actualEntries = getExpenseEntriesForDates(buildExpenseComparisonDates(plannedExpenses));

    return plannedExpenses.filter(function (plannedExpense) {
        var candidateEntries = findRecordedExpenseCandidates(plannedExpense, actualEntries);
        var shouldExclude = shouldExcludePlannedExpense(plannedExpense, candidateEntries);

        if (shouldExclude) {
            Logger.log(
                "記録済み予定を除外: " +
                    formatDate(plannedExpense.date) +
                    " title=" +
                    plannedExpense.title +
                    " memo=" +
                    plannedExpense.memo
            );
        }

        return !shouldExclude;
    });
}

function buildExpenseComparisonDates(plannedExpenses) {
    var dateMap = {};
    var dates = [];

    plannedExpenses.forEach(function (plannedExpense) {
        [-1, 0, 1].forEach(function (offset) {
            var date = new Date(plannedExpense.date);
            var key;
            date.setDate(date.getDate() + offset);
            key = date.getFullYear() + "-" + (date.getMonth() + 1) + "-" + date.getDate();
            if (!dateMap[key]) {
                dateMap[key] = true;
                dates.push(new Date(date.getFullYear(), date.getMonth(), date.getDate()));
            }
        });
    });

    return dates;
}

function findRecordedExpenseCandidates(plannedExpense, actualEntries) {
    var plannedAmounts = extractPlannedExpenseAmounts(plannedExpense);
    var plannedText = [plannedExpense.title || "", plannedExpense.memo || ""].join(" ");
    var plannedNormalized = normalizeExpenseComparisonText(plannedText);

    return actualEntries.filter(function (entry) {
        var dayDistance = Math.abs(calculateDateDistanceInDays(plannedExpense.date, entry.date));
        var actualAmount = parseFloat(entry.amount);
        var sharedTokenCount = countSharedExpenseTokens(plannedText, [entry.name || "", entry.category || ""].join(" "));
        var amountMatched = !plannedAmounts.length || plannedAmounts.some(function (plannedAmount) {
            return !isNaN(actualAmount) && Math.abs(plannedAmount - actualAmount) <= 300;
        });
        var nameNormalized = normalizeExpenseComparisonText(entry.name || "");
        var containsMatchedName = !!nameNormalized && (
            plannedNormalized.indexOf(nameNormalized) !== -1 ||
            nameNormalized.indexOf(plannedNormalized) !== -1
        );

        if (dayDistance > 1) {
            return false;
        }

        return amountMatched || sharedTokenCount > 0 || containsMatchedName;
    });
}

function shouldExcludePlannedExpense(plannedExpense, candidateEntries) {
    if (!candidateEntries.length) {
        return false;
    }

    var exactRuleMatch = findExactRecordedExpenseMatch(plannedExpense, candidateEntries);
    if (exactRuleMatch) {
        return true;
    }

    var geminiCandidates = candidateEntries.filter(function (entry) {
        return isStrongExpenseCandidateForGemini(plannedExpense, entry);
    });

    if (!geminiCandidates.length) {
        return false;
    }

    return detectRecordedPlannedExpenseWithGemini(plannedExpense, geminiCandidates);
}

function findExactRecordedExpenseMatch(plannedExpense, candidateEntries) {
    var plannedAmounts = extractPlannedExpenseAmounts(plannedExpense);
    var plannedText = [plannedExpense.title || "", plannedExpense.memo || ""].join(" ");
    var plannedNormalized = normalizeExpenseComparisonText(plannedText);

    return candidateEntries.some(function (entry) {
        var actualAmount = parseFloat(entry.amount);
        var nameNormalized = normalizeExpenseComparisonText(entry.name || "");
        var dayDistance = Math.abs(calculateDateDistanceInDays(plannedExpense.date, entry.date));
        var exactAmountMatched = plannedAmounts.length && plannedAmounts.some(function (plannedAmount) {
            return !isNaN(actualAmount) && plannedAmount === actualAmount;
        });
        var strongTextMatch = countSharedExpenseTokens(plannedText, [entry.name || "", entry.category || ""].join(" ")) >= 1 ||
            (!!nameNormalized && plannedNormalized.indexOf(nameNormalized) !== -1);

        return dayDistance === 0 && exactAmountMatched && strongTextMatch;
    });
}

function isStrongExpenseCandidateForGemini(plannedExpense, entry) {
    var plannedAmounts = extractPlannedExpenseAmounts(plannedExpense);
    var actualAmount = parseFloat(entry.amount);
    var dayDistance = Math.abs(calculateDateDistanceInDays(plannedExpense.date, entry.date));
    var sharedTokenCount = countSharedExpenseTokens(
        [plannedExpense.title || "", plannedExpense.memo || ""].join(" "),
        [entry.name || "", entry.category || ""].join(" ")
    );

    if (dayDistance > 1) {
        return false;
    }

    if (sharedTokenCount >= 1 && !plannedAmounts.length) {
        return true;
    }

    return plannedAmounts.some(function (plannedAmount) {
        return !isNaN(actualAmount) && Math.abs(plannedAmount - actualAmount) <= 300;
    });
}

function detectRecordedPlannedExpenseWithGemini(plannedExpense, candidateEntries) {
    var apiKey = getGeminiApiKey();
    var prompt;
    var responseText;
    var normalized;

    if (!apiKey) {
        return false;
    }

    prompt = [
        "以下のGoogleカレンダーの支出予定が、家計簿に記録済みの支出として扱ってよいか判定してください。",
        "保守的に判定し、自信がない場合は NO にしてください。",
        "出力は YES か NO のどちらか1語のみです。",
        "予定:",
        JSON.stringify({
            date: formatDate(plannedExpense.date),
            title: plannedExpense.title,
            memo: plannedExpense.memo
        }),
        "家計簿候補:",
        JSON.stringify(candidateEntries.map(function (entry) {
            return {
                date: formatDate(entry.date),
                category: entry.category || "",
                name: entry.name || "",
                amount: entry.amount
            };
        }))
    ].join("\n");

    responseText = generateGeminiText(apiKey, prompt, {
        temperature: 0,
        maxOutputTokens: 10
    });
    if (!responseText) {
        return false;
    }

    normalized = String(responseText).trim().toUpperCase();
    Logger.log("予定支出の記録済みGemini判定: title=" + plannedExpense.title + " response=" + normalized);
    return normalized.indexOf("YES") === 0;
}

function calculateDateDistanceInDays(leftDate, rightDate) {
    var left = new Date(leftDate.getFullYear(), leftDate.getMonth(), leftDate.getDate());
    var right = new Date(rightDate.getFullYear(), rightDate.getMonth(), rightDate.getDate());
    return Math.round((left.getTime() - right.getTime()) / (24 * 60 * 60 * 1000));
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
