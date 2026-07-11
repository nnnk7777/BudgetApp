function analyzeExpensesWithAI(dataEntries, totalAmount, adjustedBudget, percentage, baseDate, weeklyAnalysisMode, options) {
    var analysisOptions = options || {};
    var plannedExpenses = analysisOptions.plannedExpenses || getUpcomingPlannedExpenses(baseDate);
    var plannedExpenseLabel = analysisOptions.plannedExpenseLabel || "今後の予定メモ";
    var upcomingExpenseLines = plannedExpenses.map(function (entry) {
        return formatDate(entry.date) + " [" + entry.title + "] " + entry.memo;
    });
    var prompt;
    var categoryRankingLines = getCategoryRankingLines(dataEntries);
    var weeklyBudgetCarryoverMemo = getWeeklyBudgetCarryoverMemoForWeek(baseDate);
    var weeklyBudgetCarryoverGuidance = buildWeeklyBudgetCarryoverGuidanceForPrompt(weeklyBudgetCarryoverMemo);
    var weeklyAnalysisModeGuidance = buildWeeklyAnalysisModeGuidanceForPrompt(weeklyAnalysisMode);
    prompt = buildExpenseSummaryPrompt(
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

    Logger.log("AI分析プロンプト：");
    Logger.log(prompt);

    return generatePreferredAiText(prompt, {
        temperature: 0.4,
        maxOutputTokens: 1000,
        thinkingConfig: {
            thinkingBudget: 400
        }
    }, {
        logContext: "expense_summary"
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

    plannedExpenses = cleanPlannedExpenseMemosWithAI(plannedExpenses);
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

    var aiCandidates = candidateEntries.filter(function (entry) {
        return isStrongExpenseCandidateForAi(plannedExpense, entry);
    });

    if (!aiCandidates.length) {
        return false;
    }

    return detectRecordedPlannedExpenseWithAI(plannedExpense, aiCandidates);
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

function extractPlannedExpenseAmounts(plannedExpense) {
    var text = [plannedExpense.title || "", plannedExpense.memo || ""].join(" ");
    var amounts = [];
    var seen = {};
    var matches = text.match(/\d[\d,]*(?:\.\d+)?(?=\s*円)/g) || [];

    matches.forEach(function (match) {
        var amount = parseInt(String(match).replace(/,/g, ""), 10);
        if (!isNaN(amount) && !seen[amount]) {
            seen[amount] = true;
            amounts.push(amount);
        }
    });

    return amounts;
}

function normalizeExpenseComparisonText(text) {
    return normalizeFullWidthNumbers(String(text || ""))
        .toLowerCase()
        .replace(/\s+/g, "")
        .replace(/[!！?？\-ー_./／\\,，。、「」（）()【】\[\]:'"@#&]/g, "");
}

function countSharedExpenseTokens(leftText, rightText) {
    var leftTokens = buildExpenseComparisonTokens(leftText);
    var rightTokens = buildExpenseComparisonTokens(rightText);
    var rightTokenMap = {};
    var count = 0;

    rightTokens.forEach(function (token) {
        rightTokenMap[token] = true;
    });

    leftTokens.forEach(function (token) {
        if (rightTokenMap[token]) {
            count += 1;
        }
    });

    return count;
}

function buildExpenseComparisonTokens(text) {
    var normalized = normalizeFullWidthNumbers(String(text || "")).toLowerCase();
    var rawTokens = normalized.split(/[^0-9a-zA-Zぁ-んァ-ヶ一-龠]+/);
    var tokenMap = {};

    rawTokens.forEach(function (token) {
        if (token.length < 2) {
            return;
        }
        tokenMap[token] = true;
    });

    return Object.keys(tokenMap);
}

function isStrongExpenseCandidateForAi(plannedExpense, entry) {
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

function detectRecordedPlannedExpenseWithAI(plannedExpense, candidateEntries) {
    var prompt;
    var result;
    var normalized;

    prompt = buildPlannedExpenseRecordCheckPrompt(plannedExpense, candidateEntries);

    result = generatePreferredAiText(prompt, {
        temperature: 0,
        maxOutputTokens: 10
    }, {
        logContext: "planned_expense_record_check"
    });
    if (!result.text) {
        return false;
    }

    normalized = String(result.text).trim().toUpperCase();
    Logger.log("予定支出の記録済みAI判定: title=" + plannedExpense.title + " provider=" + result.provider + " response=" + normalized);
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

function cleanPlannedExpenseMemosWithAI(plannedExpenses) {
    if (!plannedExpenses.length) {
        return plannedExpenses;
    }

    var prompt = buildPlannedExpenseMemoCleanupPrompt(plannedExpenses);
    var result = generatePreferredAiText(prompt, {
        temperature: 0,
        maxOutputTokens: 800
    }, {
        logContext: "planned_expense_memo_cleanup"
    });
    var responseText = result.text;
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
