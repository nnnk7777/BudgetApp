var SPECIAL_EXPENSE_CATEGORY = "特別費";
var SPECIAL_EXPENSE_REVIEW_CACHE_PROPERTY = "SPECIAL_EXPENSE_REVIEW_CACHE";
var SPECIAL_EXPENSE_OPENAI_RETRY_COUNT = 2;

function reviewSpecialExpensesWithAI(entries) {
    var candidates = entries.filter(function (entry) {
        return String(entry.category || "").trim() === SPECIAL_EXPENSE_CATEGORY;
    });
    var cachedDecisions = getSpecialExpenseReviewCache();
    var uncachedCandidates = [];
    var decisions = [];
    var aiResult = null;

    if (!candidates.length) {
        return {
            approvedEntries: [],
            rejectedEntries: [],
            hasCandidates: false,
            aiResult: null
        };
    }

    candidates.forEach(function (entry) {
        var cachedDecision = cachedDecisions[getSpecialExpenseEntryKey(entry)];

        if (cachedDecision) {
            decisions.push({
                entry: entry,
                approved: cachedDecision.approved,
                reason: cachedDecision.reason,
                hasAiDecision: true
            });
            return;
        }

        uncachedCandidates.push(entry);
    });

    if (uncachedCandidates.length) {
        aiResult = reviewUncachedSpecialExpensesWithAI(uncachedCandidates);
        decisions = decisions.concat(aiResult.decisions);
        cacheSpecialExpenseReviewDecisions(decisions, cachedDecisions);
    }

    return {
        approvedEntries: decisions.filter(function (decision) {
            return decision.approved;
        }),
        rejectedEntries: decisions.filter(function (decision) {
            return !decision.approved;
        }),
        hasCandidates: true,
        aiResult: aiResult
    };
}

function reviewUncachedSpecialExpensesWithAI(candidates) {
    var unresolvedCandidates = candidates.slice();
    var decisionsByEntryKey = {};
    var aiResult = null;
    var attempt;

    for (attempt = 0; attempt < SPECIAL_EXPENSE_OPENAI_RETRY_COUNT && unresolvedCandidates.length; attempt++) {
        aiResult = generateSpecialExpenseReviewText(unresolvedCandidates, {
            logContext: "special_expense_review_openai_attempt_" + (attempt + 1),
            skipGemini: true
        });
        collectSpecialExpenseDecisions(decisionsByEntryKey, parseSpecialExpenseReviewResponse(aiResult.text, unresolvedCandidates));
        unresolvedCandidates = getUnresolvedSpecialExpenseCandidates(unresolvedCandidates, decisionsByEntryKey);

        if (unresolvedCandidates.length) {
            Logger.log("特別費AI判定のOpenAI応答が不完全なためリトライします: " + unresolvedCandidates.length + "件");
        }
    }

    if (unresolvedCandidates.length) {
        aiResult = generateSpecialExpenseReviewText(unresolvedCandidates, {
            logContext: "special_expense_review_gemini_fallback",
            skipOpenAi: true,
            fallbackReason: "openai_response_incomplete"
        });
        collectSpecialExpenseDecisions(decisionsByEntryKey, parseSpecialExpenseReviewResponse(aiResult.text, unresolvedCandidates));
        unresolvedCandidates = getUnresolvedSpecialExpenseCandidates(unresolvedCandidates, decisionsByEntryKey);
    }

    return {
        decisions: candidates.map(function (entry) {
            return decisionsByEntryKey[getSpecialExpenseEntryKey(entry)] || createUnresolvedSpecialExpenseDecision(entry);
        }),
        provider: aiResult.provider,
        model: aiResult.model,
        usedFallback: aiResult.usedFallback,
        fallbackReason: aiResult.fallbackReason
    };
}

function generateSpecialExpenseReviewText(candidates, options) {
    return generatePreferredAiText(buildSpecialExpenseReviewPrompt(candidates), {
        temperature: 0,
        maxOutputTokens: Math.max(80, candidates.length * 40)
    }, options);
}

function collectSpecialExpenseDecisions(decisionsByEntryKey, decisions) {
    decisions.forEach(function (decision) {
        if (decision.hasAiDecision) {
            decisionsByEntryKey[getSpecialExpenseEntryKey(decision.entry)] = decision;
        }
    });
}

function getUnresolvedSpecialExpenseCandidates(candidates, decisionsByEntryKey) {
    return candidates.filter(function (entry) {
        return !decisionsByEntryKey[getSpecialExpenseEntryKey(entry)];
    });
}

function parseSpecialExpenseReviewResponse(text, candidates) {
    var decisionByIndex = {};

    String(text || "").split("\n").forEach(function (line) {
        var parts = line.trim().split("|");
        var index = parseInt(parts[0], 10);
        var status = String(parts[1] || "").trim().toLowerCase();

        if (isNaN(index) || !candidates[index] || (status !== "approved" && status !== "rejected")) {
            return;
        }

        decisionByIndex[index] = {
            entry: candidates[index],
            approved: status === "approved",
            reason: parts.slice(2).join("|").trim() || "理由なし",
            hasAiDecision: true
        };
    });

    return candidates.map(function (entry, index) {
        return decisionByIndex[index] || createUnresolvedSpecialExpenseDecision(entry);
    });
}

function createUnresolvedSpecialExpenseDecision(entry) {
    return {
        entry: entry,
        approved: false,
        reason: "AI判定を取得できなかったため",
        hasAiDecision: false
    };
}

function getSpecialExpenseReviewCache() {
    var propertyValue;

    if (typeof PropertiesService === "undefined") {
        return {};
    }

    propertyValue = PropertiesService.getScriptProperties().getProperty(SPECIAL_EXPENSE_REVIEW_CACHE_PROPERTY);
    if (!propertyValue) {
        return {};
    }

    try {
        return JSON.parse(propertyValue);
    } catch (error) {
        Logger.log("特別費判定キャッシュの読込に失敗: " + error);
        return {};
    }
}

function cacheSpecialExpenseReviewDecisions(decisions, cachedDecisions) {
    var cache = cachedDecisions || {};

    if (typeof PropertiesService === "undefined") {
        return;
    }

    decisions.forEach(function (decision) {
        if (!decision.hasAiDecision) {
            return;
        }

        cache[getSpecialExpenseEntryKey(decision.entry)] = {
            approved: decision.approved,
            reason: decision.reason
        };
    });

    PropertiesService.getScriptProperties().setProperty(
        SPECIAL_EXPENSE_REVIEW_CACHE_PROPERTY,
        JSON.stringify(cache)
    );
}

function getBudgetTargetEntries(entries, specialExpenseReview) {
    var approvedEntryKeys = {};

    specialExpenseReview.approvedEntries.forEach(function (decision) {
        approvedEntryKeys[getSpecialExpenseEntryKey(decision.entry)] = true;
    });

    return entries.filter(function (entry) {
        return !approvedEntryKeys[getSpecialExpenseEntryKey(entry)];
    });
}

function getSpecialExpenseEntryKey(entry) {
    var date = entry.date;
    var dateKey = date.getFullYear() + "-" + (date.getMonth() + 1) + "-" + date.getDate();

    return [dateKey, entry.category || "", entry.name || "", entry.amount || ""].join("|");
}

function calculateApprovedSpecialExpenseTotal(specialExpenseReview) {
    return calculateTotalAmount(specialExpenseReview.approvedEntries.map(function (decision) {
        return decision.entry;
    }));
}

function calculateApprovedSpecialExpenseTotalForEntries(entries, specialExpenseReview) {
    return calculateTotalAmount(entries) - calculateTotalAmount(getBudgetTargetEntries(entries, specialExpenseReview));
}

function buildSpecialExpenseReviewSection(specialExpenseReview, label) {
    var lines = ["◆ " + label + "の特別費チェック"];

    if (!specialExpenseReview.hasCandidates) {
        return lines.concat(["・特別費の記録なし"]).join("\n") + "\n";
    }

    specialExpenseReview.approvedEntries.forEach(function (decision) {
        lines.push("・予算対象外: " + decision.entry.name + " " + decision.entry.amount + "円（" + decision.reason + "）");
    });
    specialExpenseReview.rejectedEntries.forEach(function (decision) {
        lines.push("・予算対象に含める: " + decision.entry.name + " " + decision.entry.amount + "円（" + decision.reason + "）");
    });

    return lines.join("\n") + "\n";
}
