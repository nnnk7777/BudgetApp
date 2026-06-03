function listUncategorizedExpenses() {
    return JSON.stringify({
        ok: true,
        items: fetchUncategorizedExpenseItems()
    });
}

function autofillUncategorizedExpenses(options) {
    options = options || {};

    var uncategorizedItems = fetchUncategorizedExpenseItems();
    var threshold = normalizeConfidenceThreshold(options.confidenceThreshold);
    var limit = normalizeLimit(options.limit);

    if (limit !== null) {
        uncategorizedItems = uncategorizedItems.slice(0, limit);
    }

    if (!uncategorizedItems.length) {
        return JSON.stringify({
            ok: true,
            updated: [],
            skipped: [],
            total: 0
        });
    }

    var suggestionResult = suggestExpenseCategories(uncategorizedItems, options);
    var updateResult = updateExpenseCategories(suggestionResult.suggestions, threshold);
    var response = {
        ok: true,
        updated: updateResult.updated,
        skipped: updateResult.skipped,
        total: uncategorizedItems.length
    };

    if (options.debug) {
        response.debug = suggestionResult.debug;
    }

    return JSON.stringify(response);
}

function suggestExpenseCategories(items, options) {
    options = options || {};
    var apiKey = getGeminiApiKey();
    if (!apiKey) {
        throw new Error('Gemini API key is not set.');
    }

    var categoryNames = getCategoryNames();
    var groupedItems = groupItemsByMonth(items);
    var suggestions = [];
    var debugEntries = [];
    var retryMissingIds = options.retryMissingIds !== false;

    Object.keys(groupedItems).sort(function (a, b) {
        return parseInt(a, 10) - parseInt(b, 10);
    }).forEach(function (monthKey) {
        var month = parseInt(monthKey, 10);
        var monthItems = groupedItems[monthKey];
        var historyItems = fetchCategoryHistoryEntriesForMonth(month);
        var monthDebug = {
            month: month,
            itemCount: monthItems.length,
            historyCount: historyItems.length
        };
        var attemptResult = requestCategorySuggestions(apiKey, monthItems, historyItems, categoryNames, month, options);
        var responseText = attemptResult.responseText;

        if (!responseText) {
            Logger.log("カテゴリ推定のGemini応答が空でした: month=" + month + " itemCount=" + monthItems.length);
            monthItems.forEach(function (item) {
                suggestions.push(buildSkippedSuggestion(
                    item,
                    null,
                    buildDetailedReason('Geminiから応答を取得できませんでした', 'Gemini response was empty.'),
                    {
                    errorCode: 'empty_response',
                    errorDetail: 'Gemini response was empty.',
                    responsePreview: ''
                    }
                ));
            });
            monthDebug.status = 'empty_response';
            if (options.debug) {
                monthDebug.rawResponse = responseText || '';
            }
            debugEntries.push(monthDebug);
            return;
        }

        var parseResult = attemptResult.parseResult;
        var parsedSuggestions = parseResult.suggestions.slice();
        var suggestionMap = {};
        parsedSuggestions.forEach(function (suggestion) {
            suggestionMap[suggestion.id] = suggestion;
        });

        var missingItems = monthItems.filter(function (item) {
            return !suggestionMap[item.id];
        });
        var retryDebug = null;

        if (retryMissingIds && missingItems.length) {
            var retryResult = requestCategorySuggestions(apiKey, missingItems, historyItems, categoryNames, month, options);
            retryDebug = {
                requestedIds: missingItems.map(function (item) {
                    return item.id;
                }),
                status: retryResult.parseResult.status,
                responsePreview: retryResult.parseResult.responsePreview,
                parseError: retryResult.parseResult.parseError || null
            };

            retryResult.parseResult.suggestions.forEach(function (suggestion) {
                suggestionMap[suggestion.id] = suggestion;
            });

            if (options.debug) {
                retryDebug.rawResponse = retryResult.responseText || '';
            }
        }

        monthItems.forEach(function (item) {
            if (suggestionMap[item.id]) {
                suggestions.push(suggestionMap[item.id]);
            } else {
                suggestions.push(buildSkippedSuggestion(
                    item,
                    null,
                    buildDetailedReason('Gemini応答に対象IDが含まれていません', 'Target ID was not included in Gemini response.'),
                    {
                    errorCode: 'missing_target_id',
                    errorDetail: 'Target ID was not included in Gemini response.',
                    responsePreview: buildResponsePreview(responseText)
                    }
                ));
            }
        });

        monthDebug.status = parseResult.status;
        monthDebug.responsePreview = parseResult.responsePreview;
        monthDebug.parseError = parseResult.parseError || null;
        monthDebug.extractedJsonPreview = parseResult.extractedJsonPreview || null;
        monthDebug.parseWarnings = parseResult.parseWarnings || [];
        if (retryDebug) {
            monthDebug.retry = retryDebug;
        }
        if (options.debug) {
            monthDebug.rawResponse = responseText;
        }
        debugEntries.push(monthDebug);
    });

    return {
        suggestions: suggestions,
        debug: debugEntries
    };
}

function updateExpenseCategories(suggestions, threshold) {
    return updateUncategorizedExpenseCategories(suggestions, threshold);
}

function groupItemsByMonth(items) {
    var grouped = {};
    items.forEach(function (item) {
        var key = String(item.month);
        if (!grouped[key]) {
            grouped[key] = [];
        }
        grouped[key].push(item);
    });
    return grouped;
}

function buildExpenseRecordId(month, row) {
    return month + "_" + row;
}

function normalizeConfidence(value) {
    var parsed = parseFloat(value);
    if (isNaN(parsed)) {
        return 0;
    }
    if (parsed < 0) {
        return 0;
    }
    if (parsed > 1) {
        return 1;
    }
    return parsed;
}

function normalizeConfidenceThreshold(value) {
    if (value === undefined || value === null || value === '') {
        return 0.9;
    }
    return normalizeConfidence(value);
}

function normalizeLimit(value) {
    if (value === undefined || value === null || value === '') {
        return null;
    }

    var parsed = parseInt(value, 10);
    if (isNaN(parsed) || parsed <= 0) {
        return null;
    }

    return parsed;
}

function normalizeAmountValue(value) {
    var normalized = parseFloat(String(normalizeFullWidthNumbers(String(value))).replace(/,/g, ""));
    return isNaN(normalized) ? value : normalized;
}

function isBlankCell(value) {
    return value === null || value === undefined || String(value).trim() === '';
}

function isSameAmount(left, right) {
    var normalizedLeft = normalizeAmountValue(left);
    var normalizedRight = normalizeAmountValue(right);
    return normalizedLeft === normalizedRight;
}

function mergeObjects(base, additions) {
    var result = {};
    Object.keys(base).forEach(function (key) {
        result[key] = base[key];
    });
    Object.keys(additions).forEach(function (key) {
        result[key] = additions[key];
    });
    return result;
}
