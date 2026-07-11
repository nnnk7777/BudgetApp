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
    if (!getOpenAiApiKey(true) && !getGeminiApiKey(true)) {
        throw new Error('OpenAI API key and Gemini API key are not set.');
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
        var attemptResult = requestCategorySuggestionsWithAI(monthItems, historyItems, categoryNames, month, options);
        var responseText = attemptResult.responseText;

        if (!responseText) {
            Logger.log("カテゴリ推定のAI応答が空でした: month=" + month + " itemCount=" + monthItems.length);
            monthItems.forEach(function (item) {
                suggestions.push(buildSkippedSuggestion(
                    item,
                    null,
                    buildDetailedReason('AIから応答を取得できませんでした', 'AI response was empty.'),
                    {
                    errorCode: 'empty_response',
                    errorDetail: 'AI response was empty.',
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
            var retryResult = requestCategorySuggestionsWithAI(missingItems, historyItems, categoryNames, month, options);
            retryDebug = {
                requestedIds: missingItems.map(function (item) {
                    return item.id;
                }),
                provider: retryResult.provider,
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
                    buildDetailedReason('AI応答に対象IDが含まれていません', 'Target ID was not included in AI response.'),
                    {
                    errorCode: 'missing_target_id',
                    errorDetail: 'Target ID was not included in AI response.',
                    responsePreview: buildResponsePreview(responseText)
                    }
                ));
            }
        });

        monthDebug.status = parseResult.status;
        monthDebug.provider = attemptResult.provider;
        monthDebug.model = attemptResult.model;
        monthDebug.usedFallback = attemptResult.usedFallback;
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
