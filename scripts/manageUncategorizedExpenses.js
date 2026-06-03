function listUncategorizedExpenses() {
    return JSON.stringify({
        ok: true,
        items: fetchUncategorizedExpenses()
    });
}

function autofillUncategorizedExpenses(options) {
    options = options || {};

    var uncategorizedItems = fetchUncategorizedExpenses();
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

function fetchUncategorizedExpenses() {
    var rows = getExpenseRowsByMonth();
    return rows.filter(function (entry) {
        return isBlankCell(entry.category) && !isBlankCell(entry.title) && !isBlankCell(entry.amount);
    }).map(function (entry) {
        return {
            id: buildExpenseRecordId(entry.month, entry.row),
            month: entry.month,
            row: entry.row,
            date: entry.date,
            title: String(entry.title).trim(),
            amount: normalizeAmountValue(entry.amount)
        };
    });
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
        var historyItems = fetchCategoryHistoryForMonth(month);
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
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("🐖 家計簿");
    if (!sheet) {
        throw new Error('シート「🐖 家計簿」が見つかりません。');
    }

    var validCategories = getCategoryNames();
    var validCategorySet = {};
    validCategories.forEach(function (name) {
        validCategorySet[name] = true;
    });

    var updated = [];
    var skipped = [];

    suggestions.forEach(function (suggestion) {
        var baseResult = {
            id: suggestion.id,
            month: suggestion.month,
            row: suggestion.row,
            title: suggestion.title,
            amount: suggestion.amount,
            confidence: suggestion.confidence,
            reason: suggestion.reason
        };

        if (!suggestion.suggestedCategory) {
            skipped.push(mergeObjects(baseResult, {
                category: null,
                reason: suggestion.reason || 'カテゴリを推定できませんでした'
            }));
            return;
        }

        if (!validCategorySet[suggestion.suggestedCategory]) {
            skipped.push(mergeObjects(baseResult, {
                category: suggestion.suggestedCategory,
                reason: 'カテゴリ一覧に存在しない候補です'
            }));
            return;
        }

        if (suggestion.confidence < threshold) {
            skipped.push(mergeObjects(baseResult, {
                category: suggestion.suggestedCategory,
                reason: 'confidence below threshold'
            }));
            return;
        }

        var monthColumns = getColumnsForMonth(suggestion.month - 1);
        var rowValues = sheet.getRange(suggestion.row, monthColumns.dateCol, 1, 4).getValues()[0];
        var existingCategory = rowValues[1];
        var existingTitle = rowValues[2];
        var existingAmount = rowValues[3];

        if (!isBlankCell(existingCategory)) {
            skipped.push(mergeObjects(baseResult, {
                category: suggestion.suggestedCategory,
                reason: 'すでにカテゴリが設定されています'
            }));
            return;
        }

        if (String(existingTitle).trim() !== String(suggestion.title).trim()) {
            skipped.push(mergeObjects(baseResult, {
                category: suggestion.suggestedCategory,
                reason: 'タイトルが一致しません'
            }));
            return;
        }

        if (!isSameAmount(existingAmount, suggestion.amount)) {
            skipped.push(mergeObjects(baseResult, {
                category: suggestion.suggestedCategory,
                reason: '金額が一致しません'
            }));
            return;
        }

        sheet.getRange(suggestion.row, monthColumns.dateCol + 1).setValue(suggestion.suggestedCategory);
        updated.push(mergeObjects(baseResult, {
            category: suggestion.suggestedCategory
        }));
    });

    return {
        updated: updated,
        skipped: skipped
    };
}

function fetchCategoryHistoryForMonth(targetMonth) {
    if (targetMonth <= 1) {
        return [];
    }

    var previousMonth = targetMonth - 1;
    var rows = getExpenseRowsForMonth(previousMonth - 1);
    return rows.filter(function (entry) {
        return !isBlankCell(entry.category) && !isBlankCell(entry.title) && !isBlankCell(entry.amount);
    }).map(function (entry) {
        return {
            date: entry.date,
            title: String(entry.title).trim(),
            amount: normalizeAmountValue(entry.amount),
            category: String(entry.category).trim()
        };
    });
}

function buildCategorySuggestionPrompt(items, historyItems, categoryNames) {
    var requiredIds = items.map(function (item) {
        return item.id;
    });

    return [
        "あなたは家計簿カテゴリ分類アシスタントです。",
        "出力は1行1件のプレーンテキストのみで、説明文・JSON・コードブロック・Markdownは禁止です。",
        "カテゴリは必ず次の一覧から選んでください。自信が低い場合は category に null を入れてください。",
        "カテゴリ一覧: " + JSON.stringify(categoryNames),
        "出力必須ID一覧:",
        requiredIds.join(", "),
        "対象の未分類レコード一覧:",
        JSON.stringify(items),
        "前月1ヶ月分の分類済み履歴一覧:",
        JSON.stringify(historyItems),
        "各要素について必ず次の形式で返してください:",
        "id|category|confidence|reason",
        "例:",
        "6_42|コンビニ・お菓子|0.95|ファミマ履歴",
        "6_43|null|0.10|情報不足",
        "ルール:",
        "- 1行1件で返す",
        "- 区切り文字は半角パイプ | のみを使う",
        "- reason は20文字以内",
        "- confidence は 0 から 1 の数値",
        "- タイトルや前月履歴だけでは判断が難しい場合は category を null にする",
        "- 同じ店名やサービス名の履歴があれば優先して参考にする",
        "- amount や title から一般常識で推定できても、自信が低ければ null にする",
        "- reason に | を含めない",
        "- 出力必須ID一覧の全IDを必ず1回ずつ出力する",
        "- 判定不能でも null 行を省略しない",
        "- 対象件数と同じ件数の行を返す",
        "- 出力順は出力必須ID一覧と同じ順にする"
    ].join("\n");
}

function requestCategorySuggestions(apiKey, items, historyItems, categoryNames, month, options) {
    var prompt = buildCategorySuggestionPrompt(items, historyItems, categoryNames);
    var responseText = generateGeminiText(apiKey, prompt, {
        temperature: 0.1,
        maxOutputTokens: 800
    });
    var parseResult;

    if (!responseText) {
        parseResult = {
            status: 'empty_response',
            responsePreview: '',
            extractedJsonPreview: null,
            suggestions: []
        };
    } else {
        parseResult = parseCategorySuggestionResponse(responseText, items, categoryNames, {
            month: month,
            debug: options.debug
        });
    }

    return {
        responseText: responseText,
        parseResult: parseResult
    };
}

function parseCategorySuggestionResponse(text, items, categoryNames, context) {
    context = context || {};
    var cleanedText = cleanGeminiResponseText(text);
    var lineParseResult = parseCategorySuggestionPipeResponse(cleanedText, items, categoryNames);
    if (lineParseResult.ok) {
        return {
            status: 'ok',
            responsePreview: buildResponsePreview(cleanedText),
            extractedJsonPreview: null,
            suggestions: lineParseResult.suggestions
        };
    }

    var parsed = tryParseJson(cleanedText);
    var extractedJsonText = null;

    if (!parsed.ok) {
        extractedJsonText = extractJsonArrayText(cleanedText);
        if (extractedJsonText && extractedJsonText !== cleanedText) {
            parsed = tryParseJson(extractedJsonText);
        }
    }

    if (!parsed.ok) {
        Logger.log(
            "カテゴリ推定応答の解析に失敗: month=" +
                (context.month || "unknown") +
                " pipeError=" +
                lineParseResult.errorDetail +
                " jsonError=" +
                parsed.errorMessage +
                " response=" +
                cleanedText
        );
        return {
            status: 'response_parse_failed',
            parseError: lineParseResult.errorDetail + " / " + parsed.errorMessage,
            responsePreview: buildResponsePreview(cleanedText),
            extractedJsonPreview: extractedJsonText ? buildResponsePreview(extractedJsonText) : null,
            suggestions: items.map(function (item) {
                return buildSkippedSuggestion(
                    item,
                    null,
                    buildDetailedReason('Gemini応答を解析できませんでした', lineParseResult.errorDetail || parsed.errorMessage),
                    {
                    errorCode: 'response_parse_failed',
                    errorDetail: (lineParseResult.errorDetail || 'Pipe parse failed.') + ' / ' + parsed.errorMessage,
                    responsePreview: buildResponsePreview(cleanedText)
                    }
                );
            })
        };
    }

    if (!Array.isArray(parsed.value)) {
        Logger.log(
            "カテゴリ推定JSONが配列形式ではありません: month=" +
                (context.month || "unknown") +
                " response=" +
                cleanedText
        );
        return {
            status: 'response_not_array',
            responsePreview: buildResponsePreview(cleanedText),
            extractedJsonPreview: extractedJsonText ? buildResponsePreview(extractedJsonText) : null,
            suggestions: items.map(function (item) {
                return buildSkippedSuggestion(
                    item,
                    null,
                    buildDetailedReason('Gemini応答が配列形式ではありません', 'Gemini response was valid JSON but not an array.'),
                    {
                    errorCode: 'response_not_array',
                    errorDetail: 'Gemini response was valid JSON but not an array.',
                    responsePreview: buildResponsePreview(cleanedText)
                    }
                );
            })
        };
    }

    var categorySet = {};
    categoryNames.forEach(function (name) {
        categorySet[name] = true;
    });

    var itemMap = {};
    items.forEach(function (item) {
        itemMap[item.id] = item;
    });

    return {
        status: 'ok',
        responsePreview: buildResponsePreview(cleanedText),
        extractedJsonPreview: extractedJsonText ? buildResponsePreview(extractedJsonText) : null,
        suggestions: parsed.value.map(function (entry) {
            var item = itemMap[entry.id];
            if (!item) {
                return null;
            }

            var suggestedCategory = entry.suggestedCategory;
            if (suggestedCategory !== null && suggestedCategory !== undefined) {
                suggestedCategory = String(suggestedCategory).trim();
                if (!categorySet[suggestedCategory]) {
                    suggestedCategory = null;
                }
            } else {
                suggestedCategory = null;
            }

            return {
                id: item.id,
                month: item.month,
                row: item.row,
                date: item.date,
                title: item.title,
                amount: item.amount,
                suggestedCategory: suggestedCategory,
                confidence: normalizeConfidence(entry.confidence),
                reason: entry.reason ? String(entry.reason).trim() : ''
            };
        }).filter(function (entry) {
            return entry !== null;
        })
    };
}

function getExpenseRowsByMonth() {
    var allRows = [];
    for (var monthIndex = 0; monthIndex < 12; monthIndex++) {
        allRows = allRows.concat(getExpenseRowsForMonth(monthIndex));
    }
    return allRows;
}

function getExpenseRowsForMonth(monthIndex) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("🐖 家計簿");
    if (!sheet) {
        throw new Error('シート「🐖 家計簿」が見つかりません。');
    }

    var columns = getColumnsForMonth(monthIndex);
    var startRow = 35;
    var endRow = 184;
    var data = sheet.getRange(startRow, columns.dateCol, endRow - startRow + 1, 4).getValues();
    var currentDateKey = null;
    var rows = [];

    for (var i = 0; i < data.length; i++) {
        var row = data[i];
        var dateCell = row[0];
        var category = row[1];
        var title = row[2];
        var amount = row[3];

        if (isFixedCostMarker(dateCell)) {
            break;
        }

        var hasContent = [dateCell, category, title, amount].some(function (cell) {
            return !isBlankCell(cell);
        });
        if (!hasContent) {
            continue;
        }

        if (!isBlankCell(dateCell)) {
            currentDateKey = toMonthDayKey(dateCell);
        }

        var hasEntryContent = !isBlankCell(title) || !isBlankCell(amount);
        if (!hasEntryContent) {
            continue;
        }

        rows.push({
            id: buildExpenseRecordId(monthIndex + 1, startRow + i),
            month: monthIndex + 1,
            row: startRow + i,
            date: currentDateKey,
            category: category,
            title: title,
            amount: amount
        });
    }

    return rows;
}

function getCategoryNames() {
    return categories.map(function (category) {
        return category.name;
    });
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

function buildSkippedSuggestion(item, suggestedCategory, reason, diagnostics) {
    var result = {
        id: item.id,
        month: item.month,
        row: item.row,
        date: item.date,
        title: item.title,
        amount: item.amount,
        suggestedCategory: suggestedCategory,
        confidence: 0,
        reason: reason
    };

    diagnostics = diagnostics || {};
    if (diagnostics.errorCode) {
        result.errorCode = diagnostics.errorCode;
    }
    if (diagnostics.errorDetail) {
        result.errorDetail = diagnostics.errorDetail;
    }
    if (diagnostics.responsePreview !== undefined) {
        result.responsePreview = diagnostics.responsePreview;
    }

    return result;
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

function tryParseJson(text) {
    try {
        return {
            ok: true,
            value: JSON.parse(text)
        };
    } catch (error) {
        return {
            ok: false,
            errorMessage: error && error.message ? String(error.message) : String(error)
        };
    }
}

function parseCategorySuggestionPipeResponse(text, items, categoryNames) {
    var itemMap = {};
    items.forEach(function (item) {
        itemMap[item.id] = item;
    });

    var categorySet = {};
    categoryNames.forEach(function (name) {
        categorySet[name] = true;
    });

    var lines = text
        .split("\n")
        .map(function (line) {
            return line.trim();
        })
        .filter(function (line) {
            return line !== '';
        })
        .filter(function (line) {
            return /^[0-9]+_[0-9]+\|/.test(line);
        });

    if (!lines.length) {
        return {
            ok: false,
            errorDetail: 'Pipe形式の応答行を検出できませんでした'
        };
    }

    var suggestions = [];
    for (var i = 0; i < lines.length; i++) {
        var parts = lines[i].split("|");
        if (parts.length < 4) {
            return {
                ok: false,
                errorDetail: 'Pipe形式の列数が不足しています: ' + lines[i]
            };
        }

        var id = parts[0].trim();
        var category = parts[1].trim();
        var confidence = parts[2].trim();
        var reason = parts.slice(3).join("|").trim();
        var item = itemMap[id];

        if (!item) {
            continue;
        }

        if (category.toLowerCase() === 'null' || category === '') {
            category = null;
        } else if (!categorySet[category]) {
            category = null;
        }

        suggestions.push({
            id: item.id,
            month: item.month,
            row: item.row,
            date: item.date,
            title: item.title,
            amount: item.amount,
            suggestedCategory: category,
            confidence: normalizeConfidence(confidence),
            reason: reason
        });
    }

    return {
        ok: true,
        suggestions: suggestions
    };
}

function extractJsonArrayText(text) {
    var startIndex = text.indexOf('[');
    var endIndex = text.lastIndexOf(']');
    if (startIndex === -1 || endIndex === -1 || endIndex < startIndex) {
        return null;
    }
    return text.substring(startIndex, endIndex + 1).trim();
}

function cleanGeminiResponseText(text) {
    return String(text)
        .replace(/```[\s\S]*?\n/g, "")
        .replace(/```/g, "")
        .trim();
}

function buildResponsePreview(text) {
    if (!text) {
        return '';
    }

    var normalized = String(text)
        .replace(/\s+/g, ' ')
        .trim();

    if (normalized.length <= 300) {
        return normalized;
    }

    return normalized.substring(0, 300) + '...';
}

function buildDetailedReason(baseReason, detail) {
    if (!detail) {
        return baseReason;
    }

    var normalizedDetail = String(detail).replace(/\s+/g, ' ').trim();
    if (!normalizedDetail) {
        return baseReason;
    }

    var maxDetailLength = 120;
    if (normalizedDetail.length > maxDetailLength) {
        normalizedDetail = normalizedDetail.substring(0, maxDetailLength) + '...';
    }

    return baseReason + '（' + normalizedDetail + '）';
}
