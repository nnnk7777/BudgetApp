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

    var suggestions = suggestExpenseCategories(uncategorizedItems);
    var updateResult = updateExpenseCategories(suggestions, threshold);

    return JSON.stringify({
        ok: true,
        updated: updateResult.updated,
        skipped: updateResult.skipped,
        total: uncategorizedItems.length
    });
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

function suggestExpenseCategories(items) {
    var apiKey = getGeminiApiKey();
    if (!apiKey) {
        throw new Error('Gemini API key is not set.');
    }

    var categoryNames = getCategoryNames();
    var groupedItems = groupItemsByMonth(items);
    var suggestions = [];

    Object.keys(groupedItems).sort(function (a, b) {
        return parseInt(a, 10) - parseInt(b, 10);
    }).forEach(function (monthKey) {
        var month = parseInt(monthKey, 10);
        var monthItems = groupedItems[monthKey];
        var historyItems = fetchCategoryHistoryForMonth(month);
        var prompt = buildCategorySuggestionPrompt(monthItems, historyItems, categoryNames);
        var responseText = generateGeminiText(apiKey, prompt, {
            temperature: 0.1,
            maxOutputTokens: 1200,
            responseMimeType: "application/json"
        });

        if (!responseText) {
            monthItems.forEach(function (item) {
                suggestions.push(buildSkippedSuggestion(item, null, 'Geminiから応答を取得できませんでした'));
            });
            return;
        }

        var parsedSuggestions = parseCategorySuggestionResponse(responseText, monthItems, categoryNames);
        var suggestionMap = {};
        parsedSuggestions.forEach(function (suggestion) {
            suggestionMap[suggestion.id] = suggestion;
        });

        monthItems.forEach(function (item) {
            if (suggestionMap[item.id]) {
                suggestions.push(suggestionMap[item.id]);
            } else {
                suggestions.push(buildSkippedSuggestion(item, null, 'Gemini応答に対象IDが含まれていません'));
            }
        });
    });

    return suggestions;
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
    return [
        "あなたは家計簿カテゴリ分類アシスタントです。",
        "出力はJSON配列のみで、説明文・コードブロック・Markdownは禁止です。",
        "カテゴリは必ず次の一覧から選んでください。自信が低い場合は suggestedCategory を null にしてください。",
        "カテゴリ一覧: " + JSON.stringify(categoryNames),
        "対象の未分類レコード一覧:",
        JSON.stringify(items),
        "前月1ヶ月分の分類済み履歴一覧:",
        JSON.stringify(historyItems),
        "各要素について次の形式で返してください:",
        '[{"id":"6_42","suggestedCategory":"コンビニ・お菓子 or null","confidence":0.0,"reason":"短い理由"}]',
        "ルール:",
        "- confidence は 0 から 1 の数値",
        "- タイトルや前月履歴だけでは判断が難しい場合は suggestedCategory を null にする",
        "- 同じ店名やサービス名の履歴があれば優先して参考にする",
        "- amount や title から一般常識で推定できても、自信が低ければ null にする"
    ].join("\n");
}

function parseCategorySuggestionResponse(text, items, categoryNames) {
    var cleanedText = text
        .replace(/```json/g, "")
        .replace(/```/g, "")
        .trim();
    var parsed;

    try {
        parsed = JSON.parse(cleanedText);
    } catch (error) {
        Logger.log("カテゴリ推定JSONの解析に失敗: " + cleanedText);
        return items.map(function (item) {
            return buildSkippedSuggestion(item, null, 'Gemini応答をJSONとして解釈できませんでした');
        });
    }

    if (!Array.isArray(parsed)) {
        return items.map(function (item) {
            return buildSkippedSuggestion(item, null, 'Gemini応答が配列形式ではありません');
        });
    }

    var categorySet = {};
    categoryNames.forEach(function (name) {
        categorySet[name] = true;
    });

    var itemMap = {};
    items.forEach(function (item) {
        itemMap[item.id] = item;
    });

    return parsed.map(function (entry) {
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
    });
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

function buildSkippedSuggestion(item, suggestedCategory, reason) {
    return {
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
