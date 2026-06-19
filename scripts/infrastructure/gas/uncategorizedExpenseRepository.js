function fetchUncategorizedExpenseItems() {
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

function updateUncategorizedExpenseCategories(suggestions, threshold) {
    var sheet = getExpenseSheet();
    var validCategories = getCategoryNames();
    var validCategorySet = {};
    var updated = [];
    var skipped = [];

    validCategories.forEach(function (name) {
        validCategorySet[name] = true;
    });

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

function fetchCategoryHistoryEntriesForMonth(targetMonth) {
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

function getExpenseRowsByMonth() {
    var allRows = [];
    for (var monthIndex = 0; monthIndex < 12; monthIndex++) {
        allRows = allRows.concat(getExpenseRowsForMonth(monthIndex));
    }
    return allRows;
}

function getExpenseRowsForMonth(monthIndex) {
    var sheet = getExpenseSheet();
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
