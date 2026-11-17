function getMonthlyExpenseEntries(year, month) {
    var sheet = getExpenseSheet();
    var startRow = 35;
    var endRow = 185;
    var columns = getColumnsForMonth(month);
    var dataRange = sheet.getRange(startRow, columns.dateCol, endRow - startRow + 1, 4);
    var data = dataRange.getValues();
    var entries = [];
    var currentDate = null;

    data.forEach(function (row, index) {
        var dateCell = row[0];
        var category = row[1];
        var name = row[2];
        var amount = row[3];
        var absoluteRow = startRow + index;

        var hasContent = [dateCell, category, name, amount].some(function (cell) {
            return cell !== null && cell.toString().trim() !== '';
        });
        if (!hasContent) {
            return;
        }

        if (dateCell && dateCell.toString().trim() !== '') {
            if (typeof dateCell === 'string') {
                currentDate = parseDate(dateCell, year);
            } else if (Object.prototype.toString.call(dateCell) === '[object Date]') {
                currentDate = new Date(year, dateCell.getMonth(), dateCell.getDate());
            }
        }

        var includeEntry = false;
        if (currentDate && currentDate.getFullYear() === year && currentDate.getMonth() === month) {
            includeEntry = true;
        } else if (!currentDate && absoluteRow >= 156) {
            currentDate = new Date(year, month, 1);
            includeEntry = true;
        }

        if (includeEntry) {
            entries.push({
                date: currentDate || new Date(year, month, 1),
                category: category || "未分類",
                name: name || "",
                amount: amount || 0
            });
        }
    });

    return entries;
}

function getMonthlyIncomeRange(month) {
    // IncomeData creates detail rows 20 through 33.
    var startRow = 20;
    var endRow = 33;
    var columns = getColumnsForMonth(month);
    return getExpenseSheet().getRange(startRow, columns.dateCol, endRow - startRow + 1, 4);
}

function appendMonthlyIncomeEntry(month, date, category, title, amount) {
    var lock = LockService.getScriptLock();
    if (!lock.tryLock(5000)) {
        throw new Error('収入欄を使用中です。しばらくしてから再試行してください');
    }

    try {
        var range = getMonthlyIncomeRange(month);
        var values = range.getValues();
        var formulas = range.getFormulas();
        var rowIndex = values.findIndex(function (row, index) {
            return row.every(function (cell, column) {
                return cell === '' && formulas[index][column] === '';
            });
        });
        if (rowIndex === -1) {
            throw new Error('この月の収入欄に空きがありません');
        }

        var row = range.offset(rowIndex, 0, 1, 4);
        var rule = row.getCell(1, 2).getDataValidation();
        if (!rule || rule.getCriteriaType() !== SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
            throw new Error('収入欄のカテゴリ設定が見つかりません');
        }
        if (!rule.getCriteriaValues()[0].includes(category)) {
            throw new Error('収入カテゴリが不正です');
        }

        // 文字列はシート上でも文字列として保存し、数式として解釈されないようにする。
        row.setValues([["'" + date, "'" + category, "'" + title, amount]]);
        SpreadsheetApp.flush();
    } finally {
        lock.releaseLock();
    }
}

function getMonthlyIncomeEntries(year, month) {
    // 「予定」などの日付ラベルも、その月の収入として読み取る。
    var data = getMonthlyIncomeRange(month).getValues();
    var entries = [];

    data.forEach(function (row) {
        var dateCell = row[0];
        var category = row[1];
        var name = row[2];
        var amount = row[3];

        var hasEntryContent = [name, amount].some(function (cell) {
            return cell !== null && cell.toString().trim() !== '';
        });
        if (!hasEntryContent) {
            return;
        }

        var entryDate = new Date(year, month, 1);
        var dateLabel = "";
        var hasActualDate = false;
        if (dateCell && dateCell.toString().trim() !== '') {
            if (typeof dateCell === 'string') {
                var parsedDate = parseDate(dateCell, year);
                if (!isNaN(parsedDate.getTime())) {
                    entryDate = parsedDate;
                    hasActualDate = true;
                } else {
                    dateLabel = dateCell.toString().trim();
                }
            } else if (Object.prototype.toString.call(dateCell) === '[object Date]') {
                entryDate = new Date(year, dateCell.getMonth(), dateCell.getDate());
                hasActualDate = true;
            }
        }

        entries.push({
            date: entryDate,
            dateLabel: dateLabel,
            hasActualDate: hasActualDate,
            category: category || "未分類",
            name: name || "",
            amount: amount || 0
        });
    });

    return entries;
}

function getIncomeEntriesForDates(dates) {
    var targetDateKeys = {};
    var targetMonths = {};
    var entries = [];

    dates.forEach(function (date) {
        var monthKey = date.getFullYear() + "-" + date.getMonth();

        targetDateKeys[getIncomeEntryDateKey(date)] = true;
        targetMonths[monthKey] = {
            year: date.getFullYear(),
            month: date.getMonth()
        };
    });

    Object.keys(targetMonths).forEach(function (monthKey) {
        var targetMonth = targetMonths[monthKey];
        entries = entries.concat(getMonthlyIncomeEntries(targetMonth.year, targetMonth.month));
    });

    return entries.filter(function (entry) {
        return entry.hasActualDate && targetDateKeys[getIncomeEntryDateKey(entry.date)];
    });
}

function getSaleIncomeEntriesForDates(dates) {
    return getIncomeEntriesForDates(dates).filter(function (entry) {
        return String(entry.category || "").trim() === "売却";
    });
}

function getIncomeEntryDateKey(date) {
    return date.getFullYear() + "-" + date.getMonth() + "-" + date.getDate();
}
