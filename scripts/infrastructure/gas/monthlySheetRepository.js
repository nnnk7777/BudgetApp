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

function getMonthlyIncomeEntries(year, month) {
    var sheet = getExpenseSheet();
    // IncomeData creates detail rows 20 through 33. The month is determined by
    // this block's columns, so an entry must not be dropped just because its
    // date cell contains a label such as "予定".
    var startRow = 20;
    var endRow = 33;
    var columns = getColumnsForMonth(month);
    var dataRange = sheet.getRange(startRow, columns.dateCol, endRow - startRow + 1, 4);
    var data = dataRange.getValues();
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
