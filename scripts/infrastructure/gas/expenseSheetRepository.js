function getExpenseEntriesForDates(dates) {
    var sheet = getExpenseSheet();
    var startRow = 35;
    var endRow = 184;
    var dataEntries = [];
    var targetDateKeys = dates.map(function (date) {
        return toMonthDayKey(date);
    });
    var targetDateKeySet = {};

    targetDateKeys.forEach(function (key) {
        targetDateKeySet[key] = true;
    });

    Logger.log("集計対象日: " + targetDateKeys.join(", "));

    var processedMonths = {};

    dates.forEach(function (date) {
        var year = date.getFullYear();
        var month = date.getMonth();
        if (processedMonths[month]) {
            return;
        }
        processedMonths[month] = true;

        var columns = getExpenseSheetColumnsForMonth(month);
        var dataRange = sheet.getRange(startRow, columns.dateCol, endRow - startRow + 1, 4);
        var data = dataRange.getValues();
        var currentDateKey = null;
        var matchedCount = 0;

        for (var i = 0; i < data.length; i++) {
            var row = data[i];
            var dateCell = row[0];
            var category = row[1];
            var name = row[2];
            var amount = row[3];

            if (isExpenseFixedCostMarker(dateCell)) {
                Logger.log(
                    "固定費マーカーを検出したため読取終了: month=" +
                        (month + 1) +
                        " row=" +
                        (startRow + i)
                );
                break;
            }

            var hasContent = [dateCell, category, name, amount].some(function (cell) {
                return cell !== null && cell.toString().trim() !== '';
            });
            if (!hasContent) {
                continue;
            }

            if (dateCell && dateCell.toString().trim() !== '') {
                currentDateKey = toMonthDayKey(dateCell);
            }

            if (currentDateKey && targetDateKeySet[currentDateKey]) {
                var hasEntryContent =
                    (name !== null && String(name).trim() !== '') ||
                    (amount !== null && String(amount).trim() !== '');
                if (!hasEntryContent) {
                    continue;
                }

                dataEntries.push({
                    date: monthDayKeyToDate(currentDateKey, year),
                    category: category,
                    name: name,
                    amount: amount
                });
                matchedCount++;
            }
        }

        Logger.log(
            "月別読取結果: month=" +
                (month + 1) +
                " cols=" +
                columns.dateCol +
                "-" +
                (columns.dateCol + 3) +
                " matched=" +
                matchedCount
        );
    });

    return dataEntries;
}

function getExpenseSheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("🐖 家計簿");
    if (!sheet) {
        throw new Error('シート「🐖 家計簿」が見つかりません。');
    }
    return sheet;
}

function getExpenseSheetColumnsForMonth(month) {
    var dateCol = 7 + month * 4;
    var amountCol = dateCol + 3;
    return { dateCol: dateCol, amountCol: amountCol };
}

function isExpenseFixedCostMarker(value) {
    if (value === null || value === undefined) {
        return false;
    }

    return String(value).trim() === '固定';
}

function getDataForDates(dates) {
    return getExpenseEntriesForDates(dates);
}

function getColumnsForMonth(month) {
    return getExpenseSheetColumnsForMonth(month);
}

function isFixedCostMarker(value) {
    return isExpenseFixedCostMarker(value);
}
