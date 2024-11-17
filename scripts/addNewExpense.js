function addExpenseRecord(title, amount) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("🐖 家計簿");
    if (!sheet) {
        throw new Error('シート「🐖 家計簿」が見つかりません。');
    }

    var currentDate;
    var testDateStr = "TEST_DATE_PLACEHOLDER"
    var isStaging = testDateStr ? true : false

    if (isStaging) {
        // テスト用の日付が指定されている場合はその日付を使用
        // YYYYMMDD フォーマットをパースして Date オブジェクトを作成
        currentDate = parseYYYYMMDD(testDateStr);
        if (!currentDate) {
            throw new Error('Invalid TEST_DATE format. Expected YYYYMMDD.');
        }
    } else {
        // 指定がない場合は現在の日付を使用
        currentDate = new Date();
    }
    const year = currentDate.getFullYear();
    const month = currentDate.getMonth(); // 0が1月
    const day = currentDate.getDate();

    // 各月の開始列を計算（1月はG列=7列目）
    const startColumn = 7 + month * 4;
    const dateColumn = startColumn;          // 日付列
    const categoryColumn = startColumn + 1;  // カテゴリ列
    const titleColumn = startColumn + 2;     // 名称列
    const amountColumn = startColumn + 3;    // 金額列

    const startRow = 35;    // 支出記録の開始行
    const maxExpenseRow = 149; // 支出記録の最大行（固定費用の手前）

    const category = ""; // カテゴリーはから文字列で登録する

    // 支出記録の最後の行を特定
    let lastRow = startRow - 1;
    for (let row = startRow; row <= maxExpenseRow; row++) {
        const rowRange = sheet.getRange(row, dateColumn, 1, 4);
        const rowValues = rowRange.getValues()[0];

        const isEmptyRow = rowValues.every(function (cell) {
            return cell === "" || cell === null || cell.toString().trim() === "";
        });

        if (isEmptyRow) {
            break; // 空白行を見つけたらループを終了
        } else {
            lastRow = row;
        }
    }

    // 最後の支出記録の日付を取得
    let lastEntryDate = null;
    if (lastRow >= startRow) {
        const numRows = lastRow - startRow + 1;
        const dataRange = sheet.getRange(startRow, dateColumn, numRows, 4);
        const dataValues = dataRange.getValues();

        let currentDate = null;
        dataValues.forEach(function (row) {
            const dateCell = row[0];
            if (dateCell && dateCell.toString().trim() !== "") {
                // 日付セルがある場合、現在の日付を更新
                currentDate = parseDateCell(dateCell, year);
            }
            lastEntryDate = currentDate;
        });
    }

    // 日付を記入するか判定
    let includeDate = true;
    if (lastEntryDate) {
        if (lastEntryDate.getFullYear() === currentDate.getFullYear() &&
            lastEntryDate.getMonth() === currentDate.getMonth() &&
            lastEntryDate.getDate() === currentDate.getDate()) {
            includeDate = false; // 日付が同じ場合、日付セルは空白
        }
    }

    const newEntryRow = lastRow + 1;
    if (newEntryRow >= maxExpenseRow + 1) {
        throw new Error('新しい支出記録を追加する位置が固定費用の行と重なっています。');
    }

    // 日付の文字列を作成
    const dateString = Utilities.formatDate(currentDate, "Asia/Tokyo", "MM/dd");
    const dateValue = includeDate ? ("'" + dateString) : "";

    // 新しい支出記録を作成
    const newRowValues = [
        dateValue,
        category,
        title,
        amount
    ];

    // シートに書き込み
    const writeRange = sheet.getRange(newEntryRow, dateColumn, 1, 4);
    writeRange.setValues([newRowValues]);

    return "Successfully registered new expense"
}

// 日付セルをDateオブジェクトに変換する関数
function parseDateCell(dateCell, year) {
    if (typeof dateCell === 'string') {
        const dateParts = dateCell.split('/');
        const monthPart = parseInt(dateParts[0], 10) - 1; // 月は0始まり
        const dayPart = parseInt(dateParts[1], 10);
        return new Date(year, monthPart, dayPart);
    } else if (Object.prototype.toString.call(dateCell) === '[object Date]') {
        return new Date(dateCell.getFullYear(), dateCell.getMonth(), dateCell.getDate());
    } else {
        throw new Error('無効な日付形式です。');
    }
}

// YYYYMMDD フォーマットの日付文字列を Date オブジェクトに変換する関数
function parseYYYYMMDD(dateStr) {
    if (!/^\d{8}$/.test(dateStr)) {
        return null;
    }
    var year = parseInt(dateStr.substring(0, 4), 10);
    var month = parseInt(dateStr.substring(4, 6), 10) - 1; // 月は0始まり
    var day = parseInt(dateStr.substring(6, 8), 10);
    return new Date(year, month, day);
}
