function onEdit(e) {
    var sheet = e.source.getActiveSheet();
    var range = e.range;
    var value = e.value;

    // 日付の列と金額の列の設定
    var dateColumns = getTargetColumns(7, 4, 12);  // G列から4列おきに12個分
    var amountColumns = getTargetColumns(10, 4, 12);  // J列から4列おきに12個分

    // 日付の入力がある場合の処理
    if (dateColumns.includes(range.getColumn())) {
        handleDateInput(sheet, range, value);
    }

    // 金額の入力がある場合の処理
    if (amountColumns.includes(range.getColumn())) {
        handleFullWidthNumberInput(sheet, range, value);
    }
}

/**
 * 指定した開始列から特定の間隔で列番号を取得する関数
 * @param {number} startColumn - 開始列番号
 * @param {number} interval - 列の間隔
 * @param {number} count - 列数
 * @return {Array} - 対象の列番号リスト
 */
function getTargetColumns(startColumn, interval, count) {
    var columns = [];
    for (var i = 0; i < count; i++) {
        columns.push(startColumn + i * interval);
    }
    return columns;
}

/**
 * 日付の入力を処理する関数
 * @param {Object} sheet - 対象のシート
 * @param {Object} range - 編集されたセルの範囲
 * @param {string} value - 編集されたセルの値
 */
function handleDateInput(sheet, range, value) {
    if (/^\d{3,4}$/.test(value)) {
        var month, day;
        if (value.length === 4) {
            month = value.substring(0, 2);
            day = value.substring(2, 4);
        } else if (value.length === 3) {
            month = value.substring(0, 1);
            day = value.substring(1, 3);
        }

        var formattedDate = ("0" + month).slice(-2) + '/' + ("0" + day).slice(-2);

        // セルをクリアしてからフォーマットした日付を設定
        sheet.getRange(range.getRow(), range.getColumn()).clearContent();
        sheet.getRange(range.getRow(), range.getColumn()).setValue("'" + formattedDate);
    }
}

/**
 * 全角数字の入力を半角に変換する関数
 * @param {Object} sheet - 対象のシート
 * @param {Object} range - 編集されたセルの範囲
 * @param {string} value - 編集されたセルの値
 */
function handleFullWidthNumberInput(sheet, range, value) {
    if (typeof value === 'string') {
        var convertedValue = value.replace(/[０-９]/g, function (s) {
            return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
        });

        if (convertedValue !== value) {
            sheet.getRange(range.getRow(), range.getColumn()).setValue(convertedValue);
        }
    }
}
