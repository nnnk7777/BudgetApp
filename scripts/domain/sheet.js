class Sheet {
    name;
    /** @type {GoogleAppsScript.Spreadsheet.Sheet} */
    spreadsheetData;
    startRow;
    maxRow;
    rowCount;
    expenseRange;
    expenseRangeValues;

    constructor(name) {
        if (!name || name === '') {
            throw new Error('Invalid name');
        }

        this.name = name;
        this.spreadsheetData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
        if (!this.spreadsheetData) {
            throw new Error(`Sheet "${name}" not found.`);
        }

        this.startRow = 35;
        this.maxRow = this.spreadsheetData.getLastRow();
        this.rowCount = rthis.maxRow - this.startRow + 1;
    }


    /**
     * 月に対応する列情報を取得するメソッド
     * @param {Number} month 
     */
    getColumnsForMonth(month) {
        // G列が7番目の列で、各月ごとに4列ずつデータがある
        // 0（1月）から始まる月を想定
        var dateCol = 7 + month * 4; // 1月は7列目
        var amountCol = dateCol + 3; // 金額列は日付列から3列後
        return { dateCol: dateCol, amountCol: amountCol };
    }

    getExpenseRangeValues(dateCol) {
        this.expenseRange = this.spreadsheetData.getRange(this.startRow, dateCol, this.rowCount, 4);
        this.expenseRangeValues = this.expenseRange.getValues();
    }

}
