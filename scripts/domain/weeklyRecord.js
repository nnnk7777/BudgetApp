class WeeklyRecord {
    dates;
    dataEntries;
    /** @type {Sheet} */
    sheet;

    totalAmount;
    budgetDifference;
    budgetPercentage;
    top5Entries;


    /**
     * @param {Sheet} sheet - スプレッドシートのSheetオブジェクト
     */
    constructor(dateList, sheet) {
        this.dates = dateList;
        this.dataEntries = [];
        this.sheet = sheet;
        this.totalAmount = 0;
    }

    getExpenseEntries() {
        // 一時的なデータ格納用
        let temporaryDataEntries = [];
        // 月ごとの列情報をキャッシュ
        let dateColumnCache = {};
        const sheet = this.sheet;

        // 各日付ごとに処理
        dates.forEach(function (date) {
            let year = date.getFullYear();
            let month = date.getMonth(); // 0始まりの月（0が1月）

            // 日付に対応する列を取得
            let columns;
            if (dateColumnCache[month]) {
                columns = dateColumnCache[month];
            } else {
                columns = sheet.getColumnsForMonth(month);
                dateColumnCache[month] = columns;
            }
            let dateCol = columns.dateCol;

            // その月のデータを取得
            let data = sheet.getExpenseRangeValues(dateCol);
            let currentDate = null;

            // 各行のデータを処理
            for (let i = 0; i < data.length; i++) {
                let row = data[i];

                // 行が空白行かどうかをチェック
                let isEmptyRow = row.every(function (cell) {
                    return cell === null || cell.toString().trim() === '';
                });

                // 空白行が検出されたらループを終了
                if (isEmptyRow) {
                    break;
                }

                let dateCell = row[0];
                let name = row[2];
                let amount = row[3];

                // 日付が空白でない場合、現在の日付を更新
                if (dateCell && dateCell.toString().trim() !== '') {
                    // 日付が文字列の場合、Dateオブジェクトに変換
                    if (typeof dateCell === 'string') {
                        currentDate = parseDate(dateCell, year);
                    } else if (Object.prototype.toString.call(dateCell) === '[object Date]') {
                        currentDate = new Date(dateCell.getFullYear(), dateCell.getMonth(), dateCell.getDate());
                    } else {
                        continue; // 日付の形式が不明な場合はスキップ
                    }
                }

                // 現在の日付が対象の日付と一致する場合、データを収集
                if (currentDate && currentDate.getTime() === date.getTime()) {
                    temporaryDataEntries.push({
                        date: currentDate,
                        name: name,
                        amount: amount
                    });
                }
            }
        });

        this.dataEntries = temporaryDataEntries;
    }

    culculateTotalAmount() {
        let total = 0;
        this.dataEntries.forEach(function (entry) {
            let amount = parseFloat(entry.amount);
            if (!isNaN(amount)) {
                total += amount;
            }
        });

        this.totalAmount = total;
    }

    calculateExpenseUsage(adjustedBudget) {
        this.budgetDifference = this.totalAmount - adjustedBudget;
        this.budgetPercentage = this.totalAmount / adjustedBudget * 100;
    }

    calculateTop5Entries() {
        let sortedEntries = this.dataEntries.sort(function (a, b) {
            return b.amount - a.amount;
        });

        this.top5Entries = sortedEntries.slice(0, 5);
    }

}