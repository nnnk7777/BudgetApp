function calculateWeeklyExpenses() {
    // 共通設定
    var budget = 45000; // 週ごとの予算
    var currentDate = new Date();

    // 現在の日付に基づいて、計算対象の列を求める
    var columns = getCurrentMonthColumns(currentDate);
    var dateCol = columns.dateCol;
    var amountCol = columns.amountCol;

    // その週の日付一覧を取得
    var datesInWeek = getDatesInWeek(currentDate);
    var startOfWeek = datesInWeek[0]; // 週の開始日（月曜日）
    var endOfWeek = datesInWeek[6];   // 週の終了日（日曜日）

    // 日付範囲の文字列を作成
    var dateRangeStr = formatDate(startOfWeek) + "〜" + formatDate(endOfWeek);

    // その週に含まれる日付内データを一覧で取得
    var dataEntries = getDataForDates(datesInWeek, dateCol, amountCol);

    // 合計金額を算出
    var totalAmount = calculateTotalAmount(dataEntries);

    // 予算との差分を計算
    var difference = totalAmount - budget;
    var percentage = (totalAmount / budget) * 100;

    // デバッグ用出力
    Logger.log("データエントリ一覧:");
    dataEntries.forEach(function (entry) {
        Logger.log("日付: " + formatDate(entry.date) + ", 名称: " + entry.name + ", 金額: " + entry.amount);
    });
    Logger.log(dateRangeStr + " の合計金額: " + totalAmount + "円");
    if (difference > 0) {
        Logger.log("予算を " + difference + " 円上回りました。");
    } else {
        Logger.log("予算を " + Math.abs(difference) + " 円下回りました。");
    }
    Logger.log("予算の " + percentage.toFixed(2) + "% を使用しました。");

    // 現在の曜日を取得（0:日曜日, 1:月曜日, ..., 6:土曜日）
    var dayOfWeek = currentDate.getDay();

    if (dayOfWeek === 0) {
        // 日曜日の場合、週次サマリーを送信
        sendWeeklySummaryEmail(dateRangeStr, totalAmount, dataEntries, difference, percentage);
    } else {
        // 日曜日以外の場合、週の開始から現在までのデータを取得し、メールで送信
        sendDailyProgressEmail(currentDate, budget, datesInWeek, dateCol, amountCol);
    }
}

// 日付を "MM/DD" の形式にフォーマットする関数
function formatDate(date) {
    var month = date.getMonth() + 1;
    var day = date.getDate();
    return month + "/" + day;
}

// 現在の日付に基づいて、計算対象の列を求めるメソッド
function getCurrentMonthColumns(currentDate) {
    var month = currentDate.getMonth(); // 0（1月）から11（12月）
    // G列が7番目の列で、各月ごとに4列ずつデータがある
    var dateCol = 7 + month * 4; // 1月は7列目
    var amountCol = dateCol + 3; // 金額列は日付列から3列後
    return { dateCol: dateCol, amountCol: amountCol };
}

// その週に含まれる日付一覧を求めるメソッド
function getDatesInWeek(date) {
    var dates = [];
    // 週の始まり（月曜日）を取得
    var day = date.getDay(); // 0（日曜）から6（土曜）
    var diff = date.getDate() - day + (day === 0 ? -6 : 1); // 日曜の場合は-6
    var monday = new Date(date);
    monday.setDate(diff);
    // 月曜日から日曜日までの日付を取得
    for (var i = 0; i < 7; i++) {
        var d = new Date(monday);
        d.setDate(monday.getDate() + i);
        dates.push(new Date(d.getFullYear(), d.getMonth(), d.getDate())); // 時間を00:00:00に設定
    }
    return dates;
}

// その週に含まれる日付内データを一覧で取得するメソッド
function getDataForDates(dates, dateCol, amountCol) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet(); // 必要に応じてシート名を指定
    var startRow = 35; // データが開始する行

    var lastRow = sheet.getLastRow();
    var numRows = lastRow - startRow + 1;

    // 日付、カテゴリ、名称、金額のデータを取得
    var dataRange = sheet.getRange(startRow, dateCol, numRows, 4);
    var data = dataRange.getValues();

    var dataEntries = [];
    var currentDate = null;

    // 各行のデータを処理
    for (var i = 0; i < data.length; i++) {
        var row = data[i];

        // 行が空白行かどうかをチェック
        var isEmptyRow = row.every(function (cell) {
            return cell === null || cell.toString().trim() === '';
        });

        // 空白行が検出されたらループを終了
        if (isEmptyRow) {
            break;
        }

        var dateCell = row[0];
        var category = row[1];
        var name = row[2];
        var amount = row[3];

        // 日付が空白でない場合、現在の日付を更新
        if (dateCell && dateCell.toString().trim() !== '') {
            // 日付が文字列の場合、Dateオブジェクトに変換
            if (typeof dateCell === 'string') {
                currentDate = parseDate(dateCell);
            } else if (Object.prototype.toString.call(dateCell) === '[object Date]') {
                currentDate = new Date(dateCell.getFullYear(), dateCell.getMonth(), dateCell.getDate());
            } else {
                continue; // 日付の形式が不明な場合はスキップ
            }
        }

        // 現在の日付が対象週の日付に含まれている場合、データを収集
        if (currentDate && isDateInDates(currentDate, dates)) {
            dataEntries.push({
                date: currentDate,
                name: name,
                amount: amount
            });
        }
    }
    return dataEntries;
}

// 日付文字列をDateオブジェクトに変換する関数
function parseDate(dateStr) {
    var dateParts = dateStr.split('/');
    var month = parseInt(dateParts[0], 10) - 1; // 月は0始まり
    var day = parseInt(dateParts[1], 10);
    var year = new Date().getFullYear(); // 必要に応じて年を調整
    return new Date(year, month, day);
}

// 日付が日付リストに含まれているか確認する関数
function isDateInDates(date, dates) {
    for (var i = 0; i < dates.length; i++) {
        if (date.getTime() === dates[i].getTime()) {
            return true;
        }
    }
    return false;
}

// 求めたデータ一覧の金額合計を算出するメソッド
function calculateTotalAmount(dataEntries) {
    var total = 0;
    dataEntries.forEach(function (entry) {
        var amount = parseFloat(entry.amount);
        if (!isNaN(amount)) {
            total += amount;
        }
    });
    return total;
}

// 週次サマリーをメールで送信するメソッド（毎週日曜日）
function sendWeeklySummaryEmail(dateRangeStr, totalAmount, dataEntries, difference, percentage) {
    var emailAddress = "electro0701+budgetapp@gmail.com"; // あなたのメールアドレスに置き換えてください

    // 予算の設定
    var budget = 45000; // 週ごとの予算

    // 予算差分の符号を設定
    var differenceSign = difference >= 0 ? "+" : "-";
    var differenceAbs = Math.abs(difference);

    // 予算割合を小数点以下2桁で表示
    var percentageStr = percentage.toFixed(2);

    // トップ5の支出を計算
    var top5Entries = dataEntries.slice(); // 配列をコピー
    top5Entries.sort(function (a, b) {
        return parseFloat(b.amount) - parseFloat(a.amount);
    });
    top5Entries = top5Entries.slice(0, 5);

    // メール本文を指定のフォーマットで作成
    var body = "";
    body += "◆ " + dateRangeStr + " の週次サマリー\n\n";
    body += "合計支出は " + totalAmount + " 円です。\n";
    body += "予算差分：" + differenceSign + differenceAbs + "円\n";
    body += "予算割合：" + percentageStr + "%\n";
    body += "支出TOP5\n\n";

    top5Entries.forEach(function (entry) {
        body += "・" + formatDate(entry.date) + " - " + entry.name + ": " + entry.amount + "円\n";
    });

    // メールを送信
    MailApp.sendEmail(emailAddress, "週次サマリー（" + dateRangeStr + "）", body);
}

// 日曜日以外に日次進捗をメールで送信するメソッド
function sendDailyProgressEmail(currentDate, budget, datesInWeek, dateCol, amountCol) {
    var emailAddress = "electro0701+budgetapp@gmail.com"; // あなたのメールアドレスに置き換えてください

    // 週の開始日から現在の日付までのデータを取得
    var datesUpToToday = datesInWeek.filter(function (date) {
        return date <= currentDate;
    });

    var dataEntries = getDataForDates(datesUpToToday, dateCol, amountCol);

    // 合計金額を算出
    var totalAmount = calculateTotalAmount(dataEntries);

    // 予算に対する割合を計算
    var percentage = (totalAmount / budget) * 100;

    // メールの件名と本文を作成
    var subject = "日次進捗（" + formatDate(datesInWeek[0]) + "〜" + formatDate(currentDate) + "）";
    var body = formatDate(datesInWeek[0]) + " から " + formatDate(currentDate) + " までの合計支出は " + totalAmount + " 円です。\n";
    body += "予算の " + percentage.toFixed(2) + "% を使用しました。\n\n";

    body += "詳細:\n";
    dataEntries.forEach(function (entry) {
        body += formatDate(entry.date) + " - " + entry.name + ": " + entry.amount + "円\n";
    });

    MailApp.sendEmail(emailAddress, subject, body);
}
