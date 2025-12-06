class Message {
    calendar;
    subject;
    body;
    stagingPrefix;;

    constructor(env) {
        this.calendar = new Calendar();
        this.stagingPrefix = env == Env.STG ? "<test>" : "";
    }

    buildDaiilySummaryMessage(
        currentDate,
        dateList,
        totalAmount,
        percentage,
        weeklyBudget
    ) {
        let temporaryBody = "";

        this.subject = this.stagingPrefix
            + "家計簿日次レポート（" + this.calendar.convertDateToStringWithSlash(currentDate) + "）";
        temporaryBody = this.calendar.convertDateToStringWithSlash(dateList[0]) + " から " + this.calendar.convertDateToStringWithSlash(currentDate) + " までの合計支出は " + totalAmount + " 円です。\n";
        temporaryBody += "予算の " + percentage.toFixed(2) + "% を使用しました。\n";
        temporaryBody += "（設定予算：" + weeklyBudget + "円）\n\n";

        temporaryBody += "詳細:\n";
        dataEntries.forEach(function (entry) {
            temporaryBody += "・" + this.calendar.convertDateToStringWithSlash(entry.date) + " - " + entry.name + ": " + entry.amount + "円\n";
        });

        this.body = temporaryBody;
    }

    buildWeeklySummaryMessage(
        dataEntries,
        dateList,
        totalAmount,
        percentage,
        weeklyBudget,
        top5Entries
    ) {
        let temporaryBody = "";

        temporaryBody += "◆ " + dateList + " の週次サマリー\n\n";
        temporaryBody += "合計支出は " + totalAmount + " 円です。\n\n";
        temporaryBody += "* 設定予算： " + weeklyBudget + " 円\n";
        temporaryBody += "* 予算差分：" + differenceSign + differenceAbs + "円\n";
        temporaryBody += "* 予算割合：" + percentage.toFixed(2) + "%\n\n";
        temporaryBody += "◆ 支出TOP5\n";

        top5Entries.forEach(function (entry) {
            temporaryBody += "・" + formatDate(entry.date) + " - " + entry.name + ": " + entry.amount + "円\n";
        });
        temporaryBody += "\n";
        temporaryBody += "◆ 支出一覧\n";

        dataEntries.forEach(function (entry) {
            temporaryBody += "・" + formatDate(entry.date) + " - " + entry.name + ": " + entry.amount + "円\n";
        });
    }

}