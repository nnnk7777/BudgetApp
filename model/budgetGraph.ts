export class BudgetGraph {
    private sheet: GoogleAppsScript.Spreadsheet.Sheet;
    private initialColumnNumber: number;
    private roughEstimateSummaryList: string[];
    private incomeMonthlySummaryList: string[];
    private outcomeMonthlySummaryList: string[];
    private monthList: string[];

    constructor(
        sheet: GoogleAppsScript.Spreadsheet.Sheet,
        initialColumnNumber: number,
        roughEstimateSummaryList: string[],
        incomeMonthlySummaryList: string[],
        outcomeMonthlySummaryList: string[],
        monthList: string[]
    ) {
        this.sheet = sheet;
        this.initialColumnNumber = initialColumnNumber;

        // 月毎の概算、収入と支出のデータ範囲を設定
        this.roughEstimateSummaryList = roughEstimateSummaryList; // 1〜12月の概算データ
        this.incomeMonthlySummaryList = incomeMonthlySummaryList; // 1〜12月の収入データ
        this.outcomeMonthlySummaryList = outcomeMonthlySummaryList; // 1〜12月の支出データ

        this.monthList = monthList;

        this.init();
    }

    private async init() {
        console.log("グラフ描画 start");
        await this.removeBudgetGraph();
        await this.initBudgetGraph();
    }

    private async removeBudgetGraph() {
        // シート内の既存のグラフを取得
        const charts = this.sheet.getCharts();

        // それぞれのグラフを削除
        for (var i = 0; i < charts.length; i++) {
            this.sheet.removeChart(charts[i]);
        }
        console.log("グラフのリセット done");
    }

    private async initBudgetGraph() {
        // 一時的なデータ範囲を作成（A1からD12まで）
        let tempRange = this.sheet.getRange(1, 1, 12, 4);
        // 一時データ表に値を設定
        await this.initTemporaryData(tempRange);
        console.log("グラフ用の一時表作成 done");

        // グラフをシートに挿入
        this.sheet.insertChart(
            this.sheet
                .newChart()
                .setChartType(Charts.ChartType.LINE)
                .addRange(this.sheet.getRange("A1:D12"))
                .setPosition(1, this.initialColumnNumber, 0, 0) // グラフの位置を設定（1行目の6列目から開始）
                .setOption("useFirstColumnAsDomain", true) // 最初の列（A列）をX軸のラベルとして使用
                .setOption("series", {
                    0: { color: "#f9da78" }, // 概算の色を黄に設定
                    1: { color: "#417be5" }, // 収入の色を青に設定
                    2: { color: "#e63a2e" }, // 支出の色を赤に設定
                })
                .setOption("pointSize", 10)
                .setOption("width", 4800) // グラフの幅を設定
                .setOption("height", 270) // グラフの高さを設定
                .setOption("vAxis.minValue", 0)
                .build()
        );
        console.log("グラフ作成 done");
    }

    private async initTemporaryData(
        tempRange: GoogleAppsScript.Spreadsheet.Range
    ): Promise<GoogleAppsScript.Spreadsheet.Range> {
        return tempRange
            .setValues([
                ...this.monthList.map((m, i) => [
                    m,
                    `=${this.roughEstimateSummaryList[i]}`,
                    `=${this.incomeMonthlySummaryList[i]}`,
                    `=${this.outcomeMonthlySummaryList[i]}`,
                ]),
            ])
            .setFontColor("#fff");
    }
}
