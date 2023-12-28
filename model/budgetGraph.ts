export class BudgetGraph {
  private sheet: GoogleAppsScript.Spreadsheet.Sheet;
  private initialColumnNumber: number;
  private initialRowNumber: number;
  private incomeMonthlySummaryList: string[];
  private outcomeMonthlySummaryList: string[];
  private monthList: string[];

  constructor(
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    initialColumnNumber: number,
    initialRowNumber: number,
    incomeMonthlySummaryList: string[],
    outcomeMonthlySummaryList: string[],
    monthList: string[]
  ) {
    this.sheet = sheet;
    this.initialColumnNumber = initialColumnNumber;
    this.initialRowNumber = initialRowNumber;

    // 収入と支出のデータ範囲を設定 (ここでは例としてA15:B26を使用)
    this.incomeMonthlySummaryList = incomeMonthlySummaryList; // 1〜12月の収入データ
    this.outcomeMonthlySummaryList = outcomeMonthlySummaryList; // 1〜12月の支出データ

    this.monthList = monthList;

    this.init();
  }

  private async init() {
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
  }

  private async initBudgetGraph() {
    // 一時的なデータ範囲を作成（A1からC13まで）
    let tempRange = this.sheet.getRange(1, 1, 12, 3);
    tempRange
      .setValues([
        ...this.monthList.map((m, i) => [
          m,
          `=${this.incomeMonthlySummaryList[i]}`,
          `=${this.outcomeMonthlySummaryList[i]}`,
        ]),
      ])
      .setFontColor("#fff");

    // グラフを作成
    let chart = this.sheet
      .newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(tempRange)
      .setOption("series", {
        0: { color: "#417be5" }, // 収入の色を青に設定
        1: { color: "#e63a2e" }, // 支出の色を赤に設定
      })
      .setPosition(1, this.initialColumnNumber, 0, 0) // グラフの位置を設定（5行目の6列目から開始）
      .setOption("width", 4600) // グラフの幅を設定
      .setOption("height", 270) // グラフの高さを設定
      .setOption("vAxis.minValue", 0)
      .setOption("series", {
        0: { dataLabel: "value" },
        1: { dataLabel: "value" },
      })
      .build();

    // グラフをシートに挿入
    this.sheet.insertChart(chart);
  }
}
