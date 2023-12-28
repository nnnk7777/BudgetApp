import { BudgetGraph } from "./budgetGraph";
import { MonthlyReport } from "./monthlyReport";

export class AnnualReport {
  private sheet: GoogleAppsScript.Spreadsheet.Sheet;
  private rowOffset: number;

  private monthlyReports!: MonthlyReport[];
  private columnOffset!: number;

  private monthNumberRow!: GoogleAppsScript.Spreadsheet.Range;
  private incomeSummaryRow!: GoogleAppsScript.Spreadsheet.Range;
  private incomeSummaryCell!: GoogleAppsScript.Spreadsheet.Range;
  private incomeSummaryLabelCell!: GoogleAppsScript.Spreadsheet.Range;
  private outcomeSummaryRow!: GoogleAppsScript.Spreadsheet.Range;
  private outcomeSummaryCell!: GoogleAppsScript.Spreadsheet.Range;
  private outcomeSummaryLabelCell!: GoogleAppsScript.Spreadsheet.Range;

  private incomeMonthlySummaryList!: string[];
  private outcomeMonthlySummaryList!: string[];
  private monthList!: string[];

  constructor(
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    columnOffset: number = 6,
    rowOffset: number = 14
  ) {
    this.sheet = sheet;
    this.columnOffset = columnOffset;
    this.rowOffset = rowOffset;

    this.init();
  }

  private async init() {
    await this.initMonthlyReport();
    await this.initAnnualReportStyle();
    await this.initAnnualSummaryCells();
    await this.initAnnualReportGraph();
  }

  private async initAnnualReportStyle() {
    // 左側の3列を固定
    this.sheet.setFrozenColumns(3);
    this.sheet.setColumnWidth(3, 150);

    // 年次レポート左部の色を指定
    this.monthNumberRow = this.sheet.getRange(15, 1, 1, this.columnOffset);
    this.monthNumberRow.setBackground("#E08A8A");

    this.incomeSummaryRow = this.sheet.getRange(17, 1, 1, this.columnOffset);
    this.incomeSummaryCell = this.sheet.getRange(17, 3, 1, 1);
    this.incomeSummaryLabelCell = this.sheet.getRange(17, 2, 1, 1);
    this.incomeSummaryRow.setBackground("#C8DEF1");
    this.incomeSummaryCell.setFontSize(14).setFontWeight("bold");
    this.incomeSummaryLabelCell
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");

    this.outcomeSummaryRow = this.sheet.getRange(27, 1, 1, this.columnOffset);
    this.outcomeSummaryCell = this.sheet.getRange(27, 3, 1, 1);
    this.outcomeSummaryLabelCell = this.sheet.getRange(27, 2, 1, 1);
    this.outcomeSummaryRow.setBackground("#F2C5C5");
    this.outcomeSummaryCell.setFontSize(14).setFontWeight("bold");
    this.outcomeSummaryLabelCell
      .setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");
  }

  private async initMonthlyReport() {
    this.monthList = [];
    this.monthlyReports = [];
    for (let i = 0; i < 12; i++) {
      this.monthlyReports.push(
        new MonthlyReport(this.sheet, i + 1, this.columnOffset + i * 4)
      );
      this.monthList.push(`${i + 1}月`);
    }
  }

  private async initAnnualSummaryCells() {
    this.incomeMonthlySummaryList = [];
    this.outcomeMonthlySummaryList = [];
    for (let i = 0; i < 12; i++) {
      this.incomeMonthlySummaryList.push(
        this.monthlyReports[i].getIncomeData().getIncomeSummaryCellString()
      );
      this.outcomeMonthlySummaryList.push(
        this.monthlyReports[i].getOutcomeData().getOutcomeSummaryCellString()
      );
    }
    // 月次レポートの合計値を算出
    const incomeRangeString = this.incomeMonthlySummaryList.join(",");
    this.incomeSummaryCell
      .setFormula(`=SUM(${incomeRangeString})`)
      .setNumberFormat("¥#,##0");
    this.incomeSummaryLabelCell.setValue("収入");
    const outcomeRangeString = this.outcomeMonthlySummaryList.join(",");
    this.outcomeSummaryCell
      .setFormula(`=SUM(${outcomeRangeString})`)
      .setNumberFormat("¥#,##0");
    this.outcomeSummaryLabelCell.setValue("支出");
  }

  private async initAnnualReportGraph() {
    new BudgetGraph(
      this.sheet,
      this.columnOffset,
      this.rowOffset,
      this.incomeMonthlySummaryList,
      this.outcomeMonthlySummaryList,
      this.monthList
    );
  }
}
