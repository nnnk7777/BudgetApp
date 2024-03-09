import { BudgetGraph } from "./budgetGraph";
import { MonthlyReport } from "./monthlyReport";

export class AnnualReport {
    private sheet: GoogleAppsScript.Spreadsheet.Sheet;
    private rowOffset: number;
    private columnOffset!: number;

    private monthlyReports!: MonthlyReport[];
    private monthNumberRow!: GoogleAppsScript.Spreadsheet.Range;
    private monthNumberRowNum: number;
    private monthList!: string[];

    private roughEstimateSummaryRowNum: number;
    private roughEstimateSummaryRow!: GoogleAppsScript.Spreadsheet.Range;
    private roughEstimateSummaryCell!: GoogleAppsScript.Spreadsheet.Range;
    private roughEstimateSummaryList!: string[];

    private incomeSummaryRowNum: number;
    private incomeSummaryRow!: GoogleAppsScript.Spreadsheet.Range;
    private incomeSummaryCell!: GoogleAppsScript.Spreadsheet.Range;
    private incomeSummaryLabelCell!: GoogleAppsScript.Spreadsheet.Range;
    private incomeMonthlySummaryList!: string[];

    private outcomeSummaryRowNum: number;
    private outcomeSummaryRow!: GoogleAppsScript.Spreadsheet.Range;
    private outcomeSummaryCell!: GoogleAppsScript.Spreadsheet.Range;
    private outcomeSummaryLabelCell!: GoogleAppsScript.Spreadsheet.Range;
    private outcomeMonthlySummaryList!: string[];

    constructor(
        sheet: GoogleAppsScript.Spreadsheet.Sheet,
        columnOffset: number = 7,
        rowOffset: number = 14
    ) {
        this.sheet = sheet;
        this.columnOffset = columnOffset;
        this.rowOffset = rowOffset;

        this.monthNumberRowNum = 15;
        this.roughEstimateSummaryRowNum = 17;
        this.incomeSummaryRowNum = 19;
        this.outcomeSummaryRowNum = 29;

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
        this.sheet.setFrozenColumns(4);
        this.sheet.setColumnWidth(3, 150);

        // 年次レポート左部の色を指定
        this.roughEstimateSummaryRow = this.sheet.getRange(
            this.roughEstimateSummaryRowNum,
            1,
            1,
            this.columnOffset
        );
        this.roughEstimateSummaryCell = this.sheet.getRange(
            this.roughEstimateSummaryRowNum,
            3,
            1,
            1
        );
        this.roughEstimateSummaryRow.setBackground("#F4F161");
        this.roughEstimateSummaryCell.setFontSize(14).setFontWeight("bold");

        this.monthNumberRow = this.sheet.getRange(
            this.monthNumberRowNum,
            1,
            1,
            this.columnOffset
        );
        this.monthNumberRow.setBackground("#E08A8A");

        this.incomeSummaryRow = this.sheet.getRange(
            this.incomeSummaryRowNum,
            1,
            1,
            this.columnOffset
        );
        this.incomeSummaryCell = this.sheet.getRange(
            this.incomeSummaryRowNum,
            3,
            1,
            1
        );
        this.incomeSummaryLabelCell = this.sheet.getRange(
            this.incomeSummaryRowNum,
            2,
            1,
            1
        );
        this.incomeSummaryRow.setBackground("#C8DEF1");
        this.incomeSummaryCell.setFontSize(14).setFontWeight("bold");
        this.incomeSummaryLabelCell
            .setFontWeight("bold")
            .setHorizontalAlignment("center")
            .setVerticalAlignment("middle");

        this.outcomeSummaryRow = this.sheet.getRange(
            this.outcomeSummaryRowNum,
            1,
            1,
            this.columnOffset
        );
        this.outcomeSummaryCell = this.sheet.getRange(
            this.outcomeSummaryRowNum,
            3,
            1,
            1
        );
        this.outcomeSummaryLabelCell = this.sheet.getRange(
            this.outcomeSummaryRowNum,
            2,
            1,
            1
        );
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
            console.log(`-- ${i + 1}月 done`);
        }
    }

    private async initAnnualSummaryCells() {
        this.roughEstimateSummaryList = [];
        this.incomeMonthlySummaryList = [];
        this.outcomeMonthlySummaryList = [];
        for (let i = 0; i < 12; i++) {
            this.roughEstimateSummaryList.push(
                this.monthlyReports[i].getRoughEstimateCell()
            );
            this.incomeMonthlySummaryList.push(
                this.monthlyReports[i]
                    .getIncomeData()
                    .getIncomeSummaryCellString()
            );
            this.outcomeMonthlySummaryList.push(
                this.monthlyReports[i]
                    .getOutcomeData()
                    .getOutcomeSummaryCellString()
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
        console.log("月次サマリとその合計値の設定 done");
    }

    private async initAnnualReportGraph() {
        new BudgetGraph(
            this.sheet,
            this.columnOffset,
            this.roughEstimateSummaryList,
            this.incomeMonthlySummaryList,
            this.outcomeMonthlySummaryList,
            this.monthList
        );
    }
}
