import { BudgetReportOption } from "../types/options";
import { BudgetGraph } from "./budgetGraph";
import { MonthlyReport } from "./monthlyReport";

export class AnnualReport {
    private sheet: GoogleAppsScript.Spreadsheet.Sheet;
    private options: BudgetReportOption;

    private monthlyReports!: MonthlyReport[];
    private monthNumberRow!: GoogleAppsScript.Spreadsheet.Range;
    private monthList: string[];

    private roughEstimateSummaryRow!: GoogleAppsScript.Spreadsheet.Range;
    private roughEstimateSummaryCell!: GoogleAppsScript.Spreadsheet.Range;
    private roughEstimateSummaryList: string[] = [];

    private incomeSummaryRow!: GoogleAppsScript.Spreadsheet.Range;
    private incomeSummaryCell!: GoogleAppsScript.Spreadsheet.Range;
    private incomeSummaryLabelCell!: GoogleAppsScript.Spreadsheet.Range;
    private incomeMonthlySummaryList: string[] = [];

    private outcomeSummaryRow!: GoogleAppsScript.Spreadsheet.Range;
    private outcomeSummaryCell!: GoogleAppsScript.Spreadsheet.Range;
    private outcomeSummaryLabelCell!: GoogleAppsScript.Spreadsheet.Range;
    private outcomeMonthlySummaryList: string[] = [];

    constructor(
        sheet: GoogleAppsScript.Spreadsheet.Sheet,
        monthlyReports: MonthlyReport[],
        monthList: string[],
        options: BudgetReportOption
    ) {
        this.sheet = sheet;
        this.monthlyReports = monthlyReports;
        this.monthList = monthList;
        this.options = options;

        this.init();
    }

    private async init() {
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
            this.options.roughEstimateSummaryRowNum,
            1,
            1,
            this.options.columnOffset
        );
        this.roughEstimateSummaryCell = this.sheet.getRange(
            this.options.roughEstimateSummaryRowNum,
            3,
            1,
            1
        );
        this.roughEstimateSummaryRow.setBackground("#F4F161");
        this.roughEstimateSummaryCell.setFontSize(14).setFontWeight("bold");

        this.monthNumberRow = this.sheet.getRange(
            this.options.monthNumberRowNum,
            1,
            1,
            this.options.columnOffset
        );
        this.monthNumberRow.setBackground("#E08A8A");

        this.incomeSummaryRow = this.sheet.getRange(
            this.options.incomeSummaryRowNum,
            1,
            1,
            this.options.columnOffset
        );
        this.incomeSummaryCell = this.sheet.getRange(
            this.options.incomeSummaryRowNum,
            3,
            1,
            1
        );
        this.incomeSummaryLabelCell = this.sheet.getRange(
            this.options.incomeSummaryRowNum,
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
            this.options.outcomeSummaryRowNum,
            1,
            1,
            this.options.columnOffset
        );
        this.outcomeSummaryCell = this.sheet.getRange(
            this.options.outcomeSummaryRowNum,
            3,
            1,
            1
        );
        this.outcomeSummaryLabelCell = this.sheet.getRange(
            this.options.outcomeSummaryRowNum,
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

    private async initAnnualSummaryCells() {
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
            this.options.columnOffset,
            this.roughEstimateSummaryList,
            this.incomeMonthlySummaryList,
            this.outcomeMonthlySummaryList,
            this.monthList
        );
    }
}
