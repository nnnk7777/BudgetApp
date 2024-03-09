import { MonthlyCategoryReportOption } from "../types/options";
import { MonthlyReport } from "./monthlyReport";
import { OutcomeCategoryList } from "./outcomeCategoryList";

export class MonthlyCategoryReport {
    private sheet: GoogleAppsScript.Spreadsheet.Sheet;
    private month: number;
    private budgetReportSheetName: string;
    private monthlyBudgetReport: MonthlyReport; // 同じ月の MonthlyReport を渡す
    private option: MonthlyCategoryReportOption;

    private outcomeCategoryNameList: string[];

    constructor(
        sheet: GoogleAppsScript.Spreadsheet.Sheet,
        month: number,
        budgetReportSheetName: string,
        monthlyBudgetReport: MonthlyReport,
        option: MonthlyCategoryReportOption
    ) {
        this.sheet = sheet;
        this.month = month;
        this.budgetReportSheetName = budgetReportSheetName;
        this.monthlyBudgetReport = monthlyBudgetReport;
        this.option = option;
        this.outcomeCategoryNameList =
            new OutcomeCategoryList().getCategoryNames();

        this.init();
    }

    private async init() {
        // 月の表示
        await this.createMonthNumberLabel();

        // カテゴリごとに SUM 関数をセット
        await this.createCategorySummary();
    }

    private async createMonthNumberLabel(): Promise<void> {
        const monthNumCell = this.sheet.getRange(
            this.option.initialRowNumber - 1,
            this.option.initialColumnNumber,
            1,
            1
        );
        monthNumCell.setValue(`${this.month}月`);
        monthNumCell
            .setFontSize(14)
            .setFontWeight("bold")
            .setHorizontalAlignment("center");
    }

    private async createCategorySummary(): Promise<void> {
        this.outcomeCategoryNameList.map((outcomeCategoryName, i) => {
            this.sheet
                .getRange(
                    this.option.initialRowNumber + i,
                    this.option.initialColumnNumber
                )
                .setFormula(
                    `=SUM(
              query(
                  '${this.budgetReportSheetName}'!${this.monthlyBudgetReport
                        .getOutcomeData()
                        .getCategoryLabelColumnStr()}:${this.monthlyBudgetReport
                        .getOutcomeData()
                        .getPriceColumnStr()},
                "select ${this.monthlyBudgetReport
                    .getOutcomeData()
                    .getPriceColumnStr()} where ${this.monthlyBudgetReport
                        .getOutcomeData()
                        .getCategoryLabelColumnStr()}='${outcomeCategoryName}'"
              )
            )`
                );
        });
    }
}
