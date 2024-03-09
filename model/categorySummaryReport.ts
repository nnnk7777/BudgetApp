import { OutcomeCategoryList } from "./outcomeCategoryList";
import { MonthlyCategoryReport } from "./monthlyCategoryReport";
import {
    CategorySummaryOption,
    MonthlyCategoryReportOption,
} from "../types/options";
import { MonthlyReport } from "./monthlyReport";

export class CategorySummaryReport {
    private sheet: GoogleAppsScript.Spreadsheet.Sheet;
    private monthlyBudgetReportList: MonthlyReport[];
    private budgetReportSheetName: string;
    private option: CategorySummaryOption;

    private outcomeCategoryNameList: string[];
    private categoryLabelRange!: GoogleAppsScript.Spreadsheet.Range;

    conditions = [
        {
            range: [0, 0],
            backgroundColor: "#ffffff",
            fontColor: "#000000",
            fontBold: false,
        },
        {
            range: [1, 4999],
            backgroundColor: "#f7cfcf",
            fontColor: "#000000",
            fontBold: false,
        },
        {
            range: [5000, 19999],
            backgroundColor: "#de7c7c",
            fontColor: "#000000",
            fontBold: true,
        },
        {
            range: [20000, 49999],
            backgroundColor: "#d84f4f",
            fontColor: "#ffffff",
            fontBold: true,
        },
        {
            range: [50000, 79999],
            backgroundColor: "#c4200e",
            fontColor: "#ffffff",
            fontBold: true,
        },
        {
            range: [80000, 139999],
            backgroundColor: "#2d57e1",
            fontColor: "#ffffff",
            fontBold: true,
        },
        {
            range: [140000, 2000000],
            backgroundColor: "#062384",
            fontColor: "#ffffff",
            fontBold: true,
        },
    ];

    constructor(
        sheet: GoogleAppsScript.Spreadsheet.Sheet,
        monthlyBudgetReportList: MonthlyReport[],
        budgetReportSheetName: string,
        option: CategorySummaryOption
    ) {
        this.sheet = sheet;
        this.monthlyBudgetReportList = monthlyBudgetReportList;
        this.budgetReportSheetName = budgetReportSheetName;
        this.option = option;

        this.outcomeCategoryNameList =
            new OutcomeCategoryList().getCategoryNames();

        this.init();
    }

    private async init() {
        console.log("カテゴリ別サマリ作成 start");
        await this.resetSheet();
        await this.createCategoryLabel();

        await this.createMontylyCategoryReport();
        await this.setMoneyRangeBackgroundColor();
    }

    private async resetSheet(): Promise<void> {
        this.sheet.clear();
        this.sheet.clearConditionalFormatRules();
        this.sheet.getRange(`A1:BD2000`).clearDataValidations();
        this.sheet.getRange(`A1:BD2000`).clearFormat();
    }

    private async createCategoryLabel(): Promise<void> {
        this.categoryLabelRange = this.sheet.getRange(
            this.option.rowOffset,
            this.option.columnOffset,
            this.outcomeCategoryNameList.length,
            1
        );
        this.categoryLabelRange.setValues(
            this.outcomeCategoryNameList.map((o) => {
                return [o];
            })
        );

        console.log("カテゴリ別サマリ作成 ラベル作成完了");
    }

    private async createMontylyCategoryReport(): Promise<void> {
        console.log("カテゴリ別月次サマリ作成 start");

        this.sheet.setColumnWidth(2, 150);
        for (let i = 0; i < 12; i++) {
            let option: MonthlyCategoryReportOption = {
                initialColumnNumber: this.option.columnOffset + i + 1,
                initialRowNumber: this.option.rowOffset,
            };

            new MonthlyCategoryReport(
                this.sheet,
                i + 1,
                this.budgetReportSheetName,
                this.monthlyBudgetReportList[i],
                option
            );
            console.log(`-- ${i + 1}月 done`);
        }
    }

    private async setMoneyRangeBackgroundColor(): Promise<void> {
        let rules = this.sheet.getConditionalFormatRules();

        const range = this.sheet.getRange(
            this.option.rowOffset,
            this.option.columnOffset + 1,
            this.outcomeCategoryNameList.length,
            12
        ); // 条件付きフォーマットを適用する範囲
        for (const condition of this.conditions) {
            let rule = SpreadsheetApp.newConditionalFormatRule()
                .whenNumberBetween(condition.range[0], condition.range[1])
                .setBackground(condition.backgroundColor)
                .setFontColor(condition.fontColor)
                .setBold(condition.fontBold)
                .setRanges([range])
                .build();

            rules.push(rule);
        }

        this.sheet.setConditionalFormatRules(rules);
    }
}
