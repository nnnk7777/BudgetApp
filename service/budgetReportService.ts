import { AnnualReport } from "../model/annualReport";
import { CategorySummaryReport } from "../model/categorySummaryReport";
import { MonthlyReport } from "../model/monthlyReport";

import { Options, MonthlyReportOption } from "../types/options";

/**
 * 家計簿機能のサービス
 */
export class BudgetReportService {
    private spreadSheetName: string;
    private budgetReportSheetName: string;
    private categorySummaryReportSheetName: string;

    private budgetReportSheet!: GoogleAppsScript.Spreadsheet.Sheet;
    private categorySummarySheet!: GoogleAppsScript.Spreadsheet.Sheet;
    private options: Options;

    private monthlyBudgetReportList: MonthlyReport[] = [];
    private monthList: string[] = [];

    constructor(
        spreadSheetName: string,
        budgetReportSheetName: string,
        categorySummaryReportSheetName: string,
        options: Options
    ) {
        this.spreadSheetName = spreadSheetName;
        this.budgetReportSheetName = budgetReportSheetName;
        this.categorySummaryReportSheetName = categorySummaryReportSheetName;
        this.options = options;

        this.init();
    }

    private async init() {
        // スプレッドシート、シートの初期化
        let spreadsheet;
        // シート検索のための初期化
        const files = DriveApp.getFilesByName(this.spreadSheetName);
        /**
         * シートの存在有無によって処理を分岐させる
         * - 存在する： 既存のシートに対してスタイルやバリデーションルールの再適用のみをする
         * - 存在しない： 新規シートを作成する
         */
        switch (files.hasNext()) {
            case true:
                console.log("同一名称のシートが存在する");

                // スプレッドシートが存在する場合、そのスプレッドシートを取得
                spreadsheet = SpreadsheetApp.open(files.next());
                this.budgetReportSheet = spreadsheet.getSheets()[0];

                // シート「🦦カテゴリ別」が存在しない場合は、新しいシートとして追加
                if (
                    !spreadsheet.getSheetByName(
                        this.categorySummaryReportSheetName
                    )
                ) {
                    this.categorySummarySheet = spreadsheet.insertSheet(
                        this.categorySummaryReportSheetName
                    );
                } else {
                    this.categorySummarySheet = spreadsheet.getSheets()[1];
                }
                break;
            case false:
                console.log("同一名称のシートが存在しない");

                /**
                 * 家計簿
                 * 1. スプレッドシートが存在しない場合、新しいスプレッドシートを作成
                 * 2. 新しいスプレッドシートにデフォルトで含まれるシートを取得
                 * 3. シート名を変更
                 */
                spreadsheet = SpreadsheetApp.create(this.spreadSheetName);
                this.budgetReportSheet = spreadsheet.getSheets()[0];
                this.budgetReportSheet.setName(this.budgetReportSheetName);
                /**
                 * カテゴリ別サマリ
                 * - スプレッドシートに対して新しいシートを挿入し、名前を設定
                 */
                this.categorySummarySheet = spreadsheet.insertSheet(
                    this.categorySummaryReportSheetName
                );
                break;
        }

        // 12ヶ月分の budgetReport を作成する。
        await this.initMonthlyReports();

        // 年次レポートの作成
        await this.initAnnualReport();

        // カテゴリ別レポートの作成
        await this.initCategorySummaryReport();
    }

    private async initMonthlyReports(): Promise<void> {
        for (let i = 0; i < 12; i++) {
            let monthlyOption: MonthlyReportOption = {
                initialColumnNumber:
                    this.options.budgetReportOption.columnOffset + i * 4,
                initialRowNumber: 15,
            };

            this.monthlyBudgetReportList.push(
                new MonthlyReport(this.budgetReportSheet, i + 1, monthlyOption)
            );
            this.monthList.push(`${i + 1}月`);
            console.log(`-- ${i + 1}月 done`);
        }
    }

    private async initAnnualReport(): Promise<void> {
        new AnnualReport(
            this.budgetReportSheet,
            this.monthlyBudgetReportList,
            this.monthList,
            this.options.budgetReportOption
        );
    }

    private async initCategorySummaryReport(): Promise<void> {
        new CategorySummaryReport(
            this.categorySummarySheet,
            this.monthlyBudgetReportList,
            this.budgetReportSheetName,
            this.options.categorySummaryOption
        );
    }
}
