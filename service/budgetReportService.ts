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
    private targetSpreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet | null;

    private monthlyBudgetReportList: MonthlyReport[] = [];
    private monthList: string[] = [];

    constructor(
        spreadSheetName: string,
        budgetReportSheetName: string,
        categorySummaryReportSheetName: string,
        options: Options,
        targetSpreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet | null = null
    ) {
        this.spreadSheetName = spreadSheetName;
        this.budgetReportSheetName = budgetReportSheetName;
        this.categorySummaryReportSheetName = categorySummaryReportSheetName;
        this.options = options;
        this.targetSpreadsheet = targetSpreadsheet;

        this.init();
    }

    private init(): void {
        // スプレッドシート、シートの初期化
        let spreadsheet = this.targetSpreadsheet;
        /**
         * シートの存在有無によって処理を分岐させる
         * - 手動実行時： 実行中のスプレッドシートに対してスタイルやバリデーションルールを再適用する
         * - 対象指定なし： 同名のスプレッドシートをDriveから検索する
         * - 存在しない： 新規シートを作成する
         */
        if (spreadsheet) {
            console.log("実行中のスプレッドシートにスタイルを再適用します");
            this.initializeExistingSheets(spreadsheet);
        } else {
            const files = DriveApp.getFilesByName(this.spreadSheetName);

            if (files.hasNext()) {
                console.log("同一名称のシートが存在する");
                spreadsheet = SpreadsheetApp.open(files.next());
                this.initializeExistingSheets(spreadsheet);
            } else {
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
            }
        }

        // 12ヶ月分の budgetReport を作成する。
        this.initMonthlyReports();

        // 年次レポートの作成
        this.initAnnualReport();

        // カテゴリ別レポートの作成
        this.initCategorySummaryReport();
    }

    private initializeExistingSheets(
        spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet
    ): void {
        this.budgetReportSheet =
            spreadsheet.getSheetByName(this.budgetReportSheetName) ||
            spreadsheet.getSheets()[0];
        this.categorySummarySheet =
            spreadsheet.getSheetByName(this.categorySummaryReportSheetName) ||
            spreadsheet.insertSheet(this.categorySummaryReportSheetName);
    }

    private initMonthlyReports(): void {
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

    private initAnnualReport(): void {
        new AnnualReport(
            this.budgetReportSheet,
            this.monthlyBudgetReportList,
            this.monthList,
            this.options.budgetReportOption
        );
    }

    private initCategorySummaryReport(): void {
        new CategorySummaryReport(
            this.categorySummarySheet,
            this.monthlyBudgetReportList,
            this.budgetReportSheetName,
            this.options.categorySummaryOption
        );
    }
}
