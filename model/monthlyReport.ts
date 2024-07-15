import { IncomeData } from "./incomeData";
import { MonthlyDataBase } from "./monthlyDataBase";
import { OutcomeData } from "./outcomeData";

import { getCellNotation } from "../util/getCellNotation";
import { columnToLetter } from "../util/columnToletter";
import { MonthlyReportOption } from "../types/options";

export class MonthlyReport extends MonthlyDataBase {
    private month: number;
    private incomeData!: IncomeData;
    private outcomeData!: OutcomeData;

    private options: MonthlyReportOption;

    constructor(
        sheet: GoogleAppsScript.Spreadsheet.Sheet,
        month: number,
        options: MonthlyReportOption
    ) {
        super(sheet);
        this.month = month;
        this.options = options;

        this.init();
    }

    // 月の表示と、支出・収入のSUM関数をセット
    private async init() {
        await this.clearSheet();

        // 月の表示
        const firstCell = this.sheet.getRange(
            this.options.initialRowNumber,
            this.options.initialColumnNumber,
            1,
            1
        );
        firstCell.setValue(`${this.month}月`);
        firstCell
            .setFontSize(14)
            .setFontWeight("bold")
            .setHorizontalAlignment("center");
        const firstRow = this.sheet.getRange(
            this.options.initialRowNumber,
            this.options.initialColumnNumber,
            1,
            4
        );
        firstRow.setBackground("#E08A8A");

        // 概算の表示
        const roughEstimateCell = this.sheet.getRange(
            this.options.initialRowNumber + 2,
            this.options.initialColumnNumber + 3,
            1,
            1
        );
        this.setRoughEstimateSum(roughEstimateCell);
        const roughEstimateRow = this.sheet.getRange(
            this.options.initialRowNumber + 2,
            this.options.initialColumnNumber,
            1,
            4
        );
        roughEstimateRow.setBackground("#F4F161");

        // 前月との境界に点線をひく
        const maxRows = this.sheet.getMaxRows();
        const range = this.sheet.getRange(
            this.options.initialRowNumber,
            this.options.initialColumnNumber,
            maxRows,
            1
        ); // 1行目から最終行までのJ列を取得
        range.setBorder(
            null,
            true,
            null,
            null,
            null,
            null,
            "black",
            SpreadsheetApp.BorderStyle.DOTTED
        );

        await this.initInOut();
        await this.setColumnsWidth();

        // カスタム日付フォーマットを設定
        const dateColumn = this.sheet.getRange(
            1,
            this.options.initialColumnNumber,
            maxRows
        );
        dateColumn
            .setNumberFormat("MM/dd")
            .setHorizontalAlignment("center")
            .setVerticalAlignment("middle");
    }

    private async initInOut() {
        const initialIncomingRowNumber = this.options.initialRowNumber + 4;
        const initialOutcomingRowNumber = initialIncomingRowNumber + 15;

        this.incomeData = new IncomeData(
            this.sheet,
            initialIncomingRowNumber,
            this.options.initialColumnNumber
        );

        this.outcomeData = new OutcomeData(
            this.sheet,
            initialOutcomingRowNumber,
            this.options.initialColumnNumber
        );
    }

    private async clearSheet() {
        this.sheet.clearConditionalFormatRules();
        this.sheet.getRange(`A1:BD2000`).clearDataValidations();
        this.sheet.getRange(`A1:BD2000`).clearFormat();
    }

    private async setColumnsWidth() {
        // 列の幅を設定する
        for (let i = 0; i < this.columnWidths.length; i++) {
            this.sheet.setColumnWidth(
                this.options.initialColumnNumber + i,
                this.columnWidths[i]
            );
        }
    }

    public getRoughEstimateCell(): string {
        return (
            columnToLetter(this.options.initialColumnNumber + 3) +
            (this.options.initialRowNumber + 2)
        );
    }

    public getIncomeData(): IncomeData {
        return this.incomeData;
    }

    public getOutcomeData(): OutcomeData {
        return this.outcomeData;
    }

    private setRoughEstimateSum(
        baseCell: GoogleAppsScript.Spreadsheet.Range
    ): void {
        // 引数で受け取ったRangeオブジェクトから、シートとセルの位置情報を取得
        var column = baseCell.getColumn();
        var row = baseCell.getRow();

        // SUM関数の計算式を構築する
        // 例: "=B1 + A3 - A13" （baseCellRangeがA1の場合）
        var formula =
            "=" +
            getCellNotation(column - 4, row) +
            " + " + // 4つ左のセル
            getCellNotation(column, row + 2) +
            " - " + // 2つ下のセル
            getCellNotation(column, row + 17); // 12個下のセル

        // SUM関数を基準セルに設定する
        baseCell.setFormula(formula);
    }
}
