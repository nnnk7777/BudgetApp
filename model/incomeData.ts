import { MonthlyDataBase } from "./monthlyDataBase";

import { columnToLetter } from "../util/columnToletter";
import { IncomeCategoryList } from "./incomeCategoryList";

export class IncomeData extends MonthlyDataBase {
    private incomeDetailRowCount: number;
    private initialColumnNumber: number;
    private initialRowNumber: number;
    private initialCalcRowNumber: number;
    private calcColumnNumber: number;
    private incomeCategoryList: IncomeCategoryList;

    constructor(
        sheet: GoogleAppsScript.Spreadsheet.Sheet,
        initialRowNumber: number,
        initialColumnNumber: number
    ) {
        super(sheet);
        this.initialRowNumber = initialRowNumber;
        this.initialCalcRowNumber = initialRowNumber + 1;
        this.initialColumnNumber = initialColumnNumber;
        this.calcColumnNumber = initialColumnNumber + 3;
        this.incomeDetailRowCount = 9;
        this.incomeCategoryList = new IncomeCategoryList();

        this.init();
    }

    private init() {
        // SUM関数の範囲を文字列として構築（例: "A1:A5"）
        const sumFormulaRange =
            columnToLetter(this.calcColumnNumber) +
            this.initialCalcRowNumber +
            ":" +
            columnToLetter(this.calcColumnNumber) +
            (this.initialCalcRowNumber + this.incomeDetailRowCount - 1);

        // SUM関数をセットするセルを指定
        const formulaCell = this.sheet.getRange(
            this.initialRowNumber,
            this.calcColumnNumber
        );
        formulaCell
            .setFormula("=SUM(" + sumFormulaRange + ")")
            .setNumberFormat("¥#,##0")
            .setFontSize(14)
            .setFontWeight("bold");

        // 合計値を表示する行の色を指定
        const incomeSummaryRow = this.sheet.getRange(
            this.initialRowNumber,
            this.calcColumnNumber - 3,
            1,
            4
        );
        incomeSummaryRow.setBackground("#C8DEF1");

        // プルダウンメニューの作成
        const rule = SpreadsheetApp.newDataValidation()
            .requireValueInList(this.incomeCategoryList.getCategoryNames())
            .build();

        let range = this.sheet.getRange(
            this.initialRowNumber + 1,
            this.initialColumnNumber + 1,
            this.incomeDetailRowCount,
            1
        );
        range.setDataValidation(rule).setFontSize(8);

        // 金額によって背景色が変わるように設定
        this.setMoneyRangeBackgroundColor(
            this.initialColumnNumber + 3,
            this.initialRowNumber + 1,
            this.incomeDetailRowCount
        );

        // ラベルによって背景色が変わるように設定
        this.setLabelBackgroundColor(
            this.incomeCategoryList,
            this.initialColumnNumber + 1,
            this.initialRowNumber + 1,
            this.incomeDetailRowCount
        );
    }

    public getIncomeSummaryCellString(): string {
        return columnToLetter(this.calcColumnNumber) + this.initialRowNumber;
    }
}
