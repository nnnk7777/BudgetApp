import { MonthlyDataBase } from "./monthlyDataBase";

import { columnToLetter } from "../util/columnToletter";
import { OutcomeCategoryList } from "./outcomeCategoryList";

export class OutcomeData extends MonthlyDataBase {
    private outcomeDetailRowCount: number;
    private initialColumnNumber: number;
    private initialRowNumber: number;
    private initialCalcRowNumber: number;
    private calcColumnNumber: number;
    private outcomeCategoryList: OutcomeCategoryList;

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
        this.outcomeDetailRowCount = 150;
        this.outcomeCategoryList = new OutcomeCategoryList();

        this.init();
    }

    private init() {
        // SUM関数の範囲を文字列として構築（例: "A1:A5"）
        const sumFormulaRange =
            columnToLetter(this.calcColumnNumber) +
            this.initialCalcRowNumber +
            ":" +
            columnToLetter(this.calcColumnNumber) +
            (this.initialCalcRowNumber + this.outcomeDetailRowCount - 1);

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
        const outcomeSummaryRow = this.sheet.getRange(
            this.initialRowNumber,
            this.calcColumnNumber - 3,
            1,
            4
        );
        outcomeSummaryRow.setBackground("#F2C5C5");

        // プルダウンメニューの作成
        const rule = SpreadsheetApp.newDataValidation()
            .requireValueInList(this.outcomeCategoryList.getCategoryNames())
            .build();

        let range = this.sheet.getRange(
            this.initialRowNumber + 1,
            this.initialColumnNumber + 1,
            this.outcomeDetailRowCount,
            1
        );
        range.setDataValidation(rule).setFontSize(8);

        // 金額によって背景色が変わるように設定
        this.setMoneyRangeBackgroundColor(
            this.initialColumnNumber + 3,
            this.initialRowNumber + 1,
            this.outcomeDetailRowCount
        );

        // ラベルによって背景色が変わるように設定
        this.setLabelBackgroundColor(
            this.outcomeCategoryList,
            this.initialColumnNumber + 1,
            this.initialRowNumber + 1,
            this.outcomeDetailRowCount
        );
    }

    public getOutcomeSummaryCellString(): string {
        return columnToLetter(this.calcColumnNumber) + this.initialRowNumber;
    }
}
