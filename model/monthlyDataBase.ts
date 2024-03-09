import { columnToLetter } from "../util/columnToletter";
import { CategoryList } from "./categoryList";

export class MonthlyDataBase {
    protected sheet: GoogleAppsScript.Spreadsheet.Sheet;
    columnWidths: number[] = [100, 70, 100, 100];
    conditions = [
        {
            range: [1500, 2499],
            backgroundColor: null,
            fontColor: null,
            fontBold: true,
        },
        {
            range: [2500, 4999],
            backgroundColor: "#ffe18f",
            fontColor: "#e39000",
            fontBold: true,
        },
        {
            range: [5000, 9999],
            backgroundColor: "#ffb9b5",
            fontColor: "#c21b15",
            fontBold: true,
        },
        {
            range: [10000, 10000000],
            backgroundColor: "#c4200e",
            fontColor: "#ffffff",
            fontBold: true,
        },
    ];

    constructor(sheet: GoogleAppsScript.Spreadsheet.Sheet) {
        this.sheet = sheet;
    }

    public setMoneyRangeBackgroundColor(
        targetColumnNumber: number,
        initialRowNumber: number,
        rowCount: number
    ): void {
        let rules = this.sheet.getConditionalFormatRules();

        const range = this.sheet.getRange(
            initialRowNumber,
            targetColumnNumber,
            rowCount,
            1
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

    public setLabelBackgroundColor(
        categoryList: CategoryList,
        targetColumnNumber: number,
        initialRowNumber: number,
        rowCount: number
    ): void {
        let rules = this.sheet.getConditionalFormatRules();
        const targetRange = this.sheet.getRange(
            initialRowNumber,
            targetColumnNumber,
            rowCount,
            2
        );

        for (const category of categoryList.getCategoryList()) {
            let rule = SpreadsheetApp.newConditionalFormatRule()
                .whenFormulaSatisfied(
                    `=$${columnToLetter(
                        targetColumnNumber
                    )}${initialRowNumber}="${category.getCategoryName()}"`
                )
                .setBackground(category.getBackgroundColor())
                .setFontColor(
                    category.getFontColor() ? category.getFontColor() : null
                )
                .setRanges([targetRange])
                .build();

            rules.push(rule);
        }
        this.sheet.setConditionalFormatRules(rules);
    }
}
