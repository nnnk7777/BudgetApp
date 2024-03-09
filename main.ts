import { BudgetReportService } from "./service/budgetReportService";
import { Options } from "./types/options";

export function main() {
    const speadSheetName = "金銭メモ2024_5";
    const budgetReportSheetName = "🐖 家計簿";
    const categorySummaryReportSheetName = "🦦 カテゴリ別";

    const options: Options = {
        budgetReportOption: {
            columnOffset: 7,
            rowOffset: 14,
            monthNumberRowNum: 15,
            roughEstimateSummaryRowNum: 17,
            incomeSummaryRowNum: 19,
            outcomeSummaryRowNum: 29,
        },
        categorySummaryOption: {
            columnOffset: 2,
            rowOffset: 25,
        },
    };

    new BudgetReportService(
        speadSheetName,
        budgetReportSheetName,
        categorySummaryReportSheetName,
        options
    );
}
