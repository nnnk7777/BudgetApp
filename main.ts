import { BudgetReportService } from "./service/budgetReportService";
import { Options } from "./types/options";

export function main() {
    const speadSheetName = "TARGET_SHEET_NAME";
    const budgetReportSheetName = "🐖 家計簿";
    const categorySummaryReportSheetName = "🦦 カテゴリ別";

    const options: Options = {
        budgetReportOption: {
            columnOffset: 7,
            rowOffset: 14,
            monthNumberRowNum: 15,
            roughEstimateSummaryRowNum: 17,
            incomeSummaryRowNum: 19,
            outcomeSummaryRowNum: 34,
        },
        categorySummaryOption: {
            columnOffset: 2,
            rowOffset: 3,
        },
    };

    new BudgetReportService(
        speadSheetName,
        budgetReportSheetName,
        categorySummaryReportSheetName,
        options
    );
}
