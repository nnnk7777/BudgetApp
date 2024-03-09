export type Options = {
    budgetReportOption: BudgetReportOption;
    categorySummaryOption: CategoryReportOption;
};

export type BudgetReportOption = {
    columnOffset: number;
    rowOffset: number;
    monthNumberRowNum: number;
    roughEstimateSummaryRowNum: number;
    incomeSummaryRowNum: number;
    outcomeSummaryRowNum: number;
};

export type CategoryReportOption = {
    columnOffset: number;
    rowOffset: number;
};

export type MonthlyReportOption = {
    initialColumnNumber: number;
    initialRowNumber: number;
};

export type CategorySummaryOption = {
    columnOffset: number;
    rowOffset: number;
};

export type MonthlyCategoryReportOption = {
    initialColumnNumber: number;
    initialRowNumber: number;
};
