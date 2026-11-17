// action: text
function expenseSummaryText() {
    return calculateExpensesSummary('text');
}

function dailySummaryText() {
    return calculateDailySummary('text');
}

function weeklySummaryText() {
    return calculateWeeklySummary('text');
}

function monthlySummaryText() {
    return calculateMonthlySummaryByUnit('text');
}

// action: mail
function expenseSummaryMail() {
    return calculateExpensesSummary('mail');
}

function dailySummaryMail() {
    return calculateDailySummary('mail');
}

function weeklySummaryMail() {
    return calculateWeeklySummary('mail');
}

// monthly summary mail
function monthlySummaryMail() {
    return calculateMonthlySummaryByUnit('mail');
}

// manual diagnostics
function runtimeDiagnostics() {
    var result = JSON.stringify(getScriptRuntimeDiagnostics());
    Logger.log(result);
}

// action: add_income
function addIncomeRecordManual() {
    var itemJson = PropertiesService.getScriptProperties().getProperty('MANUAL_INCOME_ITEM');
    if (!itemJson || !itemJson.trim()) {
        throw new Error('スクリプト プロパティ MANUAL_INCOME_ITEM に収入の JSON を設定してください');
    }

    var result = addIncomeRecord(JSON.parse(itemJson));
    Logger.log(JSON.stringify(result));
    return result;
}

// action: categories
function fetchCategoriesManual() {
    return fetchCategories();
}

// action: list_today_expenses
function listTodayExpensesManual() {
    var result = listTodayExpenses();
    Logger.log(result);
    return result;
}

// action: list_uncategorized
function listUncategorizedExpensesManual() {
    return listUncategorizedExpenses();
}

// action: autofill_uncategorized
function autofillUncategorizedExpensesManual() {
    return autofillUncategorizedExpenses({
        confidenceThreshold: 0.9,
        debug: true
    });
}
