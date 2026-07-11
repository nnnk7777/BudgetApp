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

// action: categories
function fetchCategoriesManual() {
    return fetchCategories();
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
