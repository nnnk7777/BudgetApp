// action: text
function availableExpenseSummaryText() {
    return calculateExpensesSummary('text');
}

function availableDailySummaryText() {
    return calculateDailySummary('text');
}

function availableWeeklySummaryText() {
    return calculateWeeklySummary('text');
}

function availableMonthlySummaryText() {
    return calculateMonthlySummaryByUnit('text');
}

// action: mail
function availableExpenseSummaryMail() {
    return calculateExpensesSummary('mail');
}

function availableDailySummaryMail() {
    return calculateDailySummary('mail');
}

function availableWeeklySummaryMail() {
    return calculateWeeklySummary('mail');
}

// monthly summary mail
function availableMonthlySummaryMail() {
    return calculateMonthlySummaryByUnit('mail');
}

// manual diagnostics
function availableRuntimeDiagnostics() {
    var result = JSON.stringify(getScriptRuntimeDiagnostics());
    Logger.log(result);
}

// action: categories
function availableFetchCategories() {
    return fetchCategories();
}

// action: list_uncategorized
function availableListUncategorizedExpenses() {
    return listUncategorizedExpenses();
}

// action: autofill_uncategorized
function availableAutofillUncategorizedExpenses() {
    return autofillUncategorizedExpenses({
        confidenceThreshold: 0.9,
        debug: true
    });
}
