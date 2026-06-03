// action: text
function availableExpenseSummaryText() {
    return calculateExpensesSummary('text');
}

// action: mail
function availableExpenseSummaryMail() {
    return calculateExpensesSummary('mail');
}

// monthly summary mail
function availableMonthlySummaryMail() {
    return calculateMonthlySummary('mail');
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
