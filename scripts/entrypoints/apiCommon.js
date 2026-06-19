var ApiActions = Object.freeze({
    TEXT: 'text',
    MAIL: 'mail',
    ADD: 'add',
    CATEGORIES: 'categories',
    LIST_UNCATEGORIZED: 'list_uncategorized',
    AUTOFILL_UNCATEGORIZED: 'autofill_uncategorized'
});

function parseApiRequest(e) {
    if (!e || !e.postData || !e.postData.contents) {
        throw new Error('request body is required');
    }

    return JSON.parse(e.postData.contents);
}

function verifyApiHash(hash) {
    var scriptHash = PropertiesService.getScriptProperties().getProperty("HASH");
    if (hash !== scriptHash) {
        throw new Error('hashが一致しません');
    }
}

function dispatchApiAction(data) {
    switch (data.action) {
        case ApiActions.TEXT:
        case ApiActions.MAIL:
            return calculateExpensesSummary(data.action);
        case ApiActions.ADD:
            return handleAddExpenseAction(data.item);
        case ApiActions.CATEGORIES:
            return fetchCategories();
        case ApiActions.LIST_UNCATEGORIZED:
            return listUncategorizedExpenses();
        case ApiActions.AUTOFILL_UNCATEGORIZED:
            return autofillUncategorizedExpenses(data.options || {});
        default:
            throw new Error('actionが定義されていません');
    }
}

function handleAddExpenseAction(item) {
    if (!item) {
        throw new Error('itemが定義されていません');
    }

    var amount = item.amount;
    if (amount === undefined || amount === null || amount === '') {
        amount = item.price;
    }
    if (amount === undefined || amount === null || amount === '') {
        amount = item.cost;
    }

    return addExpenseRecord(item.title, amount, item.category);
}

function buildApiErrorResponse(error) {
    return JSON.stringify({
        ok: false,
        error: error.message
    });
}

function buildJsonTextOutput(result) {
    var output = ContentService.createTextOutput(
        typeof result === 'string' ? result : JSON.stringify(result)
    );
    output.setMimeType(ContentService.MimeType.JSON);
    return output;
}
