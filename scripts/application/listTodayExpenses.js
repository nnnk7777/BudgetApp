function listTodayExpenses() {
    var currentDate = getScriptRuntimeContext().currentDate;
    var items = getExpenseEntriesForDates([currentDate]).map(function (entry) {
        return {
            category: entry.category || "未分類",
            name: entry.name || "",
            amount: entry.amount || 0
        };
    });

    return JSON.stringify({
        ok: true,
        items: items
    });
}
