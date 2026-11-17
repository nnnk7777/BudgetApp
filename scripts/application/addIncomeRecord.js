function addIncomeRecord(item) {
    if (!item || typeof item !== 'object' || Array.isArray(item)) {
        throw new Error('itemが定義されていません');
    }

    var title = typeof item.title === 'string' ? item.title.trim() : '';
    var category = typeof item.category === 'string' ? item.category.trim() : '';
    if (!title) {
        throw new Error('titleが不正です');
    }
    if (!category) {
        throw new Error('収入カテゴリを指定してください');
    }
    var amount = normalizeAddIncomeAmount(item.amount);
    var runtimeContext = getScriptRuntimeContext();
    var date = Utilities.formatDate(runtimeContext.currentDate, 'Asia/Tokyo', 'MM/dd');
    var month = Number(date.slice(0, 2)) - 1;
    appendMonthlyIncomeEntry(month, date, category, title, amount);
    return { ok: true };
}

function normalizeAddIncomeAmount(amount) {
    if (typeof amount === 'string') {
        var text = normalizeFullWidthNumbers(amount.trim())
            .replace(/^[¥￥]\s*|\s*円$/g, '');
        if (/^(?:\d+|\d{1,3}(?:,\d{3})+)$/.test(text)) {
            amount = Number(text.replace(/,/g, ''));
        }
    }

    if (!Number.isSafeInteger(amount) || amount <= 0) {
        throw new Error('amountは正の整数で指定してください');
    }
    return amount;
}
