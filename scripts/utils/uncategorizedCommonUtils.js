function groupItemsByMonth(items) {
    var grouped = {};
    items.forEach(function (item) {
        var key = String(item.month);
        if (!grouped[key]) {
            grouped[key] = [];
        }
        grouped[key].push(item);
    });
    return grouped;
}

function buildExpenseRecordId(month, row) {
    return month + "_" + row;
}

function normalizeConfidence(value) {
    var parsed = parseFloat(value);
    if (isNaN(parsed)) {
        return 0;
    }
    if (parsed < 0) {
        return 0;
    }
    if (parsed > 1) {
        return 1;
    }
    return parsed;
}

function normalizeConfidenceThreshold(value) {
    if (value === undefined || value === null || value === '') {
        return 0.9;
    }
    return normalizeConfidence(value);
}

function normalizeLimit(value) {
    if (value === undefined || value === null || value === '') {
        return null;
    }

    var parsed = parseInt(value, 10);
    if (isNaN(parsed) || parsed <= 0) {
        return null;
    }

    return parsed;
}

function normalizeAmountValue(value) {
    var normalized = parseFloat(String(normalizeFullWidthNumbers(String(value))).replace(/,/g, ""));
    return isNaN(normalized) ? value : normalized;
}

function isBlankCell(value) {
    return value === null || value === undefined || String(value).trim() === '';
}

function isSameAmount(left, right) {
    var normalizedLeft = normalizeAmountValue(left);
    var normalizedRight = normalizeAmountValue(right);
    return normalizedLeft === normalizedRight;
}

function mergeObjects(base, additions) {
    var result = {};
    Object.keys(base).forEach(function (key) {
        result[key] = base[key];
    });
    Object.keys(additions).forEach(function (key) {
        result[key] = additions[key];
    });
    return result;
}
