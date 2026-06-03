function getScriptRuntimeContext() {
    var testDateStr = "TEST_DATE_PLACEHOLDER";
    var normalizedTestDateStr = String(testDateStr || "").trim();

    if (!/^\d{8}$/.test(normalizedTestDateStr)) {
        return {
            currentDate: new Date(),
            isStaging: false
        };
    }

    var testDate = parseScriptDateYYYYMMDD(normalizedTestDateStr);
    if (!testDate) {
        throw new Error('Invalid TEST_DATE format. Expected YYYYMMDD.');
    }

    return {
        currentDate: testDate,
        isStaging: true
    };
}

function parseScriptDateYYYYMMDD(dateStr) {
    if (!/^\d{8}$/.test(dateStr)) {
        return null;
    }

    var year = parseInt(dateStr.substring(0, 4), 10);
    var month = parseInt(dateStr.substring(4, 6), 10) - 1;
    var day = parseInt(dateStr.substring(6, 8), 10);
    return new Date(year, month, day);
}
