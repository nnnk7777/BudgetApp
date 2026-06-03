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

function getTargetEmailAddress() {
    var fallbackEmailPlaceholder = "TARGET_EMAIL_ADDRESS_PLACEHOLDER";
    var emailAddress = fallbackEmailPlaceholder;

    if (emailAddress === fallbackEmailPlaceholder && typeof PropertiesService !== 'undefined') {
        emailAddress = PropertiesService.getScriptProperties().getProperty("TARGET_EMAIL_ADDRESS");
    }

    if (!emailAddress) {
        emailAddress = fallbackEmailPlaceholder;
    }

    emailAddress = String(emailAddress || "").trim();

    if (!isValidEmailAddress(emailAddress)) {
        throw new Error('無効なメール: ' + emailAddress);
    }

    return emailAddress;
}

function isValidEmailAddress(emailAddress) {
    var normalizedEmailAddress = String(emailAddress || "").trim();
    var atIndex = normalizedEmailAddress.indexOf("@");
    var lastDotIndex = normalizedEmailAddress.lastIndexOf(".");

    return (
        atIndex > 0 &&
        lastDotIndex > atIndex + 1 &&
        lastDotIndex < normalizedEmailAddress.length - 1 &&
        normalizedEmailAddress.indexOf(" ") === -1
    );
}

function getScriptRuntimeDiagnostics() {
    var runtimeContext = getScriptRuntimeContext();
    var activeSpreadsheet = null;
    var diagnostics = {
        currentDate: runtimeContext.currentDate.toISOString(),
        isStaging: runtimeContext.isStaging
    };

    if (typeof PropertiesService !== 'undefined') {
        diagnostics.scriptPropertyTargetEmailAddress =
            PropertiesService.getScriptProperties().getProperty("TARGET_EMAIL_ADDRESS") || null;
    }

    try {
        diagnostics.resolvedTargetEmailAddress = getTargetEmailAddress();
    } catch (error) {
        diagnostics.resolvedTargetEmailAddressError = String(error);
    }

    if (typeof SpreadsheetApp !== 'undefined') {
        activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    }

    if (activeSpreadsheet) {
        diagnostics.activeSpreadsheetId = activeSpreadsheet.getId();
        diagnostics.activeSpreadsheetName = activeSpreadsheet.getName();
    }

    return diagnostics;
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
