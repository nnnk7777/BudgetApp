var WEEKLY_BUDGET_CARRYOVER_MEMO_TITLE = "先週の予算差分メモ";
var WEEKLY_BUDGET_CARRYOVER_MEMO_MARKER = "WEEKLY_BUDGET_CARRYOVER_MEMO";
var WEEKLY_ANALYSIS_MODE_NORMAL = "normal";
var WEEKLY_ANALYSIS_MODE_FRUGAL = "frugal";

function getCalendarEventsInRange(startDate, endDate) {
    if (typeof CalendarApp === 'undefined') {
        Logger.log("CalendarApp is unavailable in this runtime.");
        return [];
    }

    var calendar = getTargetCalendar();
    if (!calendar) {
        return [];
    }

    return calendar.getEvents(startDate, endDate);
}

function upsertWeeklyBudgetCarryoverMemo(baseDate, difference, adjustedBudget, totalAmount, dateRangeStr) {
    if (typeof CalendarApp === 'undefined') {
        Logger.log("CalendarApp is unavailable in this runtime.");
        return null;
    }

    var calendar = getTargetCalendar();
    if (!calendar) {
        return null;
    }

    var memoDate = getNextWeekMonday(baseDate);
    var existingEvents = findWeeklyBudgetCarryoverMemoEventsForDate(memoDate);
    var title = buildWeeklyBudgetCarryoverMemoTitle(difference);
    var description = buildWeeklyBudgetCarryoverMemoDescription(
        difference,
        adjustedBudget,
        totalAmount,
        dateRangeStr
    );
    var targetEvent = existingEvents[0] || null;

    if (targetEvent) {
        targetEvent.setTitle(title);
        targetEvent.setDescription(description);
        Logger.log("前週予算差分メモを更新: " + formatDate(memoDate) + " title=" + title);
        return targetEvent;
    }

    targetEvent = calendar.createAllDayEvent(title, memoDate, {
        description: description
    });
    Logger.log("前週予算差分メモを作成: " + formatDate(memoDate) + " title=" + title);
    return targetEvent;
}

function getWeeklyBudgetCarryoverMemoForWeek(baseDate) {
    var weekStartDate = getWeekRange(baseDate).startDate;
    var existingEvents = findWeeklyBudgetCarryoverMemoEventsForDate(weekStartDate);
    var parsedMemo;

    if (!existingEvents.length) {
        return null;
    }

    parsedMemo = parseWeeklyBudgetCarryoverMemoEvent(existingEvents[0]);
    if (!parsedMemo || typeof parsedMemo.difference !== 'number') {
        Logger.log("前週予算差分メモの解析に失敗: " + weekStartDate);
        return null;
    }

    return parsedMemo;
}

function getWeeklyAnalysisMode(baseDate) {
    var weekStartDate = getWeekRange(baseDate).startDate;
    var modeEvents = findWeeklyAnalysisModeEventsForDate(weekStartDate);
    var resolvedMode = createWeeklyAnalysisModeResult(WEEKLY_ANALYSIS_MODE_NORMAL, "通常モード", "default");
    var i;
    var detectedMode;

    for (i = 0; i < modeEvents.length; i++) {
        detectedMode = detectWeeklyAnalysisModeFromEvent(modeEvents[i]);
        if (detectedMode) {
            Logger.log(
                "分析モードを検出: date=" +
                    formatDate(weekStartDate) +
                    " mode=" +
                    detectedMode.mode +
                    " source=" +
                    detectedMode.source
            );
            return detectedMode;
        }
    }

    Logger.log("分析モード指定なし: date=" + formatDate(weekStartDate));
    return resolvedMode;
}

function isWeeklyBudgetCarryoverMemo(title, description) {
    var normalizedTitle = String(title || "");
    var normalizedDescription = String(description || "");
    return (
        normalizedTitle.indexOf(WEEKLY_BUDGET_CARRYOVER_MEMO_TITLE) !== -1 ||
        normalizedDescription.indexOf(WEEKLY_BUDGET_CARRYOVER_MEMO_MARKER) !== -1
    );
}

function isWeeklyAnalysisModeEvent(title, description) {
    if (isWeeklyBudgetCarryoverMemo(title, description)) {
        return false;
    }

    return detectWeeklyAnalysisModeFromText(title, description) === WEEKLY_ANALYSIS_MODE_FRUGAL;
}

function getTargetCalendar() {
    try {
        var calendarId = PropertiesService.getScriptProperties().getProperty("CALENDAR_ID");
        if (calendarId) {
            return CalendarApp.getCalendarById(calendarId);
        }
        return CalendarApp.getDefaultCalendar();
    } catch (error) {
        Logger.log("Failed to access calendar: " + error);
        return null;
    }
}

function getUpcomingExpenseLookaheadDays() {
    try {
        var value = PropertiesService.getScriptProperties().getProperty("UPCOMING_EXPENSE_LOOKAHEAD_DAYS");
        var parsed = parseInt(value, 10);
        if (!isNaN(parsed) && parsed > 0) {
            return parsed;
        }
    } catch (error) {
        Logger.log("Failed to read UPCOMING_EXPENSE_LOOKAHEAD_DAYS: " + error);
    }
    return 14;
}

function findWeeklyBudgetCarryoverMemoEventsForDate(date) {
    if (typeof CalendarApp === 'undefined') {
        Logger.log("CalendarApp is unavailable in this runtime.");
        return [];
    }

    var calendar = getTargetCalendar();
    if (!calendar) {
        return [];
    }

    return calendar.getEventsForDay(date).filter(function (event) {
        return isWeeklyBudgetCarryoverMemo(event.getTitle(), event.getDescription());
    });
}

function findWeeklyAnalysisModeEventsForDate(date) {
    var startDate = new Date(date.getFullYear(), date.getMonth(), date.getDate());
    var endDate = new Date(startDate);
    endDate.setDate(endDate.getDate() + 1);

    return getCalendarEventsInRange(startDate, endDate).filter(function (event) {
        return !isWeeklyBudgetCarryoverMemo(event.getTitle(), event.getDescription());
    });
}

function buildWeeklyBudgetCarryoverMemoTitle(difference) {
    var sign = difference >= 0 ? "+" : "-";
    return WEEKLY_BUDGET_CARRYOVER_MEMO_TITLE + " " + sign + Math.abs(difference) + "円";
}

function buildWeeklyBudgetCarryoverMemoDescription(difference, adjustedBudget, totalAmount, dateRangeStr) {
    return [
        WEEKLY_BUDGET_CARRYOVER_MEMO_MARKER,
        "対象週: " + dateRangeStr,
        "予算差分: " + (difference >= 0 ? "+" : "-") + Math.abs(difference) + "円",
        "実支出: " + totalAmount + "円",
        "週予算: " + adjustedBudget + "円",
        "BUDGET_DIFF_YEN=" + difference,
        "TOTAL_AMOUNT_YEN=" + totalAmount,
        "ADJUSTED_BUDGET_YEN=" + adjustedBudget
    ].join("\n");
}

function parseWeeklyBudgetCarryoverMemoEvent(event) {
    var description = String(event.getDescription() || "");
    var difference = parseWeeklyBudgetCarryoverMemoInteger(description, "BUDGET_DIFF_YEN");
    var totalAmount = parseWeeklyBudgetCarryoverMemoInteger(description, "TOTAL_AMOUNT_YEN");
    var adjustedBudget = parseWeeklyBudgetCarryoverMemoInteger(description, "ADJUSTED_BUDGET_YEN");

    if (difference === null || totalAmount === null || adjustedBudget === null) {
        return null;
    }

    return {
        title: event.getTitle() || "",
        date: event.getStartTime(),
        dateRangeStr: readWeeklyBudgetCarryoverMemoValue(description, "対象週"),
        difference: difference,
        totalAmount: totalAmount,
        adjustedBudget: adjustedBudget
    };
}

function readWeeklyBudgetCarryoverMemoValue(description, key) {
    var pattern = new RegExp("^" + escapeRegExp(key) + "[:=]\\s*(.+)$", "m");
    var match = String(description || "").match(pattern);
    return match ? match[1].trim() : "";
}

function parseWeeklyBudgetCarryoverMemoInteger(description, key) {
    var parsed = parseInt(readWeeklyBudgetCarryoverMemoValue(description, key), 10);
    return isNaN(parsed) ? null : parsed;
}

function getNextWeekMonday(date) {
    var nextMonday = new Date(date);
    nextMonday.setDate(nextMonday.getDate() + 1);
    nextMonday.setHours(0, 0, 0, 0);
    return nextMonday;
}

function detectWeeklyAnalysisModeFromEvent(event) {
    var title = event.getTitle() || "";
    var description = event.getDescription() || "";
    var ruleBasedMode = detectWeeklyAnalysisModeFromText(title, description);
    var geminiMode;

    if (ruleBasedMode) {
        return createWeeklyAnalysisModeResult(ruleBasedMode, "節制モード", "rule");
    }

    geminiMode = detectWeeklyAnalysisModeWithAI(title, description);
    if (geminiMode === WEEKLY_ANALYSIS_MODE_FRUGAL) {
        return createWeeklyAnalysisModeResult(geminiMode, "節制モード", "gemini");
    }

    return null;
}

function detectWeeklyAnalysisModeFromText(title, description) {
    var normalizedCombined = normalizeWeeklyAnalysisModeText([title || "", description || ""].join(" "));
    if (!normalizedCombined) {
        return null;
    }

    if (
        normalizedCombined.indexOf("節制モード") !== -1 ||
        normalizedCombined.indexOf("節制") !== -1
    ) {
        return WEEKLY_ANALYSIS_MODE_FRUGAL;
    }

    return null;
}

function detectWeeklyAnalysisModeWithAI(title, description) {
    var result;
    var normalized;
    var prompt;

    prompt = [
        "以下のGoogleカレンダー予定が、その週の家計分析を厳しめにする『節制モード』指定かどうかを判定してください。",
        "出力は YES か NO のどちらか1語のみです。",
        "節制モードと判断してよい例: 節制モード, 節制, 引き締め週, 節約強化週間。",
        "支出予定、通常の予定、意味が曖昧で節制意図が読めないものは NO にしてください。",
        "タイトル: " + String(title || ""),
        "説明: " + String(description || "")
    ].join("\n");

    result = generatePreferredAiText(prompt, {
        temperature: 0,
        maxOutputTokens: 10
    });
    if (!result.text) {
        return null;
    }

    normalized = String(result.text).trim().toUpperCase();
    Logger.log("分析モードAI判定: title=" + title + " provider=" + result.provider + " response=" + normalized);
    return normalized.indexOf("YES") === 0 ? WEEKLY_ANALYSIS_MODE_FRUGAL : null;
}

function createWeeklyAnalysisModeResult(mode, label, source) {
    return {
        mode: mode,
        label: label,
        source: source
    };
}

function normalizeWeeklyAnalysisModeText(text) {
    return normalizeFullWidthNumbers(String(text || ""))
        .toLowerCase()
        .replace(/\s+/g, "")
        .replace(/[!！?？\-ー_./／\\,，。、「」（）()【】\[\]:'"@#&]/g, "");
}

function escapeRegExp(text) {
    return String(text).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}
