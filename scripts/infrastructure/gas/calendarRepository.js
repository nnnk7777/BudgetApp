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
