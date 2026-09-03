var AUTOMATIC_COMMUTE_EXPENSE_SOURCE = "automatic_commute";

function getAutomaticCommutePlannedExpenses(startDate, endDate, calendarEvents) {
    var holidayEvents = getJapaneseHolidayEventsInRange(startDate, endDate);

    if (holidayEvents === null) {
        Logger.log("日本の祝日カレンダーを取得できないため、通勤費の自動見込みをスキップしました");
        return [];
    }

    return buildAutomaticCommutePlannedExpenses(startDate, endDate, calendarEvents, holidayEvents);
}

function buildAutomaticCommutePlannedExpenses(startDate, endDate, calendarEvents, holidayEvents) {
    var expenses = [];
    var date = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate());
    var end = new Date(endDate.getFullYear(), endDate.getMonth(), endDate.getDate());

    while (date < end) {
        var eventsForDay;
        var hasLeave;
        var hasWorkOverride;

        if (COMMUTE_OFFICE_WEEKDAYS.indexOf(date.getDay()) === -1 || hasJapanesePublicHolidayOnDate(holidayEvents, date)) {
            date.setDate(date.getDate() + 1);
            continue;
        }

        eventsForDay = calendarEvents.filter(function (event) {
            return isCalendarEventOnDate(event, date);
        });
        hasLeave = eventsForDay.some(isCommuteLeaveEvent);
        hasWorkOverride = eventsForDay.some(isHolidayWorkEvent);

        if (!hasLeave || hasWorkOverride) {
            expenses.push({
                title: "通勤費（自動見込み）",
                date: new Date(date.getFullYear(), date.getMonth(), date.getDate()),
                memo: "通勤費 " + COMMUTE_EXPENSE_PER_DAY.toLocaleString("ja-JP") + "円",
                intent: "planned_expense",
                source: AUTOMATIC_COMMUTE_EXPENSE_SOURCE
            });
        }

        date.setDate(date.getDate() + 1);
    }

    return expenses;
}

function hasJapanesePublicHolidayOnDate(events, date) {
    return events.some(function (event) {
        return isCalendarEventOnDate(event, date) && isJapanesePublicHolidayEvent(event);
    });
}

function isJapanesePublicHolidayEvent(event) {
    var title = String(event.getTitle() || "").replace(/\s+/g, "");

    return /^(元日|成人の日|建国記念の日|天皇誕生日|春分の日|昭和の日|憲法記念日|みどりの日|こどもの日|海の日|山の日|敬老の日|秋分の日|スポーツの日|文化の日|勤労感謝の日|振替休日|国民の休日)$/.test(title);
}

function isCalendarEventOnDate(event, date) {
    var dayStart = new Date(date.getFullYear(), date.getMonth(), date.getDate());
    var dayEnd = new Date(dayStart);
    var eventStart = event.getStartTime();
    var eventEnd = event.getEndTime();

    dayEnd.setDate(dayEnd.getDate() + 1);
    return eventStart < dayEnd && eventEnd > dayStart;
}

function isCommuteLeaveEvent(event) {
    var text = getCalendarEventSearchText(event);

    return /(全休|終日休|一日休|1日休|有給休暇|有休|有給|代休|振休|公休|午前休|午後休|半日|半休|時間休)/.test(text);
}

function isHolidayWorkEvent(event) {
    return /(休日出勤|休日勤務|休出)/.test(getCalendarEventSearchText(event));
}

function getCalendarEventSearchText(event) {
    return [event.getTitle() || "", event.getDescription() || ""].join(" ").replace(/\s+/g, "");
}

function adjustAutomaticCommuteExpenseForRecordedEntries(plannedExpense, actualEntries) {
    var recordedAmount;
    var remainingAmount;

    if (plannedExpense.source !== AUTOMATIC_COMMUTE_EXPENSE_SOURCE) {
        return plannedExpense;
    }

    recordedAmount = actualEntries.filter(function (entry) {
        return calculateDateDistanceInDays(plannedExpense.date, entry.date) === 0 && isRecordedCommuteExpense(entry);
    }).reduce(function (total, entry) {
        var amount = parseFloat(entry.amount);
        return total + (isNaN(amount) ? 0 : amount);
    }, 0);
    remainingAmount = Math.max(COMMUTE_EXPENSE_PER_DAY - recordedAmount, 0);

    if (remainingAmount === 0) {
        return null;
    }

    if (remainingAmount === COMMUTE_EXPENSE_PER_DAY) {
        return plannedExpense;
    }

    return {
        title: plannedExpense.title,
        date: plannedExpense.date,
        memo: "通勤費 残り見込み " + remainingAmount.toLocaleString("ja-JP") + "円",
        intent: plannedExpense.intent,
        source: plannedExpense.source
    };
}

function isRecordedCommuteExpense(entry) {
    var category = String(entry.category || "").trim();
    var name = String(entry.name || "");

    return category === "交通費" && name.indexOf("通勤") !== -1;
}
