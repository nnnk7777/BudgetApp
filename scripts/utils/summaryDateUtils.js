function formatDate(date) {
    var month = date.getMonth() + 1;
    var day = date.getDate();
    return month + "/" + day;
}

function getDatesInWeek(date) {
    var dates = [];
    var currentYear = date.getFullYear();

    var day = date.getDay();
    var diff = date.getDate() - day + (day === 0 ? -6 : 1);
    var monday = new Date(date);
    monday.setDate(diff);
    monday.setHours(0, 0, 0, 0);

    for (var i = 0; i < 7; i++) {
        var d = new Date(monday);
        d.setDate(monday.getDate() + i);
        d.setHours(0, 0, 0, 0);

        if (d.getFullYear() === currentYear) {
            dates.push(new Date(d.getFullYear(), d.getMonth(), d.getDate()));
        }
    }
    return dates;
}

function getWeekRange(date) {
    var datesInWeek = getDatesInWeek(date);
    return {
        startDate: datesInWeek[0],
        endDate: datesInWeek[datesInWeek.length - 1]
    };
}

function parseDate(dateStr, year) {
    var dateParts = dateStr.split('/');
    var month = parseInt(dateParts[0], 10) - 1;
    var day = parseInt(dateParts[1], 10);
    return new Date(year, month, day);
}

function toMonthDayKey(value) {
    if (Object.prototype.toString.call(value) === '[object Date]') {
        return ('0' + (value.getMonth() + 1)).slice(-2) + '/' + ('0' + value.getDate()).slice(-2);
    }

    var normalized = normalizeFullWidthNumbers(String(value)).trim().replace(/^'+/, '');
    var match = normalized.match(/(\d{1,2})\/(\d{1,2})/);
    if (!match) {
        return null;
    }

    return ('0' + parseInt(match[1], 10)).slice(-2) + '/' + ('0' + parseInt(match[2], 10)).slice(-2);
}

function monthDayKeyToDate(key, year) {
    var parts = key.split('/');
    return new Date(year, parseInt(parts[0], 10) - 1, parseInt(parts[1], 10));
}

function normalizeFullWidthNumbers(text) {
    return text
        .replace(/[０-９]/g, function (char) {
            return String.fromCharCode(char.charCodeAt(0) - 0xFEE0);
        })
        .replace(/，/g, ",");
}
