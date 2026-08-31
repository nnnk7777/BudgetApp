var WEEKLY_BUDGET = 20000;

function getWeeklyBudget() {
    return WEEKLY_BUDGET;
}

function calculateBudgetForDays(days) {
    return Math.round((getWeeklyBudget() * days / 7) / 100) * 100;
}

function calculateMonthlyBudgetForDate(date) {
    var daysInMonth = new Date(date.getFullYear(), date.getMonth() + 1, 0).getDate();
    return calculateBudgetForDays(daysInMonth);
}
