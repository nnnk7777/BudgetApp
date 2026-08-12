function getDailyBudgetChartTheme(totalAmount, adjustedBudget) {
    var percentage = adjustedBudget ? (totalAmount / adjustedBudget) * 100 : 0;

    if (percentage > 100) {
        return { color: "#b42318", overBudgetColor: "#8c1d18", label: "超過" };
    }

    if (percentage >= 90) {
        return { color: "#c62828", overBudgetColor: "#c62828", label: "限界" };
    }

    if (percentage >= 80) {
        return { color: "#e53935", overBudgetColor: "#e53935", label: "危険" };
    }

    if (percentage >= 60) {
        return { color: "#f4511e", overBudgetColor: "#f4511e", label: "警戒" };
    }

    if (percentage >= 40) {
        return { color: "#f9ab00", overBudgetColor: "#f9ab00", label: "注意" };
    }

    return { color: "#188038", overBudgetColor: "#188038", label: "予算内" };
}

function buildDailyBudgetChartTitle(totalAmount, adjustedBudget, specialExpenseTotal) {
    return buildBudgetChartTitleWithSpecialExpense(totalAmount, adjustedBudget, "今週の", "週予算", specialExpenseTotal);
}

function buildWeeklySummaryMonthlyChartTitle(monthlyTotalAmount, monthlyBudget, currentDate, specialExpenseTotal) {
    var title = buildBudgetChartTitleWithSpecialExpense(monthlyTotalAmount, monthlyBudget, "今月の", "月次予算", specialExpenseTotal);

    if (!currentDate) {
        return title;
    }

    return title + "　｜　" + formatBudgetChartDate(currentDate) + "時点の目安 " + getMonthlyBudgetPacePercentage(currentDate).toFixed(1) + "%";
}

function buildWeeklySummaryWeeklyChartTitle(weeklyTotalAmount, weeklyBudget, specialExpenseTotal) {
    return buildBudgetChartTitleWithSpecialExpense(weeklyTotalAmount, weeklyBudget, "今週の", "週予算", specialExpenseTotal);
}

function formatBudgetChartDate(date) {
    return (date.getMonth() + 1) + "/" + date.getDate();
}

function buildBudgetChartTitle(totalAmount, adjustedBudget, periodPrefix, budgetLabel, expenseLabel) {
    var percentage = adjustedBudget ? (totalAmount / adjustedBudget) * 100 : 0;
    var resolvedExpenseLabel = expenseLabel || "実支出";

    if (totalAmount > adjustedBudget) {
        return "【緊急】" + budgetLabel + "の" + (totalAmount / adjustedBudget).toFixed(1) + "倍（" + (totalAmount - adjustedBudget) + "円超過）";
    }

    return periodPrefix + resolvedExpenseLabel + " " + totalAmount + "円 / " + adjustedBudget + "円（" + percentage.toFixed(1) + "%）";
}

function buildBudgetChartTitleWithSpecialExpense(totalAmount, adjustedBudget, periodPrefix, budgetLabel, specialExpenseTotal) {
    var title = buildBudgetChartTitle(totalAmount, adjustedBudget, periodPrefix, budgetLabel);
    var normalExpenseTotal = Math.max(totalAmount - specialExpenseTotal, 0);
    var normalPercentage = adjustedBudget ? normalExpenseTotal / adjustedBudget * 100 : 0;
    var specialPercentage = adjustedBudget ? specialExpenseTotal / adjustedBudget * 100 : 0;

    if (!specialExpenseTotal) {
        return title;
    }

    return title + "\n通常 " + normalExpenseTotal + "円（" + normalPercentage.toFixed(1) + "%）｜ 特別費 " + specialExpenseTotal + "円（" + specialPercentage.toFixed(1) + "%）";
}

function getDailyBudgetChartScalePercentage(totalAmount, adjustedBudget) {
    var percentage = adjustedBudget ? (totalAmount / adjustedBudget) * 100 : 0;

    return Math.max(100, Math.ceil(percentage / 25) * 25);
}

function getDailyBudgetChartTickStep(scalePercentage) {
    if (scalePercentage > 300) {
        return 100;
    }

    if (scalePercentage > 200) {
        return 50;
    }

    return 25;
}

function getDailyBudgetChartTicks(adjustedBudget, scalePercentage) {
    var percentages = [];
    var percentage;
    var tickStep = getDailyBudgetChartTickStep(scalePercentage);

    for (percentage = 0; percentage <= scalePercentage; percentage += tickStep) {
        percentages.push(percentage);
    }

    return percentages.map(function (percentage) {
        return {
            v: adjustedBudget * percentage / 100,
            f: percentage === 100 && scalePercentage > 200 ? "100% 予算上限" : percentage + "%"
        };
    });
}

function getMonthlyBudgetPacePercentage(currentDate) {
    var daysInMonth = new Date(currentDate.getFullYear(), currentDate.getMonth() + 1, 0).getDate();
    return currentDate.getDate() / daysInMonth * 100;
}

function getMonthlyBudgetChartTicks(adjustedBudget, scalePercentage) {
    var percentages = [0, 25, 50, 75, 100];
    var tickStep = getDailyBudgetChartTickStep(scalePercentage);
    var percentage;

    for (percentage = 100 + tickStep; percentage <= scalePercentage; percentage += tickStep) {
        percentages.push(percentage);
    }

    return percentages.map(function (percentage) {
        var labels = {
            0: "0%",
            25: "25%",
            50: "50%",
            75: "75%",
            100: "100% 予算上限"
        };
        return {
            v: adjustedBudget * percentage / 100,
            f: labels[percentage] || percentage + "%"
        };
    });
}

function getWeeklySummaryBudgetChartScalePercentage(monthlyTotalAmount, monthlyBudget, weeklyTotalAmount, weeklyBudget) {
    var monthlyPercentage = monthlyBudget ? monthlyTotalAmount / monthlyBudget * 100 : 0;
    var weeklyPercentage = weeklyBudget ? weeklyTotalAmount / weeklyBudget * 100 : 0;

    return Math.max(
        getDailyBudgetChartScalePercentage(monthlyPercentage, 100),
        getDailyBudgetChartScalePercentage(weeklyPercentage, 100)
    );
}

function getDailyBudgetChartAmounts(totalAmount, adjustedBudget) {
    return {
        withinBudget: Math.min(totalAmount, adjustedBudget),
        overBudget: Math.min(Math.max(totalAmount - adjustedBudget, 0), adjustedBudget / 2),
        criticalOverBudget: Math.max(totalAmount - adjustedBudget * 1.5, 0),
        remaining: Math.max(adjustedBudget - totalAmount, 0)
    };
}

function createDailyBudgetChartBlob(totalAmount, adjustedBudget, specialExpenseTotal) {
    return createBudgetChartBlob(
        totalAmount,
        adjustedBudget,
        "今週",
        buildDailyBudgetChartTitle(totalAmount, adjustedBudget, specialExpenseTotal),
        "daily-budget-chart.png",
        { specialExpenseTotal: specialExpenseTotal || 0 }
    );
}

function createWeeklySummaryMonthlyBudgetChartBlob(monthlyTotalAmount, monthlyBudget, currentDate, specialExpenseTotal, scalePercentage) {
    return createWeeklySummaryPeriodBudgetChartBlob(
        monthlyTotalAmount,
        monthlyBudget,
        "今月",
        buildWeeklySummaryMonthlyChartTitle(monthlyTotalAmount, monthlyBudget, currentDate, specialExpenseTotal),
        "monthly-budget-chart.png",
        specialExpenseTotal,
        scalePercentage
    );
}

function createWeeklySummaryWeeklyBudgetChartBlob(weeklyTotalAmount, weeklyBudget, specialExpenseTotal, scalePercentage) {
    return createWeeklySummaryPeriodBudgetChartBlob(
        weeklyTotalAmount,
        weeklyBudget,
        "今週",
        buildWeeklySummaryWeeklyChartTitle(weeklyTotalAmount, weeklyBudget, specialExpenseTotal),
        "weekly-budget-chart.png",
        specialExpenseTotal,
        scalePercentage
    );
}

function createWeeklySummaryPeriodBudgetChartBlob(totalAmount, adjustedBudget, chartLabel, chartTitle, fileName, specialExpenseTotal, scalePercentage) {
    return createBudgetChartBlob(
        totalAmount,
        adjustedBudget,
        chartLabel,
        chartTitle,
        fileName,
        {
            height: specialExpenseTotal ? 195 : 175,
            chartArea: { left: 70, top: specialExpenseTotal ? 66 : 48, width: "82%", height: "40%" },
            ticks: getMonthlyBudgetChartTicks(adjustedBudget, scalePercentage),
            specialExpenseTotal: specialExpenseTotal
        }
    );
}

function createWeeklySummaryBudgetCharts(monthlyTotalAmount, monthlyBudget, currentDate, weeklyTotalAmount, weeklyBudget, monthlySpecialExpenseTotal, weeklySpecialExpenseTotal) {
    var scalePercentage = getWeeklySummaryBudgetChartScalePercentage(monthlyTotalAmount, monthlyBudget, weeklyTotalAmount, weeklyBudget);

    return {
        monthlyBudgetChart: createWeeklySummaryMonthlyBudgetChartBlob(monthlyTotalAmount, monthlyBudget, currentDate, monthlySpecialExpenseTotal, scalePercentage),
        weeklyBudgetChart: createWeeklySummaryWeeklyBudgetChartBlob(weeklyTotalAmount, weeklyBudget, weeklySpecialExpenseTotal, scalePercentage)
    };
}

function createBudgetChartBlob(totalAmount, adjustedBudget, chartLabel, chartTitle, fileName, displayOptions) {
    var options = displayOptions || {};
    var specialExpenseTotal = options.specialExpenseTotal || 0;
    var normalExpenseTotal = Math.max(totalAmount - specialExpenseTotal, 0);
    var chartTheme = getDailyBudgetChartTheme(normalExpenseTotal, adjustedBudget);
    var chartAmounts = getDailyBudgetChartAmounts(normalExpenseTotal, adjustedBudget);
    var scalePercentage = getDailyBudgetChartScalePercentage(totalAmount, adjustedBudget);
    var chartMaximum = adjustedBudget * scalePercentage / 100;
    var chartData = Charts.newDataTable()
        .addColumn(Charts.ColumnType.STRING, "週予算")
        .addColumn(Charts.ColumnType.NUMBER, "通常支出（予算内）")
        .addColumn(Charts.ColumnType.NUMBER, "通常支出（100〜150%）")
        .addColumn(Charts.ColumnType.NUMBER, "通常支出（150%超）")
        .addColumn(Charts.ColumnType.NUMBER, "承認済み特別費")
        .addColumn(Charts.ColumnType.NUMBER, "残り")
        .addRow([chartLabel, chartAmounts.withinBudget, chartAmounts.overBudget, chartAmounts.criticalOverBudget, specialExpenseTotal, Math.max(adjustedBudget - totalAmount, 0)])
        .build();
    var chart = Charts.newBarChart()
        .setDataTable(chartData)
        .setDimensions(600, options.height || 160)
        .setOption("title", chartTitle)
        .setOption("titleTextStyle", { color: totalAmount > adjustedBudget ? "#b42318" : "#202124", fontSize: 15, bold: true })
        .setOption("legend", { position: "none" })
        .setOption("isStacked", true)
        .setOption("colors", [chartTheme.color, chartTheme.overBudgetColor, "#6b1512", "#9e88f7", "#e8eaed"])
        .setOption("hAxis", {
            viewWindow: { min: 0, max: chartMaximum },
            ticks: options.ticks || getDailyBudgetChartTicks(adjustedBudget, scalePercentage),
            gridlines: { color: "#dadce0" },
            baselineColor: "#dadce0"
        })
        .setOption("chartArea", options.chartArea || { left: 70, top: 44, width: "82%", height: "48%" })
        .build();

    return chart.getAs("image/png").setName(fileName);
}

function buildDailySummaryHtmlBody(body, totalAmount, adjustedBudget, specialExpenseTotal) {
    var chartTitle = buildDailyBudgetChartTitle(totalAmount, adjustedBudget, specialExpenseTotal || 0);

    return buildSummaryHtmlBody(body, [{ key: "dailyBudgetChart", title: chartTitle }]);
}

function buildWeeklySummaryHtmlBody(body, monthlyTotalAmount, monthlyBudget, currentDate, weeklyTotalAmount, weeklyBudget, monthlySpecialExpenseTotal, weeklySpecialExpenseTotal) {
    return buildSummaryHtmlBody(body, [
        {
            key: "monthlyBudgetChart",
            title: buildWeeklySummaryMonthlyChartTitle(monthlyTotalAmount, monthlyBudget, currentDate, monthlySpecialExpenseTotal),
            marginBottom: 4
        },
        {
            key: "weeklyBudgetChart",
            title: buildWeeklySummaryWeeklyChartTitle(weeklyTotalAmount, weeklyBudget, weeklySpecialExpenseTotal)
        }
    ]);
}

function buildSummaryHtmlBody(body, charts) {
    var chartHtml = charts.map(function (chart) {
        return (
            '<img src="cid:' + chart.key + '" alt="' +
            escapeHtmlForEmail(chart.title) +
            '" width="600" style="display:block;max-width:100%;height:auto;margin:0 0 ' + (chart.marginBottom === undefined ? 16 : chart.marginBottom) + 'px;">'
        );
    }).join("");

    return (
        chartHtml +
        '<pre style="font-family:monospace;white-space:pre-wrap;line-height:1.5;">' +
        escapeHtmlForEmail(body) +
        "</pre>"
    );
}

function escapeHtmlForEmail(value) {
    return String(value)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/\"/g, "&quot;")
        .replace(/'/g, "&#39;");
}
