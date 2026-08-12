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

function buildDailyBudgetChartTitle(totalAmount, adjustedBudget) {
    return buildBudgetChartTitle(totalAmount, adjustedBudget, "今週の", "週予算");
}

function buildWeeklySummaryBudgetChartTitle(monthlyTotalAmount, monthlyBudget, currentDate, weeklyTotalAmount, weeklyBudget) {
    var monthlyPercentage = monthlyBudget ? monthlyTotalAmount / monthlyBudget * 100 : 0;
    var weeklyPercentage = weeklyBudget ? weeklyTotalAmount / weeklyBudget * 100 : 0;
    var monthlyStatus = monthlyTotalAmount > monthlyBudget ? "・" + (monthlyTotalAmount - monthlyBudget) + "円超過" : "";
    var weeklyStatus = weeklyTotalAmount > weeklyBudget ? "・" + (weeklyTotalAmount - weeklyBudget) + "円超過" : "";
    var monthlyPace = currentDate ? "　｜　" + formatBudgetChartDate(currentDate) + "時点の目安 " + getMonthlyBudgetPacePercentage(currentDate).toFixed(1) + "%" : "";

    return (
        "今月 " + monthlyTotalAmount + "円 / " + monthlyBudget + "円（" + monthlyPercentage.toFixed(1) + "%" + monthlyStatus + "）" + monthlyPace + "\n" +
        "今週 " + weeklyTotalAmount + "円 / " + weeklyBudget + "円（" + weeklyPercentage.toFixed(1) + "%" + weeklyStatus + "）"
    );
}

function formatBudgetChartDate(date) {
    return (date.getMonth() + 1) + "/" + date.getDate();
}

function buildBudgetChartTitle(totalAmount, adjustedBudget, periodPrefix, budgetLabel) {
    var percentage = adjustedBudget ? (totalAmount / adjustedBudget) * 100 : 0;

    if (totalAmount > adjustedBudget) {
        return "【緊急】" + budgetLabel + "の" + (totalAmount / adjustedBudget).toFixed(1) + "倍（" + (totalAmount - adjustedBudget) + "円超過）";
    }

    return periodPrefix + "実支出 " + totalAmount + "円 / " + adjustedBudget + "円（" + percentage.toFixed(1) + "%）";
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
            25: "25% 予算の1/4",
            50: "50% 予算の1/2",
            75: "75% 予算の3/4",
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

function createDailyBudgetChartBlob(totalAmount, adjustedBudget) {
    return createBudgetChartBlob(totalAmount, adjustedBudget, "今週", buildDailyBudgetChartTitle(totalAmount, adjustedBudget), "daily-budget-chart.png");
}

function createWeeklySummaryBudgetChartBlob(monthlyTotalAmount, monthlyBudget, currentDate, weeklyTotalAmount, weeklyBudget) {
    var monthlyPercentage = monthlyBudget ? monthlyTotalAmount / monthlyBudget * 100 : 0;
    var weeklyPercentage = weeklyBudget ? weeklyTotalAmount / weeklyBudget * 100 : 0;
    var scalePercentage = getWeeklySummaryBudgetChartScalePercentage(monthlyTotalAmount, monthlyBudget, weeklyTotalAmount, weeklyBudget);
    var monthlyTheme = getDailyBudgetChartTheme(monthlyPercentage, 100);
    var weeklyTheme = getDailyBudgetChartTheme(weeklyPercentage, 100);
    var monthlyAmounts = getDailyBudgetChartAmounts(monthlyPercentage, 100);
    var weeklyAmounts = getDailyBudgetChartAmounts(weeklyPercentage, 100);
    var chartData = Charts.newDataTable()
        .addColumn(Charts.ColumnType.STRING, "予算")
        .addColumn(Charts.ColumnType.NUMBER, "今月の予算内支出")
        .addColumn(Charts.ColumnType.NUMBER, "今月の超過分（100〜150%）")
        .addColumn(Charts.ColumnType.NUMBER, "今月の超過分（150%超）")
        .addColumn(Charts.ColumnType.NUMBER, "今月の残り")
        .addColumn(Charts.ColumnType.NUMBER, "今週の予算内支出")
        .addColumn(Charts.ColumnType.NUMBER, "今週の超過分（100〜150%）")
        .addColumn(Charts.ColumnType.NUMBER, "今週の超過分（150%超）")
        .addColumn(Charts.ColumnType.NUMBER, "今週の残り")
        .addRow(["今月", monthlyAmounts.withinBudget, monthlyAmounts.overBudget, monthlyAmounts.criticalOverBudget, monthlyAmounts.remaining, 0, 0, 0, 0])
        .addRow(["今週", 0, 0, 0, 0, weeklyAmounts.withinBudget, weeklyAmounts.overBudget, weeklyAmounts.criticalOverBudget, weeklyAmounts.remaining])
        .build();
    var chart = Charts.newBarChart()
        .setDataTable(chartData)
        .setDimensions(600, 260)
        .setOption("title", buildWeeklySummaryBudgetChartTitle(monthlyTotalAmount, monthlyBudget, currentDate, weeklyTotalAmount, weeklyBudget))
        .setOption("titleTextStyle", {
            color: monthlyTotalAmount > monthlyBudget || weeklyTotalAmount > weeklyBudget ? "#b42318" : "#202124",
            fontSize: 14,
            bold: true
        })
        .setOption("legend", { position: "none" })
        .setOption("isStacked", true)
        .setOption("colors", [monthlyTheme.color, monthlyTheme.overBudgetColor, "#6b1512", "#e8eaed", weeklyTheme.color, weeklyTheme.overBudgetColor, "#6b1512", "#e8eaed"])
        .setOption("hAxis", {
            viewWindow: { min: 0, max: scalePercentage },
            ticks: getMonthlyBudgetChartTicks(100, scalePercentage),
            gridlines: { color: "#dadce0" },
            baselineColor: "#dadce0"
        })
        .setOption("chartArea", { left: 70, top: 76, width: "82%", height: "48%" })
        .build();

    return chart.getAs("image/png").setName("weekly-summary-budget-chart.png");
}

function createBudgetChartBlob(totalAmount, adjustedBudget, chartLabel, chartTitle, fileName, displayOptions) {
    var chartTheme = getDailyBudgetChartTheme(totalAmount, adjustedBudget);
    var chartAmounts = getDailyBudgetChartAmounts(totalAmount, adjustedBudget);
    var scalePercentage = getDailyBudgetChartScalePercentage(totalAmount, adjustedBudget);
    var chartMaximum = adjustedBudget * scalePercentage / 100;
    var options = displayOptions || {};
    var chartData = Charts.newDataTable()
        .addColumn(Charts.ColumnType.STRING, "週予算")
        .addColumn(Charts.ColumnType.NUMBER, "予算内の支出")
        .addColumn(Charts.ColumnType.NUMBER, "超過分（100〜150%）")
        .addColumn(Charts.ColumnType.NUMBER, "超過分（150%超）")
        .addColumn(Charts.ColumnType.NUMBER, "残り")
        .addRow([chartLabel, chartAmounts.withinBudget, chartAmounts.overBudget, chartAmounts.criticalOverBudget, chartAmounts.remaining])
        .build();
    var chart = Charts.newBarChart()
        .setDataTable(chartData)
        .setDimensions(600, options.height || 160)
        .setOption("title", chartTitle)
        .setOption("titleTextStyle", { color: totalAmount > adjustedBudget ? "#b42318" : "#202124", fontSize: 15, bold: true })
        .setOption("legend", { position: "none" })
        .setOption("isStacked", true)
        .setOption("colors", [chartTheme.color, chartTheme.overBudgetColor, "#6b1512", "#e8eaed"])
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

function buildDailySummaryHtmlBody(body, totalAmount, adjustedBudget) {
    var chartTitle = buildDailyBudgetChartTitle(totalAmount, adjustedBudget);

    return buildSummaryHtmlBody(body, [{ key: "dailyBudgetChart", title: chartTitle }]);
}

function buildWeeklySummaryHtmlBody(body, monthlyTotalAmount, monthlyBudget, currentDate, weeklyTotalAmount, weeklyBudget) {
    return buildSummaryHtmlBody(body, [{
        key: "weeklySummaryBudgetChart",
        title: buildWeeklySummaryBudgetChartTitle(monthlyTotalAmount, monthlyBudget, currentDate, weeklyTotalAmount, weeklyBudget)
    }]);
}

function buildSummaryHtmlBody(body, charts) {
    var chartHtml = charts.map(function (chart) {
        return (
            '<img src="cid:' + chart.key + '" alt="' +
            escapeHtmlForEmail(chart.title) +
            '" width="600" style="display:block;max-width:100%;height:auto;margin:0 0 16px;">'
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
