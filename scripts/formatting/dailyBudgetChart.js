function getDailyBudgetChartTheme(totalAmount, adjustedBudget) {
    var percentage = adjustedBudget ? (totalAmount / adjustedBudget) * 100 : 0;

    if (percentage > 100) {
        return { color: "#d93025", label: "超過" };
    }

    if (percentage >= 70) {
        return { color: "#f9ab00", label: "注意" };
    }

    return { color: "#188038", label: "予算内" };
}

function buildDailyBudgetChartTitle(totalAmount, adjustedBudget) {
    var percentage = adjustedBudget ? (totalAmount / adjustedBudget) * 100 : 0;
    var title = "今週の実支出 " + totalAmount + "円 / " + adjustedBudget + "円（" + percentage.toFixed(1) + "%）";

    if (totalAmount > adjustedBudget) {
        title += "  " + (totalAmount - adjustedBudget) + "円超過";
    }

    return title;
}

function getDailyBudgetChartScalePercentage(totalAmount, adjustedBudget) {
    var percentage = adjustedBudget ? (totalAmount / adjustedBudget) * 100 : 0;

    return Math.max(100, Math.ceil(percentage / 25) * 25);
}

function getDailyBudgetChartTicks(adjustedBudget, scalePercentage) {
    var percentages = [];
    var percentage;

    for (percentage = 0; percentage <= scalePercentage; percentage += 25) {
        percentages.push(percentage);
    }

    return percentages.map(function (percentage) {
        return {
            v: adjustedBudget * percentage / 100,
            f: percentage + "%"
        };
    });
}

function getDailyBudgetChartAmounts(totalAmount, adjustedBudget) {
    return {
        withinBudget: Math.min(totalAmount, adjustedBudget),
        overBudget: Math.max(totalAmount - adjustedBudget, 0),
        remaining: Math.max(adjustedBudget - totalAmount, 0)
    };
}

function createDailyBudgetChartBlob(totalAmount, adjustedBudget) {
    var chartTheme = getDailyBudgetChartTheme(totalAmount, adjustedBudget);
    var chartAmounts = getDailyBudgetChartAmounts(totalAmount, adjustedBudget);
    var scalePercentage = getDailyBudgetChartScalePercentage(totalAmount, adjustedBudget);
    var chartMaximum = adjustedBudget * scalePercentage / 100;
    var chartData = Charts.newDataTable()
        .addColumn(Charts.ColumnType.STRING, "週予算")
        .addColumn(Charts.ColumnType.NUMBER, "予算内の支出")
        .addColumn(Charts.ColumnType.NUMBER, "超過分")
        .addColumn(Charts.ColumnType.NUMBER, "残り")
        .addRow(["今週", chartAmounts.withinBudget, chartAmounts.overBudget, chartAmounts.remaining])
        .build();
    var chart = Charts.newBarChart()
        .setDataTable(chartData)
        .setDimensions(600, 160)
        .setOption("title", buildDailyBudgetChartTitle(totalAmount, adjustedBudget))
        .setOption("titleTextStyle", { color: "#202124", fontSize: 15, bold: true })
        .setOption("legend", { position: "none" })
        .setOption("isStacked", true)
        .setOption("colors", [chartTheme.color, "#b3261e", "#e8eaed"])
        .setOption("hAxis", {
            viewWindow: { min: 0, max: chartMaximum },
            ticks: getDailyBudgetChartTicks(adjustedBudget, scalePercentage),
            gridlines: { color: "#dadce0" },
            baselineColor: "#dadce0"
        })
        .setOption("chartArea", { left: 70, top: 44, width: "82%", height: "48%" })
        .build();

    return chart.getAs("image/png").setName("daily-budget-chart.png");
}

function buildDailySummaryHtmlBody(body, totalAmount, adjustedBudget) {
    var chartTitle = buildDailyBudgetChartTitle(totalAmount, adjustedBudget);

    return (
        '<img src="cid:dailyBudgetChart" alt="' +
        escapeHtmlForEmail(chartTitle) +
        '" width="600" style="display:block;max-width:100%;height:auto;margin:0 0 16px;">' +
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
