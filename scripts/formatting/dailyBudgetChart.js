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
    var percentage = adjustedBudget ? (totalAmount / adjustedBudget) * 100 : 0;

    if (totalAmount > adjustedBudget) {
        return "【緊急】週予算の" + (totalAmount / adjustedBudget).toFixed(1) + "倍（" + (totalAmount - adjustedBudget) + "円超過）";
    }

    return "今週の実支出 " + totalAmount + "円 / " + adjustedBudget + "円（" + percentage.toFixed(1) + "%）";
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

function getDailyBudgetChartAmounts(totalAmount, adjustedBudget) {
    return {
        withinBudget: Math.min(totalAmount, adjustedBudget),
        overBudget: Math.min(Math.max(totalAmount - adjustedBudget, 0), adjustedBudget / 2),
        criticalOverBudget: Math.max(totalAmount - adjustedBudget * 1.5, 0),
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
        .addColumn(Charts.ColumnType.NUMBER, "超過分（100〜150%）")
        .addColumn(Charts.ColumnType.NUMBER, "超過分（150%超）")
        .addColumn(Charts.ColumnType.NUMBER, "残り")
        .addRow(["今週", chartAmounts.withinBudget, chartAmounts.overBudget, chartAmounts.criticalOverBudget, chartAmounts.remaining])
        .build();
    var chart = Charts.newBarChart()
        .setDataTable(chartData)
        .setDimensions(600, 160)
        .setOption("title", buildDailyBudgetChartTitle(totalAmount, adjustedBudget))
        .setOption("titleTextStyle", { color: totalAmount > adjustedBudget ? "#b42318" : "#202124", fontSize: 15, bold: true })
        .setOption("legend", { position: "none" })
        .setOption("isStacked", true)
        .setOption("colors", [chartTheme.color, chartTheme.overBudgetColor, "#6b1512", "#e8eaed"])
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
