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

function buildMonthlySummaryBudgetChartTitle(totalAmount, adjustedBudget, specialExpenseTotal) {
    return buildBudgetChartTitleWithSpecialExpense(totalAmount, adjustedBudget, "今月の", "月次予算", specialExpenseTotal);
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

function getBudgetChartScalePercentage(totalAmount, adjustedBudget, displayOptions) {
    return (displayOptions && displayOptions.scalePercentage) || getDailyBudgetChartScalePercentage(totalAmount, adjustedBudget);
}

function getBudgetChartMarkerAmount(adjustedBudget, scalePercentage) {
    // グラフ領域は 600px × 82% = 約492px。背景のグリッド線と同程度の1px幅にする。
    return adjustedBudget * scalePercentage / 100 / 492;
}

function getBudgetChartStackSegments(totalAmount, normalExpenseTotal, specialExpenseTotal, adjustedBudget, scalePercentage, pacePercentage) {
    var normalAmounts = getDailyBudgetChartAmounts(normalExpenseTotal, adjustedBudget);
    var naturalSegments = [
        { key: "normalWithin", label: "通常支出（予算内）", amount: normalAmounts.withinBudget, color: getDailyBudgetChartTheme(normalExpenseTotal, adjustedBudget).color },
        { key: "normalOver", label: "通常支出（100〜150%）", amount: normalAmounts.overBudget, color: getDailyBudgetChartTheme(normalExpenseTotal, adjustedBudget).overBudgetColor },
        { key: "normalCritical", label: "通常支出（150%超）", amount: normalAmounts.criticalOverBudget, color: "#6b1512" },
        { key: "special", label: "承認済み特別費", amount: specialExpenseTotal, color: "#c8baff" },
        { key: "remaining", label: "残り", amount: Math.max(adjustedBudget - totalAmount, 0), color: "#e8eaed" }
    ];
    var markerAmount = getBudgetChartMarkerAmount(adjustedBudget, scalePercentage);
    var budgetMarkerAmount = markerAmount * 3;
    var markers = [{
        key: "budgetLimit",
        label: "予算上限",
        start: adjustedBudget - budgetMarkerAmount,
        end: adjustedBudget,
        color: "#202124"
    }];
    var position = 0;
    var naturalRanges = naturalSegments.map(function (segment) {
        var range = {
            key: segment.key,
            label: segment.label,
            start: position,
            end: position + segment.amount,
            color: segment.color
        };
        position = range.end;
        return range;
    });
    var boundaries = [0, position];

    if (pacePercentage > 0 && pacePercentage < scalePercentage) {
        var pacePosition = adjustedBudget * pacePercentage / 100;
        markers.push({
            key: "pace",
            label: "月内ペースの目安",
            start: pacePosition - markerAmount * 1.5,
            end: pacePosition + markerAmount * 1.5,
            color: "#1a73e8"
        });
    }

    naturalRanges.forEach(function (range) {
        boundaries.push(range.start, range.end);
    });
    markers.forEach(function (marker) {
        boundaries.push(marker.start, marker.end);
    });
    boundaries.sort(function (left, right) { return left - right; });

    return boundaries.reduce(function (segments, boundary, index) {
        var nextBoundary = boundaries[index + 1];
        var midpoint;
        var marker;
        var naturalRange;

        if (nextBoundary === undefined || nextBoundary <= boundary) {
            return segments;
        }

        midpoint = (boundary + nextBoundary) / 2;
        marker = markers.find(function (currentMarker) {
            return midpoint >= currentMarker.start && midpoint < currentMarker.end;
        });
        naturalRange = naturalRanges.find(function (range) {
            return midpoint >= range.start && midpoint < range.end;
        });

        if (marker) {
            segments.push({ key: marker.key, label: marker.label, amount: nextBoundary - boundary, color: marker.color });
        } else if (naturalRange) {
            segments.push({ key: naturalRange.key, label: naturalRange.label, amount: nextBoundary - boundary, color: naturalRange.color });
        }

        return segments;
    }, []);
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

function createMonthlySummaryBudgetChartBlob(totalAmount, adjustedBudget, specialExpenseTotal) {
    return createWeeklySummaryPeriodBudgetChartBlob(
        totalAmount,
        adjustedBudget,
        "今月",
        buildMonthlySummaryBudgetChartTitle(totalAmount, adjustedBudget, specialExpenseTotal),
        "monthly-summary-budget-chart.png",
        specialExpenseTotal || 0,
        getDailyBudgetChartScalePercentage(totalAmount, adjustedBudget),
        null
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
        scalePercentage,
        currentDate ? getMonthlyBudgetPacePercentage(currentDate) : null
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
        scalePercentage,
        null
    );
}

function createWeeklySummaryPeriodBudgetChartBlob(totalAmount, adjustedBudget, chartLabel, chartTitle, fileName, specialExpenseTotal, scalePercentage, pacePercentage) {
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
            specialExpenseTotal: specialExpenseTotal,
            scalePercentage: scalePercentage,
            pacePercentage: pacePercentage
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
    var scalePercentage = getBudgetChartScalePercentage(totalAmount, adjustedBudget, options);
    var chartMaximum = adjustedBudget * scalePercentage / 100;
    var chartSegments = getBudgetChartStackSegments(totalAmount, normalExpenseTotal, specialExpenseTotal, adjustedBudget, scalePercentage, options.pacePercentage);
    var chartDataBuilder = Charts.newDataTable()
        .addColumn(Charts.ColumnType.STRING, "週予算");
    var chartRow = [chartLabel];

    chartSegments.forEach(function (segment) {
        chartDataBuilder.addColumn(Charts.ColumnType.NUMBER, segment.label);
        chartRow.push(segment.amount);
    });

    var chartData = chartDataBuilder.addRow(chartRow).build();
    var chart = Charts.newBarChart()
        .setDataTable(chartData)
        .setDimensions(600, options.height || 160)
        .setOption("title", chartTitle)
        .setOption("titleTextStyle", { color: totalAmount > adjustedBudget ? "#b42318" : "#202124", fontSize: 15, bold: true })
        .setOption("legend", { position: "none" })
        .setOption("isStacked", true)
        .setOption("colors", chartSegments.map(function (segment) { return segment.color; }))
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

function buildMonthlySummaryHtmlBody(body, totalAmount, adjustedBudget, specialExpenseTotal) {
    return buildSummaryHtmlBody(body, [{
        key: "monthlySummaryBudgetChart",
        title: buildMonthlySummaryBudgetChartTitle(totalAmount, adjustedBudget, specialExpenseTotal || 0)
    }]);
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
