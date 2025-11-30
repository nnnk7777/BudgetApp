// 月次サマリーを生成するトリガー用
function handleCalculateMonthlySummaryTrigger() {
    calculateMonthlySummary('mail');
}

// 月次サマリーのメイン処理
function calculateMonthlySummary(action) {
    var budgetPerWeek = 45000;
    var currentDate;
    var testDateStr = "TEST_DATE_PLACEHOLDER"
    var isStaging = testDateStr ? true : false

    if (isStaging) {
        // テスト用の日付が指定されている場合、その日付を使用
        // YYYYMMDD フォーマットをパースして Date オブジェクトを作成
        currentDate = parseYYYYMMDD(testDateStr);
        if (!currentDate) {
            throw new Error('Invalid TEST_DATE format. Expected YYYYMMDD.');
        }
    } else {
        // 指定がない場合は現在の日付を使用
        currentDate = new Date();
    }
    var year = currentDate.getFullYear();
    var month = currentDate.getMonth(); // 0-based

    var startOfMonth = new Date(year, month, 1);
    var endOfMonth = new Date(year, month + 1, 0);
    var dateRangeStr = formatDate(startOfMonth) + "〜" + formatDate(endOfMonth);

    var expenseEntries = getMonthlyExpenses(year, month);
    var incomeEntries = getMonthlyIncomes(year, month);

    var totalExpenses = calculateTotalAmount(expenseEntries);
    var totalIncome = calculateTotalAmount(incomeEntries);

    var daysInMonth = endOfMonth.getDate();
    var adjustedBudget = Math.round((budgetPerWeek * daysInMonth / 7) / 100) * 100;
    var difference = totalExpenses - adjustedBudget;
    var percentage = adjustedBudget ? (totalExpenses / adjustedBudget) * 100 : 0;

    var categoryTotals = calculateCategoryTotals(expenseEntries);
    var sortedCategories = Object.keys(categoryTotals).sort(function (a, b) {
        return categoryTotals[b] - categoryTotals[a];
    });

    var subject = "家計簿月次レポート（" + (month + 1) + "月）";
    var body = "";
    body += "◆ " + dateRangeStr + " の月次サマリー\n\n";
    body += "収入合計: " + totalIncome + "円\n";
    body += "支出合計: " + totalExpenses + "円\n";
    body += "月間予算(週予算換算): " + adjustedBudget + "円\n";
    body += "予算差分: " + (difference >= 0 ? "+" : "-") + Math.abs(difference) + "円\n";
    body += "予算消化率: " + percentage.toFixed(2) + "%\n\n";

    body += "◆ カテゴリ別支出\n";
    sortedCategories.forEach(function (category) {
        body += "・" + category + ": " + categoryTotals[category] + "円\n";
    });
    body += "\n";

    body += "◆ 支出一覧\n";
    expenseEntries.forEach(function (entry) {
        body += "・" + formatDate(entry.date) + " - " + entry.category + " - " + entry.name + ": " + entry.amount + "円\n";
    });

    body += "\n◆ 収入一覧\n";
    incomeEntries.forEach(function (entry) {
        body += "・" + formatDate(entry.date) + " - " + entry.name + ": " + entry.amount + "円\n";
    });

    var geminiAnalysis = analyzeMonthlyWithGemini(expenseEntries, categoryTotals, totalExpenses, totalIncome, adjustedBudget, percentage, dateRangeStr);
    body += "\n";
    if (geminiAnalysis) {
        body += "◆ Gemini分析\n" + geminiAnalysis + "\n";
    } else {
        body += "◆ Gemini分析\n(Geminiからの回答を取得できませんでした。ログを確認してください)\n";
    }

    if (action !== 'mail') {
        throw new Error('actionはmailのみ対応しています');
    }
    MailApp.sendEmail("TARGET_EMAIL_ADDRESS", subject, body);
    return "Successfully sent monthly summary mail";
}

// YYYYMMDD フォーマットの日付文字列を Date オブジェクトに変換する関数
function parseYYYYMMDD(dateStr) {
    if (!/^\d{8}$/.test(dateStr)) {
        return null;
    }
    var year = parseInt(dateStr.substring(0, 4), 10);
    var month = parseInt(dateStr.substring(4, 6), 10) - 1; // 月は0始まり
    var day = parseInt(dateStr.substring(6, 8), 10);
    return new Date(year, month, day);
}

// 指定月の支出データを取得する
function getMonthlyExpenses(year, month) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("🐖 家計簿");
    if (!sheet) {
        throw new Error('シート「🐖 家計簿」が見つかりません。');
    }
    var startRow = 35;
    var endRow = 185;

    var columns = getColumnsForMonth(month);
    var dataRange = sheet.getRange(startRow, columns.dateCol, endRow - startRow + 1, 4);
    var data = dataRange.getValues();

    var entries = [];
    var currentDate = null;
    data.forEach(function (row, index) {
        var dateCell = row[0];
        var category = row[1];
        var name = row[2];
        var amount = row[3];
        var absoluteRow = startRow + index;

        var hasContent = [dateCell, category, name, amount].some(function (cell) {
            return cell !== null && cell.toString().trim() !== '';
        });
        if (!hasContent) {
            return;
        }

        if (dateCell && dateCell.toString().trim() !== '') {
            if (typeof dateCell === 'string') {
                currentDate = parseDate(dateCell, year);
            } else if (Object.prototype.toString.call(dateCell) === '[object Date]') {
                currentDate = new Date(dateCell.getFullYear(), dateCell.getMonth(), dateCell.getDate());
            }
        }

        var includeEntry = false;
        if (currentDate && currentDate.getFullYear() === year && currentDate.getMonth() === month) {
            includeEntry = true;
        } else if (!currentDate && absoluteRow >= 156) {
            // 156行目以降は日付が空でも固定費として当月扱い
            currentDate = new Date(year, month, 1);
            includeEntry = true;
        }

        if (includeEntry) {
            entries.push({
                date: currentDate || new Date(year, month, 1),
                category: category || "未分類",
                name: name || "",
                amount: amount || 0
            });
        }
    });

    return entries;
}

// 指定月の収入データを取得する
function getMonthlyIncomes(year, month) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("🐖 家計簿");
    if (!sheet) {
        throw new Error('シート「🐖 家計簿」が見つかりません。');
    }
    var startRow = 22;
    var endRow = 33;

    var columns = getColumnsForMonth(month);
    var dataRange = sheet.getRange(startRow, columns.dateCol, endRow - startRow + 1, 4);
    var data = dataRange.getValues();

    var entries = [];
    data.forEach(function (row) {
        var dateCell = row[0];
        var name = row[2];
        var amount = row[3];

        var hasContent = [dateCell, name, amount].some(function (cell) {
            return cell !== null && cell.toString().trim() !== '';
        });
        if (!hasContent) {
            return;
        }

        var entryDate;
        if (dateCell && dateCell.toString().trim() !== '') {
            if (typeof dateCell === 'string') {
                entryDate = parseDate(dateCell, year);
            } else if (Object.prototype.toString.call(dateCell) === '[object Date]') {
                entryDate = new Date(dateCell.getFullYear(), dateCell.getMonth(), dateCell.getDate());
            }
        } else {
            entryDate = new Date(year, month, 1);
        }

        if (entryDate.getFullYear() === year && entryDate.getMonth() === month) {
            entries.push({
                date: entryDate,
                name: name || "",
                amount: amount || 0
            });
        }
    });

    return entries;
}

// カテゴリ別合計を算出
function calculateCategoryTotals(entries) {
    var totals = {};
    entries.forEach(function (entry) {
        var key = entry.category || "未分類";
        var amount = parseFloat(entry.amount) || 0;
        if (!totals[key]) {
            totals[key] = 0;
        }
        totals[key] += amount;
    });
    return totals;
}

// 月次サマリーを Gemini で分析する
function analyzeMonthlyWithGemini(expenseEntries, categoryTotals, totalExpenses, totalIncome, adjustedBudget, percentage, dateRangeStr) {
    if (typeof PropertiesService === 'undefined') {
        return null;
    }

    var apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
    if (!apiKey) {
        Logger.log("Gemini API key is not set in script properties.");
        return null;
    }

    var modelFromProp = PropertiesService.getScriptProperties().getProperty("GEMINI_MODEL");
    if (modelFromProp && modelFromProp.indexOf("models/") === 0) {
        modelFromProp = modelFromProp.replace(/^models\//, "");
    }
    var modelsToTry = [];
    if (modelFromProp) {
        modelsToTry.push(modelFromProp);
    }
    if (typeof fetchGenerativeModels === 'function') {
        var autoModels = fetchGenerativeModels(apiKey);
        if (autoModels && autoModels.length) {
            modelsToTry = modelsToTry.concat(autoModels);
        }
    }
    if (modelsToTry.length === 0) {
        modelsToTry = [
            "gemini-2.5-flash",
            "gemini-2.5-pro",
            "gemini-2.0-flash",
            "gemini-1.5-flash"
        ];
    }
    var apiVersions = ["v1beta", "v1"];

    var expenseLines = expenseEntries.map(function (entry) {
        return formatDate(entry.date) + " [" + (entry.category || "未分類") + "] " + entry.name + " " + entry.amount + "円";
    });
    var categoryLines = Object.keys(categoryTotals).map(function (key) {
        return key + ": " + categoryTotals[key] + "円";
    });

    var prompt = [
        "あなたはプロの家計管理アドバイザー兼フィナンシャルプランナーです。金銭感覚の改善を目的としたコーチとして、冷静な分析と、時には優しく、時には厳しく指導してください。単なる分析にとどまらず、「行動に落とし込める改善提案」を重視してください。感情的にならず、客観的かつ現実的な判断で、飴と鞭を使い分けてください。",
        "対象期間: " + dateRangeStr,
        "月間予算(週予算換算): " + adjustedBudget + "円 / これまでの支出: " + totalExpenses + "円 (" + percentage.toFixed(1) + "%) / 収入: " + totalIncome + "円",
        "カテゴリ別支出: " + categoryLines.join(", "),
        "支出一覧:",
        expenseLines.join("\n"),
        "日本語で、1) 今月の傾向 2) 予算に対する評価 3) カテゴリ別の改善提案 を箇条書き最大3つずつ、プレーンテキストと絵文字のみで800文字以内にまとめてください。"
    ].join("\n");

    var payload = {
        contents: [{
            parts: [{
                text: prompt
            }]
        }],
        generationConfig: {
            temperature: 0.4,
            maxOutputTokens: 768
        }
    };

    for (var i = 0; i < apiVersions.length; i++) {
        var version = apiVersions[i];
        for (var j = 0; j < modelsToTry.length; j++) {
            var model = modelsToTry[j];
            var url = "https://generativelanguage.googleapis.com/" + version + "/models/" + model + ":generateContent?key=" + apiKey;

            try {
                var response = UrlFetchApp.fetch(url, {
                    method: "post",
                    contentType: "application/json",
                    payload: JSON.stringify(payload),
                    muteHttpExceptions: true
                });

                if (response.getResponseCode() !== 200) {
                    Logger.log("Gemini request failed (version=" + version + ", model=" + model + "): " + response.getResponseCode() + " " + response.getContentText());
                    continue;
                }

                var result = JSON.parse(response.getContentText());
                if (result && result.candidates && result.candidates.length > 0) {
                    var parts = result.candidates[0].content && result.candidates[0].content.parts;
                    if (parts && parts.length > 0) {
                        return parts.map(function (part) {
                            return part.text || "";
                        }).join("").trim();
                    }
                }
            } catch (error) {
                Logger.log("Gemini request error (version=" + version + ", model=" + model + "): " + error);
            }
        }
    }

    return null;
}
