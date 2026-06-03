// 時間設定トリガーで実行される際に、引数を渡されるようにする
function handleCalculateExpensesSummaryTrigger() {
    calculateExpensesSummary('mail');
}

// 行いたい操作を引数actionで受け取る
function calculateExpensesSummary(action) {
    // 共通設定
    var budgetPerWeek = 40000; // 週ごとの予算

    var runtimeContext = getScriptRuntimeContext();
    var currentDate = runtimeContext.currentDate;
    var isStaging = runtimeContext.isStaging;

    // その週の日付一覧を取得し、年内の日付のみを含める
    var datesInWeek = getDatesInWeek(currentDate);
    var startOfWeek = datesInWeek[0]; // 週の開始日（月曜日）
    var endOfWeek = datesInWeek[datesInWeek.length - 1];   // 週の終了日

    // 日付範囲の文字列を作成
    var dateRangeStr = formatDate(startOfWeek) + "〜" + formatDate(endOfWeek);

    // その週に含まれる日付内データを一覧で取得
    var dataEntries = getDataForDates(datesInWeek);

    // 合計金額を算出
    var totalAmount = calculateTotalAmount(dataEntries);

    // 予算を含まれる日数に応じて調整
    var numberOfDays = datesInWeek.length;
    var adjustedBudget = Math.round((budgetPerWeek * numberOfDays / 7) / 100) * 100; // 100円単位で丸め込み

    // 予算との差分を計算
    var difference = totalAmount - adjustedBudget;
    var percentage = (totalAmount / adjustedBudget) * 100;

    // デバッグ用出力
    Logger.log("データエントリ一覧:");
    dataEntries.forEach(function (entry) {
        Logger.log("日付: " + formatDate(entry.date) + ", 名称: " + entry.name + ", 金額: " + entry.amount);
    });
    Logger.log(dateRangeStr + " の合計金額: " + totalAmount + "円");
    if (difference > 0) {
        Logger.log("予算を " + difference + " 円上回りました。");
    } else {
        Logger.log("予算を " + Math.abs(difference) + " 円下回りました。");
    }
    Logger.log("予算の " + percentage.toFixed(2) + "% を使用しました。");

    // 現在の曜日を取得（0:日曜日, 1:月曜日, ..., 6:土曜日）
    var dayOfWeek = currentDate.getDay();

    if (dayOfWeek === 0) {
        // 日曜日の場合、週次サマリーを送信
        return sendWeeklySummaryEmail(dateRangeStr, totalAmount, dataEntries, difference, percentage, adjustedBudget, isStaging, action, currentDate);
    } else {
        // 日曜日以外の場合、週の開始から現在までのデータを取得し、メールで送信
        return sendDailyProgressEmail(currentDate, datesInWeek, adjustedBudget, isStaging, action);
    }
}

// 日付を "MM/DD" の形式にフォーマットする関数
function formatDate(date) {
    var month = date.getMonth() + 1;
    var day = date.getDate();
    return month + "/" + day;
}

// その週に含まれる日付一覧を求めるメソッド（年内の日付のみを含める）
function getDatesInWeek(date) {
    var dates = [];
    var currentYear = date.getFullYear();

    // 週の始まり（月曜日）を取得
    var day = date.getDay(); // 0（日曜）から6（土曜）
    var diff = date.getDate() - day + (day === 0 ? -6 : 1); // 日曜の場合は-6
    var monday = new Date(date);
    monday.setDate(diff);
    monday.setHours(0, 0, 0, 0);

    // 月曜日から日曜日までの日付を取得
    for (var i = 0; i < 7; i++) {
        var d = new Date(monday);
        d.setDate(monday.getDate() + i);
        d.setHours(0, 0, 0, 0);

        // 年が同じ場合のみ追加
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

// その週に含まれる日付内データを一覧で取得するメソッド
function getDataForDates(dates) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("🐖 家計簿");
    if (!sheet) {
        throw new Error('シート「🐖 家計簿」が見つかりません。');
    }
    var startRow = 35; // データが開始する行
    var endRow = 184; // 支出明細の終端行

    var dataEntries = [];
    var targetDateKeys = dates.map(function (date) {
        return toMonthDayKey(date);
    });
    var targetDateKeySet = {};
    targetDateKeys.forEach(function (key) {
        targetDateKeySet[key] = true;
    });

    Logger.log("集計対象日: " + targetDateKeys.join(", "));

    var processedMonths = {};

    dates.forEach(function (date) {
        var year = date.getFullYear();
        var month = date.getMonth(); // 0始まりの月（0が1月）
        if (processedMonths[month]) {
            return;
        }
        processedMonths[month] = true;

        var columns = getColumnsForMonth(month);
        var dataRange = sheet.getRange(
            startRow,
            columns.dateCol,
            endRow - startRow + 1,
            4
        );
        var data = dataRange.getValues();
        var currentDateKey = null;
        var matchedCount = 0;

        for (var i = 0; i < data.length; i++) {
            var row = data[i];
            var dateCell = row[0];
            var category = row[1];
            var name = row[2];
            var amount = row[3];

            if (isFixedCostMarker(dateCell)) {
                Logger.log(
                    "固定費マーカーを検出したため読取終了: month=" +
                        (month + 1) +
                        " row=" +
                        (startRow + i)
                );
                break;
            }

            var hasContent = [dateCell, category, name, amount].some(function (cell) {
                return cell !== null && cell.toString().trim() !== '';
            });
            if (!hasContent) {
                continue;
            }

            if (dateCell && dateCell.toString().trim() !== '') {
                currentDateKey = toMonthDayKey(dateCell);
            }

            if (currentDateKey && targetDateKeySet[currentDateKey]) {
                var hasEntryContent =
                    (name !== null && String(name).trim() !== '') ||
                    (amount !== null && String(amount).trim() !== '');
                if (!hasEntryContent) {
                    continue;
                }

                dataEntries.push({
                    date: monthDayKeyToDate(currentDateKey, year),
                    category: category,
                    name: name,
                    amount: amount
                });
                matchedCount++;
            }
        }

        Logger.log(
            "月別読取結果: month=" +
                (month + 1) +
                " cols=" +
                columns.dateCol +
                "-" +
                (columns.dateCol + 3) +
                " matched=" +
                matchedCount
        );
    });

    return dataEntries;
}

// 月に対応する列情報を取得するメソッド
function getColumnsForMonth(month) {
    // G列が7番目の列で、各月ごとに4列ずつデータがある
    // 0（1月）から始まる月を想定
    var dateCol = 7 + month * 4; // 1月は7列目
    var amountCol = dateCol + 3; // 金額列は日付列から3列後
    return { dateCol: dateCol, amountCol: amountCol };
}

// 日付文字列をDateオブジェクトに変換する関数
function parseDate(dateStr, year) {
    var dateParts = dateStr.split('/');
    var month = parseInt(dateParts[0], 10) - 1; // 月は0始まり
    var day = parseInt(dateParts[1], 10);
    return new Date(year, month, day);
}

function isFixedCostMarker(value) {
    if (value === null || value === undefined) {
        return false;
    }

    return String(value).trim() === '固定';
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

// 求めたデータ一覧の金額合計を算出するメソッド
function calculateTotalAmount(dataEntries) {
    var total = 0;
    dataEntries.forEach(function (entry) {
        var amount = parseFloat(entry.amount);
        if (!isNaN(amount)) {
            total += amount;
        }
    });
    return total;
}

function calculateCategoryTotals(dataEntries) {
    var totals = {};
    dataEntries.forEach(function (entry) {
        var key = entry.category || "未分類";
        var amount = parseFloat(entry.amount);
        if (isNaN(amount)) {
            return;
        }
        if (!totals[key]) {
            totals[key] = 0;
        }
        totals[key] += amount;
    });
    return totals;
}

function getCategoryRankingLines(dataEntries) {
    var categoryTotals = calculateCategoryTotals(dataEntries);
    return Object.keys(categoryTotals)
        .sort(function (a, b) {
            return categoryTotals[b] - categoryTotals[a];
        })
        .map(function (category, index) {
            return (
                "・" +
                (index + 1) +
                "位 " +
                category +
                ": " +
                categoryTotals[category] +
                "円"
            );
        });
}

function countUncategorizedEntries(dataEntries) {
    return dataEntries.filter(function (entry) {
        return !entry.category || String(entry.category).trim() === "";
    }).length;
}

// Gemini に支出傾向の簡易分析を依頼するメソッド
function analyzeExpensesWithGemini(dataEntries, totalAmount, adjustedBudget, percentage, baseDate) {
    var apiKey = getGeminiApiKey();
    if (!apiKey) {
        return null;
    }
    var expenseLines = dataEntries.map(function (entry) {
        return formatDate(entry.date) + " [" + (entry.category || "未分類") + "] " + entry.name + " " + entry.amount + "円";
    });
    var upcomingPlannedExpenses = getUpcomingPlannedExpenses(baseDate);
    var upcomingExpenseLines = upcomingPlannedExpenses.map(function (entry) {
        return formatDate(entry.date) + " [" + entry.title + "] " + entry.memo;
    });
    var categoryRankingLines = getCategoryRankingLines(dataEntries);

    var prompt = [
        "あなたはプロの家計管理アドバイザーです。挨拶や自己紹介は禁止です。金銭感覚の改善を目的としたコーチとして、冷静な分析と、時には優しく、時には厳しく指導してください。1週間分の支出について、予算を超えないようアドバイスをください。カジュアルな敬語て対応してください。",
        "レシートは保管していませんが、代わりに全ての支出・収入をスプレッドシートに記録しています。単なる分析にとどまらず、「行動に落とし込める改善提案」を重視してください。感情的にならず、客観的かつ現実的な判断で、飴と鞭を使い分けてください。",
        "通勤時（勤務地：永田町）の通勤定期はありませんが、給与で補填されます。食事はスーパーでまとめ買いした上でほぼ自炊しており、外食は友人と会う時が多いです。",
        "Googleカレンダーの今後の予定メモに書かれた支出予定も考慮して助言してください。近いうちに大きな支出予定があるなら、今週の節約を強めに促してください。",
        "1週間は月曜始まり・日曜終わりで考えて、今日までの傾向を数個と、次に意識すべき点を数個、箇条書きでまとめてください。箇条書きは最大5個までにし、それぞれに補足を付けてください。Markdown記法は使わず、プレーンな文字と絵文字のみで出力し、全体で500文字以内に収めてください。",
        "週予算: " + adjustedBudget + "円 / これまでの支出: " + totalAmount + "円 (" + percentage.toFixed(1) + "%)",
        "カテゴリ別支出ランキング: " + (categoryRankingLines.length ? categoryRankingLines.join(" / ") : "なし"),
        "今後の予定メモ: " + (upcomingExpenseLines.length ? upcomingExpenseLines.join(" / ") : "なし"),
        "以下は支出一覧(日付、カテゴリ、名称、金額)です。名称だけで内容が不明瞭な場合はカテゴリから内容を推定してください:",
        expenseLines.join("\n")
    ].join("\n");

    Logger.log("Geminiプロンプト：")
    Logger.log(prompt)

    return generateGeminiText(apiKey, prompt, {
        temperature: 0.4,
        maxOutputTokens: 1000,
        thinkingConfig: {
            thinkingBudget: 400
        }
    });
}

function getUpcomingPlannedExpenses(baseDate) {
    var startDate = new Date(baseDate.getFullYear(), baseDate.getMonth(), baseDate.getDate());
    var lookaheadDays = getUpcomingExpenseLookaheadDays();
    var endDate = new Date(startDate);
    endDate.setDate(endDate.getDate() + lookaheadDays);
    return getPlannedExpensesInRange(startDate, endDate);
}

function getPlannedExpensesForCurrentWeek(baseDate) {
    var range = getWeekRange(baseDate);
    var startDate = new Date(baseDate.getFullYear(), baseDate.getMonth(), baseDate.getDate());
    var endDateExclusive = new Date(range.endDate);
    endDateExclusive.setDate(endDateExclusive.getDate() + 1);
    return getPlannedExpensesInRange(startDate, endDateExclusive);
}

function getPlannedExpensesInRange(startDate, endDate) {
    if (typeof CalendarApp === 'undefined') {
        Logger.log("CalendarApp is unavailable in this runtime.");
        return [];
    }

    var calendar = getTargetCalendar();
    if (!calendar) {
        return [];
    }

    var events = calendar.getEvents(startDate, endDate);
    var plannedExpenses = [];
    Logger.log(
        "予定支出検索: start=" +
            formatDate(startDate) +
            " end=" +
            formatDate(endDate) +
            " events=" +
            events.length
    );

    events.forEach(function (event) {
        var title = event.getTitle() || "予定";
        var description = sanitizePlannedExpenseMemo(event.getDescription() || "");
        if (!hasPlannedExpenseMemo(description, title)) {
            return;
        }

        plannedExpenses.push({
            title: title,
            date: event.getStartTime(),
            memo: description
        });
    });

    plannedExpenses = cleanPlannedExpenseMemosWithGemini(plannedExpenses);

    plannedExpenses.sort(function (a, b) {
        return a.date.getTime() - b.date.getTime();
    });

    Logger.log("予定支出取得件数: " + plannedExpenses.length);

    return plannedExpenses;
}

function formatUpcomingPlannedExpenseLines(plannedExpenses) {
    return plannedExpenses.map(function (entry) {
        return "・" + formatDate(entry.date) + " - " + entry.title + ": " + entry.memo;
    });
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

function cleanPlannedExpenseMemosWithGemini(plannedExpenses) {
    if (!plannedExpenses.length) {
        return plannedExpenses;
    }

    var apiKey = getGeminiApiKey();
    if (!apiKey) {
        return plannedExpenses;
    }

    var prompt = [
        "以下はGoogleカレンダー予定のタイトルとメモです。",
        "目的は、支出予定として意味のある情報だけを残し、Googleの自動追記やURLや案内文など無関係な文を除去することです。",
        "各要素について cleanedMemo を返してください。",
        "ルール:",
        "- 支出予定として意味がある部分だけ残す",
        "- URL、Googleの案内文、自動生成の説明、予約メール由来の定型文は削除する",
        "- 金額がなくても、予定に関係する買い物や外食のメモなら残してよい",
        "- 予定に関係する情報が何も残らないなら cleanedMemo を空文字にする",
        "- 出力は各行 `index|cleanedMemo` のみ",
        "- 説明文、Markdown、コードブロック、jsonという語は出力しない",
        "- 例: `0|予定：服4000円`",
        JSON.stringify(plannedExpenses.map(function (entry, index) {
            return {
                index: index,
                title: entry.title,
                memo: entry.memo
            };
        }))
    ].join("\n");

    var responseText = generateGeminiText(apiKey, prompt, {
        temperature: 0,
        maxOutputTokens: 800
    });
    if (!responseText) {
        return plannedExpenses;
    }

    var cleanedByIndex = {};
    var parsedLines = parsePipeResponse(responseText);
    if (!parsedLines.length) {
        Logger.log("予定メモ整形の行解析に失敗: " + responseText);
        return plannedExpenses;
    }

    parsedLines.forEach(function (item) {
        if (typeof item.index === 'number' && typeof item.cleanedMemo === 'string') {
            cleanedByIndex[item.index] = item.cleanedMemo.trim();
        }
    });

    var cleanedExpenses = plannedExpenses.map(function (entry, index) {
        var cleanedMemo = cleanedByIndex.hasOwnProperty(index)
            ? cleanedByIndex[index]
            : entry.memo;

        return {
            title: entry.title,
            date: entry.date,
            memo: cleanedMemo
        };
    }).filter(function (entry) {
        return entry.memo !== "";
    });

    Logger.log("予定メモ整形後件数: " + cleanedExpenses.length);
    return cleanedExpenses;
}

function hasPlannedExpenseMemo(text, title) {
    var combined = [title || "", text || ""].join(" ").trim();
    if (!combined) {
        return false;
    }

    if (/予定/.test(combined)) {
        return true;
    }

    if (/[¥￥]/.test(combined) || /\d[\d,]*\s*円/.test(combined)) {
        return true;
    }

    return /(交通費|食費|外食|買い物|ランチ|ディナー|飲み|会費|チケット|ホテル|宿|タクシー|電車|バス|飛行機)/.test(combined);
}

function sanitizePlannedExpenseMemo(text) {
    return text
        .replace(/\r\n/g, "\n")
        .replace(/\r/g, "\n")
        .split("\n")
        .map(function (line) {
            return line.trim();
        })
        .filter(function (line) {
            return line !== "";
        })
        .join(" / ");
}

function normalizeFullWidthNumbers(text) {
    return text
        .replace(/[０-９]/g, function (char) {
            return String.fromCharCode(char.charCodeAt(0) - 0xFEE0);
        })
        .replace(/，/g, ",");
}

function getGeminiApiKey() {
    var apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
    if (!apiKey) {
        Logger.log("Gemini API key is not set in script properties.");
        return null;
    }
    return apiKey;
}

function getGeminiModelsToTry(apiKey) {
    var modelFromProp = PropertiesService.getScriptProperties().getProperty("GEMINI_MODEL");
    if (modelFromProp && modelFromProp.indexOf("models/") === 0) {
        modelFromProp = modelFromProp.replace(/^models\//, "");
    }

    var modelsToTry = [];
    if (modelFromProp) {
        modelsToTry.push(modelFromProp);
    } else {
        modelsToTry = modelsToTry.concat(fetchGenerativeModels(apiKey));
    }

    if (modelsToTry.length === 0) {
        modelsToTry = [
            "gemini-2.5-flash",
            "gemini-2.5-pro",
            "gemini-2.5-pro-preview-06-05",
            "gemini-2.0-flash",
            "gemini-1.5-flash"
        ];
    }

    return modelsToTry;
}

function generateGeminiText(apiKey, prompt, generationConfig) {
    var modelsToTry = getGeminiModelsToTry(apiKey);
    var apiVersions = ["v1beta", "v1"];
    var payload = {
        contents: [{
            parts: [{
                text: prompt
            }]
        }],
        generationConfig: generationConfig
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

function parsePipeResponse(text) {
    return text
        .replace(/```[\s\S]*?\n/g, "")
        .replace(/```/g, "")
        .split("\n")
        .map(function (line) {
            return line.trim();
        })
        .filter(function (line) {
            return /^\d+\|/.test(line);
        })
        .map(function (line) {
            var separatorIndex = line.indexOf("|");
            return {
                index: parseInt(line.substring(0, separatorIndex), 10),
                cleanedMemo: line.substring(separatorIndex + 1).trim()
            };
        })
        .filter(function (item) {
            return !isNaN(item.index);
        });
}

function calculatePlannedExpenseTotal(plannedExpenses) {
    var total = 0;

    plannedExpenses.forEach(function (entry) {
        var memo = entry.memo || "";
        var matches = memo.match(/([0-9,]+)\s*円/g) || [];

        matches.forEach(function (match) {
            var amount = parseInt(match.replace(/[^\d]/g, ""), 10);
            if (!isNaN(amount)) {
                total += amount;
            }
        });
    });

    return total;
}

// listModels から generateContent 可能なモデル一覧を取得する
function fetchGenerativeModels(apiKey) {
    try {
        var listUrl = "https://generativelanguage.googleapis.com/v1beta/models?key=" + apiKey;
        var listResponse = UrlFetchApp.fetch(listUrl, { method: "get", muteHttpExceptions: true });
        if (listResponse.getResponseCode() !== 200) {
            Logger.log("Failed to list models: status=" + listResponse.getResponseCode() + " body=" + listResponse.getContentText());
            return [];
        }
        var parsed = JSON.parse(listResponse.getContentText());
        if (!parsed.models || !parsed.models.length) {
            Logger.log("No models found in listModels response.");
            return [];
        }
        var candidates = parsed.models.filter(function (model) {
            return model.supportedGenerationMethods && model.supportedGenerationMethods.indexOf("generateContent") !== -1;
        }).map(function (model) {
            return model.name.replace(/^models\//, "");
        });
        Logger.log("generateContent-capable models (auto-detected): " + candidates.join(", "));
        return candidates;
    } catch (error) {
        Logger.log("Error while listing models: " + error);
        return [];
    }
}

// 週次サマリーをメールで送信するメソッド（毎週日曜日）
function sendWeeklySummaryEmail(dateRangeStr, totalAmount, dataEntries, difference, percentage, adjustedBudget, isStaging, action, currentDate) {
    var emailAddress = "TARGET_EMAIL_ADDRESS";
    var upcomingPlannedExpenses = getPlannedExpensesForCurrentWeek(currentDate);
    var upcomingExpenseLines = formatUpcomingPlannedExpenseLines(upcomingPlannedExpenses);
    var plannedExpenseTotal = calculatePlannedExpenseTotal(upcomingPlannedExpenses);

    // 予算差分の符号を設定
    var differenceSign = difference >= 0 ? "+" : "-";
    var differenceAbs = Math.abs(difference);

    // 予算割合を小数点以下2桁で表示
    var percentageStr = percentage.toFixed(2);
    var projectedPercentage = adjustedBudget
        ? (((totalAmount + plannedExpenseTotal) / adjustedBudget) * 100).toFixed(2)
        : "0.00";
    var categoryRankingLines = getCategoryRankingLines(dataEntries);

    // トップ5の支出を計算
    var top5Entries = dataEntries.slice(); // 配列をコピー
    top5Entries.sort(function (a, b) {
        return parseFloat(b.amount) - parseFloat(a.amount);
    });
    top5Entries = top5Entries.slice(0, 5);

    // メール本文を指定のフォーマットで作成
    var body = "";
    body += "◆ " + dateRangeStr + " の週次サマリー\n\n";
    body += "++++++++++ 💸 予算サマリー 💸 ++++++++++\n";
    body += dateRangeStr + " の実支出: " + totalAmount + " 円\n";
    body += "今後の予定金額: " + plannedExpenseTotal + " 円\n";
    body += "支出＋予定の合計見込み: " + (totalAmount + plannedExpenseTotal) + " 円\n";
    body += "\n";
    body += "予算に対して\n";
    body += "・実支出: " + percentageStr + "%\n";
    body += "・合計見込み: " + projectedPercentage + "%\n";
    body += "（設定予算：" + adjustedBudget + "円）\n";
    body += "++++++++++++++++++++++++++++++++++++\n";
    body += "* 予算差分：" + differenceSign + differenceAbs + "円\n\n";
    body += "◆ カテゴリ別支出ランキング\n";
    if (categoryRankingLines.length) {
        categoryRankingLines.forEach(function (line) {
            body += line + "\n";
        });
    } else {
        body += "・なし\n";
    }
    body += "\n";
    body += "◆ 支出TOP5\n";
    top5Entries.forEach(function (entry) {
        body += "・" + formatDate(entry.date) + " - " + entry.name + ": " + entry.amount + "円\n";
    });
    body += "\n";
    body += "◆ 支出一覧\n";
    dataEntries.forEach(function (entry) {
        body += "・" + formatDate(entry.date) + " - " + entry.name + ": " + entry.amount + "円\n";
    });
    body += "\n";

    body += "◆ 直近の予定支出\n";
    if (upcomingExpenseLines.length) {
        upcomingExpenseLines.forEach(function (line) {
            body += line + "\n";
        });
    } else {
        body += "・なし\n";
    }
    body += "\n";

    var geminiAnalysis = analyzeExpensesWithGemini(dataEntries, totalAmount, adjustedBudget, percentage, currentDate);
    if (geminiAnalysis) {
        body += "◆ Gemini分析\n" + geminiAnalysis + "\n";
    } else {
        body += "◆ Gemini分析\n(Geminiからの回答を取得できませんでした。ログを確認してください)\n";
    }

    switch (action) {
        case 'mail':
            // メールを送信
            var subject = (isStaging ? "<test>" : "")
                + "家計簿週次レポート" + "（" + dateRangeStr + "）";
            MailApp.sendEmail(emailAddress, subject, body);
            return "Successfully sent mail";
        case 'text':
            return body;
        default:
            throw new Error('actionが定義されていません');
    }
}

// 日曜日以外に日次進捗をメールで送信するメソッド
function sendDailyProgressEmail(currentDate, datesInWeek, adjustedBudget, isStaging, action) {
    var emailAddress = "TARGET_EMAIL_ADDRESS";
    var upcomingPlannedExpenses = getPlannedExpensesForCurrentWeek(currentDate);
    var upcomingExpenseLines = formatUpcomingPlannedExpenseLines(upcomingPlannedExpenses);
    var plannedExpenseTotal = calculatePlannedExpenseTotal(upcomingPlannedExpenses);

    // 週の開始日から現在の日付までのデータを取得
    var datesUpToToday = datesInWeek.filter(function (date) {
        return date <= currentDate;
    });

    var dataEntries = getDataForDates(datesUpToToday).reverse().map(entry => {
        if (entry.name.length >= 16) {
            entry.name = entry.name.substring(0, 14) + "...";
        }
        return entry;
    });
    // 合計金額を算出
    var totalAmount = calculateTotalAmount(dataEntries);

    // 予算に対する割合を計算
    var percentage = (totalAmount / adjustedBudget) * 100;
    var projectedPercentage = adjustedBudget
        ? (((totalAmount + plannedExpenseTotal) / adjustedBudget) * 100).toFixed(2)
        : "0.00";
    var categoryRankingLines = getCategoryRankingLines(dataEntries);
    var uncategorizedCount = countUncategorizedEntries(dataEntries);

    // Gemini分析
    var geminiAnalysis = analyzeExpensesWithGemini(dataEntries, totalAmount, adjustedBudget, percentage, currentDate);


    // メールの件名と本文を作成
    var subject = (isStaging ? "<test>" : "")
        + "家計簿日次レポート（" + formatDate(currentDate) + "）";
    var body = "++++++++++ 💸 予算サマリー 💸 ++++++++++\n";
    body += formatDate(datesInWeek[0]) + " から " + formatDate(currentDate) + " までの実支出: " + totalAmount + " 円\n";
    body += "今後の予定金額: " + plannedExpenseTotal + " 円\n";
    body += "支出＋予定の合計見込み: " + (totalAmount + plannedExpenseTotal) + " 円\n";
    body += "\n";
    body += "予算に対して\n";
    body += "・実支出: " + percentage.toFixed(2) + "%\n";
    body += "・合計見込み: " + projectedPercentage + "%\n";
    body += "（設定予算：" + adjustedBudget + "円）\n";
    if (uncategorizedCount > 0) {
        body += "・未分類の支出: " + uncategorizedCount + "件\n";
    }
    body += "++++++++++++++++++++++++++++++++++++\n\n";

    body += "◆ カテゴリ別支出ランキング\n";
    if (categoryRankingLines.length) {
        categoryRankingLines.forEach(function (line) {
            body += line + "\n";
        });
    } else {
        body += "・なし\n";
    }
    body += "\n";

    body += "詳細:\n";
    dataEntries.forEach(function (entry) {
        body += "・" + formatDate(entry.date) + " - " + entry.name + ": " + entry.amount + "円\n";
    });
    body += "\n";

    body += "◆ 直近の予定支出\n";
    if (upcomingExpenseLines.length) {
        upcomingExpenseLines.forEach(function (line) {
            body += line + "\n";
        });
    } else {
        body += "・なし\n";
    }
    body += "\n";

    if (geminiAnalysis) {
        body += "◆ Gemini分析\n" + geminiAnalysis + "\n";
    } else {
        body += "◆ Gemini分析\n(Geminiからの回答を取得できませんでした。ログを確認してください)\n";
    }

    switch (action) {
        case 'mail':
            // メールを送信
            MailApp.sendEmail(emailAddress, subject, body);
            return "Successfully sent mail";
        case 'text':
            return body;
        default:
            throw new Error('actionが定義されていません');
    }
}
