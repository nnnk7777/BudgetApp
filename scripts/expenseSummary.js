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
    var dataEntries = getExpenseEntriesForDates(datesInWeek);

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
        return handleWeeklySummaryResult(dateRangeStr, totalAmount, dataEntries, difference, percentage, adjustedBudget, isStaging, action, currentDate);
    } else {
        // 日曜日以外の場合、週の開始から現在までのデータを取得し、メールで送信
        return handleDailySummaryResult(currentDate, datesInWeek, adjustedBudget, isStaging, action);
    }
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
