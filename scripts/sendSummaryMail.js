// 時間設定トリガーで実行される際に、引数を渡されるようにする
function handleCalculateExpensesSummaryTrigger() {
    calculateExpensesSummary('mail');
}

// 行いたい操作を引数actionで受け取る
function calculateExpensesSummary(action) {
    // 共通設定
    var budgetPerWeek = 45000; // 週ごとの予算

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
        return sendWeeklySummaryEmail(dateRangeStr, totalAmount, dataEntries, difference, percentage, adjustedBudget, isStaging, action);
    } else {
        // 日曜日以外の場合、週の開始から現在までのデータを取得し、メールで送信
        return sendDailyProgressEmail(currentDate, datesInWeek, adjustedBudget, isStaging, action);
    }
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

// その週に含まれる日付内データを一覧で取得するメソッド
function getDataForDates(dates) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("🐖 家計簿");
    if (!sheet) {
        throw new Error('シート「🐖 家計簿」が見つかりません。');
    }
    var startRow = 35; // データが開始する行

    var lastRow = sheet.getLastRow();
    var numRows = lastRow - startRow + 1;

    var dataEntries = [];
    var dateColumnCache = {}; // 月ごとの列情報をキャッシュ

    // 各日付ごとに処理
    dates.forEach(function (date) {
        var year = date.getFullYear();
        var month = date.getMonth(); // 0始まりの月（0が1月）
        var day = date.getDate();

        // 日付に対応する列を取得
        var columns;
        if (dateColumnCache[month]) {
            columns = dateColumnCache[month];
        } else {
            columns = getColumnsForMonth(month);
            dateColumnCache[month] = columns;
        }

        var dateCol = columns.dateCol;
        var amountCol = columns.amountCol;

        // その月のデータを取得
        var dataRange = sheet.getRange(startRow, dateCol, numRows, 4);
        var data = dataRange.getValues();

        var currentDate = null;

        // 各行のデータを処理
        for (var i = 0; i < data.length; i++) {
            var row = data[i];

            // 行が空白行かどうかをチェック
            var isEmptyRow = row.every(function (cell) {
                return cell === null || cell.toString().trim() === '';
            });

            // 空白行が検出されたらループを終了
            if (isEmptyRow) {
                break;
            }

            var dateCell = row[0];
            var category = row[1];
            var name = row[2];
            var amount = row[3];

            // 日付が空白でない場合、現在の日付を更新
            if (dateCell && dateCell.toString().trim() !== '') {
                // 日付が文字列の場合、Dateオブジェクトに変換
                if (typeof dateCell === 'string') {
                    currentDate = parseDate(dateCell, year);
                } else if (Object.prototype.toString.call(dateCell) === '[object Date]') {
                    currentDate = new Date(dateCell.getFullYear(), dateCell.getMonth(), dateCell.getDate());
                } else {
                    continue; // 日付の形式が不明な場合はスキップ
                }
            }

            // 現在の日付が対象の日付と一致する場合、データを収集
            if (currentDate && currentDate.getTime() === date.getTime()) {
                dataEntries.push({
                    date: currentDate,
                    category: category,
                    name: name,
                    amount: amount
                });
            }
        }
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

// Gemini に支出傾向の簡易分析を依頼するメソッド
function analyzeExpensesWithGemini(dataEntries, totalAmount, adjustedBudget, percentage) {
    var apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
    if (!apiKey) {
        Logger.log("Gemini API key is not set in script properties.");
        return null;
    }

    // モデルは Script Properties で上書き可能。なければ listModels から generateContent 可能なモデルを自動取得。
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
    var apiVersions = ["v1beta", "v1"];
    var expenseLines = dataEntries.map(function (entry) {
        return formatDate(entry.date) + " " + entry.name + " " + entry.amount + "円";
    });

    var prompt = [
        "あなたはプロの家計管理アドバイザー兼フィナンシャルプランナーです。金銭感覚の改善を目的としたコーチとして、冷静な分析と、時には優しく、時には厳しく指導してください。単なる分析にとどまらず、**「行動に落とし込める改善提案」**を重視してください。感情的にならず、客観的かつ現実的な判断で、飴と鞭を使い分けてください。以下は家計簿の支出データです。日本語で今日までの傾向を数個と、次に意識すべき点を数個、箇条書きでまとめてください。Markdownのフォーマットは使わず、プレーンな文字と絵文字のみで出力してください。出力は800文字以内でコンパクトにお願いします。",
        "週予算: " + adjustedBudget + "円 / これまでの支出: " + totalAmount + "円 (" + percentage.toFixed(1) + "%)",
        "支出一覧:",
        expenseLines.join("\n")
    ].join("\n");

    var payload = {
        contents: [{
            parts: [{
                text: prompt
            }]
        }],
        generationConfig: {
            temperature: 0.4,
            maxOutputTokens: 256
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
function sendWeeklySummaryEmail(dateRangeStr, totalAmount, dataEntries, difference, percentage, adjustedBudget, isStaging, action) {
    var emailAddress = "TARGET_EMAIL_ADDRESS";

    // 予算差分の符号を設定
    var differenceSign = difference >= 0 ? "+" : "-";
    var differenceAbs = Math.abs(difference);

    // 予算割合を小数点以下2桁で表示
    var percentageStr = percentage.toFixed(2);

    // トップ5の支出を計算
    var top5Entries = dataEntries.slice(); // 配列をコピー
    top5Entries.sort(function (a, b) {
        return parseFloat(b.amount) - parseFloat(a.amount);
    });
    top5Entries = top5Entries.slice(0, 5);

    // メール本文を指定のフォーマットで作成
    var body = "";
    body += "◆ " + dateRangeStr + " の週次サマリー\n\n";
    body += "合計支出は " + totalAmount + " 円です。\n\n";
    body += "* 設定予算： " + adjustedBudget + " 円\n";
    body += "* 予算差分：" + differenceSign + differenceAbs + "円\n";
    body += "* 予算割合：" + percentageStr + "%\n\n";
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

    var geminiAnalysis = analyzeExpensesWithGemini(dataEntries, totalAmount, adjustedBudget, percentage);
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

    // Gemini分析
    var geminiAnalysis = analyzeExpensesWithGemini(dataEntries, totalAmount, adjustedBudget, percentage);


    // メールの件名と本文を作成
    var subject = (isStaging ? "<test>" : "")
        + "家計簿日次レポート（" + formatDate(currentDate) + "）";
    var body = formatDate(datesInWeek[0]) + " から " + formatDate(currentDate) + " までの合計支出は " + totalAmount + " 円です。\n";
    body += "予算の " + percentage.toFixed(2) + "% を使用しました。\n";
    body += "（設定予算：" + adjustedBudget + "円）\n\n";

    body += "詳細:\n";
    dataEntries.forEach(function (entry) {
        body += "・" + formatDate(entry.date) + " - " + entry.name + ": " + entry.amount + "円\n";
    });
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
