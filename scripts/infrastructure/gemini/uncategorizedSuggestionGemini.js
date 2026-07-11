function buildCategorySuggestionPrompt(items, historyItems, categoryNames) {
    var requiredIds = items.map(function (item) {
        return item.id;
    });

    return [
        "あなたは家計簿カテゴリ分類アシスタントです。",
        "出力は1行1件のプレーンテキストのみで、説明文・JSON・コードブロック・Markdownは禁止です。",
        "カテゴリは必ず次の一覧から選んでください。自信が低い場合は category に null を入れてください。",
        "カテゴリ一覧: " + JSON.stringify(categoryNames),
        "出力必須ID一覧:",
        requiredIds.join(", "),
        "対象の未分類レコード一覧:",
        JSON.stringify(items),
        "前月1ヶ月分の分類済み履歴一覧:",
        JSON.stringify(historyItems),
        "各要素について必ず次の形式で返してください:",
        "id|category|confidence|reason",
        "例:",
        "6_42|コンビニ・お菓子|0.95|ファミマ履歴",
        "6_43|null|0.10|情報不足",
        "ルール:",
        "- 1行1件で返す",
        "- 区切り文字は半角パイプ | のみを使う",
        "- reason は10文字以内",
        "- confidence は 0 から 1 の数値",
        "- タイトルや前月履歴だけでは判断が難しい場合は category を null にする",
        "- 同じ店名やサービス名の履歴があれば優先して参考にする",
        "- amount や title から一般常識で推定できても、自信が低ければ null にする",
        "- reason に | を含めない",
        "- 出力必須ID一覧の全IDを必ず1回ずつ出力する",
        "- 判定不能でも null 行を省略しない",
        "- 対象件数と同じ件数の行を返す",
        "- 出力順は出力必須ID一覧と同じ順にする"
    ].join("\n");
}

function requestCategorySuggestionsWithAI(items, historyItems, categoryNames, month, options) {
    var prompt = buildCategorySuggestionPrompt(items, historyItems, categoryNames);
    var aiResult = generatePreferredAiText(prompt, {
        temperature: 0.1,
        maxOutputTokens: 1200
    }, {
        logContext: "uncategorized_suggestion_month_" + month
    });
    var responseText = aiResult.text;
    var parseResult;

    if (!responseText) {
        parseResult = {
            status: 'empty_response',
            responsePreview: '',
            extractedJsonPreview: null,
            suggestions: []
        };
    } else {
        parseResult = parseCategorySuggestionResponse(responseText, items, categoryNames, {
            month: month,
            debug: options.debug
        });
    }

    return {
        provider: aiResult.provider,
        model: aiResult.model,
        usedFallback: aiResult.usedFallback,
        responseText: responseText,
        parseResult: parseResult
    };
}

function parseCategorySuggestionResponse(text, items, categoryNames, context) {
    context = context || {};
    var cleanedText = cleanGeminiResponseText(text);
    var lineParseResult = parseCategorySuggestionPipeResponse(cleanedText, items, categoryNames);
    if (lineParseResult.ok) {
        return {
            status: lineParseResult.warnings.length ? 'ok_with_warnings' : 'ok',
            responsePreview: buildResponsePreview(cleanedText),
            extractedJsonPreview: null,
            suggestions: lineParseResult.suggestions,
            parseWarnings: lineParseResult.warnings
        };
    }

    var parsed = tryParseJson(cleanedText);
    var extractedJsonText = null;

    if (!parsed.ok) {
        extractedJsonText = extractJsonArrayText(cleanedText);
        if (extractedJsonText && extractedJsonText !== cleanedText) {
            parsed = tryParseJson(extractedJsonText);
        }
    }

    if (!parsed.ok) {
        Logger.log(
            "カテゴリ推定応答の解析に失敗: month=" +
                (context.month || "unknown") +
                " pipeError=" +
                lineParseResult.errorDetail +
                " jsonError=" +
                parsed.errorMessage +
                " response=" +
                cleanedText
        );
        return {
            status: 'response_parse_failed',
            parseError: lineParseResult.errorDetail + " / " + parsed.errorMessage,
            responsePreview: buildResponsePreview(cleanedText),
            extractedJsonPreview: extractedJsonText ? buildResponsePreview(extractedJsonText) : null,
            parseWarnings: lineParseResult.warnings || [],
            suggestions: items.map(function (item) {
                return buildSkippedSuggestion(
                    item,
                    null,
                    buildDetailedReason('AI応答を解析できませんでした', lineParseResult.errorDetail || parsed.errorMessage),
                    {
                        errorCode: 'response_parse_failed',
                        errorDetail: (lineParseResult.errorDetail || 'Pipe parse failed.') + ' / ' + parsed.errorMessage,
                        responsePreview: buildResponsePreview(cleanedText)
                    }
                );
            })
        };
    }

    if (!Array.isArray(parsed.value)) {
        Logger.log(
            "カテゴリ推定JSONが配列形式ではありません: month=" +
                (context.month || "unknown") +
                " response=" +
                cleanedText
        );
        return {
            status: 'response_not_array',
            responsePreview: buildResponsePreview(cleanedText),
            extractedJsonPreview: extractedJsonText ? buildResponsePreview(extractedJsonText) : null,
            parseWarnings: lineParseResult.warnings || [],
            suggestions: items.map(function (item) {
                return buildSkippedSuggestion(
                    item,
                    null,
                    buildDetailedReason('AI応答が配列形式ではありません', 'AI response was valid JSON but not an array.'),
                    {
                        errorCode: 'response_not_array',
                        errorDetail: 'AI response was valid JSON but not an array.',
                        responsePreview: buildResponsePreview(cleanedText)
                    }
                );
            })
        };
    }

    var categorySet = {};
    categoryNames.forEach(function (name) {
        categorySet[name] = true;
    });

    var itemMap = {};
    items.forEach(function (item) {
        itemMap[item.id] = item;
    });

    return {
        status: 'ok',
        responsePreview: buildResponsePreview(cleanedText),
        extractedJsonPreview: extractedJsonText ? buildResponsePreview(extractedJsonText) : null,
        suggestions: parsed.value.map(function (entry) {
            var item = itemMap[entry.id];
            if (!item) {
                return null;
            }

            var suggestedCategory = entry.suggestedCategory;
            if (suggestedCategory !== null && suggestedCategory !== undefined) {
                suggestedCategory = String(suggestedCategory).trim();
                if (!categorySet[suggestedCategory]) {
                    suggestedCategory = null;
                }
            } else {
                suggestedCategory = null;
            }

            return {
                id: item.id,
                month: item.month,
                row: item.row,
                date: item.date,
                title: item.title,
                amount: item.amount,
                suggestedCategory: suggestedCategory,
                confidence: normalizeConfidence(entry.confidence),
                reason: entry.reason ? String(entry.reason).trim() : ''
            };
        }).filter(function (entry) {
            return entry !== null;
        })
    };
}

function buildSkippedSuggestion(item, suggestedCategory, reason, diagnostics) {
    var result = {
        id: item.id,
        month: item.month,
        row: item.row,
        date: item.date,
        title: item.title,
        amount: item.amount,
        suggestedCategory: suggestedCategory,
        confidence: 0,
        reason: reason
    };

    diagnostics = diagnostics || {};
    if (diagnostics.errorCode) {
        result.errorCode = diagnostics.errorCode;
    }
    if (diagnostics.errorDetail) {
        result.errorDetail = diagnostics.errorDetail;
    }
    if (diagnostics.responsePreview !== undefined) {
        result.responsePreview = diagnostics.responsePreview;
    }

    return result;
}

function tryParseJson(text) {
    try {
        return {
            ok: true,
            value: JSON.parse(text)
        };
    } catch (error) {
        return {
            ok: false,
            errorMessage: error && error.message ? String(error.message) : String(error)
        };
    }
}

function parseCategorySuggestionPipeResponse(text, items, categoryNames) {
    var itemMap = {};
    items.forEach(function (item) {
        itemMap[item.id] = item;
    });

    var categorySet = {};
    categoryNames.forEach(function (name) {
        categorySet[name] = true;
    });

    var lines = text
        .split("\n")
        .map(function (line) {
            return line.trim();
        })
        .filter(function (line) {
            return line !== '';
        })
        .filter(function (line) {
            return /^[0-9]+_[0-9]+\|/.test(line);
        });

    if (!lines.length) {
        return {
            ok: false,
            errorDetail: 'Pipe形式の応答行を検出できませんでした',
            warnings: []
        };
    }

    var suggestions = [];
    var warnings = [];
    for (var i = 0; i < lines.length; i++) {
        var parts = lines[i].split("|");
        if (parts.length < 4) {
            warnings.push('Pipe形式の列数が不足しています: ' + lines[i]);
            continue;
        }

        var id = parts[0].trim();
        var category = parts[1].trim();
        var confidence = parts[2].trim();
        var reason = parts.slice(3).join("|").trim();
        var item = itemMap[id];

        if (!item) {
            continue;
        }

        if (category.toLowerCase() === 'null' || category === '') {
            category = null;
        } else if (!categorySet[category]) {
            category = null;
        }

        suggestions.push({
            id: item.id,
            month: item.month,
            row: item.row,
            date: item.date,
            title: item.title,
            amount: item.amount,
            suggestedCategory: category,
            confidence: normalizeConfidence(confidence),
            reason: reason
        });
    }

    if (!suggestions.length) {
        return {
            ok: false,
            errorDetail: warnings.length
                ? warnings[0]
                : 'Pipe形式の有効な応答行を解釈できませんでした',
            warnings: warnings
        };
    }

    return {
        ok: true,
        suggestions: suggestions,
        warnings: warnings
    };
}

function extractJsonArrayText(text) {
    var startIndex = text.indexOf('[');
    var endIndex = text.lastIndexOf(']');
    if (startIndex === -1 || endIndex === -1 || endIndex < startIndex) {
        return null;
    }
    return text.substring(startIndex, endIndex + 1).trim();
}

function cleanGeminiResponseText(text) {
    return String(text)
        .replace(/```[\s\S]*?\n/g, "")
        .replace(/```/g, "")
        .trim();
}

function buildResponsePreview(text) {
    if (!text) {
        return '';
    }

    var normalized = String(text)
        .replace(/\s+/g, ' ')
        .trim();

    if (normalized.length <= 300) {
        return normalized;
    }

    return normalized.substring(0, 300) + '...';
}

function buildDetailedReason(baseReason, detail) {
    if (!detail) {
        return baseReason;
    }

    var normalizedDetail = String(detail).replace(/\s+/g, ' ').trim();
    if (!normalizedDetail) {
        return baseReason;
    }

    var maxDetailLength = 120;
    if (normalizedDetail.length > maxDetailLength) {
        normalizedDetail = normalizedDetail.substring(0, maxDetailLength) + '...';
    }

    return baseReason + '（' + normalizedDetail + '）';
}
