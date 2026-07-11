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
    var cleanedText = cleanAiResponseText(text);
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
