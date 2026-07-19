function buildExpenseSummaryPrompt(
    dataEntries,
    totalAmount,
    adjustedBudget,
    percentage,
    baseDate,
    categoryRankingLines,
    weeklyAnalysisMode,
    weeklyAnalysisModeGuidance,
    weeklyBudgetCarryoverMemo,
    weeklyBudgetCarryoverGuidance,
    plannedExpenseLabel,
    upcomingExpenseLines,
    contextualMemoLabel,
    contextualMemoLines
) {
    var expenseLines = dataEntries.map(function (entry) {
        return (
            "日付: " +
            formatDate(entry.date) +
            " / カテゴリ: " +
            (entry.category || "未分類") +
            " / 名称: " +
            (entry.name || "") +
            " / 金額: " +
            entry.amount +
            "円"
        );
    });
    var analysisDateContext = buildAnalysisDateContext(baseDate);
    var categoryRankingText = categoryRankingLines.length ? categoryRankingLines.join(" / ") : "なし";
    var weeklyAnalysisModeText = formatWeeklyAnalysisModeForPrompt(weeklyAnalysisMode);
    var weeklyBudgetCarryoverMemoText = formatWeeklyBudgetCarryoverMemoForPrompt(weeklyBudgetCarryoverMemo);
    var upcomingExpenseText = upcomingExpenseLines.length ? upcomingExpenseLines.join(" / ") : "なし";
    var contextualMemoText = contextualMemoLines.length ? contextualMemoLines.join(" / ") : "なし";
    var prompt = `あなたはプロの家計管理アドバイザーです。挨拶や自己紹介は禁止です。金銭感覚の改善を目的としたコーチとして、冷静な分析と、時には優しく、時には厳しく指導してください。1週間分の支出について、予算を超えないようアドバイスをください。カジュアルな敬語て対応してください。
レシートは保管していませんが、代わりに全ての支出・収入をスプレッドシートに記録しています。単なる分析にとどまらず、「行動に落とし込める改善提案」を重視してください。感情的にならず、客観的かつ現実的な判断で、飴と鞭を使い分けてください。
通勤の交通費は給与で補填されます。食事はスーパーでまとめ買いした上でほぼ自炊しており、外食は友人と会う時が多いです。
Googleカレンダーの今後の予定メモに書かれた支出予定も考慮して助言してください。近いうちに大きな支出予定があるなら、今週の節約を強めに促してください。
ユーザーの補足メモには、支出に対する後悔や不安、忘れていた定期購入、週予算と別管理にしたい支出についての迷いなど、内面的な事情や管理方針の相談が書かれます。メモのニュアンスを受け止め、責めずに具体的な次の行動へつながる助言に反映してください。
補足メモだけを根拠に支出を週予算の対象外と断定したり、金額を推測したりしてはいけません。別管理が適切そうな場合も、確定事項として扱わず、管理方法の選択肢として提案してください。
前週の予算差分メモがあれば補助情報として参照してください。特に前週が大きく超過していた場合は、その影響を締めの一言だけで済ませず、今週の傾向分析と次に意識すべき点の両方に反映してください。
前週超過がある場合は、裁量支出や先送りできる支出への姿勢を普段より一段引き締めて提案してください。ただし今週の週予算、実支出、予定支出を優先し、前週差分だけで過度に断定したり、今週の予算を実質的に減額したような言い方はしないでください。
前週が予算内に収まっていた場合でも、安心しすぎる助言にはせず、今週の数値を主軸に冷静に判断してください。
${analysisDateContext}
与えられた分析基準日・対象期間・支出一覧だけを根拠に判断してください。曜日や週の進捗を勝手に推測しないでください。
特に、基準日が火曜日なのに『今日が日曜日』と書いたり、対象週が月をまたぐのに月末で週が終わったかのように扱うのは禁止です。
1週間は月曜始まり・日曜終わりで考えて、今日までの傾向と次に意識すべき点を、見出し付きの短いセクションでまとめてください。
出力形式のルール:
- 見出しは絵文字付きにする。例: 📊 現状 / ⚠️ 注意点 / ✅ 次の行動
- セクションは2〜3個まで
- 各セクションの本文は1〜2個の箇条書きまで
- 箇条書きは最大5個までに収め、それぞれに短い補足を付ける
- Markdown記法は使わず、プレーンな文字と絵文字のみで出力する
- 全体で500文字以内に収める
週予算: ${adjustedBudget}円 / これまでの支出: ${totalAmount}円 (${percentage.toFixed(1)}%)
カテゴリ別支出ランキング: ${categoryRankingText}
今週の分析モード: ${weeklyAnalysisModeText}
分析モードの反映方針: ${weeklyAnalysisModeGuidance}
前週の予算差分メモ: ${weeklyBudgetCarryoverMemoText}
前週差分の反映方針: ${weeklyBudgetCarryoverGuidance}
${plannedExpenseLabel}: ${upcomingExpenseText}
${contextualMemoLabel}: ${contextualMemoText}
以下は支出一覧です。各行には日付・カテゴリ・名称・金額を含みます。名称だけで内容が不明瞭な場合はカテゴリから内容を推定してください:
${expenseLines.join("\n")}`;

    return prompt;
}

function buildPlannedExpenseRecordCheckPrompt(plannedExpense, candidateEntries) {
    var plannedExpenseJson = JSON.stringify({
        date: formatDate(plannedExpense.date),
        title: plannedExpense.title,
        memo: plannedExpense.memo
    });
    var candidateEntriesJson = JSON.stringify(candidateEntries.map(function (entry) {
        return {
            date: formatDate(entry.date),
            category: entry.category || "",
            name: entry.name || "",
            amount: entry.amount
        };
    }));

    return `以下のGoogleカレンダーの支出予定が、家計簿に記録済みの支出として扱ってよいか判定してください。
保守的に判定し、自信がない場合は NO にしてください。
出力は YES か NO のどちらか1語のみです。
予定:
${plannedExpenseJson}
家計簿候補:
${candidateEntriesJson}`;
}

function buildCalendarMemoClassificationPrompt(calendarMemos) {
    var calendarMemosJson = JSON.stringify(calendarMemos.map(function (entry, index) {
        return {
            index: index,
            title: entry.title,
            memo: entry.memo
        };
    }));

    return `以下はGoogleカレンダー予定のタイトルとメモです。
各予定の意図を分類し、家計サマリに役立つ情報だけを短く整形してください。
各要素について intent と cleanedMemo を返してください。
ルール:
- intent は planned_expense、contextual_note、reservation_info、ignore のいずれかにする
- planned_expense: 今後発生しそうな支出。金額がなくても、外食・買い物・交通・宿泊など支出意図が明確なら該当する
- タイトルまたはメモに明示的な金額（例: 15000円）がある予定は、定期購入を忘れていた・支出として痛い・別管理したいという相談が併記されていても、必ず planned_expense にする。金額は cleanedMemo に必ず残す
- contextual_note: ユーザーが書いた事情・感情・意図・管理方針の相談。支出額には含めないが、助言の文脈として役立つ情報。場所や同行者のメモも含む
- reservation_info: レストランや美容院などの予約情報。日時・店名・場所・サービス名が有用なら cleanedMemo に残す。そうでなければ空文字にする
- ignore: Googleの自動追記、URL、案内文、会議通知など、家計サマリに不要な情報
- URL、Googleの案内文、自動生成の説明、予約メール由来の定型文は削除する
- contextual_note は、ユーザーの言い回しにある後悔、不安、別管理したい意図などのニュアンスを落とさず、短く整形する。事実や金額を補完・推測しない
- 金額、店名、場所、支出目的、ユーザーの事情など意味のある情報だけを残し、タイトルと同じ内容は重複させない
- 出力は各行 \`index|intent|cleanedMemo\` のみ
- 説明文、Markdown、コードブロック、jsonという語は出力しない
- 例: \`0|planned_expense|ランチ 1200円くらい\`
- 例: \`1|planned_expense|定期購入プロテイン 15000円。忘れていたため支出として痛い\`
- 例: \`2|contextual_note|ふるさと納税のため週予算とは別管理にしたい\`
- 例: \`3|reservation_info|美容院 カット\`
${calendarMemosJson}`;
}

function buildWeeklyAnalysisModePrompt(title, description) {
    var safeTitle = String(title || "");
    var safeDescription = String(description || "");

    return `以下のGoogleカレンダー予定が、その週の家計分析を厳しめにする『節制モード』指定かどうかを判定してください。
出力は YES か NO のどちらか1語のみです。
節制モードと判断してよい例: 節制モード, 節制, 引き締め週, 節約強化週間。
支出予定、通常の予定、意味が曖昧で節制意図が読めないものは NO にしてください。
タイトル: ${safeTitle}
説明: ${safeDescription}`;
}

function formatWeeklyAnalysisModeForPrompt(modeResult) {
    if (!modeResult || modeResult.mode !== WEEKLY_ANALYSIS_MODE_FRUGAL) {
        return "通常モード";
    }

    return modeResult.label;
}

function buildWeeklyAnalysisModeGuidanceForPrompt(modeResult) {
    if (!modeResult || modeResult.mode !== WEEKLY_ANALYSIS_MODE_FRUGAL) {
        return "通常モード。必要以上に厳しくしすぎず、数値と予定支出に基づいて冷静に評価する。";
    }

    return "節制モード。通常より明確に厳しめに評価し、裁量支出、先送りできる支出、習慣化すると危険な出費には甘くしない。数値がまだ良くても油断を促す言い方は避け、今のうちに削れる支出や抑えるべき行動を具体的に指摘する。ただし感情的な説教にはせず、改善行動が明確になる実務的で手厳しい助言を優先する。";
}

function formatWeeklyBudgetCarryoverMemoForPrompt(memo) {
    if (!memo) {
        return "なし";
    }

    return `${memo.dateRangeStr || "対象週不明"} / 予算差分 ${memo.difference >= 0 ? "+" : "-"}${Math.abs(memo.difference)}円 / 実支出 ${memo.totalAmount}円 / 週予算 ${memo.adjustedBudget}円`;
}

function buildWeeklyBudgetCarryoverGuidanceForPrompt(memo) {
    var overrunRatio;

    if (!memo) {
        return "前週差分なし。今週の実績と予定支出だけで判断する。";
    }

    if (memo.difference <= 0) {
        return "前週は予算内。評価としては軽く触れる程度に留め、今週を緩めすぎない。";
    }

    overrunRatio = memo.adjustedBudget ? (memo.difference / memo.adjustedBudget) : 0;
    if (overrunRatio >= 0.25) {
        return "前週は大きな予算超過。今週は全体的に一段厳しめのトーンで、節約余地や回避策を複数箇所で具体的に指摘する。ただし今週の数値が良好なら前週だけで悲観しすぎない。";
    }

    if (overrunRatio >= 0.1) {
        return "前週はやや大きめの予算超過。今週は通常より引き締め寄りに評価し、無理なく削れる支出を明確に示す。ただし今週の実績が落ち着いていれば過度には引きずらない。";
    }

    return "前週は軽度の予算超過。分析のどこかで一度は触れつつ、今週の数値を主軸にバランスよく助言する。";
}

function buildAnalysisDateContext(baseDate) {
    var weekRange = getWeekRange(baseDate);
    var weekday = getJapaneseWeekday(baseDate.getDay());
    var isWeekClosed = baseDate.getDay() === 0;
    var weekStatus = isWeekClosed
        ? "今週は日曜まで終了した確定値として扱う。"
        : "今週はまだ途中であり、" + weekday + "時点の途中経過として扱う。";

    return `分析基準日: ${formatDate(baseDate)} (${weekday}) / 分析対象の週: ${formatDate(weekRange.startDate)}〜${formatDate(weekRange.endDate)} の1週間 / 週の扱い: 1週間は月曜始まり・日曜終わり。月をまたいでも同じ週として扱う。年をまたぐ場合だけ、その年に含まれる日付までを対象にする。 / 分析時点: ${weekStatus}`;
}

function getJapaneseWeekday(dayIndex) {
    return ["日曜日", "月曜日", "火曜日", "水曜日", "木曜日", "金曜日", "土曜日"][dayIndex] || "";
}
