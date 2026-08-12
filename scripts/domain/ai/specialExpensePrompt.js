function buildSpecialExpenseReviewPrompt(entries) {
    var entryLines = entries.map(function (entry, index) {
        return [
            index,
            formatDate(entry.date),
            entry.name || "",
            entry.amount
        ].join("|");
    });

    return `あなたは家計簿の「特別費」カテゴリを厳格に監査するアシスタントです。
判定するのは支出そのものの善悪や必要性ではなく、「通常の生活費予算とは別枠にしてよい支出か」だけです。

承認してよいもの:
- 旅行の交通・宿泊・旅行に直接付随する費用
- 年額課金など、年に一度または低頻度で、あらかじめ予定化しやすい支出
- 医療、故障による修理・買い替えなど、緊急性または不可避性が高い低頻度支出

承認しないもの:
- 高額というだけの服、趣味、外食、ガジェット、Amazon購入
- 予定外だった、忘れていた、使いすぎたという事情だけの支出
- 名称から内容が判断できない支出

不明な場合は必ず rejected にする。支出を特別費にすることで通常予算を軽く見せないよう、厳しめに判定する。
各行は必ず次の形式で返す:
index|approved または rejected|理由（20文字以内）
説明文、Markdown、コードブロック、JSONは出力しない。

対象の特別費候補:
${entryLines.join("\n")}`;
}
