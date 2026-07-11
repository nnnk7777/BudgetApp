function buildCategorySuggestionPrompt(items, historyItems, categoryNames) {
    var requiredIds = items.map(function (item) {
        return item.id;
    });
    var categoryNamesJson = JSON.stringify(categoryNames);
    var requiredIdsText = requiredIds.join(", ");
    var itemsJson = JSON.stringify(items);
    var historyItemsJson = JSON.stringify(historyItems);

    return `あなたは家計簿カテゴリ分類アシスタントです。
出力は1行1件のプレーンテキストのみで、説明文・JSON・コードブロック・Markdownは禁止です。
カテゴリは必ず次の一覧から選んでください。自信が低い場合は category に null を入れてください。
カテゴリ一覧: ${categoryNamesJson}
出力必須ID一覧:
${requiredIdsText}
対象の未分類レコード一覧:
${itemsJson}
前月1ヶ月分の分類済み履歴一覧:
${historyItemsJson}
各要素について必ず次の形式で返してください:
id|category|confidence|reason
例:
6_42|コンビニ・お菓子|0.95|ファミマ履歴
6_43|null|0.10|情報不足
ルール:
- 1行1件で返す
- 区切り文字は半角パイプ | のみを使う
- reason は10文字以内
- confidence は 0 から 1 の数値
- タイトルや前月履歴だけでは判断が難しい場合は category を null にする
- 同じ店名やサービス名の履歴があれば優先して参考にする
- amount や title から一般常識で推定できても、自信が低ければ null にする
- reason に | を含めない
- 出力必須ID一覧の全IDを必ず1回ずつ出力する
- 判定不能でも null 行を省略しない
- 対象件数と同じ件数の行を返す
- 出力順は出力必須ID一覧と同じ順にする`;
}
