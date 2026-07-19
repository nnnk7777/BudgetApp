# BudgetApp

## 概要

Google スプレッドシートの家計簿を生成・管理するための GAS スクリプト

## ディレクトリ構成

-   Typescript のファイルは、`build.ts`で Javascript に変換した上で GAS にデプロイされます。
-   `/scripts`直下のファイルはそのまま GAS にデプロイされ、API や編集時のトリガーとして利用されます。

```
.
├── main.ts
├── build.ts
├── Makefile
├── scripts
│   ├── entrypoints
│   │   ├── handleApi.js
│   │   ├── apiCommon.js
│   │   ├── 0_manualEntryPoints.js
│   │   ├── 1_reapplySheetStyle.js
│   │   ├── scheduledSummaryTriggers.js
│   │   └── formatDateAndPriceNumbers.js
│   ├── application
│   │   ├── addExpenseRecord.js
│   │   ├── fetchCategories.js
│   │   ├── expenseSummary.js
│   │   ├── monthlySummary.js
│   │   └── uncategorizedExpenses.js
│   ├── domain
│   │   └── ai
│   │       ├── expenseSummaryPrompts.js
│   │       ├── monthlySummaryPrompt.js
│   │       └── categorySuggestionPrompt.js
│   ├── infrastructure
│   │   ├── ai
│   │   │   ├── aiClient.js
│   │   │   ├── aiResponseParsers.js
│   │   │   ├── expenseSummaryAi.js
│   │   │   ├── monthlySummaryAi.js
│   │   │   ├── categorySuggestionAi.js
│   │   │   └── weeklyAnalysisModeAi.js
│   │   ├── gas
│   │   │   ├── scriptRuntime.js
│   │   │   ├── expenseSheetRepository.js
│   │   │   ├── monthlySheetRepository.js
│   │   │   ├── uncategorizedExpenseRepository.js
│   │   │   └── calendarRepository.js
│   │   ├── openai
│   │   │   └── openaiClient.js
│   │   └── gemini
│   │       └── geminiClient.js
│   ├── formatting
│   │   ├── summaryMessageFormatter.js
│   │   └── monthlySummaryFormatter.js
│   ├── utils
│   │   ├── expenseSummaryUtils.js
│   │   ├── summaryDateUtils.js
│   │   └── uncategorizedCommonUtils.js
│   └── deployment
│       ├── setup_claspjson.sh
│       └── setup_clasprcjson.sh
├── service
├── model
├── types
├── config
└── util
```

## 開発準備

-   家計簿をつけるスプレッドシートを作成しておいてください。また、シート名を指定してください。
    -   例. `金銭メモ2024`
-   シートの「拡張機能 -> App Script」から GAS プロジェクトを作成しておいてください。
-   clasp と Typescript をインストールしてください。
-   ローカルでスタイルを再適用する場合は、対象のスプレッドシートに紐づくGASプロジェクトから実行してください。`main.ts` は実行中のスプレッドシートを対象にします。
-   ローカルで `clasp login` を実行しておき、`~/.clasprc.json` が生成されていることを確認してください。
-   `~/.clasprc.json` 内の以下の値をそれぞれ Github の Secret として登録してください。
-   これらの値は [prod.yml](./.github/workflows/prod.yml) と [stg.yml](./.github/workflows/stg.yml) で読み込まれて利用されます。

```json
{
    "token": {
        "access_token": "ACCESS_TOKEN として登録",
        "refresh_token": "REFRESH_TOKEN として登録",
        "scope": "xxxxxxxxxx",
        "token_type": "xxxxxxxxxx",
        "id_token": "ID_TOKEN として登録",
        "expiry_date": 000000000
    },
    "oauth2ClientSettings": {
        "clientId": "CLIENTID として登録",
        "clientSecret": "CLIENTSECRET として登録",
        "redirectUri": "http://localhost"
    },
    "isLocalCreds": false
}
```

-   作成済みの GAS プロジェクトの URL に含まれる script_id を Secret として登録してください。

    -   例. `https://script.google.com/u/0/home/projects/<SCRIPT_ID として登録>`
    -   この ID で指定したシートに対してデプロイされます。

-   GAS プロジェクトの設定から、スクリプトプロパティを設定してください。
-   後述の API 実行の際に利用される値です。
    -   プロパティ：`HASH`
    -   値：`<ハッシュとして利用する値を設定>`

- OpenAI をメイン利用するため、OpenAI Platform で API Key を発行し、GAS プロジェクトの設定にスクリプトプロパティを追加してください。
- https://platform.openai.com/api-keys
    -   プロパティ：`OPENAI_API_KEY`
    -   値：`<発行したAPIキーを設定>`
    -   任意: `OPENAI_MODEL`
    -   例: `gpt-5.4-mini`

- Gemini は OpenAI 障害時のフォールバックとして残せます。必要なら以下も設定してください。
- https://aistudio.google.com/
    -   プロパティ：`GEMINI_API_KEY`
    -   値：`<発行したAPIキーを設定>`
    -   任意: `GEMINI_MODEL`

## 開発・デプロイ方法

-   `feat/**` または `fix/**` ブランチへプッシュすると、[stg.yml](./.github/workflows/stg.yml) がstg用GASプロジェクトへビルド・プッシュします。
-   `main` へマージすると、[prod.yml](./.github/workflows/prod.yml) が本番用GASプロジェクトへビルド・プッシュします。
-   ローカルでGAS成果物だけを確認する場合は `make build` を実行します。
-   ローカルからGASへ反映する場合は、clasp設定後に `make build-and-push` を実行します。

## API 利用方法

-   デプロイしたプロジェクトに対して、様々な処理を実行できます。
-   以下の body を利用して、POST リクエストを送信してください。

```json
{
    "action": "mail",
    "hash": "<ハッシュとして利用する値を設定>"
}
```

action に渡せる値は以下のとおりです。

| `action` | 詳細 |
| :------: | --- |
| `mail` | 日次サマリをメール送信します。日曜日は週次、月末は月次も追加送信します。 |
| `text` | `mail` と同じ自動判定のサマリを文字列で返します。 |
| `daily_mail` | 日次サマリだけをメール送信します。 |
| `weekly_mail` | 週次サマリだけをメール送信します。 |
| `monthly_mail` | 月次サマリだけをメール送信します。 |
| `daily_text` | 日次サマリだけを文字列で返します。 |
| `weekly_text` | 週次サマリだけを文字列で返します。 |
| `monthly_text` | 月次サマリだけを文字列で返します。 |
| `add` | 支出レコードを追加します。 |
| `categories` | カテゴリ一覧を返します。 |
| `list_uncategorized` | カテゴリ未設定の支出一覧を返します。 |
| `autofill_uncategorized` | OpenAIを優先し、失敗時はGeminiでカテゴリを推定して高信頼のものだけ自動反映します。 |

`add`の場合は、以下のような body を渡してください。

```json
{
    "action": "add",
    "item" : {
        "title": "<支出のタイトル>",
        "amount": <支出金額>,
        "category": "<カテゴリー名>"
    },
    "hash": "<ハッシュとして利用する値を設定>"
}
```

`amount` を優先して使用します。互換入力として `price` または `cost` も指定できます。

`list_uncategorized` の場合は、以下のような body を渡してください。

```json
{
    "action": "list_uncategorized",
    "hash": "<ハッシュとして利用する値を設定>"
}
```

`autofill_uncategorized` の場合は、以下のような body を渡してください。

```json
{
    "action": "autofill_uncategorized",
    "options": {
        "confidenceThreshold": 0.9,
        "limit": 20,
        "debug": true,
        "retryMissingIds": true
    },
    "hash": "<ハッシュとして利用する値を設定>"
}
```

`debug` を `true` にすると、AI 応答解析に失敗した際の詳細情報がレスポンスの `debug` および `skipped[*]` に含まれます。`retryMissingIds` は、AI応答から欠落したIDだけを再問い合わせするかを指定し、省略時は `true` です。

## 手動実行

-   GASの実行画面では [0_manualEntryPoints.js](./scripts/entrypoints/0_manualEntryPoints.js) のサマリ・未分類支出操作を実行できます。
-   [1_reapplySheetStyle.js](./scripts/entrypoints/1_reapplySheetStyle.js) の `reapplySheetStyleManual` は、現在開いているスプレッドシートの書式・入力規則・条件付き書式をクリアしてから再構築します。

## scripts構成

`/scripts` は現在、主に以下の責務で分かれています。

-   `entrypoints`
    -   GAS から直接呼ばれる入口
    -   `handleApi.js`
    -   `0_manualEntryPoints.js`
    -   `1_reapplySheetStyle.js`
    -   `scheduledSummaryTriggers.js`
    -   `formatDateAndPriceNumbers.js`
-   `application`
    -   action ごとの本処理
    -   `addExpenseRecord.js`
    -   `fetchCategories.js`
    -   `expenseSummary.js`
    -   `monthlySummary.js`
    -   `uncategorizedExpenses.js`
-   `domain/ai`
    -   AI 向けのプロンプト生成と入力整形
    -   `expenseSummaryPrompts.js`
    -   `monthlySummaryPrompt.js`
    -   `categorySuggestionPrompt.js`
-   `infrastructure/ai`
    -   AI プロバイダに依存しない呼び出し制御、応答解析、AI利用フロー
    -   `aiClient.js`
    -   `aiResponseParsers.js`
    -   `expenseSummaryAi.js`
    -   `monthlySummaryAi.js`
    -   `categorySuggestionAi.js`
    -   `weeklyAnalysisModeAi.js`
-   `infrastructure/gas`
    -   SpreadsheetApp / PropertiesService / runtime 依存
    -   `scriptRuntime.js`
    -   `expenseSheetRepository.js`
    -   `monthlySheetRepository.js`
    -   `calendarRepository.js`
    -   `uncategorizedExpenseRepository.js`
-   `infrastructure/openai`
    -   OpenAI Responses API client
    -   `openaiClient.js`
-   `infrastructure/gemini`
    -   Gemini API client。OpenAI障害時のフォールバックに利用する。
    -   `geminiClient.js`
-   `formatting`
    -   mail / text の本文生成
    -   `summaryMessageFormatter.js`
    -   `monthlySummaryFormatter.js`
-   `utils`
    -   純粋関数寄りの共通処理
    -   `expenseSummaryUtils.js`
    -   `summaryDateUtils.js`
    -   `uncategorizedCommonUtils.js`
