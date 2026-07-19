# BudgetApp

## 概要

Googleスプレッドシートを家計簿として使うためのGoogle Apps Scriptプロジェクトです。支出・収入の記録、カテゴリ別集計、週次・月次の予算サマリを1つのスプレッドシートで管理します。

- Web APIから支出を追加し、カテゴリ一覧や未分類支出を取得できます。
- 日次・週次・月次のサマリをテキストまたはメールで出力できます。
- OpenAIを優先し、Geminiをフォールバックとして、家計サマリの助言・カレンダーメモの整理・未分類支出のカテゴリ推定に利用します。
- Googleカレンダーの予定支出とユーザー補足メモをサマリへ反映します。
- GASの手動実行から、スプレッドシートのレイアウト・入力規則・条件付き書式を再適用できます。

## ディレクトリ構成

-   TypeScriptのレイアウト生成処理は [src/layout/main.ts](./src/layout/main.ts) を入口に `build.ts` でGAS用JavaScriptへ変換します。
-   `src/config/` はレイアウト生成とGAS実行時処理で共有する設定を管理します。
-   `scripts/` 配下はAPI、トリガー、AI分析などのGAS処理で、デプロイ時にそのまま配置されます。
-   `scripts/` の責務分割は [scripts/README.md](./scripts/README.md) を参照してください。

## 開発準備

-   家計簿をつけるスプレッドシートを作成しておいてください。また、シート名を指定してください。
    -   例. `金銭メモ2024`
-   シートの「拡張機能 -> App Script」から GAS プロジェクトを作成しておいてください。
-   プロジェクト直下で以下を実行し、`clasp` と `typescript` を含む依存パッケージをインストールしてください。グローバルインストールは不要です。

```sh
npm ci
```
-   ローカルでスタイルを再適用する場合は、対象のスプレッドシートに紐づくGASプロジェクトから実行してください。`src/layout/main.ts` は実行中のスプレッドシートを対象にします。
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

## AI設定

支出サマリ、カレンダーメモの分類・整形、未分類支出のカテゴリ推定でAIを利用します。OpenAIを優先し、APIキー未設定・APIエラー・空応答の場合はGeminiを試します。

- OpenAIを利用する場合は、[OpenAI Platform](https://platform.openai.com/api-keys) でAPIキーを発行し、GASのスクリプトプロパティに追加してください。
    -   プロパティ：`OPENAI_API_KEY`
    -   値：`<発行したAPIキーを設定>`
    -   任意: `OPENAI_MODEL`
    -   例: `gpt-5.4-mini`

- Geminiをフォールバックとして利用する場合は、[Google AI Studio](https://aistudio.google.com/) でAPIキーを発行し、GASのスクリプトプロパティに追加してください。
    -   プロパティ：`GEMINI_API_KEY`
    -   値：`<発行したAPIキーを設定>`
    -   任意: `GEMINI_MODEL`

両方のプロバイダから応答を取得できない場合、サマリはAI助言なしで返します。`autofill_uncategorized` は、OpenAI・Geminiの両方のキーがない場合にエラーを返します。

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
