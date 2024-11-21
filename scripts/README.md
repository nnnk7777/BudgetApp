# 概要

## API

### [handleApi.js](handleApi.js)

-   `doPost`で POST リクエストを受け付けて、指定されたパラメータによって後続処理を決定する。

### [addNewExpense.js](addNewExpense.js)

-   `add`アクションによって実行される。
-   金額と内容を受け取って該当の日付の直下にデータを追加する。

### [sendSummaryMail.js](sendSummaryMail.js)

-   `mail`または`text`アクションによって実行される。
-   予算（45000 円/週）に対する割合や差分を取得する。
-   メールとして送信または文字列をレスポンスとして返却する。

### [fetchCategories.js](fetchCategories.js)

-   `categoreis`アクションによって実行される。
-   `categories.js`内のデータから名称の一覧を取得する

## 編集をトリガーとした処理

### [formatDateAndPriceNumbers.js](formatDateAndPriceNumbers.js)

-   onEdit を利用して、日付と金額の列に修正があった際に必要に応じて自動でフォーマットを行う
-   日付
    -   `mmdd` で入力された場合、それを `mm/dd` に修正する
-   金額
    -   全角数字で入力されたものを半角に変換する

## デプロイ用スクリプト

### /deployment

Github Actions 上から利用されるスクリプト

-   クレデンシャルな情報を Github の Secret から取得した上で設定ファイルを生成する
