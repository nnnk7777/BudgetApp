# 概要

## [sendSummaryMail.js](sendSummaryMail.js)

-   予算（45000 円/週）に対する割合や差分を取得する。
-   メールの設定は GAS の Web 画面側から手動で設定する。

## [formatDateAndPriceNumbers.js](formatDateAndPriceNumbers.js)

-   onEdit を利用して、日付と金額の列に修正があった際に必要に応じて自動でフォーマットを行う
-   日付
    -   `mmdd` で入力された場合、それを `mm/dd` に修正する
-   金額
    -   全角数字で入力されたものを半角に変換する

## /deployment

Github Actions 上から利用されるスクリプト

-   クレデンシャルな情報を Github の Secret から取得した上で設定ファイルを生成する
