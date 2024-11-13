# BudgetApp

Google スプレッドシートの家計簿を生成・管理するための GAS スクリプト

## 準備

-   家計簿をつけるスプレッドシートを作成しておいてください。また、シート名を指定してください。
    -   例. `金銭メモ2024`
-   シートの「拡張機能 -> App Script」から GAS プロジェクトを作成しておいてください。
-   clasp と Typescript をインストールしてください。
-   `main.ts` の `speadSheetName` を作成したシートと同様の名称にしてください。
    -   例. `金銭メモ2024`
-   ローカルで `clasp login` を実行しておき、`~/.clasprc.json` が生成されていることを確認してください。
-   `~/.clasprc.json` 内の以下の値をそれぞれ Github の Secret として登録してください。
    -   これらの値が `.github/workflows/build-and-deploy-on-main.yml` で読み込まれて利用されます。

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
        "clientId": "CLIENT_ID として登録",
        "clientSecret": "CLIENT_SECRET として登録",
        "redirectUri": "http://localhost"
    },
    "isLocalCreds": false
}
```

-   作成済みの GAS プロジェクトの URL に含まれる script_id を Secret として登録してください。

    -   例. `https://script.google.com/u/0/home/projects/<SCRIPT_ID として登録>`
    -   この ID で指定したシートに対してデプロイされます。

-   GAS プロジェクトの設定から、スクリプトプロパティを設定してください。
    -   プロパティ：`HASH`
    -   値：`<ハッシュとして利用する値を設定>`

## 開発・デプロイ方法

-   `feat/**`ブランチを切って作業をしてください。
    -   `feat/**`ブランチの PR にコミットがされると、自動でアプリがビルドされます。ビルドが成功することを確認してください。
-   作成した PR を`main`ブランチにマージすると、SCRIPT_ID で指定した GAS プロジェクトにデプロイされます。
