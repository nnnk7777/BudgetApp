function doPost(e) {
    let result = "init";

    try {
        const jsonString = e.postData.contents;
        const data = JSON.parse(jsonString);

        const hash = data.hash;
        const action = data.action;

        const scriptHash = PropertiesService.getScriptProperties().getProperty("HASH");

        // 受け取ったハッシュが想定通りの値だった場合、メールサマリ生成を実行
        if (hash === scriptHash) {
            switch (action) {
                case 'text':
                case 'mail':
                    result = calculateWeeklyExpenses(action);
                    break;
                case 'add':
                    item = data.item;

                    title = item.title;
                    amount = item.amount;

                    result = addExpenseRecord(title, amount);
                    break;
                default:
                    throw new Error('actionが定義されていません');
            }
        }
    } catch (error) {
        result = error.message;
        Logger.log(error);
    } finally {
        // レスポンスを作成
        var output = ContentService.createTextOutput(result);
        output.setMimeType(ContentService.MimeType.JSON);

        return output;
    }
}
