function doPost(e) {
    const request = new ApiRequest(JSON.parse(e.postData.contents));
    const response = new ApiResponse();

    try {
        const scriptHash = PropertiesService.getScriptProperties().getProperty("HASH");

        // 受け取ったハッシュが想定通りの値だった場合、メールサマリ生成を実行
        if (request.hash === scriptHash) {
            switch (request.action.name) {
                case Actions.TEXT:
                case Actions.MAIL:
                    result = calculateExpensesSummary(action);
                    break;
                case Actions.ADD:
                    result = addExpenseRecord(request.registerData);
                    break;
                case Actions.CATEGORIES:
                    result = fetchCategories();
                    break;
                default:
                    throw new Error('Invalid action');
            }
        }
    } catch (error) {
        response.setError(error);
        Logger.log(response.error);
    } finally {
        response.setOutput();

        return response.output;
    }
}
