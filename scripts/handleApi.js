function doPost(e) {
    let result = "init";

    try {
        const data = parseApiRequest(e);
        verifyApiHash(data.hash);
        result = dispatchApiAction(data);
    } catch (error) {
        result = buildApiErrorResponse(error);
        Logger.log(error);
    } finally {
        return buildJsonTextOutput(result);
    }
}
