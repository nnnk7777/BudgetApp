class ApiResponse {
    /** @type {string} */
    message;
    /** @type {Error} */
    error;
    /** @type {GoogleAppsScript.Content.TextOutput} */
    output;

    constructor() {
        this.message = 'response message is not initialized';
        this.error = null;
    }

    /**
     * @param {Error} error 
     */
    setError(error) {
        this.error = error;
        this.message = error.message;
    }

    setOutput() {
        this.output = ContentService.createTextOutput(this.message);
        this.output.setMimeType(ContentService.MimeType.JSON);
    }

}