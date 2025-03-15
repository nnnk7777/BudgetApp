/**
 * @typedef {Object} RegisterData
 * @property {string} title - 登録データのタイトル
 * @property {number} amount - 登録データの金額
 * @property {string} category - 登録データのカテゴリ
 */

class ApiRequest {
    /** @type {string} */
    body;
    /** @type {string} */
    hash;
    /** @type {Action} */
    action;
    /** @type {RegisterData} */
    registerData = {};

    constructor(body) {
        if (!body) {
            throw new Error('body is required');
        }

        this.body = body;
        this.hash = body.hash;
        this.action = new Action(body.action);
        if (this.action.name === 'add') {
            this.setRegisterValues(body);
        }
    }

    setRegisterValues(body) {
        this.registerData.title = body.title;
        this.registerData.amount = body.amount;
        this.registerData.category = body.category;
    }

}
