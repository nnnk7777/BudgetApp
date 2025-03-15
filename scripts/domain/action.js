const Actions = Object.freeze({
    ADD: 'add',
    CATEGORIES: 'categories',
    TEXT: 'text',
    MAIL: 'mail'
});


class Action {
    /** @type {string} */
    name;

    constructor(name) {
        if (!Object.values(Actions).includes(name) || !name) {
            throw new Error('Invalid action');
        }

        this.name = name;
    }

}
