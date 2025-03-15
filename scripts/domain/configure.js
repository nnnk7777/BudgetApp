const Env = Object.freeze({
    STG: 'staging',
    PRD: 'production'
});

class Configure {
    env;
    mail;
    budget = 45000;

    constructor(env, mail) {
        if (!Object.values(Env).includes(env) || !env) {
            throw new Error('Invalid environment');
        }
        if (!mail || mail === '') {
            throw new Error('Invalid mail');
        }
        this.env = env;
        this.mail = mail;
    }

    adjustedBudget(dateList) {
        const numberOfDays = dateList.length;
        const adjustedBudget = Math.round((this.budget * numberOfDays / 7) / 100) * 100;

        this.budget = adjustedBudget;
    }

}
