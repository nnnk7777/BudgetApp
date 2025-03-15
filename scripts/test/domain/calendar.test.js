const Calendar = require('../../domain/calendar');

describe('calendar#getDatesInWeek', () => {
    let calendar = new Calendar();

    test('平日の入力（例：2023-08-16(水)）の場合、月曜～日曜の7日分が返る', () => {
        /**
         * given:
         */
        const input = new Date('2023-08-16T00:00:00+09:00');

        const expectedDates = [
            '2023-08-14',
            '2023-08-15',
            '2023-08-16',
            '2023-08-17',
            '2023-08-18',
            '2023-08-19',
            '2023-08-20'
        ].map(dateStr => new Date(`${dateStr}T00:00:00+09:00`));

        /**
         * when:
         */
        const result = calendar.getDatesInWeek(input);

        /**
         * then:
         */
        const actualStr = result.map(d => d.toISOString().substring(0, 10));
        const expectedStr = expectedDates.map(d => d.toISOString().substring(0, 10));
        expect(actualStr).toEqual(expectedStr);
    });

    test('日曜の入力で週の一部が前年の場合、同じ年のものだけ返す（例：2023-01-01（日））', () => {
        /**
         * given:
         *  2023-01-01 の場合、週の月曜は2022-12-27～2022-12-31は前年なので、2023年のは2023-01-01, 2023-01-02のみ
         */
        const input = new Date('2023-01-01T00:00:00+09:00'); // 日曜日

        const expectedDates = [
            '2023-01-01'
        ].map(dateStr => new Date(`${dateStr}T00:00:00+09:00`));

        /**
         * when:
         */
        const result = calendar.getDatesInWeek(input);

        /**
         * then:
         */
        const actualStr = result.map(d => d.toISOString().substring(0, 10));
        const expectedStr = expectedDates.map(d => d.toISOString().substring(0, 10));
        expect(actualStr).toEqual(expectedStr);
    });

    test('年末の日曜入力の場合、正しい週の日付が返る（例：2023-12-31（日））', () => {
        /**
         * given:
         * 2023-12-31の場合、週の月曜は2023-12-25～2023-12-31となる
         */
        const input = new Date('2023-12-31T00:00:00+09:00'); // 日曜日

        const expectedDates = [
            '2023-12-25',
            '2023-12-26',
            '2023-12-27',
            '2023-12-28',
            '2023-12-29',
            '2023-12-30',
            '2023-12-31'
        ].map(dateStr => new Date(`${dateStr}T00:00:00+09:00`));

        /**
         * when:
         */
        const result = calendar.getDatesInWeek(input);

        /**
         * then:
         */
        const actualStr = result.map(d => d.toISOString().substring(0, 10));
        const expectedStr = expectedDates.map(d => d.toISOString().substring(0, 10));
        expect(actualStr).toEqual(expectedStr);
    });
});
