class Calendar {
    /**@type {Date} */
    currentDate;
    /**@type {array} */
    data;

    constructor() {
        const now = new Date();
        const jstString = now.toLocaleString('en-US', { timeZone: 'Asia/Tokyo' });
        this.currentDate = new Date(jstString);

        this.data = this.getDatesInWeek(this.currentDate);
    }

    /**
     * その週に含まれる日付一覧を求めるメソッド（年内の日付のみを含める）
     * 
     * @param {Date} date 
     */
    getDatesInWeek(date) {
        let dateList = [];
        let currentYear = date.getFullYear();

        // 週の始まり（月曜日）を取得
        let day = date.getDay(); // 0（日曜）から6（土曜）
        let diff = date.getDate() - day + (day === 0 ? -6 : 1); // 日曜の場合は-6
        let monday = new Date(date);
        monday.setDate(diff);
        monday.setHours(0, 0, 0, 0);

        // 月曜日から日曜日までの日付を取得
        for (let i = 0; i < 7; i++) {
            let d = new Date(monday);
            d.setDate(monday.getDate() + i);
            d.setHours(0, 0, 0, 0);

            // 年が同じ場合のみ追加
            if (d.getFullYear() === currentYear) {
                dateList.push(new Date(d.getFullYear(), d.getMonth(), d.getDate()));
            }
        }
        return dateList;
    }

    /**
     * 日付文字列をDate型に変換するメソッド
     * 
     * @param {string} dateStr
     * @param {number} year
     * 
     * @return {Date}
     */
    consertDateStringsToDate(dateStr, year) {
        let dateParts = dateStr.split('/');
        let month = parseInt(dateParts[0], 10) - 1; // 月は0始まり
        let day = parseInt(dateParts[1], 10);
        return new Date(year, month, day);
    }

    /**
     * YYYYMMDD文字列をDateオブジェクトに変換するメソッド
     * 
     * @param {string} dateStr 
     */
    convertYYYYMMDDToDate(dateStr) {
        if (!/^\d{8}$/.test(dateStr)) {
            return null;
        }
        let year = parseInt(dateStr.substring(0, 4), 10);
        let month = parseInt(dateStr.substring(4, 6), 10) - 1; // 月は0始まり
        let day = parseInt(dateStr.substring(6, 8), 10);

        return new Date(year, month, day);
    }

    /**
     * Date型をMM/DD文字列に変換するメソッド
     * 
     * @param {Date} date 
     */
    convertDateToStringWithSlash(date) {
        let month = date.getMonth() + 1;
        let day = date.getDate();
        return month + "/" + day;
    }

}

module.exports = Calendar;
