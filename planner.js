// ======================================
// TourPlan AI Planner v1.0
// ======================================

const Planner = {

    monthMap: {
        January:0,
        February:1,
        March:2,
        April:3,
        May:4,
        June:5,
        July:6,
        August:7,
        September:8,
        October:9,
        November:10,
        December:11
    },

    days:[
        "SUNDAY",
        "MONDAY",
        "TUESDAY",
        "WEDNESDAY",
        "THURSDAY",
        "FRIDAY",
        "SATURDAY"
    ],

    updateRow(row, month, year){

        if(!row[0]) return row;

        let oldDate = new Date(row[0]);

        if(isNaN(oldDate)) return row;

        let day = oldDate.getDate();

        let newDate = new Date(
            Number(year),
            this.monthMap[month],
            day
        );

        row[0] = newDate;

        row[1] = this.days[newDate.getDay()];

        return row;

    }

};
