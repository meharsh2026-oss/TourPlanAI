// ==========================================
// TourPlan AI - excel.js
// ==========================================

const TourPlanEngine = {

    workbook: null,
    sheetName: null,
    sheet: null,
    rows: null,

    load(workbook){

        this.workbook = workbook;

        this.sheetName = workbook.SheetNames[0];

        this.sheet = workbook.Sheets[this.sheetName];

        this.rows = XLSX.utils.sheet_to_json(
            this.sheet,
            {header:1}
        );

    },

    updateDates(monthName, year){

        const months = {
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
        };

        const days = [
            "SUNDAY",
            "MONDAY",
            "TUESDAY",
            "WEDNESDAY",
            "THURSDAY",
            "FRIDAY",
            "SATURDAY"
        ];

        const month = months[monthName];

        const lastDay = new Date(
            Number(year),
            month + 1,
            0
        ).getDate();

        for(let r=5;r<this.rows.length;r++){

            const row = this.rows[r];

            if(!row || !row[0]) continue;

            const oldDate = new Date(row[0]);

            if(isNaN(oldDate)) continue;

            const date = oldDate.getDate();

            if(date>lastDay){

                this.rows.splice(r,1);

                r--;

                continue;

            }

            const newDate = new Date(
                Number(year),
                month,
                date
            );

            row[0]=newDate;

            row[1]=days[newDate.getDay()];

            if(newDate.getDay()==0){

                for(let c=2;c<row.length;c++){

                    row[c]="";

                }

            }

        }

    },

    save(){

        this.sheet =
            XLSX.utils.aoa_to_sheet(
                this.rows
            );

        this.workbook.Sheets[
            this.sheetName
        ]=this.sheet;

        return this.workbook;

    }

};
