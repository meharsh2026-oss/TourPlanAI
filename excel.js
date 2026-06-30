// ======================================
// TourPlan AI - Excel Engine V2
// ======================================

const TourPlanEngine = {

    workbook: null,
    sheetName: null,
    sheet: null,
    rows: null,

    load(workbook){

        this.workbook = workbook;

        this.findSheet();

        this.readRows();

    },

    findSheet(){

    this.sheetName = this.workbook.SheetNames[0];

    this.sheet = this.workbook.Sheets[this.sheetName];

}

        throw new Error("Tour Plan sheet not found.");

    },

    readRows(){

        this.rows = XLSX.utils.sheet_to_json(
            this.sheet,
            {header:1}
        );

    },

    updateDates(month,year){

        const days=[
            "SUNDAY",
            "MONDAY",
            "TUESDAY",
            "WEDNESDAY",
            "THURSDAY",
            "FRIDAY",
            "SATURDAY"
        ];

        const monthNumber={
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

        const targetMonth=monthNumber[month];

        this.rows[r] = Planner.updateRow(
    this.rows[r],
    month,
    year
);

            const cell=this.rows[r][0];

            if(!cell) continue;

            let date;

            if(cell instanceof Date){

                date=new Date(cell);

            }else{

                date=new Date(cell);

            }

            if(isNaN(date)) continue;

            const day=date.getDate();

            const newDate=new Date(
                Number(year),
                targetMonth,
                day
            );

            this.rows[r][0]=newDate;

            this.rows[r][1]=
            days[newDate.getDay()];

        }

    },

    save(){
        this.rows = ValueEngine.update(this.rows);

        const newSheet =
            XLSX.utils.aoa_to_sheet(this.rows);

        this.workbook.Sheets[
            this.sheetName
        ] = newSheet;

        return this.workbook;

    }

};
