
// =========================================
// TourPlan AI - Excel Engine v1.0
// =========================================

const MONTHS = {
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

const DAYS=[
"Sunday",
"Monday",
"Tuesday",
"Wednesday",
"Thursday",
"Friday",
"Saturday"
];

function generateTourPlan(workbook,monthName,year){

    const sheetName=workbook.SheetNames[0];

    const sheet=workbook.Sheets[sheetName];

    const rows=XLSX.utils.sheet_to_json(sheet,{
        header:1
    });

    const month=MONTHS[monthName];

    for(let i=5;i<rows.length;i++){

        if(!rows[i][0]) continue;

        const dayNumber=i-4;

        const newDate=new Date(
            Number(year),
            month,
            dayNumber
        );

        rows[i][0]=newDate;

        rows[i][1]=
        DAYS[newDate.getDay()].toUpperCase();

    }

    const newSheet=
    XLSX.utils.aoa_to_sheet(rows);

    workbook.Sheets[sheetName]=newSheet;

    return workbook;

}
