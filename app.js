// ====================================================
// TourPlan AI - app.js
// ====================================================

let workbook = null;

const fileInput = document.getElementById("excelFile");
const monthSelect = document.getElementById("month");
const yearInput = document.getElementById("year");
const generateBtn = document.getElementById("generateBtn");

const fileName = document.getElementById("fileName");
const sheetName = document.getElementById("sheetName");
const rowCount = document.getElementById("rowCount");
const selectedMonth = document.getElementById("selectedMonth");

const aiSteps = document.getElementById("aiSteps");
const progressBar = document.getElementById("progressBar");

function updateAI(message, progress){

    if(aiSteps){
        aiSteps.innerHTML = "🤖 " + message;
    }

    if(progressBar){
        progressBar.style.width = progress + "%";
    }

}

fileInput.addEventListener("change", async function(e){

    alert("File change detected");

    const file = e.target.files[0];

    if(!file) return;

    try{

        updateAI("Reading Workbook...",20);

        const data = await file.arrayBuffer();

        workbook = XLSX.read(data,{type:"array"});

        const sheet = workbook.SheetNames[0];

        const rows = XLSX.utils.sheet_to_json(
            workbook.Sheets[sheet],
            {header:1}
        );

        if(fileName) fileName.textContent=file.name;
        if(sheetName) sheetName.textContent=sheet;
        if(rowCount) rowCount.textContent=rows.length;
        if(selectedMonth) selectedMonth.textContent=monthSelect.value+" "+yearInput.value;

        updateAI("Workbook Loaded Successfully",100);

        alert("Workbook Loaded Successfully");

    }catch(err){

        console.error(err);

        alert(err.message);

    }

});

generateBtn.addEventListener("click", () => {

    console.log(workbook);

    ...
});
    if (!workbook) {
        alert("Please upload an Excel file first.");
        return;
    }

    try {

        updateAI("Reading Workbook...",20);

        TourPlanEngine.load(workbook);

        updateAI("Updating Dates...",40);

        TourPlanEngine.updateDates(
            monthSelect.value,
            yearInput.value
        );

        updateAI("Applying VALUE Rules...",70);

        const newWorkbook = TourPlanEngine.save();

        updateAI("Generating Excel...",90);

        XLSX.writeFile(
            newWorkbook,
            `TourPlan_${monthSelect.value}_${yearInput.value}.xlsx`
        );

        updateAI("Completed Successfully ✅",100);

        alert("Tour Plan Generated Successfully!");

    } catch(err){

        console.error(err);

        alert(err.message);

    }

});

});
