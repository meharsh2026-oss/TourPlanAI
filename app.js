// ==========================================
// TourPlan AI - app.js (Part 1)
// ==========================================

// Global Variables
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

function updateWorkbookInfo(file, sheet, rows){

    if(fileName)
        fileName.textContent = file;

    if(sheetName)
        sheetName.textContent = sheet;

    if(rowCount)
        rowCount.textContent = rows;

    if(selectedMonth)
        selectedMonth.textContent =
            monthSelect.value + " " + yearInput.value;

}

fileInput.addEventListener("change", async (e)=>{

    const file = e.target.files[0];

    if(!file) return;

    try{

        updateAI("Reading Workbook...",20);

        const data = await file.arrayBuffer();

        workbook = XLSX.read(data,{
            type:"array"
        });

        const sheet =
            workbook.SheetNames[0];

        const rows =
            XLSX.utils.sheet_to_json(
                workbook.Sheets[sheet],
                {header:1}
            );

        updateWorkbookInfo(
            file.name,
            sheet,
            rows.length
        );

        updateAI(
            "Workbook Loaded Successfully",
            100
        );

    }catch(err){

        console.error(err);

        alert(err.message);

    }

});
// ==========================================
// TourPlan AI - app.js (Part 2)
// ==========================================

generateBtn.addEventListener("click", () => {

    if (!workbook) {
        alert("Please upload an Excel file first.");
        return;
    }

    try {

        updateAI("Loading Tour Plan...", 20);

        TourPlanEngine.load(workbook);

        updateAI("Updating Dates...", 40);

        TourPlanEngine.updateDates(
            monthSelect.value,
            yearInput.value
        );

        updateAI("Saving Workbook...", 70);

        const newWorkbook = TourPlanEngine.save();

        updateAI("Downloading Excel...", 90);

        const outputFile =
            `TourPlan_${monthSelect.value}_${yearInput.value}.xlsx`;

        XLSX.writeFile(newWorkbook, outputFile);

        updateAI("Completed Successfully ✅", 100);

        alert("Tour Plan Generated Successfully!");

    } catch (err) {

        console.error(err);

        alert("Error: " + err.message);

        updateAI("Generation Failed ❌", 0);

    }

});

// Update month display automatically
monthSelect.addEventListener("change", () => {

    if (selectedMonth) {
        selectedMonth.textContent =
            monthSelect.value + " " + yearInput.value;
    }

});

yearInput.addEventListener("input", () => {

    if (selectedMonth) {
        selectedMonth.textContent =
            monthSelect.value + " " + yearInput.value;
    }

});
