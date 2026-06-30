// ====================================================
// TourPlan AI - app.js
// Version 1.0
// ====================================================

// Global variables
let workbook = null;
let worksheet = null;
let jsonData = [];

const fileInput = document.getElementById("excelFile");
const monthSelect = document.getElementById("month");
const yearInput = document.getElementById("year");
const generateBtn = document.getElementById("generateBtn");

// Workbook info
const fileName = document.getElementById("fileName");
const sheetName = document.getElementById("sheetName");
const rowCount = document.getElementById("rowCount");
const selectedMonth = document.getElementById("selectedMonth");

// AI panel
const aiSteps = document.getElementById("aiSteps");
const progressBar = document.getElementById("progressBar");

function updateAI(message, progress) {
    aiSteps.innerHTML = `<p>🤖 ${message}</p>`;
    progressBar.style.width = progress + "%";
}

fileInput.addEventListener("change", loadWorkbook);

async function loadWorkbook(e) {

    const file = e.target.files[0];

    if (!file) return;

    updateAI("Reading Workbook...", 15);

    const data = await file.arrayBuffer();

    workbook = XLSX.read(data);

    const firstSheet = workbook.SheetNames[0];

    worksheet = workbook.Sheets[firstSheet];

    jsonData = XLSX.utils.sheet_to_json(worksheet, {
        header: 1
    });

    updateWorkbookInfo(file.name, firstSheet, jsonData.length);

}

function updateWorkbookInfo(file, sheet, rows){

    fileName.textContent = file;

    sheetName.textContent = sheet;

    rowCount.textContent = rows;

    selectedMonth.textContent =
        monthSelect.value + " " + yearInput.value;

    updateAI("Workbook Loaded Successfully",100);

}
// ===============================
// Generate Tour Plan Button
// ===============================

generateBtn.addEventListener("click", () => {

    if (!workbook) {
        alert("Please upload an Excel file first.");
        return;
    }

    try {

        updateAI("Loading Tour Plan...", 10);

        TourPlanEngine.load(workbook);

        updateAI("Updating Dates...", 40);

        TourPlanEngine.updateDates(
            monthSelect.value,
            yearInput.value
        );

        updateAI("Creating Workbook...", 70);

        const newWorkbook = TourPlanEngine.save();

        updateAI("Downloading...", 90);

        XLSX.writeFile(
            newWorkbook,
            `TourPlan_${monthSelect.value}_${yearInput.value}.xlsx`
        );

        updateAI("Completed Successfully ✅", 100);

    } catch (err) {

        console.error(err);

        alert(err.message);

        updateAI("Error ❌", 0);

    }

});
