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
