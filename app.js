const fileInput = document.getElementById("excelFile");
const generateBtn = document.getElementById("generateBtn");
const statusText = document.getElementById("status");

let workbook = null;

fileInput.addEventListener("change", handleFile);

function handleFile(e) {
    const file = e.target.files[0];

    if (!file) return;

    statusText.innerHTML = "📂 Reading Excel...";

    const reader = new FileReader();

    reader.onload = function(event){

        const data = new Uint8Array(event.target.result);

        workbook = XLSX.read(data,{
            type:"array"
        });

        statusText.innerHTML =
        "✅ Workbook Loaded<br><br>" +
        "Sheets:<br>" +
        workbook.SheetNames.join("<br>");
    };

    reader.readAsArrayBuffer(file);
}

generateBtn.addEventListener("click",()=>{

    if(!workbook){

        alert("Please upload an Excel file first.");

        return;
    }

    statusText.innerHTML="🤖 AI is analysing workbook...";

    setTimeout(()=>{

        statusText.innerHTML=
        "✅ Workbook Ready<br>" +
        "Next update will modify dates automatically.";

    },2000);

});
