const fileInput = document.getElementById("excelFile");
const generateBtn = document.getElementById("generateBtn");
const status = document.getElementById("status");

let workbook = null;

// Read Excel File
fileInput.addEventListener("change", (e) => {

    const file = e.target.files[0];

    if (!file) return;

    status.innerHTML = "📂 Reading Excel file...";

    const reader = new FileReader();

    reader.onload = function(event){

        const data = new Uint8Array(event.target.result);

        workbook = XLSX.read(data,{
            type:"array"
        });

        status.innerHTML =
        "✅ Excel Loaded Successfully<br><br>" +
        "📑 Sheets:<br>" +
        workbook.SheetNames.join("<br>");
    };

    reader.readAsArrayBuffer(file);

});

// Generate Button
generateBtn.addEventListener("click",()=>{

    if(workbook==null){

        alert("Please upload an Excel file first.");

        return;

    }

    status.innerHTML="🤖 AI is analysing your workbook...";

    setTimeout(()=>{

        status.innerHTML+="<br><br>✅ Workbook Ready";

    },2000);

});
