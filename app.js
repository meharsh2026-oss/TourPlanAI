alert("New app.js loaded");
const fileInput = document.getElementById("excelFile");
const generateBtn = document.getElementById("generateBtn");
const status = document.getElementById("status");
const month = document.getElementById("month");
const year = document.getElementById("year");

let workbook = null;

fileInput.addEventListener("change", function(e){

    const file = e.target.files[0];

    if(!file) return;

    status.innerHTML="📂 Reading Excel...";

    const reader=new FileReader();

    reader.onload=function(event){

        const data=new Uint8Array(event.target.result);

        workbook=XLSX.read(data,{type:"array"});

        status.innerHTML=
        "✅ Excel Loaded Successfully<br><br>" +
        "Sheets:<br>" +
        workbook.SheetNames.join("<br>");

    };

    reader.readAsArrayBuffer(file);

});

generateBtn.addEventListener("click",function(){

alert("Generate button clicked");
    if(!workbook){

        alert("Please upload Excel first.");

        return;

    }

    const sheetName=workbook.SheetNames[0];

    const worksheet=workbook.Sheets[sheetName];

    const json=XLSX.utils.sheet_to_json(worksheet,{
        header:1
    });

    console.log(json);

    let preview="<br><br>";

for(let i=0;i<10 && i<json.length;i++){

    preview+=JSON.stringify(json[i])+"<br>";

}

status.innerHTML=
"🤖 AI Analysing Workbook...<br><br>"+
"Rows : "+json.length+
"<br>Sheet : "+sheetName+
preview;
});
