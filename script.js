let globalPOData = {};

function processFile() {

  const file = document.getElementById('fileUpload').files[0];
  const reader = new FileReader();

  reader.onload = function(e) {

    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, {type:'array'});

    const marksData = XLSX.utils.sheet_to_json(workbook.Sheets["Sheet1"]);
    const ipccData = XLSX.utils.sheet_to_json(workbook.Sheets["IPCC"], {header:1});

    calculateAttainment(marksData, ipccData);
  };

  reader.readAsArrayBuffer(file);
}

function calculateAttainment(marksData, ipccData) {

  // -------- Robust CO detection --------
  let headers = Object.keys(marksData[0]);

  let coColumns = headers.filter(h =>
      h.toLowerCase().includes("co") && !isNaN(parseFloat(marksData[0][h]))
  );

  if(coColumns.length === 0){
      alert("CO columns not detected. Check column names.");
      return;
  }

  let total = marksData.length;
  let coLevels = {};
  let coPercent = {};

  coColumns.forEach(co => {

    let count = marksData.filter(s => parseFloat(s[co]) >= 60).length;
    let percent = (count/total)*100;
    coPercent[co] = percent;

    if(percent >= 80) coLevels[co] = 3;
    else if(percent >= 70) coLevels[co] = 2;
    else if(percent >= 60) coLevels[co] = 1;
    else coLevels[co] = 0;
  });

  displayCO(coPercent, coLevels);

  // -------- Robust Matrix Extraction --------
  let start = ipccData.findIndex(r =>
      r.some(cell => String(cell).toLowerCase().includes("co"))
  );

  if(start === -1){
      alert("CO-PO Matrix not detected in IPCC sheet.");
      return;
  }

  let header = ipccData[start];
  let matrix = {};

  for(let i = start+1; i < ipccData.length; i++){

    let row = ipccData[i];
    if(!row || !row[0]) continue;

    if(String(row[0]).toLowerCase().includes("co")){
        matrix[row[0]] = {};

        for(let j=1; j<header.length; j++){
            matrix[row[0]][header[j]] = parseFloat(row[j]) || 0;
        }
    }
  }

  calculatePO(coLevels, matrix);
}

function calculatePO(coLevels, matrix){

  let poResults = {};

  Object.keys(matrix).forEach(co => {

    if(!coLevels[co]) return;

    Object.keys(matrix[co]).forEach(po => {

      if(!poResults[po]) poResults[po] = {sum:0, weight:0};

      poResults[po].sum += coLevels[co] * matrix[co][po];
      poResults[po].weight += matrix[co][po];
    });
  });

  let finalPO = {};
  Object.keys(poResults).forEach(po=>{
    finalPO[po] = poResults[po].weight === 0 ? 0 :
                  poResults[po].sum/poResults[po].weight;
  });

  globalPOData = finalPO;
  displayPO(finalPO);
}

function displayCO(percent, levels){

  let table="<table><tr><th>CO</th><th>% ≥60</th><th>Level</th></tr>";

  Object.keys(percent).forEach(co=>{
    table+=`<tr>
      <td>${co}</td>
      <td>${percent[co].toFixed(2)}%</td>
      <td>${levels[co]}</td>
    </tr>`;
  });

  table+="</table>";
  document.getElementById("coTable").innerHTML=table;
}

function displayPO(poData){

  let table="<table><tr><th>PO</th><th>Level</th></tr>";

  Object.keys(poData).forEach(po=>{
    table+=`<tr>
      <td>${po}</td>
      <td>${poData[po].toFixed(2)}</td>
    </tr>`;
  });

  table+="</table>";
  document.getElementById("poTable").innerHTML=table;
}