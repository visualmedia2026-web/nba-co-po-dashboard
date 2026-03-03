let globalPOData = {};
let globalCOData = {};

function processFile() {

  const file = document.getElementById('fileUpload').files[0];
  const reader = new FileReader();

  reader.onload = function(e) {

    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, {type:'array'});

    const marksSheet = workbook.Sheets["Sheet1"];
    const marksData = XLSX.utils.sheet_to_json(marksSheet);

    const ipccSheet = workbook.Sheets["IPCC"];
    const ipccData = XLSX.utils.sheet_to_json(ipccSheet, {header:1});

    calculateAttainment(marksData, ipccData);
  };

  reader.readAsArrayBuffer(file);
}

function calculateAttainment(marksData, ipccData) {

  let coColumns = Object.keys(marksData[0]).filter(key => key.includes("CO"));
  let totalStudents = marksData.length;

  let coLevels = {};
  let coPercent = {};

  coColumns.forEach(co => {

    let count = 0;
    marksData.forEach(student => {
      if(student[co] >= 60) count++;
    });

    let percentage = (count / totalStudents) * 100;
    coPercent[co] = percentage;

    if(percentage >= 80) coLevels[co] = 3;
    else if(percentage >= 70) coLevels[co] = 2;
    else if(percentage >= 60) coLevels[co] = 1;
    else coLevels[co] = 0;
  });

  globalCOData = coLevels;
  displayCO(coPercent, coLevels);

  let matrixStartRow = ipccData.findIndex(row => row.includes("CO"));
  let header = ipccData[matrixStartRow];

  let matrix = {};

  for(let i = matrixStartRow + 1; i < matrixStartRow + 6; i++) {
    let row = ipccData[i];
    matrix[row[0]] = {};

    for(let j = 1; j < header.length; j++) {
      matrix[row[0]][header[j]] = row[j] || 0;
    }
  }

  calculatePO(coLevels, matrix);
}

function calculatePO(coLevels, matrix) {

  let poResults = {};

  Object.keys(matrix).forEach(co => {
    Object.keys(matrix[co]).forEach(po => {

      if(!poResults[po]) poResults[po] = {sum:0, weight:0};

      poResults[po].sum += coLevels[co] * matrix[co][po];
      poResults[po].weight += matrix[co][po];
    });
  });

  let finalPO = {};

  Object.keys(poResults).forEach(po => {
    finalPO[po] = poResults[po].weight === 0 ? 0 :
                  (poResults[po].sum / poResults[po].weight).toFixed(2);
  });

  globalPOData = finalPO;
  displayPO(finalPO);
}

function displayCO(percent, levels) {

  let table = "<table><tr><th>CO</th><th>% ≥60</th><th>Level</th></tr>";

  Object.keys(percent).forEach(co => {

    table += `<tr>
                <td>${co}</td>
                <td>${percent[co].toFixed(2)}%</td>
                <td class="level${levels[co]}">${levels[co]}</td>
              </tr>`;
  });

  table += "</table>";
  document.getElementById("coTable").innerHTML = table;
}

function displayPO(poData) {

  let table = "<table><tr><th>PO</th><th>Attainment</th><th>Status</th></tr>";

  Object.keys(poData).forEach(po => {

    let status = poData[po] >= 2 ? "Attained" : "Needs Improvement";

    table += `<tr>
                <td>${po}</td>
                <td>${poData[po]}</td>
                <td>${status}</td>
              </tr>`;
  });

  table += "</table>";
  document.getElementById("poTable").innerHTML = table;

  new Chart(document.getElementById("poChart"), {
    type: 'bar',
    data: {
      labels: Object.keys(poData),
      datasets: [{
        label: "PO Attainment",
        data: Object.values(poData)
      }]
    }
  });
}

function generatePDF() {

  const { jsPDF } = window.jspdf;
  const doc = new jsPDF();

  doc.text("NBA CO-PO Attainment Report", 20, 20);

  let y = 30;
  Object.keys(globalPOData).forEach(po => {
    doc.text(`${po} : ${globalPOData[po]}`, 20, y);
    y += 10;
  });

  doc.save("NBA_Attainment_Report.pdf");
}