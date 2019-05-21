const Excel = require("exceljs");
const workbook = new Excel.Workbook();
const fs = require("fs");
const filename = "output.xlsx";
let json = fs.readFileSync("convert.json", "utf-8");
json = JSON.parse(json);
for (let i in json) {
  workbook.addWorksheet(i);
  worksheet = workbook.getWorksheet(i);
  firstRow = json[i][0];
  let columns = [];
  for (l in firstRow) {
    let tempObj = {};
    tempObj["header"] = l;
    tempObj["key"] = l;
    columns.push(tempObj);
  }
  worksheet.columns = columns;
  for (let k = 0; k < json[i].length; k++) {
    worksheet.addRow(json[i][k]);
  }
}

const stream = fs.createWriteStream(filename);
workbook.xlsx
  .write(stream)
  .then(function() {
    console.log(`File: ${filename} saved!`);
    stream.end();
  })
  .catch(error => {
    console.err(`File: ${filename} save failed: `, error);
  });
