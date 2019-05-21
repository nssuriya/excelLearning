const Excel = require("exceljs");
const workbook = new Excel.Workbook();
const fs = require("fs");
const filename = "test.xlsx";
const sheetNames = ["A", "B", "C", "D"];

sheetNames.forEach(sheetName => {
  workbook.addWorksheet(sheetName);
});

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
