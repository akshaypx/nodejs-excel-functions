const XLSX = require("xlsx");
const workbook = XLSX.readFile("./Output.xlsx");
const sheetName = workbook.SheetNames[2];
console.log(sheetName);
const worksheet = workbook.Sheets[sheetName];
const data = XLSX.utils.sheet_to_json(worksheet);

let offshoreSum = {};
let onsiteSum = {};

for (let row of data) {
  for (let i = 2; i < Object.keys(row).length; i++) {
    const colName = Object.keys(row)[i];
    const cellValue = row[colName];
    const cellValueNum = Number(cellValue);

    // Check the value of the cell in the second column (column B)
    const offshore = row["Ofshore"];
    if (offshore === "Y") {
      if (colName in offshoreSum) {
        offshoreSum[colName] += cellValueNum || 0;
      } else {
        offshoreSum[colName] = cellValueNum || 0;
      }
    } else if (offshore === "N") {
      if (colName in onsiteSum) {
        onsiteSum[colName] += cellValueNum || 0;
      } else {
        onsiteSum[colName] = cellValueNum || 0;
      }
    }
  }
}

const offshoreTotal = Object.values(offshoreSum).reduce(
  (total, value) => total + value,
  0
);

const onsiteTotal = Object.values(onsiteSum).reduce(
  (total, value) => total + value,
  0
);

console.log(`Offshore sum: ${JSON.stringify(offshoreSum)}`);
console.log(`Onsite sum: ${JSON.stringify(onsiteSum)}`);

console.log("Offshore total sum:", offshoreTotal);
console.log("Onsite total sum:", onsiteTotal);

//UPDATE EXCEL DATA

const workbook2 = XLSX.readFile("./Output.xlsx");
const sheetName2 = "JP-EV";
const cellAddressOffShore = "C2";
const cellAddressOnSite = "C3";

// Update the cell value
const worksheet2 = workbook2.Sheets[sheetName2];
worksheet2[cellAddressOffShore].v = offshoreTotal;
worksheet2[cellAddressOnSite].v = onsiteTotal;

// Save the updated workbook to a new file
const newWorkbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(newWorkbook, worksheet2, sheetName2);
XLSX.writeFile(newWorkbook, "updated.xlsx");
