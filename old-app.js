// const xlsx = require("xlsx");
// const workbook = xlsx.readFile("./userdata.xlsx");
// const sheet_name_list = workbook.SheetNames;
// const data = xlsx.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);
// console.log(data);

//working
const xlsx = require("xlsx");
const workbook = xlsx.readFile("./userdata.xlsx");
const sheet_name_list = workbook.SheetNames;
const data = xlsx.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);

//sum of all cells under each column, starting from
//third column (i=2),
const columnSums = {};
const offshoreColSum = {};
const onsiteColSum = {};
// for (let row of data) {
//   //console.log(row);
//   for (let i = 2; i < Object.keys(row).length; i++) {
//     const colName = Object.keys(row)[i];
//     const site = Object.keys(row)[1];
//     console.log(site, row[site]);
//     if (columnSums[colName]) {
//       columnSums[colName] += Number(row[colName]) || 0;
//     } else {
//       columnSums[colName] = Number(row[colName]) || 0;
//     }
//   }
// }

for (let row of data) {
  for (let i = 2; i < Object.keys(row).length; i++) {
    const colName = Object.keys(row)[i];
    const site = Object.keys(row)[1];
    //clsconsole.log(site, row[site]);
    //may-20:
    if (row[site] == "Y") {
      offshoreColSum[colName] += Number(row[colName]) || 0;
    } else {
      onsiteColSum[colName] = Number(row[colName]) || 0;
    }
  }
}

console.log("Ofshore\n", offshoreColSum);
console.log("Onsite\n", onsiteColSum);

//try 3 offshore onsite
// const xlsx = require("xlsx");
// const workbook = xlsx.readFile("./userdata.xlsx");
// const sheet_name_list = workbook.SheetNames;
// const data = xlsx.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);

// const onsiteSum = {};
// const offshoreSum = {};
// for (let row of data) {
//   for (let i = 2; i < Object.keys(row).length; i++) {
//     const colName = Object.keys(row)[i];
//     const cellValue = Number(row[colName]) || 0;
//     if (colName === "Offshore" && row[colName] === "Y") {
//       if (offshoreSum[colName]) {
//         offshoreSum[colName] += cellValue;
//       } else {
//         offshoreSum[colName] = cellValue;
//       }
//     } else if (colName !== "Offshore" || row[colName] !== "Y") {
//       if (onsiteSum[colName]) {
//         onsiteSum[colName] += cellValue;
//       } else {
//         onsiteSum[colName] = cellValue;
//       }
//     }
//   }
// }

// console.log("Onsite Sum:", onsiteSum);
// console.log("Offshore Sum:", offshoreSum);
