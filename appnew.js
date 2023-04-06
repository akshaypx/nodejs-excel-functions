// Require the xlsx library
const xlsx = require("xlsx");

// Load the Excel file and get the first sheet
const workbook = xlsx.readFile("./userdata.xlsx");
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

// Convert the sheet data to a 2D array
const data = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

// Use reduce to calculate the sum of values in each column
const columnSums = data[0].slice(2).reduce((acc, colName) => {
  // Get the index of the current column
  const colIndex = data[0].indexOf(colName);

  // Use reduce to calculate the sum of offshore values in the current column
  const offshoreSum = data.slice(1).reduce((acc, row) => {
    // Get the value of the "Ofshore" column for the current row
    const offshore = row[1];
    // If the "Ofshore" value is "Y", add the value in the current cell to the accumulator
    if (offshore === "Y") {
      const cellValue = Number(row[colIndex]);
      return acc + (isNaN(cellValue) ? 0 : cellValue);
    }
    // Otherwise, just return the accumulator unchanged
    return acc;
  }, 0);

  // Add the offshore sum to the result object with the appropriate key
  acc[`offshore-sum-${colName}`] = offshoreSum;

  // Use reduce to calculate the sum of onsite values in the current column
  const onsiteSum = data.slice(1).reduce((acc, row) => {
    // Get the value of the "Ofshore" column for the current row
    const offshore = row[1];
    // If the "Ofshore" value is "N", add the value in the current cell to the accumulator
    if (offshore === "N") {
      const cellValue = Number(row[colIndex]);
      return acc + (isNaN(cellValue) ? 0 : cellValue);
    }
    // Otherwise, just return the accumulator unchanged
    return acc;
  }, 0);

  // Add the onsite sum to the result object with the appropriate key
  acc[`onsite-sum-${colName}`] = onsiteSum;

  // Return the accumulator for the next iteration of the reduce loop
  return acc;
}, {});

// Log the result object to the console
console.log(columnSums);
