const fs = require("fs");
const jsonpath = require("jsonpath");
const XLSX = require("xlsx");
const path = require("path");

// === Configuration ===
const columnLetter = "B"; // Column where JSONPaths are present
const startRow = 2;       // Start from this row
const inputExcel = "paths.xlsx"; // Will overwrite this file
const inputJSON = "data.json";   // JSON file for validation
// const inputExcel = path.resolve("C:/Users/praveen/Downloads/paths.xlsx");
// const inputJSON = path.resolve("./data.json");


// === Load Excel and JSON ===
const workbook = XLSX.readFile(inputExcel, { cellStyles: true });
const sheetName = workbook.SheetNames[0];
const sheet = workbook.Sheets[sheetName];
const json = JSON.parse(fs.readFileSync(inputJSON, "utf-8"));

// === Fill color definitions ===
const fills = {
  green: { patternType: "solid", fgColor: { rgb: "C6EFCE" } }, // valid and found
  red:   { patternType: "solid", fgColor: { rgb: "FFC7CE" } }, // valid but not found
  blue:  { patternType: "solid", fgColor: { rgb: "D9E1F2" } }  // invalid syntax
};

// === Loop through cells in the column ===
const range = XLSX.utils.decode_range(sheet["!ref"]);
for (let row = startRow - 1; row <= range.e.r; row++) {
  const cellAddress = `${columnLetter}${row + 1}`;
  const cell = sheet[cellAddress];
  if (!cell || !cell.v) continue;

  const path = cell.v.toString().trim();

  try {
    const result = jsonpath.query(json, path);
    if (result.length > 0) {
      cell.s = { fill: fills.green };
    } else {
      cell.s = { fill: fills.red };
    }
  } catch (err) {
    cell.s = { fill: fills.blue };
  }
}

// === Overwrite the same file ===
XLSX.writeFile(workbook, inputExcel, { cellStyles: true });

console.log(`✅ Excel file '${inputExcel}' updated with validation colors.`);




{
  "user": {
    "id": 101,
    "name": "Alice",
    "email": "alice@example.com",
    "address": {
      "city": "Bangalore",
      "pincode": 560001
    }
  },
  "active": true
}

$.user.name
$.user.address.city
$.user.age
$.user.address[city]

{
  "orders": [
    { "orderId": 1001, "amount": 250, "status": "shipped" },
    { "orderId": 1002, "amount": 150, "status": "processing" }
  ],
  "customer": {
    "name": "Bob",
    "primeMember": false
  }
}

$.orders[*].orderId
$.orders[0].status
$.customer.primeMember
$.orders[*].invalidField
$.orders[*].status[
