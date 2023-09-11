// deno-lint-ignore-file no-inferrable-types no-explicit-any
import XLSX from "npm:xlsx@0.18.5";

// import mysql from "npm:mysql2@^2.3.3/promise";

// Prompt the user for input
let excelFileName: string = "";
while (!excelFileName.trim()) {
  excelFileName = prompt(
    "Please enter the Excel file name (with extension):"
  ) as string;
}

let idColumn: string = "";
while (!idColumn.trim()) {
  idColumn = prompt("Enter the column with the ID (e.g. 'B'):") as string;
}

let dataColumnsInput: string = "";
while (!dataColumnsInput.trim()) {
  dataColumnsInput = prompt(
    "Enter the columns for data (comma separated, e.g. 'C,D,E'):"
  ) as string;
}
const dataColumns = dataColumnsInput.split(",");

let startRowInput: string = "";
let startRow = 0;
while (!startRow) {
  startRowInput = prompt(
    "Enter the row number where the data starts:"
  ) as string;
  startRow = parseInt(startRowInput);
  if (isNaN(startRow)) {
    console.log("Please enter a valid row number.");
    startRow = 0; // Reset startRow to ensure the loop continues
  }
}

// Load the Excel workbook
const workbook = XLSX.readFile(excelFileName);
const worksheet = workbook.Sheets[workbook.SheetNames[0]];

// Create a map to store rows associated with each ID
const idRowMap = new Map();

// Loop through the idColumn to gather unique IDs and their associated rows
for (let i = startRow; ; i++) {
  const cellAddress = `${idColumn}${i}`;
  const cell = worksheet[cellAddress];
  if (!cell) break; // Reached the end of the data

  const idValue = cell.v;
  if (!idRowMap.has(idValue)) {
    idRowMap.set(idValue, []);
  }
  idRowMap.get(idValue).push(i);
}

// Process each unique ID
for (const [id, rows] of idRowMap.entries()) {
  const columnDataMap = new Map();

  // Initialize columnDataMap with Sets
  for (const col of dataColumns) {
    columnDataMap.set(col, new Set());
  }

  // Collect data from the dataColumns for each row this ID appears in
  for (const row of rows) {
    for (const col of dataColumns) {
      const cell = worksheet[`${col}${row}`];
      if (cell) {
        columnDataMap.get(col).add(cell.v);
      }
    }
  }

  // Propagate data if unique for this ID
  for (const col of dataColumns) {
    const uniqueValues = Array.from(columnDataMap.get(col));
    if (uniqueValues.length === 1) {
      // Only one unique value for this ID
      for (const row of rows) {
        if (!worksheet[`${col}${row}`]) {
          worksheet[`${col}${row}`] = {}; // Create cell if it doesn't exist
        }
        worksheet[`${col}${row}`].v = uniqueValues[0];
      }
    }
  }
}

// Save the modified workbook
XLSX.writeFile(workbook, excelFileName);
