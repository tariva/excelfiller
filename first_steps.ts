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

// Extract unique IDs from the ID column
const uniqueIds = new Set();
for (let i = startRow; ; i++) {
  const cellAddress = `${idColumn}${i}`;
  const cell = worksheet[cellAddress];
  if (!cell) break; // Reached the end of the data
  uniqueIds.add(cell.v);
}

// For each unique ID, gather the values from the data columns
for (const id of uniqueIds) {
  const rowsWithId = [];
  const dataForColumns: any = {};

  for (let i = startRow; ; i++) {
    const cellAddress = `${idColumn}${i}`;
    const cell = worksheet[cellAddress];
    if (!cell) break; // Reached the end of the data

    if (cell.v === id) {
      rowsWithId.push(i);
      for (const dataColumn of dataColumns) {
        if (!dataForColumns[dataColumn]) {
          dataForColumns[dataColumn] = new Set();
        }
        const dataCell = worksheet[`${dataColumn}${i}`];
        if (dataCell) dataForColumns[dataColumn].add(dataCell.v);
      }
    }
  }

  // Update the data columns for rows with the same ID
  for (const row of rowsWithId) {
    for (const dataColumn of dataColumns) {
      const uniqueValues = Array.from(dataForColumns[dataColumn]);
      if (uniqueValues.length === 1) {
        worksheet[`${dataColumn}${row}`].v = uniqueValues[0];
      }
    }
  }
}

// Save the modified workbook
XLSX.writeFile(workbook, excelFileName);
