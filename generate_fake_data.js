// Step 1: Install the required packages
// Run the following command in your terminal
// npm install xlsx yargs

// Step 2: Parse command line arguments using yargs
const yargs = require('yargs/yargs');
const { hideBin } = require('yargs/helpers');
const argv = yargs(hideBin(process.argv))
  .option('n', {
    alias: 'rows',
    description: 'Number of rows to generate',
    type: 'number',
    demandOption: true
  })
  .option('i', {
    alias: 'input',
    description: 'Input XLSX template file',
    type: 'string',
    demandOption: true
  })
  .option('o', {
    alias: 'output',
    description: 'Output XLSX file path',
    type: 'string',
    demandOption: true
  }).argv;

// Step 3: Read the input XLSX file using SheetJS
const XLSX = require('xlsx');
const workbook = XLSX.readFile(argv.input);
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

// Step 4: Extract the headers (column names) from the input XLSX
const headers = [];
const range = XLSX.utils.decode_range(worksheet['!ref']);
for (let col = range.s.c; col <= range.e.c; ++col) {
  const cell = worksheet[XLSX.utils.encode_cell({ c: col, r: range.s.r })];
  headers.push(cell ? cell.v : '');
}

// Step 5: Generate the specified number of rows of fake data with random strings
const data = [];
for (let i = 0; i < argv.rows; i++) {
  const row = [];
  for (let j = 0; j < headers.length; j++) {
    row.push(Math.random().toString(36).slice(-10));
  }
  data.push(row);
}

// Step 6: Create a new XLSX file with the generated data
const newWorksheet = XLSX.utils.aoa_to_sheet([headers, ...data]);
const newWorkbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, sheetName);

// Step 7: Save the output XLSX file to the specified path
XLSX.writeFile(newWorkbook, argv.output);
