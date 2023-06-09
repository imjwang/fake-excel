// Step 1: Install the required packages
// Run the following command in your terminal
// npm install xlsx yargs

// Step 2: Parse command line arguments using yargs
const yargs = require('yargs/yargs');
const { hideBin } = require('yargs/helpers');
const ProgressBar = require('progress');
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

  // Step 2: Create a generateData() function that generates data based on the provided type
function generateData(type) {
  switch (type) {
    case 'int':
      return Math.floor(Math.random() * 10000)+1;
    case 'String':
      return Math.random().toString(36).slice(-10);
    default:
      return type; // Returns a custom string value
  }
}

// Add the following lines after parsing the command line arguments
// Step 3: Create a progress bar instance with a custom format
const progressBar = new ProgressBar('Generating data [:bar] :percent :etas', {
  complete: '=',
  incomplete: ' ',
  width: 20,
  total: argv.rows
});

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
// Step 1: Read the second row of the input XLSX for data type information
const dataTypes = [];
for (let col = range.s.c; col <= range.e.c; ++col) {
  const cell = worksheet[XLSX.utils.encode_cell({ c: col, r: range.s.r + 1 })];
  dataTypes.push(cell ? cell.v : 'String'); // Default to 'String' if undefined
}
// Modify the data generation loop in Step 5
// Step 4: Update the progress bar while generating fake data
const data = [];
// Modify the data generation loop in Step 5
// Step 3: Generate data based on the specified type for each column
for (let i = 0; i < argv.rows; i++) {
  const row = [];
  for (let j = 0; j < headers.length; j++) {
    row.push(generateData(dataTypes[j]));
  }
  data.push(row);
  progressBar.tick(); // Update the progress bar
}
// Step 6: Create a new XLSX file with the generated data
const newWorksheet = XLSX.utils.aoa_to_sheet([headers, ...data]);
const newWorkbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, sheetName);

// Step 7: Save the output XLSX file to the specified path
XLSX.writeFile(newWorkbook, argv.output);
progressBar.terminate();
