import XLSX from 'xlsx';

const workbook = XLSX.readFile('./test.xlsx');

const worksheet = workbook.Sheets.Sheet1!;

XLSX.utils.sheet_add_aoa(worksheet, [['stop using excel as a database']], { origin: 'E5' });

console.log(worksheet);

XLSX.writeFile(workbook, './test.xlsx');
