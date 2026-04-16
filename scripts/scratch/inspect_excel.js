const XLSX = require('xlsx');

function inspectExcel(filePath) {
    console.log(`--- Inspecting ${filePath} ---`);
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet);
    console.log(`Sheet Name: ${sheetName}`);
    console.log(`Total Rows: ${data.length}`);
    if (data.length > 0) {
        console.log('Columns:', Object.keys(data[0]));
        console.log('First 2 rows:', JSON.stringify(data.slice(0, 2), null, 2));
    }
}

inspectExcel('tema.xlsx');
inspectExcel('subtema.xlsx');
