const XLSX = require('xlsx');

function verifyOutput(filePath) {
    console.log(`--- Verifying ${filePath} ---`);
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
    console.log(`Total Rows: ${data.length}`);
    console.log('Sample data (first 3 rows):');
    data.slice(0, 3).forEach((row, i) => {
        console.log(`Row ${i+1}:`);
        console.log(`  Subtema: ${row['Subtema']}`);
        console.log(`  Descrição: ${row['Descrição do Subtema']}`);
    });
}

verifyOutput('saida.xlsx');
