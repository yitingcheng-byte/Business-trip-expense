const ExcelJS = require('exceljs');
const fs = require('fs');

async function analyzeTemplate() {
    const workbook = new ExcelJS.Workbook();
    try {
        console.log("Current working directory: ", process.cwd());
        const filePath = '../public/templates/expense_template.xlsx';
        console.log("File exists: ", fs.existsSync(filePath));

        await workbook.xlsx.readFile(filePath);
        const ws = workbook.worksheets[0];
        
        console.log(`Sheet Name: ${ws.name}`);
        console.log(`Total Rows: ${ws.rowCount}`);
        
        for (let r = 1; r <= 35; r++) {
            const row = ws.getRow(r);
            let rowData = [];
            row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                let val = cell.value;
                if (val !== null && val !== undefined) {
                  if (typeof val === 'object' && val !== null) {
                      if (val.richText) val = val.richText.map(rt => rt.text).join('');
                      else val = JSON.stringify(val);
                  }
                  rowData.push(`[${cell.address}]: ${val}`);
                }
            });
            if (rowData.length > 0) {
                console.log(`Row ${r}: ` + rowData.join(', '));
            }
        }
        
        console.log("--- Merged Cells ---");
        for (let m in ws._merges) {
            console.log(m);
        }
    } catch(e) {
        console.error("Error reading file", e);
    }
}
analyzeTemplate();
