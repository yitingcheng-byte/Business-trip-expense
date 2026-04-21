const ExcelJS = require('exceljs');

async function analyzeTemplate() {
    const workbook = new ExcelJS.Workbook();
    try {
        await workbook.xlsx.readFile('public/templates/expense_template.xlsx');
        const ws = workbook.worksheets[0];
        
        console.log(`Sheet Name: ${ws.name}`);
        console.log(`Total Rows: ${ws.rowCount}`);
        
        // Let's print out the first 30 rows to see what text is where
        for (let r = 1; r <= 30; r++) {
            const row = ws.getRow(r);
            let rowData = [];
            row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
                rowData.push(`[${cell.address}]: ${cell.value}`);
            });
            if (rowData.length > 0) {
                console.log(`Row ${r}: ` + rowData.join(', '));
            }
        }
        
        console.log("--- Merged Cells ---");
        const merges = Object.values(ws._merges || {});
        for (let m in ws._merges) {
            console.log(m);
        }
    } catch(e) {
        console.error("Error reading file", e);
    }
}
analyzeTemplate();
