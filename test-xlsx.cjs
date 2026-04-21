const ExcelJS = require('exceljs');
async function test() {
    const wb = new ExcelJS.Workbook();
    try {
        await wb.xlsx.readFile('./expense_template.xlsx');
        const ws = wb.worksheets[0];
        console.log("Images: ", ws.getImages().length);
        console.log("Header/Footer: ", ws.headerFooter);
    } catch(e) { console.error(e); }
}
test();
