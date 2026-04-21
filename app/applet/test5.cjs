const ExcelJS = require('exceljs');
async function test() {
    const wb = new ExcelJS.Workbook();
    // I don't have the template on disk! I'm blind!
    console.log("No template here!");
}
test();
