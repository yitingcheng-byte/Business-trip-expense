const XlsxPopulate = require('xlsx-populate');

async function test() {
    const wb = await XlsxPopulate.fromBlankAsync();
    const ws = wb.sheet(0);
    ws.range("A5:B5").merged(true);
    console.log(ws._mergeCells);
}
test();
