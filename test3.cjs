const XlsxPopulate = require('xlsx-populate');

async function test() {
    const wb = await XlsxPopulate.fromBlankAsync();
    const ws = wb.sheet(0);
    console.log(Object.keys(ws.row(1).__proto__));
}
test();
