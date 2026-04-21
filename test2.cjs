const XlsxPopulate = require('xlsx-populate');
async function test() {
    const wb = await XlsxPopulate.fromBlankAsync();
    const ws = wb.sheet(0);
    // test if moveTo exists or anything on range
    console.log(Object.keys(ws.range("A1:C1")));
    if (ws.range("A1:C1").move) console.log("move exists");
}
test();
