const XlsxPopulate = require('xlsx-populate');

async function test() {
    const wb = await XlsxPopulate.fromBlankAsync();
    const ws = wb.sheet(0);
    ws.cell("A1").value("Test1");
    ws.cell("A2").value("Test2");
    ws.range("B2:C2").merged(true);
    
    // insert row at 2
    ws.row(2).insert();
    
    console.log("A1:", ws.cell("A1").value());
    console.log("A2:", ws.cell("A2").value());
    console.log("A3:", ws.cell("A3").value());
    console.log("Merge shifted?", ws.cell("B3").merged());
}
test();
