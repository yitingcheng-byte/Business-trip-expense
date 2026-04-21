const XlsxPopulate = require('xlsx-populate');

async function test() {
    const wb = await XlsxPopulate.fromBlankAsync();
    const ws = wb.sheet(0);
    ws.range("A5:B5").merged(true);
    
    console.log("Sheet Keys:", Object.keys(ws).filter(k => k.startsWith('_')));
    if (ws._sheetDataNode) {
        console.log("sheetDataNode children:", ws._sheetDataNode.children.length);
    }
    if (ws._mergeCellsNode) {
        console.log("mergeCellsNode children:", ws._mergeCellsNode.children);
    }
}
test();
