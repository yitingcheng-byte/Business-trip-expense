const ExcelJS = require('exceljs');
const fs = require('fs');

async function createTemplate() {
  const workbook = new ExcelJS.Workbook();
  const ws = workbook.addWorksheet('報銷明細');

  ws.getCell('A13').value = "出差人：";
  ws.getCell('E13').value = "工號：";
  ws.getCell('I13').value = "單位：";
  ws.getCell('P13').value = "部門：";
  ws.getCell('V13').value = "出差期間：";

  ws.getCell('A14').value = "日期";
  ws.getCell('C14').value = "行程";
  ws.getCell('M14').value = "報支金額";
  ws.getCell('Y14').value = "專案代號 (備註6)";
  ws.getCell('AB14').value = "交通工具";

  ws.getCell('C15').value = "地點";
  ws.getCell('G15').value = "費用說明(備註5)";
  ws.getCell('M15').value = "幣別";
  ws.getCell('O15').value = "交通費";
  ws.getCell('Q15').value = "住宿費";
  ws.getCell('S15').value = "膳雜費";
  ws.getCell('U15').value = "交際費";
  ws.getCell('W15').value = "其他費用";

  ws.mergeCells('C13:D13');
  ws.mergeCells('G13:H13');
  ws.mergeCells('K13:O13');
  ws.mergeCells('R13:U13');
  ws.mergeCells('Y13:AD13');

  ws.mergeCells('A14:B15');
  ws.mergeCells('C14:L14');
  ws.mergeCells('C15:F15');
  ws.mergeCells('G15:L15');
  ws.mergeCells('M14:X14');
  ws.mergeCells('M15:N15');
  ws.mergeCells('O15:P15');
  ws.mergeCells('Q15:R15');
  ws.mergeCells('S15:T15');
  ws.mergeCells('U15:V15');
  ws.mergeCells('W15:X15');
  ws.mergeCells('Y14:AA15');
  ws.mergeCells('AB14:AD15');

  ws.getCell('A18').value = "費用報支合計";
  ws.getCell('F18').value = "幣別";
  ws.getCell('H18').value = "合計";
  ws.getCell('K18').value = "已先預支費用";
  ws.getCell('P18').value = "幣別";
  ws.getCell('R18').value = "合計";
  ws.getCell('U18').value = "應付員工或員工繳回";
  ws.getCell('Z18').value = "幣別";
  ws.getCell('AB18').value = "合計";

  ws.mergeCells('A18:E18');
  ws.mergeCells('F18:G18');
  ws.mergeCells('H18:J18');
  ws.mergeCells('K18:O18');
  ws.mergeCells('P18:Q18');
  ws.mergeCells('R18:T18');
  ws.mergeCells('U18:Y18');
  ws.mergeCells('Z18:AA18');
  ws.mergeCells('AB18:AD18');

  ws.mergeCells('A19:E19');
  ws.mergeCells('F19:G19');
  ws.mergeCells('H19:J19');
  ws.mergeCells('K19:O19');
  ws.mergeCells('P19:Q19');
  ws.mergeCells('R19:T19');
  ws.mergeCells('U19:Y19');
  ws.mergeCells('Z19:AA19');
  ws.mergeCells('AB19:AD19');

  if (!fs.existsSync('public/templates')) {
    fs.mkdirSync('public/templates', { recursive: true });
  }

  await workbook.xlsx.writeFile('public/templates/expense_template.xlsx');
  console.log('Template created!');
}

createTemplate().catch(console.error);
