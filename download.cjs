const https = require('https');
const fs = require('fs');
const file = fs.createWriteStream("expense_template.xlsx");
https.get("https://ais-dev-ven5pp67sbo57hxthtkcdq-433672822595.asia-northeast1.run.app/Business-trip-expense/templates/expense_template.xlsx", function(response) {
  response.pipe(file);
});
