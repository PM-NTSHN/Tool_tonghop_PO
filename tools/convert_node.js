const X=require(process.env.XLSX_PATH || 'xlsx-js-style');const C=require('../src/converter.js');const fs=require('fs');
const wb=X.read(fs.readFileSync(process.argv[2]),{type:'buffer'});
const {rows:raw}=C.sheetToRows(X,wb.Sheets[wb.SheetNames[0]]);
const r=C.convert(raw);console.log('rows',r.rows.length);r.warnings.forEach(w=>console.log('WARN',w));
fs.writeFileSync(process.argv[3],C.writeClean(X,r.rows));
