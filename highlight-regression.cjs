const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const XLSX = require(process.env.SHEETJS_MODULE || 'xlsx');
const ExcelJS = require(process.env.EXCELJS_MODULE || 'exceljs');
const html = fs.readFileSync(process.argv[2] || path.join(__dirname, 'index.html'), 'utf8');
const fixture = process.argv[3] || path.join(__dirname, 'gửi thịnh.xlsx');
const script = html.match(/<script>\s*('use strict';[\s\S]*?)<\/script>/)[1];
// Run application and ExcelJS in the same realm, as they run together in a browser.
const pureCode = script.slice(0, script.indexOf('const state='))
  + script.slice(script.indexOf('function toTSV('), script.indexOf('function render('))
  + script.slice(script.indexOf('async function exportXlsx('), script.indexOf('async function download('));
const api = new Function('XLSX','ExcelJS',pureCode+';return {FIELDS,parseWorkbook,toTSV,exportXlsx,parseVNDate,declarationInfo};')(XLSX,ExcelJS);
const context = vm.createContext(api);
async function main() {
  const wb = XLSX.read(fs.readFileSync(fixture), {type:'buffer', cellDates:true});
  context.input = wb;
  const parsed = vm.runInContext('parseWorkbook(input)', context);
  assert.equal(parsed.errors.length, 0);
  assert.equal(parsed.rows.length, 4);
  const row = parsed.rows[0];
  const expected = {
    soTK:['108723643450','E4'],loaiHinh:['C11','P6'],coQuanHQ:['HQTTHUAN','L7'],
    nguoiXK:['DONGGUAN BAIYOU ELECTRONIC LIMITED','H23'],luuKho:['06B1A04','U30'],
    soKien:[8,'K36'],dvtKien:['PK','P36'],tongGross:[68,'K37'],dvtGross:['KGM','P37'],
    maHS:['94054190','G149'],sl1:[35,'V153'],dvt1:['SET','AE153'],sl2:[35,'V154'],dvt2:['PCE','AE154'],
    triGia:[577.5,'I155'],donGia:[16.5,'V155'],tienTe:['USD','AC155'],dvtDonGia:['SET','AE155'],
    maXuatXu:['CN','X160'],tenXuatXu:['CHINA','Z160'],maKemXuatXu:['B30','AC160']
  };
  for(const [key,[value,addr]] of Object.entries(expected)) {
    assert.equal(row.values[key],value,`${key}: highlighted value`);
    assert.equal(row.src[key],addr,`${key}: source address`);
  }
  assert.equal(row.src.ngayDK,'G8');
  assert.equal(row.src.moTa,'G150');
  assert.ok(row.values.moTa.startsWith('PO# 66011#&'));
  const expectedInfo='FROM SAVIMEX\nTK 108723643450 - C11\nDATE: 15/10/2026\nHQTTHUAN / 06B1A04';
  assert.equal(row.values.thongTin,expectedInfo);
  assert.equal(vm.runInContext('FIELDS.at(-1).key',context),'thongTin');
  context.rows=[row];
  const tsv=vm.runInContext("toTSV(rows,false,FIELDS.filter(f=>f.key==='thongTin'))",context);
  assert.equal(tsv,'"'+expectedInfo+'"','Copy must retain four lines in one quoted cell');
  context.rows=[{values:{moTa:'A "quoted"\nitem'}}];
  assert.equal(vm.runInContext("toTSV(rows,false,FIELDS.filter(f=>f.key==='moTa'))",context),'"A ""quoted""\nitem"');
  context.rows=parsed.rows;
  const buffer=await vm.runInContext('exportXlsx(rows)',context);
  const out=new ExcelJS.Workbook();await out.xlsx.load(buffer);
  const ws=out.getWorksheet('Output');
  const fields=vm.runInContext('FIELDS',context);
  assert.equal(ws.getRow(1).getCell(fields.length).value,'Thông tin');
  assert.equal(ws.getRow(2).getCell(fields.length).value,expectedInfo);
  assert.equal(ws.getRow(2).getCell(fields.length).alignment.wrapText,true);
  assert.ok(ws.getRow(2).height>=64,'Four-line information must fit the row');
  for(const [key,[value]] of Object.entries(expected))assert.equal(ws.getCell(2,fields.findIndex(f=>f.key===key)+1).value,value);
  assert.equal(vm.runInContext("declarationInfo({soTK:'309014359050',loaiHinh:'E62 9',ngayDK:parseVNDate('30/09/2026'),coQuanHQ:'HQTTHUAN',luuKho:'02PJOZZ'})",context),'FROM SAVIMEX\nTK 309014359050 - E62\nDATE: 30/09/2026\nHQTTHUAN / 02PJOZZ');
  // Relocated values must still follow their labels rather than original addresses.
  const moved=XLSX.read(fs.readFileSync(fixture),{type:'buffer',cellDates:true});
  const sheet=moved.Sheets.TKN;
  for(const [from,to] of [['K36','L36'],['K37','L37'],['X160','W160']]){
    const points=[from,to].map(XLSX.utils.decode_cell);
    sheet['!merges']=(sheet['!merges']||[]).filter(m=>!points.some(p=>p.r>=m.s.r&&p.r<=m.e.r&&p.c>=m.s.c&&p.c<=m.e.c));
    sheet[to]=sheet[from];delete sheet[from];
  }
  context.input=moved;
  const relocated=vm.runInContext('parseWorkbook(input)',context).rows[0];
  assert.equal(relocated.values.soKien,8);assert.equal(relocated.src.soKien,'L36');
  assert.equal(relocated.values.tongGross,68);assert.equal(relocated.src.tongGross,'L37');
  assert.equal(relocated.values.maXuatXu,'CN');assert.equal(relocated.src.maXuatXu,'W160');
  // Missing header quantity must not accidentally read the adjacent section label.
  for(const a of ['L36','P36'])delete sheet[a];
  const missing=vm.runInContext('parseWorkbook(input)',context).rows[0];
  assert.equal(missing.values.soKien,null);assert.equal(missing.values.dvtKien,'');
  console.log(`PASS: highlighted fields and source cells; ${fields.length} output columns; fixed FROM with dynamic declaration; multiline clipboard and XLSX; moved and missing inputs.`);
}
main().catch(e=>{console.error(e);process.exitCode=1;});
