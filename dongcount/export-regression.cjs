// Run with ExcelJS 4.4.0 and JSZip 3.10.1 installed, or set their module paths.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const ExcelJS = require(process.env.EXCELJS_MODULE || 'exceljs');
const JSZip = require(process.env.JSZIP_MODULE || 'jszip');
const html = fs.readFileSync(process.argv[2] || path.join(__dirname, 'DongCont.html'), 'utf8');
const context = vm.createContext({ ExcelJS, JSZip, Uint8Array, ArrayBuffer, Date });
for (const id of ['embedded-template', 'core']) {
  vm.runInContext(html.match(new RegExp(`<script id="${id}">([\\s\\S]*?)</script>`))[1], context);
}
async function main() {
  const template = new ExcelJS.Workbook();
  await template.xlsx.load(Buffer.from(vm.runInContext('TEMPLATE_B64', context), 'base64'));
  context.template = template;
  const wb = vm.runInContext(`fillTemplate(template, [{id:1,group:0,po:'65395',name:'Regression item',code:'TEST',qty:2,pkgs:2,L:100,W:200,H:300}], readHeader(template))`, context);
  const raw = await wb.xlsx.writeBuffer();
  const rawZip = await JSZip.loadAsync(raw);
  const names = Object.keys(rawZip.files).filter(n => /^xl\/worksheets\/sheet\d+\.xml$/.test(n));
  assert.equal(names.length, 3);
  let reproduced = 0;
  for (const name of names) {
    if (/<pageSetUpPr\b[^>]*\/>\s*<outlinePr\b/.test(await rawZip.file(name).async('string'))) reproduced++;
  }
  assert.equal(reproduced, 3, 'Pinned ExcelJS must reproduce the defect on all three export sheets');
  context.wb = wb;
  // Before the fix, inspect the actual raw export so this test fails on its bad order.
  const fixed = vm.runInContext('typeof writeExcelCompatibleBuffer', context) === 'function'
    ? await vm.runInContext('writeExcelCompatibleBuffer(wb)', context) : raw;
  const fixedZip = await JSZip.loadAsync(fixed);
  assert.deepEqual(Object.keys(fixedZip.files).sort(), Object.keys(rawZip.files).sort());
  for (const name of names) {
    const xml = await fixedZip.file(name).async('string');
    assert.match(xml, /<outlinePr\b[^>]*\/>\s*<pageSetUpPr\b/, `${name}: outlinePr must precede pageSetUpPr`);
  }
  for (const name of Object.keys(rawZip.files)) {
    if (rawZip.files[name].dir) continue;
    const before = await rawZip.file(name).async('string');
    const after = await fixedZip.file(name).async('string');
    if (names.includes(name)) {
      const expected = before.replace(/(<pageSetUpPr\b[^>]*\/>)(\s*)(<outlinePr\b[^>]*\/>)/g, '$3$2$1');
      assert.equal(after, expected, `${name}: only property order may change`);
    } else {
      // Core timestamps can differ between the two independent ExcelJS serializations.
      if (name !== 'docProps/core.xml') assert.equal(after, before, `${name}: unrelated part changed`);
    }
  }
  const reloaded = new ExcelJS.Workbook();
  await reloaded.xlsx.load(fixed);
  assert.equal(reloaded.getWorksheet('DO').getCell('B17').value, 'Regression item');
  assert.equal(reloaded.getWorksheet('DO').getCell('K17').formula, '(H17*I17*J17)/10^9');
  // Non-default outline settings must survive; a workaround must not drop them.
  const custom = new ExcelJS.Workbook();
  custom.addWorksheet('Custom', {properties:{outlineProperties:{summaryBelow:false,summaryRight:false}},pageSetup:{fitToPage:true}});
  context.custom = custom;
  const customZip = await JSZip.loadAsync(await vm.runInContext('writeExcelCompatibleBuffer(custom)', context));
  assert.match(await customZip.file('xl/worksheets/sheet1.xml').async('string'), /<outlinePr summaryBelow="0" summaryRight="0"\/><pageSetUpPr/);
  assert.match(html, /buffer=await writeExcelCompatibleBuffer\(wb\)/, 'Download handler must use the corrected writer');
  console.log('PASS: real template export; three sheet orders; content preservation; reload; non-default outline settings; download wiring.');
}
main().catch(error => { console.error(error.message); process.exitCode = 1; });
