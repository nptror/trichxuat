// Run with ExcelJS 4.4.0 installed, or set EXCELJS_MODULE. Checks parseCbm finds columns by header in both CBM layouts.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const ExcelJS = require(process.env.EXCELJS_MODULE || 'exceljs');
const html = fs.readFileSync(process.argv[2] || path.join(__dirname, 'DongCont.html'), 'utf8');
const context = vm.createContext({ ExcelJS, Uint8Array, ArrayBuffer, Date });
for (const id of ['embedded-template', 'core']) {
  const start = html.indexOf(`<script id="${id}">`) + `<script id="${id}">`.length;
  vm.runInContext(html.slice(start, html.indexOf('</script>', start)), context);
}
// JSON round-trip: arrays from the vm realm fail deepStrictEqual against this realm's arrays.
const parse = wb => { context.wb = wb; return JSON.parse(vm.runInContext('JSON.stringify(parseCbm(wb))', context)); };
const book = rows => { const wb = new ExcelJS.Workbook(), ws = wb.addWorksheet('CBM'); rows.forEach(r => ws.addRow(r)); return [wb, ws]; };

// Cont layout: Sets/Carton pair per cont, label row ("TOTAL", "CONT 1") under the Sets headers, PO heading rows.
{
  const [wb] = book([
    ['CBM PO# 65395 SV + 65116 NHAT HUY'],
    ['ARTICLE NAME', 'CODE', 'L', 'W', 'H', 'SETS', 'CARTON', 'SETS', 'CARTON'],
    [null, null, null, null, null, 'TOTAL', null, 'CONT 1\nEBKG123', null],
    ['PO# 65395'],
    ['Chair', 'C1', 100, 200, 300, 2, 2, 1, 1],
    ['BOX 2/2', null, 50, 60, 70, null, 2, null, 1],
    ['PO# 65116'],
    ['Crate', 'Q1', 10, 20, 30, 1, 1, 0, 0],
    ['TOTAL'],
  ]);
  const p = parse(wb);
  assert.deepEqual(p.conts.map(c => c.name), ['TOTAL', 'CONT 1']);
  assert.deepEqual(p.rows.map(r => [r.po, r.group, r.code, r.qty, r.pkgs, r.auxiliary]),
    [['65395', 0, 'C1', 2, 2, false], ['65395', 0, '', null, 2, true], ['65116', 1, 'Q1', 1, 1, false]]);
}

// Flat layout of one container: columns reordered, PO column, "CTN SIZE" merged over L/W/H, merged name/code.
{
  const [wb, ws] = book([
    ['CBM PO# 65144 NHF + 65789 SKVN'],
    ['FAC', 'PO NO:', 'Qty\nPCS', 'PART', 'Description', 'No of carton', 'unit/pack', 'CTN SIZE(mm)\n (WxDxH)', null, null, 'CBM'],
    ['NHF', 65144, 2, 'LC5491-A', 'Privacy Chair', 2, 1, 1030, 1030, 1300],
    ['SKVN', '65789 Knoxville', 120, '86P801XC', 'Vanity Base', 60, 2, 1134, 366, 582],
    [null, '65789 Knoxville', 2, null, null, 1, 2, 1134, 366, 582],
    ['TOTAL', null, null, null, null, 63],
  ]);
  ws.mergeCells('H2:J2'); ws.mergeCells('D4:D5'); ws.mergeCells('E4:E5');
  const p = parse(wb);
  assert.deepEqual(p.conts, [{ name: 'TOTAL', sets: 'C', carton: 'F', booking: '', seal: '' }]);
  assert.deepEqual(p.rows.map(r => [r.po, r.group, r.code, r.name, r.qty, r.pkgs, r.L, r.W, r.H]), [
    ['65144', 0, 'LC5491-A', 'Privacy Chair', 2, 2, 1030, 1030, 1300],
    ['65789', 1, '86P801XC', 'Vanity Base', 120, 60, 1134, 366, 582],
    ['65789', 1, '86P801XC', 'Vanity Base', 2, 1, 1134, 366, 582],
  ]);
  assert.deepEqual(p.totals, { qty: 124, pkgs: 63 });
  // An unmerged size header must not silently read the neighbouring columns as dimensions.
  ws.unMergeCells('H2:J2');
  assert.throws(() => parse(wb), /gộp đúng 3 cột/);
}
console.log('PASS: cont layout (label row, PO headings, BOX 2/2); flat layout (header-only, PO column, merged CTN SIZE); unmerged size rejected.');
