// Usage: npm run build && node examples/sample.mjs [path]
// Produces a small workbook that exercises every v0.1 feature; open it in Excel, Numbers and Google Sheets.
import { createWriteStream } from 'node:fs';
import { Writable } from 'node:stream';
import { createWorkbook } from '../dist/index.js';

const path = process.argv[2] ?? 'sample.xlsx';
const wb = createWorkbook();
const saved = wb.readable.pipeTo(Writable.toWeb(createWriteStream(path)));

const header = wb.style({ font: { bold: true, color: '#FFFFFF' }, fill: '#4472C4', border: 'thin', align: { horizontal: 'center' } });
const money = wb.style({ numFmt: '#,##0.00 €' });
const percent = wb.style({ numFmt: '0.0%', font: { italic: true } });
const wrapped = wb.style({ align: { wrap: true, vertical: 'top' } });
const title = wb.style({ font: { size: 16, bold: true } });

const sales = wb.addSheet('Sales', {
  columns: [{ width: 28 }, { width: 14, style: money }, { width: 10, style: percent }, { width: 14 }, { width: 40, style: wrapped }],
  freeze: { rows: 2 },
});
await sales.writeRow(['Quarterly sales report'], { style: title, height: 28 });
sales.merge('A1:E1');
await sales.writeRow(['Product', 'Revenue', 'Share', 'Updated', 'Notes'], { style: header });
const products = [
  ['Coffee & tea', 12500.5, 0.42, new Date(Date.UTC(2026, 9, 1)), 'Line one\nLine two — wrapped text'],
  ['Books "Classics"', 8300, 0.28, new Date(Date.UTC(2026, 9, 2, 14, 30)), '  leading and trailing spaces  '],
  ['Stationery <misc>', 4200.25, 0.14, new Date(Date.UTC(2026, 9, 3)), 'Unicode: Привет, 你好, 😀'],
  ['Returns', -1200, -0.04, new Date(Date.UTC(2026, 9, 4)), { link: 'https://example.com/returns', value: 'Return policy' }],
];
for (const row of products) await sales.writeRow(row);
const last = sales.nextRow - 1;
await sales.writeRow(['Total', { formula: `SUM(B3:B${last})` }, { formula: `SUM(C3:C${last})` }, null, { formula: `COUNTA(A3:A${last})&" products"` }], { style: header });
sales.autoFilter(`A2:E${last}`);

const misc = wb.addSheet('Edge cases', { columns: [{ width: 24 }, { width: 24 }] });
await misc.writeRow(['NaN becomes #NUM!', NaN]);
await misc.writeRow(['Infinity becomes #NUM!', Infinity]);
await misc.writeRow(['Invalid date becomes #NUM!', new Date(NaN)]);
await misc.writeRow(['Date before 1900 becomes #NUM!', new Date(Date.UTC(1899, 0, 1)), 'x']);
await misc.writeRow(['Booleans', true, false]);
await misc.writeRow(['Large number', 1e21]);
await misc.writeRow(['Empty styled cell →', { value: null, style: header }]);
await misc.writeRow(['Control chars stripped', 'a\u0000b\u0001c']);

await wb.close();
await saved;
console.log(`wrote ${path}`);
