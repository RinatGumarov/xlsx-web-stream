// Usage: npm run build && node bench/million-rows.mjs [rows]
// Writes `rows` × 10 cells to a discarding sink and reports throughput and heap usage.
import { createWorkbook } from '../dist/index.js';

const ROWS = Number(process.argv[2] ?? 1_000_000);
const COLUMNS = 10;

const wb = createWorkbook();
let outputBytes = 0;
const done = wb.readable.pipeTo(
  new WritableStream({
    write(chunk) {
      outputBytes += chunk.byteLength;
    },
  }),
);

const bold = wb.style({ font: { bold: true } });
const money = wb.style({ numFmt: '#,##0.00' });
const sheet = wb.addSheet('Bench', {
  columns: Array.from({ length: COLUMNS }, () => ({ width: 16 })),
  freeze: { rows: 1 },
});
await sheet.writeRow(Array.from({ length: COLUMNS }, (_, i) => `Column ${i + 1}`), { style: bold });

const heapSamples = [];
const started = performance.now();
for (let i = 1; i <= ROWS; i++) {
  await sheet.writeRow([
    i,
    `Item ${i}`,
    { value: i * 1.5, style: money },
    i % 2 === 0,
    new Date(Date.UTC(2026, 0, 1 + (i % 365))),
    'Lorem ipsum dolor sit amet',
    i * 3,
    `Code-${i % 1000}`,
    i / 7,
    i % 10 === 0 ? null : 'end',
  ]);
  if (i % 100_000 === 0) {
    const heapMb = process.memoryUsage().heapUsed / 1024 / 1024;
    heapSamples.push(heapMb);
    console.log(`${i.toLocaleString()} rows, heap ${heapMb.toFixed(1)} MB`);
  }
}
await wb.close();
await done;
const seconds = (performance.now() - started) / 1000;

console.log('---');
console.log(`rows: ${ROWS.toLocaleString()} × ${COLUMNS} cells`);
console.log(`time: ${seconds.toFixed(1)} s (${Math.round(ROWS / seconds).toLocaleString()} rows/s)`);
console.log(`output: ${(outputBytes / 1024 / 1024).toFixed(1)} MB`);
console.log(`heap: min ${Math.min(...heapSamples).toFixed(1)} MB, max ${Math.max(...heapSamples).toFixed(1)} MB`);
