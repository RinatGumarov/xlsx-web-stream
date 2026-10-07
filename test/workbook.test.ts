import * as ExcelJS from 'exceljs';
import { describe, expect, it } from 'vitest';
import { createWorkbook } from '../src/index.js';
import { collect, parseXml, unzipText } from './helpers.js';

async function readBack(bytes: Uint8Array): Promise<ExcelJS.Workbook> {
  const workbook = new ExcelJS.Workbook();
  // exceljs wants a Node Buffer (tests may use Node APIs, src/ may not). Its typings declare a global `Buffer`
  // that clashes with @types/node's, hence the cast.
  await workbook.xlsx.load(Buffer.from(bytes) as unknown as Parameters<ExcelJS.Workbook['xlsx']['load']>[0]);
  return workbook;
}

describe('createWorkbook end to end', () => {
  it('produces a package with every required part, readable by an independent library', async () => {
    const wb = createWorkbook();
    const header = wb.style({ font: { bold: true }, fill: '#EEEEEE', border: 'thin' });
    const money = wb.style({ numFmt: '#,##0.00', align: { horizontal: 'right' } });
    const output = collect(wb.readable);

    const report = wb.addSheet('Report', { columns: [{ width: 30 }, { width: 14, style: money }, { width: 12 }], freeze: { rows: 1 } });
    await report.writeRow(['Item', 'Amount', 'Date'], { style: header });
    await report.writeRow(['Coffee & tea', 12.5, new Date(Date.UTC(2026, 9, 6))]);
    await report.writeRow(['Books', 1000, new Date(Date.UTC(2026, 0, 1, 12))]);
    await report.writeRow(['Total', { formula: 'SUM(B2:B3)' }, { link: 'https://example.com', value: 'Details' }]);
    await report.writeRow([true, false, null]);
    report.autoFilter('A1:C1');
    report.merge('A6:C6');
    await report.writeRow(['Merged title']);

    const second = wb.addSheet('Second');
    await second.writeRow([NaN, 'plain']);
    await wb.close();

    const bytes = await output;
    const files = unzipText(bytes);
    expect(Object.keys(files).sort()).toEqual(
      ['[Content_Types].xml', '_rels/.rels', 'xl/_rels/workbook.xml.rels', 'xl/styles.xml', 'xl/workbook.xml', 'xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml'].sort(),
    );
    const contentTypes = parseXml(files['[Content_Types].xml']!);
    expect(contentTypes.Types.Override.map((o: any) => o.PartName)).toContain('/xl/worksheets/sheet2.xml');
    const workbookXml = parseXml(files['xl/workbook.xml']!);
    expect(workbookXml.workbook.sheets.sheet.map((s: any) => s.name)).toEqual(['Report', 'Second']);
    expect(files['xl/workbook.xml']).toContain(`'Report'!$A$1:$C$1`);
    expect(files['xl/worksheets/sheet1.xml']).toContain('<autoFilter ref="A1:C1"/><mergeCells count="1"><mergeCell ref="A6:C6"/></mergeCells>');
    expect(files['xl/worksheets/sheet2.xml']).toContain('<c r="A1" t="e"><v>#NUM!</v></c>');

    const excel = await readBack(bytes);
    expect(excel.worksheets.map((ws) => ws.name)).toEqual(['Report', 'Second']);
    const ws = excel.getWorksheet('Report')!;
    expect(ws.getCell('A1').value).toBe('Item');
    expect(ws.getCell('A1').font.bold).toBe(true);
    expect(ws.getCell('A1').fill).toMatchObject({ type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEEEEEE' } });
    expect(ws.getCell('A1').border.top?.style).toBe('thin');
    expect(ws.getCell('A2').value).toBe('Coffee & tea');
    expect(ws.getCell('B2').value).toBe(12.5);
    expect(ws.getCell('B2').numFmt).toBe('#,##0.00');
    expect(ws.getCell('B2').alignment.horizontal).toBe('right');
    expect(ws.getCell('C2').value).toBeInstanceOf(Date);
    expect((ws.getCell('C2').value as Date).getTime()).toBe(Date.UTC(2026, 9, 6));
    expect((ws.getCell('C3').value as Date).getTime()).toBe(Date.UTC(2026, 0, 1, 12));
    expect(ws.getCell('C2').numFmt).toBe('yyyy-mm-dd');
    expect(ws.getCell('B4').value).toMatchObject({ formula: 'SUM(B2:B3)' });
    expect(ws.getCell('C4').value).toMatchObject({ formula: 'HYPERLINK("https://example.com","Details")', result: 'Details' });
    expect(ws.getCell('C4').font.underline).toBeTruthy();
    expect(ws.getCell('A5').value).toBe(true);
    expect(ws.getCell('B5').value).toBe(false);
    expect(ws.getCell('C5').value).toBeNull();
    expect(ws.getColumn(1).width).toBe(30);
    expect(ws.getColumn(2).width).toBe(14);
    expect(ws.views[0]).toMatchObject({ state: 'frozen', ySplit: 1 });
    expect(ws.autoFilter).toBe('A1:C1');
    expect(ws.model.merges).toContain('A6:C6');
    expect(excel.getWorksheet('Second')!.getCell('A1').value).toMatchObject({ error: '#NUM!' });
    expect(excel.getWorksheet('Second')!.getCell('B1').value).toBe('plain');
  });

  it('supports local date mode and a custom default date format', async () => {
    const wb = createWorkbook({ dates: 'local', dateFormat: 'dd.mm.yyyy hh:mm' });
    const output = collect(wb.readable);
    const sheet = wb.addSheet('Dates');
    await sheet.writeRow([new Date(2026, 9, 6, 15, 30)]);
    await wb.close();
    const excel = await readBack(await output);
    const cell = excel.getWorksheet('Dates')!.getCell('A1');
    expect((cell.value as Date).getTime()).toBe(Date.UTC(2026, 9, 6, 15, 30));
    expect(cell.numFmt).toBe('dd.mm.yyyy hh:mm');
  });

  it('closes the previous sheet implicitly when a new one is added', async () => {
    const wb = createWorkbook();
    const output = collect(wb.readable);
    const a = wb.addSheet('A');
    void a.writeRow([1]);
    const b = wb.addSheet('B');
    await b.writeRow([2]);
    await wb.close();
    const files = unzipText(await output);
    expect(files['xl/worksheets/sheet1.xml']).toContain('<c r="A1"><v>1</v></c>');
    expect(files['xl/worksheets/sheet2.xml']).toContain('<c r="A1"><v>2</v></c>');
    expect(() => a.writeRow([3])).toThrow('already closed');
  });

  it('validates sheet names, including control characters', () => {
    const wb = createWorkbook();
    wb.addSheet('Report');
    expect(() => wb.addSheet('report')).toThrow(/already used/);
    expect(() => wb.addSheet('REPORT')).toThrow(/already used/);
    for (const bad of ['', 'x'.repeat(32), 'a[b', 'a]b', 'a:b', 'a*b', 'a?b', 'a/b', 'a\\b', "'quoted", "quoted'", 'tab\there', 'line\nbreak', 'cr\rhere', 'nul\u0000', 'unit\u001Fsep']) {
      expect(() => wb.addSheet(bad), JSON.stringify(bad)).toThrow(TypeError);
    }
    expect(() => wb.addSheet('x'.repeat(31))).not.toThrow();
    wb.abort();
  });

  it('handles 100k rows with constant-size chunks and a valid result', async () => {
    const wb = createWorkbook();
    let chunks = 0;
    let largest = 0;
    const output = wb.readable.pipeTo(
      new WritableStream({
        write(chunk) {
          chunks++;
          largest = Math.max(largest, chunk.byteLength);
        },
      }),
    );
    const sheet = wb.addSheet('Big');
    for (let i = 1; i <= 100_000; i++) await sheet.writeRow([i, `row ${i}`, i % 7 === 0]);
    expect(sheet.nextRow).toBe(100_001);
    await wb.close();
    await output;
    expect(chunks).toBeGreaterThan(100);
    expect(largest).toBeLessThanOrEqual(256 * 1024);
  });
});

describe('failure propagation', () => {
  it('a sheet added after the consumer cancelled rejects with the cancel reason', async () => {
    const wb = createWorkbook();
    const reader = wb.readable.getReader();
    const a = wb.addSheet('A');
    // Enough XML to force flushes, i.e. an open ZIP entry. Writes still pending at cancel time reject; observe them
    // up front so they do not surface as unhandled rejections.
    const pending = Promise.allSettled(Array.from({ length: 2000 }, () => a.writeRow(['x'.repeat(100)])));
    await reader.cancel('download stopped');
    await pending;
    const b = wb.addSheet('B');
    await expect(b.writeRow([1])).rejects.toBe('download stopped');
    await expect(wb.close()).rejects.toBe('download stopped');
  });

  it('abort() followed by close() on an empty workbook rejects with the abort reason', async () => {
    const wb = createWorkbook();
    const reason = new Error('aborted by caller');
    wb.abort(reason);
    await expect(wb.close()).rejects.toBe(reason);
  });

  it('a failed addSheet leaves the previous sheet writable', async () => {
    const wb = createWorkbook();
    const output = collect(wb.readable);
    const a = wb.addSheet('A');
    expect(() => wb.addSheet('B', { freeze: { rows: -1 } })).toThrow(RangeError);
    await a.writeRow([1]);
    await wb.close();
    const files = unzipText(await output);
    expect(files['xl/worksheets/sheet1.xml']).toContain('<c r="A1"><v>1</v></c>');
    expect(Object.keys(files)).not.toContain('xl/worksheets/sheet2.xml');
  });
});
