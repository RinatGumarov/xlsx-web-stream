import { describe, expect, it } from 'vitest';
import type { CellContext } from '../src/cell.js';
import { MAX_ROWS } from '../src/cell.js';
import { SheetWriter } from '../src/sheet.js';
import type { Style } from '../src/types.js';
import { ZipWriter } from '../src/zip/zip-writer.js';
import { MemoryEntry, MemorySink, unzipText } from './helpers.js';

const style = (id: number): Style => ({ id });
const ctx: CellContext = { dates: 'utc', styleId: (s) => s?.id, dateStyle: (s) => (s === 0 ? 99 : s), linkStyle: () => 7 };

const HEADER_START = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';

function sheet(options = {}, first = true) {
  const entry = new MemoryEntry();
  return { entry, sheet: new SheetWriter('S', options, { openEntry: () => entry, ctx, first }) };
}

describe('SheetWriter header', () => {
  it('writes a minimal header for the first sheet', async () => {
    const { entry, sheet: s } = sheet();
    await s.close();
    expect(entry.text()).toBe(
      `${HEADER_START}<sheetViews><sheetView workbookViewId="0" tabSelected="1"></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData></sheetData></worksheet>`,
    );
    expect(entry.closed).toBe(true);
  });

  it('marks only the first sheet as selected', async () => {
    const { entry, sheet: s } = sheet({}, false);
    await s.close();
    expect(entry.text()).toContain('<sheetView workbookViewId="0"></sheetView>');
  });

  it('freezes rows and/or columns with the right active pane', async () => {
    const both = sheet({ freeze: { rows: 1, columns: 2 } });
    await both.sheet.close();
    expect(both.entry.text()).toContain('<pane xSplit="2" ySplit="1" topLeftCell="C2" activePane="bottomRight" state="frozen"/>');
    const rows = sheet({ freeze: { rows: 3 } });
    await rows.sheet.close();
    expect(rows.entry.text()).toContain('<pane ySplit="3" topLeftCell="A4" activePane="bottomLeft" state="frozen"/>');
    const cols = sheet({ freeze: { columns: 1 } });
    await cols.sheet.close();
    expect(cols.entry.text()).toContain('<pane xSplit="1" topLeftCell="B1" activePane="topRight" state="frozen"/>');
    const none = sheet({ freeze: { rows: 0 } });
    await none.sheet.close();
    expect(none.entry.text()).not.toContain('<pane');
  });

  it('writes column widths and styles', async () => {
    const { entry, sheet: s } = sheet({ columns: [{ width: 30 }, {}, { style: style(4) }, { width: 12.5, style: style(2) }] });
    await s.close();
    expect(entry.text()).toContain(
      '<cols><col min="1" max="1" width="30" customWidth="1"/><col min="3" max="3" width="8.43" style="4"/><col min="4" max="4" width="12.5" customWidth="1" style="2"/></cols>',
    );
  });

  it('validates options', () => {
    expect(() => sheet({ columns: 'wide' })).toThrow(TypeError);
    expect(() => sheet({ columns: [{ width: 300 }] })).toThrow(RangeError);
    expect(() => sheet({ columns: [{ width: -1 }] })).toThrow(RangeError);
    expect(() => sheet({ freeze: { rows: 1.5 } })).toThrow(RangeError);
    expect(() => sheet({ freeze: { columns: -1 } })).toThrow(RangeError);
    expect(() => sheet({ freeze: 'top' })).toThrow(TypeError);
  });
});

describe('SheetWriter rows', () => {
  it('writes rows with increasing numbers and tracks nextRow', async () => {
    const { entry, sheet: s } = sheet();
    expect(s.nextRow).toBe(1);
    await s.writeRow(['a', 1]);
    await s.writeRow([null, true]);
    expect(s.nextRow).toBe(3);
    await s.close();
    expect(entry.text()).toContain(
      '<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>a</t></is></c><c r="B1"><v>1</v></c></row><row r="2"><c r="B2" t="b"><v>1</v></c></row></sheetData>',
    );
  });

  it('applies row style, column style and cell style with the right precedence', async () => {
    const { entry, sheet: s } = sheet({ columns: [{ style: style(3) }] });
    await s.writeRow([1, 2]); // column style on A only
    await s.writeRow([1, 2, { value: 3, style: style(8) }], { style: style(5) }); // row style beats column style, cell beats row
    await s.close();
    const xml = entry.text();
    expect(xml).toContain('<row r="1"><c r="A1" s="3"><v>1</v></c><c r="B1"><v>2</v></c></row>');
    expect(xml).toContain('<row r="2" s="5" customFormat="1"><c r="A2" s="5"><v>1</v></c><c r="B2" s="5"><v>2</v></c><c r="C2" s="8"><v>3</v></c></row>');
  });

  it('lets a row style that resolves to the default xf override the column style', async () => {
    const { entry, sheet: s } = sheet({ columns: [{ style: style(3) }] });
    await s.writeRow([1], { style: style(0) }); // xf 0 is still an explicit row style: cell ?? row ?? column
    await s.close();
    const xml = entry.text();
    expect(xml).toContain('<row r="1"><c r="A1"><v>1</v></c></row>');
    expect(xml).not.toContain('s="3"');
  });

  it('writes custom row heights and validates them', async () => {
    const { entry, sheet: s } = sheet();
    await s.writeRow(['x'], { height: 30 });
    expect(() => s.writeRow(['x'], { height: 500 })).toThrow(RangeError);
    expect(() => s.writeRow(['x'], { height: 'tall' as never })).toThrow(RangeError);
    await s.close();
    expect(entry.text()).toContain('<row r="1" ht="30" customHeight="1">');
  });

  it('writeRows consumes sync and async iterables', async () => {
    const { entry, sheet: s } = sheet();
    await s.writeRows([[1], [2]]);
    async function* more() {
      yield [3];
      yield [4];
    }
    await s.writeRows(more());
    await s.close();
    expect(entry.text()).toMatch(/<row r="1">.*<row r="2">.*<row r="3">.*<row r="4">/);
  });

  it('flushes to the entry in chunks before close', async () => {
    const { entry, sheet: s } = sheet();
    const wide = Array.from({ length: 50 }, (_, i) => `cell number ${i} with some padding text`);
    for (let i = 0; i < 200; i++) await s.writeRow(wide);
    expect(entry.chunks.length).toBeGreaterThan(1);
    await s.close();
    expect(entry.text()).toContain('<row r="200">');
  });

  it('rejects non-array rows, too many cells and too many rows', () => {
    const { sheet: s } = sheet();
    expect(() => s.writeRow('abc' as never)).toThrow(TypeError);
    expect(() => s.writeRow(new Array(16_385).fill(1))).toThrow(RangeError);
    for (let i = 0; i < MAX_ROWS; i++) void s.writeRow([]);
    expect(() => s.writeRow([1])).toThrow(RangeError);
  });
});

describe('SheetWriter merges, filters and lifecycle', () => {
  it('does not open its ZIP entry until the first flush', async () => {
    const entry = new MemoryEntry();
    let opened = 0;
    const s = new SheetWriter('S', {}, { openEntry: () => { opened++; return entry; }, ctx, first: true });
    expect(opened).toBe(0);
    await s.writeRow(['a']); // far below the flush threshold
    expect(opened).toBe(0);
    await s.close();
    expect(opened).toBe(1);
    expect(entry.text()).toContain('<row r="1">');
  });

  it('rejects invalid options without opening the entry', () => {
    const entry = new MemoryEntry();
    let opened = 0;
    expect(
      () => new SheetWriter('S', { columns: [{ width: 300 }] }, { openEntry: () => { opened++; return entry; }, ctx, first: true }),
    ).toThrow(RangeError);
    expect(opened).toBe(0);
  });

  it('writes autoFilter before mergeCells after sheetData', async () => {
    const { entry, sheet: s } = sheet();
    await s.writeRow(['a', 'b', 'c']);
    s.merge('C3:A3');
    s.autoFilter('A1:C1');
    s.merge('A5:A6');
    await s.close();
    expect(entry.text()).toContain(
      '</sheetData><autoFilter ref="A1:C1"/><mergeCells count="2"><mergeCell ref="A3:C3"/><mergeCell ref="A5:A6"/></mergeCells></worksheet>',
    );
    expect(s.autoFilterRange).toEqual({ c1: 1, r1: 1, c2: 3, r2: 1 });
  });

  it('rejects single-cell merges, malformed ranges and a second auto filter', () => {
    const { sheet: s } = sheet();
    expect(() => s.merge('A1:A1')).toThrow(TypeError);
    expect(() => s.merge('A1')).toThrow(TypeError);
    s.autoFilter('A1:B1');
    expect(() => s.autoFilter('A1:C1')).toThrow('already has an auto filter');
  });

  it('rejects a duplicate merge and a merge that overlaps an existing one', async () => {
    const { entry, sheet: s } = sheet();
    s.merge('A1:B2');
    expect(() => s.merge('A1:B2')).toThrow(Error);
    expect(() => s.merge('B2:A1')).toThrow('Merge range A1:B2 overlaps the existing merge A1:B2'); // corner order is normalized
    expect(() => s.merge('B2:C3')).toThrow('Merge range B2:C3 overlaps the existing merge A1:B2');
    expect(() => s.merge('A2:A5')).toThrow(/overlaps the existing merge A1:B2/); // shares a column edge cell
    expect(() => s.merge('A1:XFD1')).toThrow(/overlaps/);
    s.merge('D4:E5');
    expect(() => s.merge('C3:D4')).toThrow('Merge range C3:D4 overlaps the existing merge D4:E5');
    expect(() => s.merge('A1:E5')).toThrow(/overlaps/); // fully contains both
    await s.close();
    // Rejected merges leave no trace in the output.
    expect(entry.text()).toContain('<mergeCells count="2"><mergeCell ref="A1:B2"/><mergeCell ref="D4:E5"/></mergeCells>');
  });

  it('accepts merges that only touch edges without sharing a cell', async () => {
    const { entry, sheet: s } = sheet();
    s.merge('A1:B2');
    s.merge('C1:D2'); // adjacent to the right
    s.merge('A3:B4'); // adjacent below
    s.merge('C3:D4'); // diagonal neighbour
    await s.close();
    expect(entry.text()).toContain('<mergeCells count="4">');
  });

  it('records 100k non-overlapping row merges in well under a second', { timeout: 30_000 }, async () => {
    // Row-by-row merges are the common case; the overlap check must not rescan every earlier merge for each of them.
    const { entry, sheet: s } = sheet();
    const start = performance.now();
    for (let row = 1; row <= 100_000; row++) s.merge(`A${row}:C${row}`);
    expect(performance.now() - start).toBeLessThan(1000);
    await s.close();
    expect(entry.text()).toContain('<mergeCells count="100000"><mergeCell ref="A1:C1"/>');
  });

  it('still rejects overlaps that start at or above the bottom row of earlier merges', () => {
    const { sheet: s } = sheet();
    s.merge('A1:B2');
    s.merge('A10:B12'); // the bottom-most merge so far ends at row 12
    expect(() => s.merge('A5:A11')).toThrow('Merge range A5:A11 overlaps the existing merge A10:B12'); // reaches down into it
    expect(() => s.merge('B2:C3')).toThrow('Merge range B2:C3 overlaps the existing merge A1:B2'); // overlaps an older, higher merge
    expect(() => s.merge('A12:C12')).toThrow('Merge range A12:C12 overlaps the existing merge A10:B12'); // starts exactly at the bottom row
    expect(() => s.merge('A1:A20')).toThrow(/overlaps the existing merge A1:B2/); // spans everything
    s.merge('C12:D12'); // same bottom row, no shared cell
    s.merge('A13:B13'); // first row below every earlier merge
    s.merge('A3:B9'); // fits between the two without touching either
    expect(() => s.merge('A13:A14')).toThrow('Merge range A13:A14 overlaps the existing merge A13:B13'); // bottom row moved to 13
  });

  it('refuses every operation after close and keeps close idempotent', async () => {
    const { sheet: s } = sheet();
    const first = s.close();
    expect(s.close()).toBe(first);
    await first;
    expect(() => s.writeRow([1])).toThrow('already closed');
    expect(() => s.merge('A1:B1')).toThrow('already closed');
    expect(() => s.autoFilter('A1:B1')).toThrow('already closed');
  });

  it('fail() turns later operations into rejections with the given reason', async () => {
    const { sheet: s } = sheet();
    const reason = new Error('stream gone');
    s.fail(reason);
    await expect(s.writeRow([1])).rejects.toBe(reason);
    await expect(s.close()).rejects.toBe(reason);
    expect(() => s.merge('A1:B1')).toThrow(reason);
  });

  it('remembers a failed flush and rejects subsequent writes', async () => {
    const entry = new MemoryEntry();
    const boom = new Error('entry failed');
    entry.write = () => Promise.reject(boom);
    const s = new SheetWriter('S', {}, { openEntry: () => entry, ctx, first: true });
    await expect(s.close()).rejects.toBe(boom);
    await expect(s.writeRow([1])).rejects.toBe(boom);
  });

  it('lets the next sheet begin its ZIP entry as soon as close() is called', async () => {
    const sink = new MemorySink();
    const zip = new ZipWriter(sink);
    const a = new SheetWriter('A', {}, { openEntry: () => zip.beginEntry('a.xml'), ctx, first: true });
    const b = new SheetWriter('B', {}, { openEntry: () => zip.beginEntry('b.xml'), ctx, first: false });
    void a.writeRow([1]);
    void a.close(); // not awaited: the workbook closes the previous sheet implicitly
    await b.writeRow([2]);
    await b.close();
    await zip.finish();
    const files = unzipText(sink.bytes());
    expect(files['a.xml']).toContain('<c r="A1"><v>1</v></c>');
    expect(files['b.xml']).toContain('<c r="A1"><v>2</v></c>');
  });

  it('turns a failing openEntry into a rejection and remembers it', async () => {
    const boom = new Error('cannot begin entry');
    const openEntry = (): never => {
      throw boom;
    };
    const closing = new SheetWriter('S', {}, { openEntry, ctx, first: true });
    await expect(closing.close()).rejects.toBe(boom);
    await expect(closing.writeRow([1])).rejects.toBe(boom);

    // The threshold flush inside writeRow() triggers the failure: rejected promises, never a synchronous throw.
    const writing = new SheetWriter('S', {}, { openEntry, ctx, first: true });
    const wide = Array.from({ length: 50 }, (_, i) => `cell number ${i} with some padding text`);
    const promises: Promise<void>[] = [];
    for (let i = 0; i < 200; i++) promises.push(writing.writeRow(wide));
    await expect(Promise.all(promises)).rejects.toBe(boom);
    await expect(writing.close()).rejects.toBe(boom);
  });
});
