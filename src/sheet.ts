import { cellRef, cellXml, formatRange, MAX_COLS, MAX_ROWS, parseRange, type CellContext, type Range } from './cell.js';
import type { CellValue, RowOptions, Sheet, SheetOptions } from './types.js';
import { NS_MAIN, NS_REL, XML_DECL } from './xml.js';
import type { ZipEntry } from './zip/zip-writer.js';

/** Buffered XML (UTF-16 code units) that triggers a flush to the compressor. */
const FLUSH_THRESHOLD = 64 * 1024;
const DEFAULT_COLUMN_WIDTH = 8.43;
const MAX_COLUMN_WIDTH = 255;
const MAX_ROW_HEIGHT = 409;
const ENCODER = new TextEncoder();
const RESOLVED = Promise.resolve();
const noop = (): void => {};

export interface SheetDeps {
  /** Begins the sheet's ZIP entry. Called once, on the first flush, so that constructor validation can fail without touching the archive. */
  openEntry: () => ZipEntry;
  ctx: CellContext;
  /** Only the first sheet of a workbook is marked as selected. */
  first: boolean;
}

function assertCount(value: unknown, what: string, max: number): asserts value is number {
  if (typeof value !== 'number' || !Number.isInteger(value) || value < 0 || value > max) {
    throw new RangeError(`${what} must be an integer between 0 and ${max}`);
  }
}

/** Streams one worksheet's XML into a ZIP entry. */
export class SheetWriter implements Sheet {
  readonly name: string;
  private readonly openEntry: () => ZipEntry;
  private entry: ZipEntry | null = null;
  private readonly ctx: CellContext;
  private readonly columnStyles: number[] = [];
  private readonly merges: Range[] = [];
  private filter: Range | undefined;
  private buffer: string;
  private row = 1;
  private closePromise: Promise<void> | null = null;
  private failure: { reason: unknown } | null = null;

  constructor(name: string, options: SheetOptions, deps: SheetDeps) {
    this.name = name;
    this.openEntry = deps.openEntry;
    this.ctx = deps.ctx;
    this.buffer = this.header(options, deps.first);
  }

  get nextRow(): number {
    return this.row;
  }

  /** @internal The auto-filter range, needed by workbook.xml's defined names. */
  get autoFilterRange(): Range | undefined {
    return this.filter;
  }

  writeRow(cells: readonly CellValue[], options?: RowOptions): Promise<void> {
    if (this.failure) return Promise.reject(this.failure.reason);
    this.assertOpen('writeRow');
    if (!Array.isArray(cells)) throw new TypeError('writeRow() expects an array of cell values');
    const row = this.row;
    if (row > MAX_ROWS) throw new RangeError(`Sheet "${this.name}" already has ${MAX_ROWS} rows, the maximum Excel supports`);
    if (cells.length > MAX_COLS) throw new RangeError(`Row ${row} has ${cells.length} cells; Excel supports at most ${MAX_COLS} columns`);

    // `undefined` means "no row style"; an explicit style that resolves to xf 0 still overrides the column style.
    const rowId = this.ctx.styleId(options?.style);
    let xml = `<row r="${row}"`;
    if (rowId !== undefined && rowId !== 0) xml += ` s="${rowId}" customFormat="1"`;
    if (options?.height !== undefined) {
      const height = options.height;
      if (typeof height !== 'number' || !(height >= 0 && height <= MAX_ROW_HEIGHT)) {
        throw new RangeError(`Row height must be a number between 0 and ${MAX_ROW_HEIGHT} points`);
      }
      xml += ` ht="${height}" customHeight="1"`;
    }
    xml += '>';
    for (let i = 0; i < cells.length; i++) {
      const inherited = rowId ?? (this.columnStyles[i] ?? 0);
      xml += cellXml(i + 1, row, cells[i], inherited, this.ctx);
    }
    xml += '</row>';

    this.row = row + 1;
    this.buffer += xml;
    return this.buffer.length >= FLUSH_THRESHOLD ? this.flush() : RESOLVED;
  }

  async writeRows(rows: Iterable<readonly CellValue[]> | AsyncIterable<readonly CellValue[]>): Promise<void> {
    for await (const row of rows) await this.writeRow(row);
  }

  merge(range: string): void {
    this.assertUsable('merge');
    const r = parseRange(range);
    if (r.c1 === r.c2 && r.r1 === r.r2) throw new TypeError(`Cannot merge the single cell ${JSON.stringify(range)}`);
    // Excel repairs files whose merged ranges overlap.
    for (const other of this.merges) {
      if (r.c1 <= other.c2 && other.c1 <= r.c2 && r.r1 <= other.r2 && other.r1 <= r.r2) {
        throw new Error(`Merge range ${formatRange(r)} overlaps the existing merge ${formatRange(other)}`);
      }
    }
    this.merges.push(r);
  }

  autoFilter(range: string): void {
    this.assertUsable('autoFilter');
    if (this.filter) throw new Error(`Sheet "${this.name}" already has an auto filter`);
    this.filter = parseRange(range);
  }

  close(): Promise<void> {
    if (this.closePromise) return this.closePromise;
    if (this.failure) return (this.closePromise = Promise.reject(this.failure.reason));
    let tail = '</sheetData>';
    if (this.filter) tail += `<autoFilter ref="${formatRange(this.filter)}"/>`;
    if (this.merges.length > 0) {
      tail += `<mergeCells count="${this.merges.length}">${this.merges.map((m) => `<mergeCell ref="${formatRange(m)}"/>`).join('')}</mergeCells>`;
    }
    this.buffer += `${tail}</worksheet>`;
    const text = this.buffer;
    this.buffer = '';
    let entry: ZipEntry;
    try {
      entry = this.entry ??= this.openEntry();
    } catch (reason) {
      this.fail(reason);
      return (this.closePromise = Promise.reject(reason));
    }
    // Both calls are sequenced by the ZipWriter, and close() marks the entry as closing right away, so the
    // next sheet may begin its entry without awaiting this one. A write failure surfaces through the combined promise.
    const written = entry.write(ENCODER.encode(text));
    const closed = entry.close();
    this.closePromise = Promise.all([written, closed]).then(noop, (reason: unknown) => {
      this.fail(reason);
      throw reason;
    });
    return this.closePromise;
  }

  /** @internal Marks the sheet as failed; later operations reject/throw with `reason`. */
  fail(reason: unknown): void {
    if (!this.failure) this.failure = { reason };
  }

  private assertOpen(method: string): void {
    if (this.closePromise) throw new Error(`Cannot call ${method}() on sheet "${this.name}": it is already closed`);
  }

  private assertUsable(method: string): void {
    if (this.failure) throw this.failure.reason;
    this.assertOpen(method);
  }

  private flush(): Promise<void> {
    const text = this.buffer;
    this.buffer = '';
    let entry: ZipEntry;
    try {
      entry = this.entry ??= this.openEntry();
    } catch (reason) {
      this.fail(reason);
      return Promise.reject(reason);
    }
    return entry.write(ENCODER.encode(text)).catch((reason: unknown) => {
      this.fail(reason);
      throw reason;
    });
  }

  private header(options: SheetOptions, first: boolean): string {
    const { columns = [], freeze } = options;
    if (!Array.isArray(columns)) throw new TypeError('columns must be an array');
    if (columns.length > MAX_COLS) throw new RangeError(`columns has ${columns.length} entries; Excel supports at most ${MAX_COLS}`);

    let xml = `${XML_DECL}<worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_REL}"><sheetViews><sheetView workbookViewId="0"${first ? ' tabSelected="1"' : ''}>`;
    if (freeze !== undefined) {
      if (freeze === null || typeof freeze !== 'object') throw new TypeError('freeze must be an object like { rows: 1 }');
      const rows = freeze.rows ?? 0;
      const cols = freeze.columns ?? 0;
      assertCount(rows, 'freeze.rows', MAX_ROWS - 1);
      assertCount(cols, 'freeze.columns', MAX_COLS - 1);
      if (rows > 0 || cols > 0) {
        const pane = rows > 0 && cols > 0 ? 'bottomRight' : rows > 0 ? 'bottomLeft' : 'topRight';
        xml += `<pane${cols > 0 ? ` xSplit="${cols}"` : ''}${rows > 0 ? ` ySplit="${rows}"` : ''} topLeftCell="${cellRef(cols + 1, rows + 1)}" activePane="${pane}" state="frozen"/>`;
      }
    }
    xml += '</sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/>';

    let cols = '';
    columns.forEach((column, i) => {
      if (column === null || typeof column !== 'object') throw new TypeError(`columns[${i}] must be an object`);
      const styleId = this.ctx.styleId(column.style) ?? 0;
      this.columnStyles[i] = styleId;
      const { width } = column;
      if (width !== undefined && !(typeof width === 'number' && width >= 0 && width <= MAX_COLUMN_WIDTH)) {
        throw new RangeError(`columns[${i}].width must be a number between 0 and ${MAX_COLUMN_WIDTH}`);
      }
      if (width === undefined && styleId === 0) return;
      cols += `<col min="${i + 1}" max="${i + 1}" width="${width ?? DEFAULT_COLUMN_WIDTH}"${width !== undefined ? ' customWidth="1"' : ''}${styleId !== 0 ? ` style="${styleId}"` : ''}/>`;
    });
    if (cols) xml += `<cols>${cols}</cols>`;
    return `${xml}<sheetData>`;
  }
}
