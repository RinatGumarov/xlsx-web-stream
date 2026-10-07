import { MAX_COLS, type CellContext } from './cell.js';
import { contentTypesXml, rootRelsXml, workbookRelsXml, workbookXml } from './parts.js';
import { SheetWriter } from './sheet.js';
import { StreamSink } from './sink.js';
import { StyleRegistry } from './styles.js';
import type { Sheet, SheetOptions, Style, StyleDefinition, Workbook, WorkbookOptions } from './types.js';
import { ZipWriter } from './zip/zip-writer.js';

const ENCODER = new TextEncoder();
const MAX_SHEET_NAME_LENGTH = 31;
const FORBIDDEN_SHEET_NAME_CHARS = /[[\]:*?/\\]/;
// An attribute value normalizes TAB/LF/CR to a space while element text keeps them, so the sheet name in
// workbook.xml and the one in the _FilterDatabase defined name would diverge. Reject all control characters.
// eslint-disable-next-line no-control-regex
const CONTROL_CHARS = /[\u0000-\u001F]/;
const noop = (): void => {};

function validateSheetName(name: unknown, taken: ReadonlySet<string>): asserts name is string {
  if (typeof name !== 'string' || name.length === 0) throw new TypeError('Sheet name must be a non-empty string');
  if (name.length > MAX_SHEET_NAME_LENGTH) {
    throw new TypeError(`Sheet name ${JSON.stringify(name)} is longer than ${MAX_SHEET_NAME_LENGTH} characters`);
  }
  if (FORBIDDEN_SHEET_NAME_CHARS.test(name)) {
    throw new TypeError(`Sheet name ${JSON.stringify(name)} contains a forbidden character ([ ] : * ? / \\)`);
  }
  if (CONTROL_CHARS.test(name)) {
    throw new TypeError(`Sheet name ${JSON.stringify(name)} contains a control character`);
  }
  if (name.startsWith("'") || name.endsWith("'")) {
    throw new TypeError(`Sheet name ${JSON.stringify(name)} must not start or end with an apostrophe`);
  }
  if (taken.has(name.toLowerCase())) {
    throw new Error(`Sheet name ${JSON.stringify(name)} is already used (sheet names are case-insensitive)`);
  }
}

export class WorkbookImpl implements Workbook {
  readonly readable: ReadableStream<Uint8Array>;
  private readonly sink: StreamSink;
  private readonly zip: ZipWriter;
  private readonly styles: StyleRegistry;
  private readonly ctx: CellContext;
  private readonly sheets: SheetWriter[] = [];
  private readonly names = new Set<string>();
  private current: SheetWriter | null = null;
  private closePromise: Promise<void> | null = null;
  private closed = false;

  constructor(options: WorkbookOptions = {}) {
    if (options === null || typeof options !== 'object') throw new TypeError('Workbook options must be an object');
    const dates = options.dates ?? 'utc';
    if (dates !== 'utc' && dates !== 'local') throw new TypeError(`dates must be "utc" or "local", got ${JSON.stringify(dates)}`);
    const dateFormat = options.dateFormat ?? 'yyyy-mm-dd';
    if (typeof dateFormat !== 'string' || dateFormat === '') throw new TypeError('dateFormat must be a non-empty string');

    this.sink = new StreamSink();
    this.sink.onFailure = (reason) => this.fail(reason);
    this.readable = this.sink.readable;
    this.zip = new ZipWriter(this.sink);
    this.styles = new StyleRegistry(this, dateFormat);
    this.ctx = {
      dates,
      styleId: (style) => this.styles.idOf(style),
      dateStyle: (id) => this.styles.dateVariant(id),
      linkStyle: () => this.styles.linkStyle(),
    };
  }

  style(definition: StyleDefinition): Style {
    this.assertOpen('style');
    return this.styles.register(definition);
  }

  addSheet(name: string, options: SheetOptions = {}): Sheet {
    this.assertOpen('addSheet');
    validateSheetName(name, this.names);
    if (options === null || typeof options !== 'object') throw new TypeError('Sheet options must be an object');
    if (options.columns !== undefined && Array.isArray(options.columns) && options.columns.length > MAX_COLS) {
      throw new RangeError(`columns has ${options.columns.length} entries; Excel supports at most ${MAX_COLS}`);
    }
    // Close the previous sheet implicitly; its failure, if any, resurfaces through the ZIP chain.
    if (this.current) this.current.close().catch(noop);
    const index = this.sheets.length;
    // The entry is begun lazily on the sheet's first flush, so option validation can fail without touching the archive.
    const sheet = new SheetWriter(name, options, {
      openEntry: () => this.zip.beginEntry(`xl/worksheets/sheet${index + 1}.xml`),
      ctx: this.ctx,
      first: index === 0,
    });
    this.sheets.push(sheet);
    this.names.add(name.toLowerCase());
    this.current = sheet;
    return sheet;
  }

  close(): Promise<void> {
    if (this.closePromise) return this.closePromise;
    this.closed = true;
    if (this.sheets.length === 0) {
      const error = new Error('Workbook must contain at least one sheet');
      this.fail(error);
      this.sink.error(error);
      return (this.closePromise = Promise.reject(error));
    }
    this.closePromise = this.finish().catch((reason: unknown) => {
      this.fail(reason);
      this.sink.error(reason);
      throw reason;
    });
    return this.closePromise;
  }

  abort(reason: unknown = new Error('Workbook was aborted')): void {
    this.closed = true;
    this.fail(reason);
    this.sink.error(reason);
  }

  private async finish(): Promise<void> {
    if (this.current) await this.current.close();
    await this.writePart('xl/styles.xml', this.styles.toXml());
    await this.writePart('xl/workbook.xml', workbookXml(this.sheets.map((s) => ({ name: s.name, autoFilter: s.autoFilterRange }))));
    await this.writePart('xl/_rels/workbook.xml.rels', workbookRelsXml(this.sheets.length));
    await this.writePart('_rels/.rels', rootRelsXml());
    await this.writePart('[Content_Types].xml', contentTypesXml(this.sheets.length));
    await this.zip.finish();
    this.sink.close();
  }

  private async writePart(name: string, xml: string): Promise<void> {
    const entry = this.zip.beginEntry(name);
    await entry.write(ENCODER.encode(xml));
    await entry.close();
  }

  private fail(reason: unknown): void {
    this.zip.fail(reason);
    for (const sheet of this.sheets) sheet.fail(reason);
  }

  private assertOpen(method: string): void {
    if (this.closed) throw new Error(`Cannot call ${method}() after the workbook was closed`);
  }
}
