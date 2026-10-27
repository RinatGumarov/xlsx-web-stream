/** Options accepted by `createWorkbook()`. */
export interface WorkbookOptions {
  /** Interpret `Date` values by their UTC components (default) or by local components. */
  dates?: 'utc' | 'local';
  /** Number format applied to `Date` cells whose style has no `numFmt`. Default `'yyyy-mm-dd'`. */
  dateFormat?: string;
}

/** `'#RRGGBB'` or `'RRGGBB'`, case-insensitive. */
export type Color = string;

export type BorderStyle = 'thin' | 'medium' | 'thick' | 'dashed' | 'dotted' | 'double' | 'hair';

export interface BorderSide {
  style: BorderStyle;
  color?: Color;
}

export type BorderSideSpec = BorderStyle | BorderSide;

export type HorizontalAlignment =
  | 'left'
  | 'center'
  | 'right'
  | 'fill'
  | 'justify'
  | 'centerContinuous'
  | 'distributed';

export type VerticalAlignment = 'top' | 'center' | 'bottom' | 'justify' | 'distributed';

export interface FontDefinition {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  /** Font size in points. Default 11. */
  size?: number;
  color?: Color;
  /** Font name. Default Calibri. */
  name?: string;
}

export interface BorderDefinition {
  top?: BorderSideSpec;
  right?: BorderSideSpec;
  bottom?: BorderSideSpec;
  left?: BorderSideSpec;
}

export interface AlignmentDefinition {
  horizontal?: HorizontalAlignment;
  vertical?: VerticalAlignment;
  wrap?: boolean;
}

export interface StyleDefinition {
  font?: FontDefinition;
  /** Solid fill color. */
  fill?: Color;
  /** One style for all four sides, or per-side settings. */
  border?: BorderStyle | BorderDefinition;
  align?: AlignmentDefinition;
  /** Excel number format code, e.g. `'#,##0.00'` or `'yyyy-mm-dd'`. */
  numFmt?: string;
}

/** Opaque handle returned by `workbook.style()`. Valid only within the workbook that created it. */
export interface Style {
  readonly id: number;
}

export type ScalarValue = string | number | boolean | Date | null | undefined;

export type CellValue =
  | ScalarValue
  | { value: ScalarValue; style?: Style }
  | { style: Style; value?: ScalarValue }
  | { formula: string; value?: string | number | boolean; style?: Style }
  | { link: string; value?: string; style?: Style };

export interface ColumnOptions {
  /** Width in Excel character units (0–255). */
  width?: number;
  /** Default style for cells in this column. */
  style?: Style;
}

export interface SheetOptions {
  columns?: ColumnOptions[];
  /** Freeze the first `rows` rows and/or `columns` columns. */
  freeze?: { rows?: number; columns?: number };
}

export interface RowOptions {
  /** Default style for cells in this row (overrides column styles). */
  style?: Style;
  /** Row height in points (0–409). */
  height?: number;
}

export interface Sheet {
  readonly name: string;
  /** 1-based number of the next row `writeRow()` will produce. */
  readonly nextRow: number;
  writeRow(cells: readonly CellValue[], options?: RowOptions): Promise<void>;
  writeRows(rows: Iterable<readonly CellValue[]> | AsyncIterable<readonly CellValue[]>): Promise<void>;
  /** Merge a range such as `'A1:C1'`. May be called any time before the sheet is closed. */
  merge(range: string): void;
  /** Add an auto filter to a range such as `'A1:C1'`. At most one per sheet. */
  autoFilter(range: string): void;
  close(): Promise<void>;
}

export interface Workbook {
  /** The XLSX bytes. Pipe it anywhere a Web ReadableStream is accepted. */
  readonly readable: ReadableStream<Uint8Array>;
  style(definition: StyleDefinition): Style;
  /** Starts a new sheet. The previous sheet, if still open, is closed implicitly. */
  addSheet(name: string, options?: SheetOptions): Sheet;
  /** Closes the current sheet, writes workbook metadata and ends the stream. Idempotent. */
  close(): Promise<void>;
  /** Fails the stream and every pending or future operation with `reason`. */
  abort(reason?: unknown): void;
}
