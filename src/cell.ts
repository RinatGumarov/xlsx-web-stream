import type { CellValue, Style } from './types.js';
import { escapeText, sanitize } from './xml.js';

export const MAX_ROWS = 1_048_576;
export const MAX_COLS = 16_384;
export const MAX_TEXT_LENGTH = 32_767;

export type DateMode = 'utc' | 'local';

const letterCache: string[] = [];

/** 1-based column index → letters: 1 → "A", 27 → "AA", 16384 → "XFD". */
export function columnLetter(col: number): string {
  const cached = letterCache[col];
  if (cached !== undefined) return cached;
  let n = col;
  let out = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    out = String.fromCharCode(65 + r) + out;
    n = (n - 1 - r) / 26;
  }
  letterCache[col] = out;
  return out;
}

export function cellRef(col: number, row: number): string {
  return columnLetter(col) + row;
}

const CELL_REF = /^([A-Z]{1,3})([0-9]{1,7})$/;

export function parseCellRef(ref: string): { col: number; row: number } | null {
  const match = CELL_REF.exec(ref.toUpperCase());
  if (!match) return null;
  let col = 0;
  for (const ch of match[1]!) col = col * 26 + ch.charCodeAt(0) - 64;
  const row = Number(match[2]);
  if (col > MAX_COLS || row < 1 || row > MAX_ROWS) return null;
  return { col, row };
}

export interface Range {
  c1: number;
  r1: number;
  c2: number;
  r2: number;
}

/** Parses "A1:C3" (corners in any order) into a normalized range. */
export function parseRange(range: string): Range {
  const parts = typeof range === 'string' ? range.split(':') : [];
  const a = parts.length === 2 ? parseCellRef(parts[0]!) : null;
  const b = parts.length === 2 ? parseCellRef(parts[1]!) : null;
  if (!a || !b) throw new TypeError(`Invalid range ${JSON.stringify(range)}: expected the form "A1:C3"`);
  return {
    c1: Math.min(a.col, b.col),
    r1: Math.min(a.row, b.row),
    c2: Math.max(a.col, b.col),
    r2: Math.max(a.row, b.row),
  };
}

export function formatRange(r: Range): string {
  return `${cellRef(r.c1, r.r1)}:${cellRef(r.c2, r.r2)}`;
}

/** "$A$1:$C$3" — the form used in defined names. */
export function formatAbsoluteRange(r: Range): string {
  return `$${columnLetter(r.c1)}$${r.r1}:$${columnLetter(r.c2)}$${r.r2}`;
}

const EXCEL_EPOCH = Date.UTC(1899, 11, 30);
const MS_PER_DAY = 86_400_000;

/** Excel serial number of a date, or null when Excel cannot represent it (invalid, before 1900-01-01 or after 9999-12-31). */
export function dateToSerial(date: Date, mode: DateMode): number | null {
  const time = date.getTime();
  if (Number.isNaN(time)) return null;
  let utc = time;
  if (mode === 'local') {
    if (date.getFullYear() < 100) return null; // Date.UTC would map years 0–99 to 1900–1999
    utc = Date.UTC(
      date.getFullYear(),
      date.getMonth(),
      date.getDate(),
      date.getHours(),
      date.getMinutes(),
      date.getSeconds(),
      date.getMilliseconds(),
    );
  }
  let serial = (utc - EXCEL_EPOCH) / MS_PER_DAY;
  if (!(serial >= 2 && serial < 2958466)) return null; // before 1900-01-01, after 9999-12-31, or NaN
  if (serial < 61) serial -= 1; // Excel treats 1900 as a leap year; skip its phantom Feb 29
  return serial;
}

export interface CellContext {
  readonly dates: DateMode;
  /** Resolves a style handle to its xf index; throws if the handle belongs to another workbook. */
  styleId(style: Style | undefined): number | undefined;
  /** xf index to use for a Date value whose resolved style has no number format. */
  dateStyle(styleId: number): number;
  /** xf index of the built-in hyperlink style. */
  linkStyle(): number;
}

interface RichCell {
  value?: unknown;
  style?: Style;
  formula?: unknown;
  link?: unknown;
}

const ERROR_VALUE = '#NUM!';
const LEADING_OR_TRAILING_SPACE = /^\s|\s$/;

function sAttr(s: number): string {
  return s === 0 ? '' : ` s="${s}"`;
}

function errorCell(ref: string, s: number): string {
  return `<c r="${ref}"${sAttr(s)} t="e"><v>${ERROR_VALUE}</v></c>`;
}

function text(value: string, ref: string): string {
  if (value.length > MAX_TEXT_LENGTH) {
    throw new RangeError(`Text in cell ${ref} is ${value.length} characters long; Excel allows at most ${MAX_TEXT_LENGTH}`);
  }
  return escapeText(sanitize(value));
}

/** Quotes a string literal for use inside a formula: "say ""hi""". */
function formulaString(value: string): string {
  return `"${value.replace(/"/g, '""')}"`;
}

/**
 * Serializes one cell. `inherited` is the xf index coming from the row or column (0 = none).
 * Returns '' when the cell is empty and unstyled.
 */
export function cellXml(col: number, row: number, value: CellValue, inherited: number, ctx: CellContext): string {
  const ref = cellRef(col, row);
  let v: unknown = value;
  let s = inherited;
  let ownStyle = false;
  let formula: string | undefined;
  let link: string | undefined;

  if (v !== null && typeof v === 'object' && !(v instanceof Date)) {
    if (Array.isArray(v)) throw new TypeError(`Unsupported value in cell ${ref}: array`);
    const rich = v as RichCell;
    const id = ctx.styleId(rich.style);
    if (id !== undefined) {
      s = id;
      ownStyle = true;
    }
    if (rich.formula !== undefined) {
      if (typeof rich.formula !== 'string' || rich.formula.trim() === '') {
        throw new TypeError(`Formula in cell ${ref} must be a non-empty string`);
      }
      formula = rich.formula;
    } else if (rich.link !== undefined) {
      if (typeof rich.link !== 'string' || rich.link === '') {
        throw new TypeError(`Link in cell ${ref} must be a non-empty string`);
      }
      link = rich.link;
    } else if (!('value' in rich) && !('style' in rich)) {
      throw new TypeError(`Unsupported value in cell ${ref}: expected { value, style }, { formula } or { link }`);
    }
    v = rich.value;
  }

  if (link !== undefined) {
    const label = v === undefined || v === null ? link : String(v);
    if (!ownStyle && s === 0) s = ctx.linkStyle();
    const f = `HYPERLINK(${formulaString(link)},${formulaString(label)})`;
    return `<c r="${ref}"${sAttr(s)} t="str"><f>${text(f, ref)}</f><v>${text(label, ref)}</v></c>`;
  }

  if (formula !== undefined) {
    const f = `<f>${text(formula.startsWith('=') ? formula.slice(1) : formula, ref)}</f>`;
    if (v === undefined || v === null) return `<c r="${ref}"${sAttr(s)}>${f}</c>`;
    if (typeof v === 'number') {
      return Number.isFinite(v)
        ? `<c r="${ref}"${sAttr(s)}>${f}<v>${v}</v></c>`
        : `<c r="${ref}"${sAttr(s)} t="e">${f}<v>${ERROR_VALUE}</v></c>`;
    }
    if (typeof v === 'boolean') return `<c r="${ref}"${sAttr(s)} t="b">${f}<v>${v ? 1 : 0}</v></c>`;
    if (typeof v === 'string') return `<c r="${ref}"${sAttr(s)} t="str">${f}<v>${text(v, ref)}</v></c>`;
    throw new TypeError(`Cached formula value in cell ${ref} must be a string, number or boolean`);
  }

  switch (typeof v) {
    case 'string':
      if (v === '') break;
      return `<c r="${ref}"${sAttr(s)} t="inlineStr"><is><t${LEADING_OR_TRAILING_SPACE.test(v) ? ' xml:space="preserve"' : ''}>${text(v, ref)}</t></is></c>`;
    case 'number':
      return Number.isFinite(v) ? `<c r="${ref}"${sAttr(s)}><v>${v}</v></c>` : errorCell(ref, s);
    case 'boolean':
      return `<c r="${ref}"${sAttr(s)} t="b"><v>${v ? 1 : 0}</v></c>`;
    case 'undefined':
      break;
    case 'object':
      if (v === null) break;
      if (v instanceof Date) {
        const serial = dateToSerial(v, ctx.dates);
        return serial === null ? errorCell(ref, s) : `<c r="${ref}" s="${ctx.dateStyle(s)}"><v>${serial}</v></c>`;
      }
      throw new TypeError(`Unsupported value in cell ${ref}: ${Object.prototype.toString.call(v)}`);
    default:
      throw new TypeError(`Unsupported value in cell ${ref}: ${typeof v}`);
  }
  return s === 0 ? '' : `<c r="${ref}" s="${s}"/>`;
}
