import { describe, expect, it } from 'vitest';
import {
  cellRef,
  cellXml,
  columnLetter,
  dateToSerial,
  formatAbsoluteRange,
  formatRange,
  parseCellRef,
  parseRange,
  type CellContext,
} from '../src/cell.js';
import type { Style } from '../src/types.js';

const style = (id: number): Style => ({ id });

/** Fake context: handles resolve to their id, the date variant of style 0 is 99, hyperlink style is 7. */
const ctx: CellContext = {
  dates: 'utc',
  styleId: (s) => s?.id,
  dateStyle: (s) => (s === 0 ? 99 : s),
  linkStyle: () => 7,
};

describe('column letters and references', () => {
  it('converts 1-based column indexes to letters', () => {
    expect(columnLetter(1)).toBe('A');
    expect(columnLetter(26)).toBe('Z');
    expect(columnLetter(27)).toBe('AA');
    expect(columnLetter(52)).toBe('AZ');
    expect(columnLetter(53)).toBe('BA');
    expect(columnLetter(702)).toBe('ZZ');
    expect(columnLetter(703)).toBe('AAA');
    expect(columnLetter(16384)).toBe('XFD');
  });

  it('builds and parses A1 references', () => {
    expect(cellRef(3, 12)).toBe('C12');
    expect(parseCellRef('C12')).toEqual({ col: 3, row: 12 });
    expect(parseCellRef('xfd1048576')).toEqual({ col: 16384, row: 1048576 });
    expect(parseCellRef('XFE1')).toBeNull();
    expect(parseCellRef('A0')).toBeNull();
    expect(parseCellRef('A1048577')).toBeNull();
    expect(parseCellRef('1A')).toBeNull();
    expect(parseCellRef('')).toBeNull();
  });
});

describe('ranges', () => {
  it('normalizes corner order', () => {
    expect(parseRange('C3:A1')).toEqual({ c1: 1, r1: 1, c2: 3, r2: 3 });
    expect(formatRange(parseRange('C3:A1'))).toBe('A1:C3');
    expect(formatAbsoluteRange(parseRange('A1:C3'))).toBe('$A$1:$C$3');
  });

  it('rejects malformed ranges', () => {
    for (const bad of ['A1', 'A1:B', 'A1:B2:C3', 'A1-B2', '', 'XFE1:XFE2']) {
      expect(() => parseRange(bad), bad).toThrow(TypeError);
    }
    expect(() => parseRange(42 as unknown as string)).toThrow(TypeError);
  });
});

describe('dateToSerial', () => {
  it('matches Excel serial numbers including the 1900 leap-year bug', () => {
    expect(dateToSerial(new Date(Date.UTC(1900, 0, 1)), 'utc')).toBe(1);
    expect(dateToSerial(new Date(Date.UTC(1900, 1, 28)), 'utc')).toBe(59);
    expect(dateToSerial(new Date(Date.UTC(1900, 2, 1)), 'utc')).toBe(61);
    expect(dateToSerial(new Date(Date.UTC(2000, 0, 1)), 'utc')).toBe(36526);
    expect(dateToSerial(new Date(Date.UTC(2026, 9, 6)), 'utc')).toBe(46301);
    expect(dateToSerial(new Date(Date.UTC(2026, 9, 6, 12)), 'utc')).toBe(46301.5);
  });

  it('uses local components in local mode', () => {
    expect(dateToSerial(new Date(2026, 9, 6, 12), 'local')).toBe(46301.5);
    expect(dateToSerial(new Date(2026, 9, 6, 6), 'local')).toBe(46301.25);
  });

  it('returns null for dates Excel cannot represent', () => {
    expect(dateToSerial(new Date(NaN), 'utc')).toBeNull();
    expect(dateToSerial(new Date(Date.UTC(1899, 11, 31)), 'utc')).toBeNull();
    const ancient = new Date(2000, 0, 1);
    ancient.setFullYear(50);
    expect(dateToSerial(ancient, 'local')).toBeNull();
  });

  it('returns null for dates past Excel\'s 9999-12-31 ceiling or beyond the Date range', () => {
    expect(dateToSerial(new Date(8.64e15), 'local')).toBeNull();
    expect(dateToSerial(new Date(8.64e15), 'utc')).toBeNull();
    expect(dateToSerial(new Date(Date.UTC(10000, 0, 1)), 'utc')).toBeNull();
    expect(dateToSerial(new Date(Date.UTC(9999, 11, 31)), 'utc')).toBe(2958465);
  });
});

describe('cellXml', () => {
  it('writes strings inline, preserving significant whitespace', () => {
    expect(cellXml(1, 1, 'hi', 0, ctx)).toBe('<c r="A1" t="inlineStr"><is><t>hi</t></is></c>');
    expect(cellXml(2, 3, ' lead', 0, ctx)).toBe('<c r="B3" t="inlineStr"><is><t xml:space="preserve"> lead</t></is></c>');
    expect(cellXml(1, 1, 'a<b>&', 0, ctx)).toBe('<c r="A1" t="inlineStr"><is><t>a&lt;b&gt;&amp;</t></is></c>');
    expect(cellXml(1, 1, 'x\u0000y', 0, ctx)).toBe('<c r="A1" t="inlineStr"><is><t>xy</t></is></c>');
  });

  it('writes numbers, booleans and non-finite numbers', () => {
    expect(cellXml(1, 1, 12.5, 0, ctx)).toBe('<c r="A1"><v>12.5</v></c>');
    expect(cellXml(1, 1, -0, 0, ctx)).toBe('<c r="A1"><v>0</v></c>');
    expect(cellXml(1, 1, 1e21, 0, ctx)).toBe('<c r="A1"><v>1e+21</v></c>');
    expect(cellXml(1, 1, true, 0, ctx)).toBe('<c r="A1" t="b"><v>1</v></c>');
    expect(cellXml(1, 1, false, 0, ctx)).toBe('<c r="A1" t="b"><v>0</v></c>');
    expect(cellXml(1, 1, NaN, 0, ctx)).toBe('<c r="A1" t="e"><v>#NUM!</v></c>');
    expect(cellXml(1, 1, Infinity, 3, ctx)).toBe('<c r="A1" s="3" t="e"><v>#NUM!</v></c>');
  });

  it('writes dates with the date variant of the resolved style', () => {
    expect(cellXml(1, 1, new Date(Date.UTC(2026, 9, 6)), 0, ctx)).toBe('<c r="A1" s="99"><v>46301</v></c>');
    expect(cellXml(1, 1, new Date(Date.UTC(2026, 9, 6)), 5, ctx)).toBe('<c r="A1" s="5"><v>46301</v></c>');
    expect(cellXml(1, 1, new Date(NaN), 0, ctx)).toBe('<c r="A1" t="e"><v>#NUM!</v></c>');
    expect(cellXml(1, 1, new Date(8.64e15), 0, ctx)).toBe('<c r="A1" t="e"><v>#NUM!</v></c>');
  });

  it('omits empty cells unless they carry a style', () => {
    expect(cellXml(1, 1, null, 0, ctx)).toBe('');
    expect(cellXml(1, 1, undefined, 0, ctx)).toBe('');
    expect(cellXml(1, 1, '', 0, ctx)).toBe('');
    expect(cellXml(1, 1, null, 4, ctx)).toBe('<c r="A1" s="4"/>');
    expect(cellXml(1, 1, { value: null, style: style(6) }, 0, ctx)).toBe('<c r="A1" s="6"/>');
    expect(cellXml(1, 1, { style: style(6) }, 0, ctx)).toBe('<c r="A1" s="6"/>');
  });

  it('lets a cell style override the inherited one', () => {
    expect(cellXml(1, 1, { value: 1, style: style(2) }, 9, ctx)).toBe('<c r="A1" s="2"><v>1</v></c>');
    expect(cellXml(1, 1, { value: 'x' }, 9, ctx)).toBe('<c r="A1" s="9" t="inlineStr"><is><t>x</t></is></c>');
  });

  it('writes formulas with optional cached values', () => {
    expect(cellXml(1, 1, { formula: '=SUM(A1:A2)' }, 0, ctx)).toBe('<c r="A1"><f>SUM(A1:A2)</f></c>');
    expect(cellXml(1, 1, { formula: 'A1*2', value: 4 }, 0, ctx)).toBe('<c r="A1"><f>A1*2</f><v>4</v></c>');
    expect(cellXml(1, 1, { formula: 'A1&"x"', value: 'yx' }, 0, ctx)).toBe('<c r="A1" t="str"><f>A1&amp;"x"</f><v>yx</v></c>');
    expect(cellXml(1, 1, { formula: 'A1>0', value: true }, 0, ctx)).toBe('<c r="A1" t="b"><f>A1&gt;0</f><v>1</v></c>');
    expect(cellXml(1, 1, { formula: '1/0', value: NaN }, 0, ctx)).toBe('<c r="A1" t="e"><f>1/0</f><v>#NUM!</v></c>');
  });

  it('writes hyperlinks as HYPERLINK formulas with the built-in link style', () => {
    expect(cellXml(1, 1, { link: 'https://x.com/?a=1&b=2', value: 'Site' }, 0, ctx)).toBe(
      '<c r="A1" s="7" t="str"><f>HYPERLINK("https://x.com/?a=1&amp;b=2","Site")</f><v>Site</v></c>',
    );
    expect(cellXml(1, 1, { link: 'https://x.com' }, 0, ctx)).toBe(
      '<c r="A1" s="7" t="str"><f>HYPERLINK("https://x.com","https://x.com")</f><v>https://x.com</v></c>',
    );
    expect(cellXml(1, 1, { link: 'https://x.com', value: 'say "hi"' }, 0, ctx)).toContain('HYPERLINK("https://x.com","say ""hi""")');
    expect(cellXml(1, 1, { link: 'https://x.com', value: 'x', style: style(3) }, 0, ctx)).toContain(' s="3" ');
    expect(cellXml(1, 1, { link: 'https://x.com', value: 'x' }, 5, ctx)).toContain(' s="5" ');
  });

  it('accepts a HYPERLINK() link and label of exactly 255 characters', () => {
    const link = `https://x.com/${'a'.repeat(241)}`;
    expect(link).toHaveLength(255);
    const xml = cellXml(1, 1, { link, value: 'l'.repeat(255) }, 0, ctx);
    expect(xml).toContain(`HYPERLINK("${link}","${'l'.repeat(255)}")`);
  });

  it('rejects a HYPERLINK() link over 255 characters, which Excel cannot store in a formula', () => {
    const link = `https://x.com/${'a'.repeat(242)}`;
    expect(link).toHaveLength(256);
    expect(() => cellXml(1, 1, { link, value: 'Site' }, 0, ctx)).toThrow(RangeError);
    expect(() => cellXml(1, 1, { link, value: 'Site' }, 0, ctx)).toThrow(/Link in cell A1 is 256 characters long.*HYPERLINK\(\).*255.*plain string/);
    // Without an explicit label the link doubles as the label; the link check fires first.
    expect(() => cellXml(1, 1, { link }, 0, ctx)).toThrow(/Link in cell A1/);
  });

  it('rejects a HYPERLINK() label over 255 characters', () => {
    expect(() => cellXml(2, 3, { link: 'https://x.com', value: 'l'.repeat(256) }, 0, ctx)).toThrow(RangeError);
    expect(() => cellXml(2, 3, { link: 'https://x.com', value: 'l'.repeat(256) }, 0, ctx)).toThrow(
      /Link label in cell B3 is 256 characters long.*HYPERLINK\(\).*255.*Shorten the label/,
    );
  });

  it('accepts formulas of 8192 characters and rejects longer ones', () => {
    const exact = 'A'.repeat(8192);
    expect(cellXml(1, 1, { formula: exact }, 0, ctx)).toBe(`<c r="A1"><f>${exact}</f></c>`);
    // The leading '=' is not counted.
    expect(cellXml(1, 1, { formula: `=${exact}` }, 0, ctx)).toBe(`<c r="A1"><f>${exact}</f></c>`);
    expect(() => cellXml(1, 1, { formula: 'A'.repeat(8193) }, 0, ctx)).toThrow(RangeError);
    expect(() => cellXml(1, 1, { formula: `=${'A'.repeat(8193)}` }, 0, ctx)).toThrow(
      /Formula in cell A1 is 8193 characters long; Excel allows at most 8192/,
    );
  });

  it('rejects values it cannot represent', () => {
    expect(() => cellXml(1, 1, 'x'.repeat(32_768), 0, ctx)).toThrow(RangeError);
    expect(() => cellXml(1, 1, {} as never, 0, ctx)).toThrow(TypeError);
    expect(() => cellXml(1, 1, 10n as never, 0, ctx)).toThrow(TypeError);
    expect(() => cellXml(1, 1, { formula: '' } as never, 0, ctx)).toThrow(TypeError);
    expect(() => cellXml(1, 1, { link: '' } as never, 0, ctx)).toThrow(TypeError);
    expect(() => cellXml(1, 1, { formula: 'A1', value: new Date() } as never, 0, ctx)).toThrow(TypeError);
    expect(() => cellXml(1, 1, [1, 2] as never, 0, ctx)).toThrow(TypeError);
  });
});
