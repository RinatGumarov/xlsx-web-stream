import { describe, expect, it } from 'vitest';
import { normalizeColor, StyleHandle, StyleRegistry } from '../src/styles.js';
import { parseXml } from './helpers.js';

const owner = {};
const registry = (): StyleRegistry => new StyleRegistry(owner, 'yyyy-mm-dd');

describe('normalizeColor', () => {
  it('accepts #RRGGBB and RRGGBB in any case and returns ARGB', () => {
    expect(normalizeColor('#eeeeee', 'fill')).toBe('FFEEEEEE');
    expect(normalizeColor('0563C1', 'fill')).toBe('FF0563C1');
  });

  it('rejects anything else', () => {
    for (const bad of ['#fff', 'red', '#GGGGGG', '', 123, null]) {
      expect(() => normalizeColor(bad, 'fill'), String(bad)).toThrow(TypeError);
    }
  });
});

describe('StyleRegistry.register', () => {
  it('returns handles owned by the registry owner and deduplicates equal definitions', () => {
    const r = registry();
    const a = r.register({ font: { bold: true }, fill: '#EEEEEE' });
    const b = r.register({ fill: '#eeeeee', font: { bold: true } });
    const c = r.register({ font: { bold: true } });
    expect(a).toBeInstanceOf(StyleHandle);
    expect(a.owner).toBe(owner);
    expect(a.id).toBe(b.id);
    expect(a.id).not.toBe(c.id);
    expect(a.id).toBeGreaterThan(0);
  });

  it('maps an empty definition and default-only fonts to xf 0', () => {
    const r = registry();
    expect(r.register({}).id).toBe(0);
    expect(r.register({ font: {} }).id).toBe(0);
    expect(r.register({ numFmt: 'General' }).id).toBe(0);
  });

  it('rejects unknown keys and invalid values', () => {
    const r = registry();
    expect(() => r.register({ numFormat: '0' } as never)).toThrow(TypeError);
    expect(() => r.register({ font: { weight: 'bold' } } as never)).toThrow(TypeError);
    expect(() => r.register({ font: { bold: 'yes' } } as never)).toThrow(TypeError);
    expect(() => r.register({ font: { size: -1 } })).toThrow(TypeError);
    expect(() => r.register({ fill: 'red' })).toThrow(TypeError);
    expect(() => r.register({ border: 'bold' as never })).toThrow(TypeError);
    expect(() => r.register({ border: { top: { style: 'thin', color: 'nope' } } })).toThrow(TypeError);
    expect(() => r.register({ border: { diagonal: 'thin' } } as never)).toThrow(TypeError);
    expect(() => r.register({ align: { horizontal: 'middle' as never } })).toThrow(TypeError);
    expect(() => r.register({ align: { vertical: 'up' as never } })).toThrow(TypeError);
    expect(() => r.register({ numFmt: '' })).toThrow(TypeError);
    expect(() => r.register(null as never)).toThrow(TypeError);
  });
});

describe('StyleRegistry.idOf', () => {
  it('resolves own handles, passes undefined through and rejects foreign handles', () => {
    const r = registry();
    const h = r.register({ font: { italic: true } });
    expect(r.idOf(h)).toBe(h.id);
    expect(r.idOf(undefined)).toBeUndefined();
    expect(() => r.idOf({ id: 1 })).toThrow(TypeError);
    expect(() => r.idOf(new StyleHandle(1, {}))).toThrow(TypeError);
  });
});

describe('date variants and link style', () => {
  it('creates a date variant once per style and keeps styles that already have a numFmt', () => {
    const r = registry();
    const bold = r.register({ font: { bold: true } });
    const money = r.register({ numFmt: '#,##0.00' });
    const dateDefault = r.dateVariant(0);
    const dateBold = r.dateVariant(bold.id);
    expect(dateDefault).not.toBe(0);
    expect(dateBold).not.toBe(bold.id);
    expect(dateBold).not.toBe(dateDefault);
    expect(r.dateVariant(0)).toBe(dateDefault);
    expect(r.dateVariant(money.id)).toBe(money.id);
    expect(r.register({ numFmt: 'yyyy-mm-dd' }).id).toBe(dateDefault);
  });

  it('registers the hyperlink style lazily and only once', () => {
    const r = registry();
    const id = r.linkStyle();
    expect(id).toBeGreaterThan(0);
    expect(r.linkStyle()).toBe(id);
    expect(r.register({ font: { color: '#0563C1', underline: true } }).id).toBe(id);
  });
});

describe('StyleRegistry.toXml', () => {
  it('emits the mandatory default fills, fonts and borders first', () => {
    const xml = registry().toXml();
    const doc = parseXml(xml);
    expect(xml.startsWith('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')).toBe(true);
    expect(doc.styleSheet.fills.count).toBe('2');
    expect(doc.styleSheet.fills.fill[0].patternFill.patternType).toBe('none');
    expect(doc.styleSheet.fills.fill[1].patternFill.patternType).toBe('gray125');
    expect(doc.styleSheet.fonts.count).toBe('1');
    expect(doc.styleSheet.borders.count).toBe('1');
    expect(doc.styleSheet.cellXfs.count).toBe('1');
    expect(doc.styleSheet.cellStyles.cellStyle.name).toBe('Normal');
    expect(xml).not.toContain('<numFmts');
  });

  it('serializes registered components in the order Excel expects', () => {
    const r = registry();
    const h = r.register({
      font: { bold: true, italic: true, underline: true, strike: true, size: 12, color: '#FF0000', name: 'Arial' },
      fill: '#EEEEEE',
      border: { top: 'thin', bottom: { style: 'double', color: '#00FF00' } },
      align: { horizontal: 'center', vertical: 'top', wrap: true },
      numFmt: '#,##0.00',
    });
    const xml = r.toXml();
    expect(xml).toContain('<numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0.00"/></numFmts>');
    expect(xml).toContain('<font><b/><i/><strike/><u/><sz val="12"/><color rgb="FFFF0000"/><name val="Arial"/><family val="2"/></font>');
    expect(xml).toContain('<fill><patternFill patternType="solid"><fgColor rgb="FFEEEEEE"/><bgColor indexed="64"/></patternFill></fill>');
    expect(xml).toContain('<border><left/><right/><top style="thin"/><bottom style="double"><color rgb="FF00FF00"/></bottom><diagonal/></border>');
    expect(xml).toContain(
      `<xf numFmtId="164" fontId="1" fillId="2" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="top" wrapText="1"/></xf>`,
    );
    expect(h.id).toBe(1);
  });

  it('applies a border style string to all four sides and escapes format codes', () => {
    const r = registry();
    r.register({ border: 'medium', numFmt: '"<"0' });
    const xml = r.toXml();
    expect(xml).toContain('<border><left style="medium"/><right style="medium"/><top style="medium"/><bottom style="medium"/><diagonal/></border>');
    expect(xml).toContain('formatCode="&quot;&lt;&quot;0"');
  });
});
