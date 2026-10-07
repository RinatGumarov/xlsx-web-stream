import type {
  AlignmentDefinition,
  BorderDefinition,
  BorderSideSpec,
  FontDefinition,
  Style,
  StyleDefinition,
} from './types.js';
import { NS_MAIN, XML_DECL, xmlAttr } from './xml.js';

const BORDER_STYLES: ReadonlySet<string> = new Set(['thin', 'medium', 'thick', 'dashed', 'dotted', 'double', 'hair']);
const HORIZONTAL: ReadonlySet<string> = new Set(['left', 'center', 'right', 'fill', 'justify', 'centerContinuous', 'distributed']);
const VERTICAL: ReadonlySet<string> = new Set(['top', 'center', 'bottom', 'justify', 'distributed']);
const STYLE_KEYS: ReadonlySet<string> = new Set(['font', 'fill', 'border', 'align', 'numFmt']);
const FONT_KEYS: ReadonlySet<string> = new Set(['bold', 'italic', 'underline', 'strike', 'size', 'color', 'name']);
const BORDER_KEYS: ReadonlySet<string> = new Set(['top', 'right', 'bottom', 'left']);
const ALIGN_KEYS: ReadonlySet<string> = new Set(['horizontal', 'vertical', 'wrap']);
const SIDES = ['left', 'right', 'top', 'bottom'] as const;

const FIRST_CUSTOM_NUMFMT = 164;
// No theme colors or font scheme: the package has no xl/theme/theme1.xml, so a theme reference could resolve to anything.
const DEFAULT_FONT = '<font><sz val="11"/><name val="Calibri"/><family val="2"/></font>';
const DEFAULT_FILLS = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
const DEFAULT_BORDER = '<border><left/><right/><top/><bottom/><diagonal/></border>';
const LINK_COLOR = '#0563C1';

interface Xf {
  numFmtId: number;
  fontId: number;
  fillId: number;
  borderId: number;
  alignment: string;
}

const DEFAULT_XF: Xf = { numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, alignment: '' };

/** Converts '#RRGGBB' / 'RRGGBB' to the ARGB form used in styles.xml. */
export function normalizeColor(color: unknown, what: string): string {
  if (typeof color !== 'string' || !/^#?[0-9a-fA-F]{6}$/.test(color)) {
    throw new TypeError(`Invalid ${what} ${JSON.stringify(color)}: expected "#RRGGBB"`);
  }
  return `FF${(color.startsWith('#') ? color.slice(1) : color).toUpperCase()}`;
}

/** Opaque handle returned by `workbook.style()`. */
export class StyleHandle implements Style {
  constructor(
    readonly id: number,
    readonly owner: object,
  ) {}
}

function assertObject(value: unknown, what: string): asserts value is object {
  if (value === null || typeof value !== 'object' || Array.isArray(value)) throw new TypeError(`${what} must be an object`);
}

function assertKeys(value: object, allowed: ReadonlySet<string>, what: string): void {
  for (const key of Object.keys(value)) {
    if (!allowed.has(key)) throw new TypeError(`Unknown ${what} property "${key}"`);
  }
}

function assertOptionalBoolean(value: unknown, what: string): void {
  if (value !== undefined && typeof value !== 'boolean') throw new TypeError(`${what} must be a boolean`);
}

function fontXml(font: FontDefinition): string {
  assertObject(font, 'font');
  assertKeys(font, FONT_KEYS, 'font');
  for (const key of ['bold', 'italic', 'underline', 'strike'] as const) assertOptionalBoolean(font[key], `font.${key}`);
  if (font.size !== undefined && !(typeof font.size === 'number' && font.size > 0 && font.size <= 409)) {
    throw new TypeError('font.size must be a number between 0 and 409');
  }
  if (font.name !== undefined && (typeof font.name !== 'string' || font.name === '')) {
    throw new TypeError('font.name must be a non-empty string');
  }
  let xml = '<font>';
  if (font.bold) xml += '<b/>';
  if (font.italic) xml += '<i/>';
  if (font.strike) xml += '<strike/>';
  if (font.underline) xml += '<u/>';
  xml += `<sz val="${font.size ?? 11}"/>`;
  if (font.color !== undefined) xml += `<color rgb="${normalizeColor(font.color, 'font.color')}"/>`;
  xml += `<name val="${xmlAttr(font.name ?? 'Calibri')}"/><family val="2"/>`;
  return `${xml}</font>`;
}

function fillXml(color: string): string {
  return `<fill><patternFill patternType="solid"><fgColor rgb="${normalizeColor(color, 'fill')}"/><bgColor indexed="64"/></patternFill></fill>`;
}

function borderXml(border: BorderDefinition | string): string {
  let sides: BorderDefinition;
  if (typeof border === 'string') {
    sides = { left: border as BorderSideSpec, right: border as BorderSideSpec, top: border as BorderSideSpec, bottom: border as BorderSideSpec };
  } else {
    assertObject(border, 'border');
    assertKeys(border, BORDER_KEYS, 'border');
    sides = border;
  }
  let xml = '<border>';
  for (const side of SIDES) {
    const spec = sides[side];
    if (spec === undefined) {
      xml += `<${side}/>`;
      continue;
    }
    let style: unknown;
    let color: unknown;
    if (typeof spec === 'string') {
      style = spec;
    } else {
      assertObject(spec, `border.${side}`);
      style = spec.style;
      color = spec.color;
    }
    if (typeof style !== 'string' || !BORDER_STYLES.has(style)) {
      throw new TypeError(`Invalid border style ${JSON.stringify(style)} for border.${side}`);
    }
    xml +=
      color === undefined
        ? `<${side} style="${style}"/>`
        : `<${side} style="${style}"><color rgb="${normalizeColor(color, `border.${side}.color`)}"/></${side}>`;
  }
  return `${xml}<diagonal/></border>`;
}

function alignmentXml(align: AlignmentDefinition): string {
  assertObject(align, 'align');
  assertKeys(align, ALIGN_KEYS, 'align');
  assertOptionalBoolean(align.wrap, 'align.wrap');
  if (align.horizontal !== undefined && !HORIZONTAL.has(align.horizontal)) {
    throw new TypeError(`Invalid align.horizontal ${JSON.stringify(align.horizontal)}`);
  }
  if (align.vertical !== undefined && !VERTICAL.has(align.vertical)) {
    throw new TypeError(`Invalid align.vertical ${JSON.stringify(align.vertical)}`);
  }
  let attrs = '';
  if (align.horizontal) attrs += ` horizontal="${align.horizontal}"`;
  if (align.vertical) attrs += ` vertical="${align.vertical}"`;
  if (align.wrap) attrs += ' wrapText="1"';
  return attrs === '' ? '' : `<alignment${attrs}/>`;
}

function xfXml(xf: Xf): string {
  const apply =
    `${xf.numFmtId ? ' applyNumberFormat="1"' : ''}${xf.fontId ? ' applyFont="1"' : ''}` +
    `${xf.fillId ? ' applyFill="1"' : ''}${xf.borderId ? ' applyBorder="1"' : ''}${xf.alignment ? ' applyAlignment="1"' : ''}`;
  const open = `<xf numFmtId="${xf.numFmtId}" fontId="${xf.fontId}" fillId="${xf.fillId}" borderId="${xf.borderId}" xfId="0"${apply}`;
  return xf.alignment ? `${open}>${xf.alignment}</xf>` : `${open}/>`;
}

function xfKey(xf: Xf): string {
  return `${xf.numFmtId}/${xf.fontId}/${xf.fillId}/${xf.borderId}/${xf.alignment}`;
}

/** Returns the index of `xml` in `list`, appending it when new. */
function intern(list: string[], ids: Map<string, number>, xml: string): number {
  let id = ids.get(xml);
  if (id === undefined) {
    id = list.length;
    list.push(xml);
    ids.set(xml, id);
  }
  return id;
}

/** Deduplicating registry of cell formats; produces styles.xml. */
export class StyleRegistry {
  private readonly numFmts: string[] = [];
  private readonly numFmtIds = new Map<string, number>();
  private readonly fonts = [DEFAULT_FONT];
  private readonly fontIds = new Map([[DEFAULT_FONT, 0]]);
  private readonly fills = [...DEFAULT_FILLS];
  private readonly fillIds = new Map(DEFAULT_FILLS.map((xml, i) => [xml, i] as const));
  private readonly borders = [DEFAULT_BORDER];
  private readonly borderIds = new Map([[DEFAULT_BORDER, 0]]);
  private readonly xfs: Xf[] = [DEFAULT_XF];
  private readonly xfIds = new Map([[xfKey(DEFAULT_XF), 0]]);
  private readonly dateVariants = new Map<number, number>();
  private linkId: number | undefined;

  constructor(
    private readonly owner: object,
    private readonly dateFormat: string,
  ) {}

  register(definition: StyleDefinition): StyleHandle {
    assertObject(definition, 'Style definition');
    assertKeys(definition, STYLE_KEYS, 'style');
    const xf: Xf = {
      numFmtId: this.numFmtId(definition.numFmt),
      fontId: definition.font === undefined ? 0 : intern(this.fonts, this.fontIds, fontXml(definition.font)),
      fillId: definition.fill === undefined ? 0 : intern(this.fills, this.fillIds, fillXml(definition.fill)),
      borderId: definition.border === undefined ? 0 : intern(this.borders, this.borderIds, borderXml(definition.border)),
      alignment: definition.align === undefined ? '' : alignmentXml(definition.align),
    };
    return new StyleHandle(this.xfId(xf), this.owner);
  }

  /** Resolves a handle to its xf index. Throws for handles made by another workbook. */
  idOf(style: Style | undefined): number | undefined {
    if (style === undefined) return undefined;
    if (!(style instanceof StyleHandle) || style.owner !== this.owner) {
      throw new TypeError('Style must be created with style() of the same workbook');
    }
    return style.id;
  }

  /** The same style with the default date format applied, for Date cells whose style has no numFmt. */
  dateVariant(xfId: number): number {
    const xf = this.xfs[xfId]!;
    if (xf.numFmtId !== 0) return xfId;
    let id = this.dateVariants.get(xfId);
    if (id === undefined) {
      id = this.xfId({ ...xf, numFmtId: this.numFmtId(this.dateFormat) });
      this.dateVariants.set(xfId, id);
    }
    return id;
  }

  /** Blue underlined text, applied to hyperlink cells without an explicit style. */
  linkStyle(): number {
    this.linkId ??= this.register({ font: { color: LINK_COLOR, underline: true } }).id;
    return this.linkId;
  }

  toXml(): string {
    const numFmts =
      this.numFmts.length === 0
        ? ''
        : `<numFmts count="${this.numFmts.length}">${this.numFmts
            .map((code, i) => `<numFmt numFmtId="${FIRST_CUSTOM_NUMFMT + i}" formatCode="${xmlAttr(code)}"/>`)
            .join('')}</numFmts>`;
    return (
      `${XML_DECL}<styleSheet xmlns="${NS_MAIN}">${numFmts}` +
      `<fonts count="${this.fonts.length}">${this.fonts.join('')}</fonts>` +
      `<fills count="${this.fills.length}">${this.fills.join('')}</fills>` +
      `<borders count="${this.borders.length}">${this.borders.join('')}</borders>` +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      `<cellXfs count="${this.xfs.length}">${this.xfs.map(xfXml).join('')}</cellXfs>` +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
      '</styleSheet>'
    );
  }

  private numFmtId(code: string | undefined): number {
    if (code === undefined || code === 'General') return 0;
    if (typeof code !== 'string' || code === '') throw new TypeError('numFmt must be a non-empty string');
    let id = this.numFmtIds.get(code);
    if (id === undefined) {
      id = FIRST_CUSTOM_NUMFMT + this.numFmts.length;
      this.numFmts.push(code);
      this.numFmtIds.set(code, id);
    }
    return id;
  }

  private xfId(xf: Xf): number {
    const key = xfKey(xf);
    let id = this.xfIds.get(key);
    if (id === undefined) {
      id = this.xfs.length;
      this.xfs.push(xf);
      this.xfIds.set(key, id);
    }
    return id;
  }
}
