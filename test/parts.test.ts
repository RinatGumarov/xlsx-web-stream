import { describe, expect, it } from 'vitest';
import { parseRange } from '../src/cell.js';
import { contentTypesXml, rootRelsXml, workbookRelsXml, workbookXml } from '../src/parts.js';
import { parseXml } from './helpers.js';

describe('contentTypesXml', () => {
  it('lists defaults, the workbook, every sheet and the styles part', () => {
    const doc = parseXml(contentTypesXml(2));
    const overrides = doc.Types.Override.map((o: any) => o.PartName);
    expect(overrides).toEqual(['/xl/workbook.xml', '/xl/worksheets/sheet1.xml', '/xl/worksheets/sheet2.xml', '/xl/styles.xml']);
    expect(doc.Types.Default).toHaveLength(2);
    expect(doc.Types.xmlns).toBe('http://schemas.openxmlformats.org/package/2006/content-types');
  });
});

describe('relationships', () => {
  it('points the package at xl/workbook.xml', () => {
    const doc = parseXml(rootRelsXml());
    expect(doc.Relationships.Relationship[0]).toMatchObject({
      Id: 'rId1',
      Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
      Target: 'xl/workbook.xml',
    });
  });

  it('relates sheets by index and styles last', () => {
    const rels = parseXml(workbookRelsXml(2)).Relationships.Relationship;
    expect(rels.map((r: any) => [r.Id, r.Target])).toEqual([
      ['rId1', 'worksheets/sheet1.xml'],
      ['rId2', 'worksheets/sheet2.xml'],
      ['rId3', 'styles.xml'],
    ]);
    expect(rels[2].Type).toBe('http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles');
  });
});

describe('workbookXml', () => {
  it('declares sheets, auto-filter defined names and full recalculation on load', () => {
    const xml = workbookXml([
      { name: 'Data & "more"', autoFilter: parseRange('A1:C1') },
      { name: "O'Brien", autoFilter: undefined },
      { name: 'Third', autoFilter: parseRange('B2:D9') },
    ]);
    const doc = parseXml(xml);
    expect(doc.workbook.sheets.sheet.map((s: any) => [s.name, s.sheetId, s['r:id']])).toEqual([
      ['Data & "more"', '1', 'rId1'],
      ["O'Brien", '2', 'rId2'],
      ['Third', '3', 'rId3'],
    ]);
    expect(xml).toContain(`<definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">'Data &amp; "more"'!$A$1:$C$1</definedName>`);
    expect(xml).toContain(`<definedName name="_xlnm._FilterDatabase" localSheetId="2" hidden="1">'Third'!$B$2:$D$9</definedName>`);
    expect(doc.workbook.definedNames.definedName).toHaveLength(2);
    expect(doc.workbook.calcPr.fullCalcOnLoad).toBe('1');
    expect(xml.indexOf('<definedNames>')).toBeLessThan(xml.indexOf('<calcPr'));
  });

  it('doubles apostrophes in sheet names inside defined names and omits definedNames without filters', () => {
    const xml = workbookXml([{ name: "O'Brien", autoFilter: parseRange('A1:A1') }]);
    expect(xml).toContain(`'O''Brien'!$A$1:$A$1`);
    expect(workbookXml([{ name: 'Plain', autoFilter: undefined }])).not.toContain('definedNames');
  });
});
