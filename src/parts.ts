import { formatAbsoluteRange, type Range } from './cell.js';
import { NS_MAIN, NS_REL, XML_DECL, xmlAttr, xmlText } from './xml.js';

const NS_CONTENT_TYPES = 'http://schemas.openxmlformats.org/package/2006/content-types';
const NS_PACKAGE_RELS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CT_RELS = 'application/vnd.openxmlformats-package.relationships+xml';
const CT_WORKBOOK = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml';
const CT_WORKSHEET = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml';
const CT_STYLES = 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml';
const REL_OFFICE_DOCUMENT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
const REL_WORKSHEET = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';
const REL_STYLES = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles';

export interface SheetInfo {
  name: string;
  autoFilter: Range | undefined;
}

export function contentTypesXml(sheetCount: number): string {
  let xml =
    `${XML_DECL}<Types xmlns="${NS_CONTENT_TYPES}">` +
    `<Default Extension="rels" ContentType="${CT_RELS}"/>` +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    `<Override PartName="/xl/workbook.xml" ContentType="${CT_WORKBOOK}"/>`;
  for (let i = 1; i <= sheetCount; i++) {
    xml += `<Override PartName="/xl/worksheets/sheet${i}.xml" ContentType="${CT_WORKSHEET}"/>`;
  }
  return `${xml}<Override PartName="/xl/styles.xml" ContentType="${CT_STYLES}"/></Types>`;
}

export function rootRelsXml(): string {
  return (
    `${XML_DECL}<Relationships xmlns="${NS_PACKAGE_RELS}">` +
    `<Relationship Id="rId1" Type="${REL_OFFICE_DOCUMENT}" Target="xl/workbook.xml"/>` +
    '</Relationships>'
  );
}

export function workbookRelsXml(sheetCount: number): string {
  let xml = `${XML_DECL}<Relationships xmlns="${NS_PACKAGE_RELS}">`;
  for (let i = 1; i <= sheetCount; i++) {
    xml += `<Relationship Id="rId${i}" Type="${REL_WORKSHEET}" Target="worksheets/sheet${i}.xml"/>`;
  }
  return `${xml}<Relationship Id="rId${sheetCount + 1}" Type="${REL_STYLES}" Target="styles.xml"/></Relationships>`;
}

/** Sheet names inside formulas are single-quoted with apostrophes doubled. */
function quoteSheetName(name: string): string {
  return `'${name.replace(/'/g, "''")}'`;
}

export function workbookXml(sheets: readonly SheetInfo[]): string {
  let xml = `${XML_DECL}<workbook xmlns="${NS_MAIN}" xmlns:r="${NS_REL}"><workbookPr/><bookViews><workbookView/></bookViews><sheets>`;
  sheets.forEach((sheet, i) => {
    xml += `<sheet name="${xmlAttr(sheet.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`;
  });
  xml += '</sheets>';
  let definedNames = '';
  sheets.forEach((sheet, i) => {
    if (!sheet.autoFilter) return;
    definedNames +=
      `<definedName name="_xlnm._FilterDatabase" localSheetId="${i}" hidden="1">` +
      `${xmlText(quoteSheetName(sheet.name))}!${formatAbsoluteRange(sheet.autoFilter)}</definedName>`;
  });
  if (definedNames) xml += `<definedNames>${definedNames}</definedNames>`;
  return `${xml}<calcPr fullCalcOnLoad="1"/></workbook>`;
}
