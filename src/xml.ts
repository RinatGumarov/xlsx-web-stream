export const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
export const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
export const NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

// Characters outside the XML 1.0 Char production (except TAB, LF, CR). Excel refuses files that contain them.
const INVALID_CHARS = /[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/;
const INVALID_CHARS_GLOBAL = /[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g;
const TEXT_SPECIALS = /[&<>]/;
const TEXT_SPECIALS_GLOBAL = /[&<>]/g;
const ATTR_SPECIALS = /[&<>"']/;
const ATTR_SPECIALS_GLOBAL = /[&<>"']/g;

const REPLACEMENTS: Record<string, string> = {
  '&': '&amp;',
  '<': '&lt;',
  '>': '&gt;',
  '"': '&quot;',
  "'": '&apos;',
};

const replace = (ch: string): string => REPLACEMENTS[ch]!;

/** Removes characters that are not allowed in an XML 1.0 document. */
export function sanitize(text: string): string {
  return INVALID_CHARS.test(text) ? text.replace(INVALID_CHARS_GLOBAL, '') : text;
}

/** Escapes a string for use as element content. */
export function escapeText(text: string): string {
  return TEXT_SPECIALS.test(text) ? text.replace(TEXT_SPECIALS_GLOBAL, replace) : text;
}

/** Escapes a string for use inside a double- or single-quoted attribute value. */
export function escapeAttr(text: string): string {
  return ATTR_SPECIALS.test(text) ? text.replace(ATTR_SPECIALS_GLOBAL, replace) : text;
}

/** Sanitizes and escapes user-provided element content. */
export function xmlText(text: string): string {
  return escapeText(sanitize(text));
}

/** Sanitizes and escapes a user-provided attribute value. */
export function xmlAttr(text: string): string {
  return escapeAttr(sanitize(text));
}
