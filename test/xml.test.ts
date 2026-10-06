import { describe, expect, it } from 'vitest';
import { escapeAttr, escapeText, sanitize, xmlAttr, xmlText, XML_DECL } from '../src/xml.js';

describe('escapeText', () => {
  it('escapes ampersand and angle brackets only', () => {
    expect(escapeText('a & b < c > "d" \'e\'')).toBe('a &amp; b &lt; c &gt; "d" \'e\'');
  });

  it('returns the same string when nothing needs escaping', () => {
    const s = 'plain text';
    expect(escapeText(s)).toBe(s);
  });
});

describe('escapeAttr', () => {
  it('escapes quotes as well', () => {
    expect(escapeAttr('a & "b" \'c\' <d>')).toBe('a &amp; &quot;b&quot; &apos;c&apos; &lt;d&gt;');
  });
});

describe('sanitize', () => {
  it('removes control characters that XML 1.0 forbids', () => {
    expect(sanitize('a\u0000b\u0001c\u001Fd￾e')).toBe('abcde');
  });

  it('keeps tab, newline, carriage return and non-ASCII text', () => {
    const s = 'tab\there\nnew\r\nline Привет 😀';
    expect(sanitize(s)).toBe(s);
  });
});

describe('xmlText / xmlAttr', () => {
  it('sanitizes and escapes in one go', () => {
    expect(xmlText('x\u0000 < y')).toBe('x &lt; y');
    expect(xmlAttr('x\u0000 "y"')).toBe('x &quot;y&quot;');
  });
});

describe('constants', () => {
  it('declares UTF-8 standalone XML', () => {
    expect(XML_DECL).toBe('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
  });
});
