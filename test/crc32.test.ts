import { describe, expect, it } from 'vitest';
import { crc32 } from '../src/zip/crc32.js';

const bytes = (s: string): Uint8Array => new TextEncoder().encode(s);

describe('crc32', () => {
  it('matches the standard check value for "123456789"', () => {
    expect(crc32(bytes('123456789'))).toBe(0xcbf43926);
  });

  it('returns 0 for empty input', () => {
    expect(crc32(new Uint8Array(0))).toBe(0);
  });

  it('matches a known single-byte value and is unsigned', () => {
    const value = crc32(bytes('a'));
    expect(value).toBe(0xe8b7be43);
    expect(value).toBeGreaterThan(0);
  });

  it('can be computed incrementally by passing the previous result as seed', () => {
    const partial = crc32(bytes('1234'));
    expect(crc32(bytes('56789'), partial)).toBe(0xcbf43926);
  });
});
