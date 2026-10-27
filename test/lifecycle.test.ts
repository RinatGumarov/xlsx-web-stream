import { describe, expect, it } from 'vitest';
import { createWorkbook } from '../src/index.js';
import { collect } from './helpers.js';

describe('workbook options', () => {
  it('validates options', () => {
    expect(() => createWorkbook({ dates: 'server' as never })).toThrow(TypeError);
    expect(() => createWorkbook({ dateFormat: '' })).toThrow(TypeError);
    expect(() => createWorkbook(null as never)).toThrow(TypeError);
  });
});

describe('sheet names', () => {
  it('rejects invalid and duplicate names', () => {
    const wb = createWorkbook();
    expect(() => wb.addSheet('')).toThrow(TypeError);
    expect(() => wb.addSheet('x'.repeat(32))).toThrow(TypeError);
    for (const ch of '[]:*?/\\') expect(() => wb.addSheet(`a${ch}b`), ch).toThrow(TypeError);
    expect(() => wb.addSheet("'quoted")).toThrow(TypeError);
    expect(() => wb.addSheet("quoted'")).toThrow(TypeError);
    wb.addSheet('Data');
    expect(() => wb.addSheet('data')).toThrow('already used');
    expect(() => wb.addSheet('Other', 'opts' as never)).toThrow(TypeError);
    wb.abort();
  });

  it('accepts 31-character names with apostrophes inside', () => {
    const wb = createWorkbook();
    expect(() => wb.addSheet(`O'Brien ${'x'.repeat(23)}`)).not.toThrow();
    wb.abort();
  });
});

describe('closing', () => {
  it('rejects closing a workbook without sheets and fails the stream', async () => {
    const wb = createWorkbook();
    const reader = wb.readable.getReader();
    await expect(wb.close()).rejects.toThrow('at least one sheet');
    await expect(reader.read()).rejects.toThrow('at least one sheet');
  });

  it('is idempotent and refuses new sheets or styles afterwards', async () => {
    const wb = createWorkbook();
    const output = collect(wb.readable);
    wb.addSheet('S');
    const first = wb.close();
    expect(wb.close()).toBe(first);
    await first;
    expect(() => wb.addSheet('T')).toThrow('after the workbook was closed');
    expect(() => wb.style({ font: { bold: true } })).toThrow('after the workbook was closed');
    expect((await output).byteLength).toBeGreaterThan(0);
  });

  it('rejects styles from another workbook', () => {
    const a = createWorkbook();
    const b = createWorkbook();
    const foreign = a.style({ font: { bold: true } });
    const sheet = b.addSheet('S');
    expect(() => sheet.writeRow([{ value: 1, style: foreign }])).toThrow(TypeError);
    expect(() => sheet.writeRow([1], { style: foreign })).toThrow(TypeError);
    expect(() => b.addSheet('T', { columns: [{ style: foreign }] })).toThrow(TypeError);
    a.abort();
    b.abort();
  });
});

describe('abort and cancellation', () => {
  it('abort() fails the stream and every pending or later operation', async () => {
    const wb = createWorkbook();
    const reader = wb.readable.getReader();
    const sheet = wb.addSheet('S');
    await sheet.writeRow([1]);
    const reason = new Error('user navigated away');
    wb.abort(reason);
    await expect(reader.read()).rejects.toBe(reason);
    await expect(sheet.writeRow([2])).rejects.toBe(reason);
    await expect(sheet.close()).rejects.toBe(reason);
    await expect(wb.close()).rejects.toBe(reason);
    expect(() => wb.addSheet('T')).toThrow('after the workbook was closed');
  });

  it('abort() without a reason uses a descriptive error', async () => {
    const wb = createWorkbook();
    const reader = wb.readable.getReader();
    wb.addSheet('S');
    wb.abort();
    await expect(reader.read()).rejects.toThrow('aborted');
  });

  it('consumer cancellation rejects the producer with the cancel reason', async () => {
    const wb = createWorkbook();
    const reader = wb.readable.getReader();
    const sheet = wb.addSheet('S');
    await sheet.writeRow(['before cancel']);
    await reader.cancel('download stopped');
    await expect(sheet.writeRow(['after cancel'])).rejects.toBe('download stopped');
    await expect(wb.close()).rejects.toBe('download stopped');
  });

  it('a failing close() errors the stream for the consumer', async () => {
    const wb = createWorkbook();
    const reader = wb.readable.getReader();
    const sheet = wb.addSheet('S');
    await sheet.writeRow([1]);
    await reader.cancel('gone');
    await expect(wb.close()).rejects.toBe('gone');
    await expect(reader.closed).resolves.toBeUndefined();
  });
});
