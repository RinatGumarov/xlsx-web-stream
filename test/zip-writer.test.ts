import { unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { StreamSink } from '../src/sink.js';
import { crc32 } from '../src/zip/crc32.js';
import {
  centralDirectoryHeader,
  endOfCentralDirectory,
  zip64EndOfCentralDirectory,
  zip64Locator,
  ZipWriter,
  type EntryRecord,
  type ZipEntry,
} from '../src/zip/zip-writer.js';
import { decode, encode, FailingSink, littleEndian, MemorySink, randomBytes } from './helpers.js';

/** Minimal parser of the end-of-central-directory record and central directory entries. */
function readCentralDirectory(zip: Uint8Array) {
  const view = littleEndian(zip);
  const eocd = zip.byteLength - 22;
  expect(view.getUint32(eocd, true)).toBe(0x06054b50);
  const count = view.getUint16(eocd + 10, true);
  const cdOffset = view.getUint32(eocd + 16, true);
  const entries = [];
  let pos = cdOffset;
  for (let i = 0; i < count; i++) {
    expect(view.getUint32(pos, true)).toBe(0x02014b50);
    const nameLength = view.getUint16(pos + 28, true);
    const extraLength = view.getUint16(pos + 30, true);
    entries.push({
      flags: view.getUint16(pos + 8, true),
      method: view.getUint16(pos + 10, true),
      crc: view.getUint32(pos + 16, true),
      compressedSize: view.getUint32(pos + 20, true),
      uncompressedSize: view.getUint32(pos + 24, true),
      localOffset: view.getUint32(pos + 42, true),
      name: decode(zip.subarray(pos + 46, pos + 46 + nameLength)),
    });
    pos += 46 + nameLength + extraLength;
  }
  return entries;
}

describe('ZipWriter', () => {
  it('produces an archive that an independent unzipper restores byte for byte', async () => {
    const sink = new MemorySink();
    const zip = new ZipWriter(sink);
    const text = zip.beginEntry('a.txt');
    await text.write(encode('hello '));
    await text.write(encode('world'));
    await text.close();
    const data = randomBytes(100_000);
    const binary = zip.beginEntry('dir/b.bin');
    await binary.write(data);
    await binary.close();
    await zip.finish();

    const files = unzipSync(sink.bytes());
    expect(decode(files['a.txt']!)).toBe('hello world');
    expect(files['dir/b.bin']).toEqual(data);
  });

  it('writes streaming local headers, a data descriptor and a matching central directory', async () => {
    const sink = new MemorySink();
    const zip = new ZipWriter(sink);
    const payload = encode('x'.repeat(1000));
    const entry = zip.beginEntry('f.txt');
    await entry.write(payload);
    await entry.close();
    await zip.finish();
    const bytes = sink.bytes();
    const view = littleEndian(bytes);

    // Local header: signature, version 20, flags = data descriptor (bit 3) + UTF-8 (bit 11), deflate, zero sizes.
    expect(view.getUint32(0, true)).toBe(0x04034b50);
    expect(view.getUint16(4, true)).toBe(20);
    expect(view.getUint16(6, true)).toBe(0x0808);
    expect(view.getUint16(8, true)).toBe(8);
    expect(view.getUint32(14, true)).toBe(0);
    expect(view.getUint32(18, true)).toBe(0);
    expect(view.getUint32(22, true)).toBe(0);
    expect(decode(bytes.subarray(30, 35))).toBe('f.txt');

    const [record] = readCentralDirectory(bytes);
    expect(record).toMatchObject({ name: 'f.txt', flags: 0x0808, method: 8, localOffset: 0, uncompressedSize: 1000, crc: crc32(payload) });
    expect(record!.compressedSize).toBeLessThan(1000);

    // Data descriptor sits right after the compressed data.
    const descriptor = 30 + 5 + record!.compressedSize;
    expect(view.getUint32(descriptor, true)).toBe(0x08074b50);
    expect(view.getUint32(descriptor + 4, true)).toBe(crc32(payload));
    expect(view.getUint32(descriptor + 8, true)).toBe(record!.compressedSize);
    expect(view.getUint32(descriptor + 12, true)).toBe(1000);
  });

  it('sequences entries: the next entry may begin as soon as close() of the previous one is called', async () => {
    const sink = new MemorySink();
    const zip = new ZipWriter(sink);
    const first = zip.beginEntry('1.txt');
    void first.write(encode('one'));
    void first.close();
    const second = zip.beginEntry('2.txt');
    await second.write(encode('two'));
    await second.close();
    await zip.finish();
    const files = unzipSync(sink.bytes());
    expect(decode(files['1.txt']!)).toBe('one');
    expect(decode(files['2.txt']!)).toBe('two');
    expect(readCentralDirectory(sink.bytes()).map((e) => e.name)).toEqual(['1.txt', '2.txt']);
  });

  it('refuses to begin an entry while another one is open', () => {
    const zip = new ZipWriter(new MemorySink());
    zip.beginEntry('a');
    expect(() => zip.beginEntry('b')).toThrow('still open');
  });

  it('keeps guarding the open entry after the previous entry finishes closing', async () => {
    const sink = new MemorySink();
    const zip = new ZipWriter(sink);
    const a = zip.beginEntry('a.txt');
    void a.write(encode('1'));
    const closeA = a.close();
    const b = zip.beginEntry('b.txt'); // allowed: a is already closing
    await closeA; // a completes while b is open and must not clear the guard
    expect(() => zip.beginEntry('c.txt')).toThrow('still open');
    await b.write(encode('2'));
    await b.close();
    await zip.finish();
    const files = unzipSync(sink.bytes());
    expect(decode(files['a.txt']!)).toBe('1');
    expect(decode(files['b.txt']!)).toBe('2');
    expect(readCentralDirectory(sink.bytes()).map((e) => e.name)).toEqual(['a.txt', 'b.txt']);
  });

  it('begins no entry after a failure and surfaces the failure reason, even while an entry is still open', () => {
    const zip = new ZipWriter(new MemorySink());
    zip.beginEntry('a'); // left open on purpose
    const reason = new Error('sink died');
    zip.fail(reason);
    expect(() => zip.beginEntry('b')).toThrow(reason);
  });

  it('refuses writes to a closed entry and operations after finish()', async () => {
    const zip = new ZipWriter(new MemorySink());
    const entry = zip.beginEntry('a');
    void entry.close();
    await expect(entry.write(encode('x'))).rejects.toThrow('already closed');
    await zip.finish();
    expect(() => zip.beginEntry('b')).toThrow('already finished');
  });

  it('rejects an entry that grows beyond the size limit and fails the archive', async () => {
    const zip = new ZipWriter(new MemorySink(), { entrySizeLimit: 10 });
    const entry = zip.beginEntry('big');
    await expect(entry.write(new Uint8Array(11))).rejects.toThrow(RangeError);
    await expect(entry.close()).rejects.toThrow(/4 GiB/);
    await expect(zip.finish()).rejects.toThrow(/4 GiB/);
  });

  it('reports the first failure through onFailure exactly once (size limit)', async () => {
    const zip = new ZipWriter(new MemorySink(), { entrySizeLimit: 10 });
    const reasons: unknown[] = [];
    zip.onFailure = (reason) => reasons.push(reason);
    const entry = zip.beginEntry('big');
    const rejection = await entry.write(new Uint8Array(11)).catch((reason: unknown) => reason);
    expect(rejection).toBeInstanceOf(RangeError);
    await expect(entry.close()).rejects.toBe(rejection);
    await expect(zip.finish()).rejects.toBe(rejection);
    zip.fail(new Error('a later failure')); // already failed: ignored
    expect(reasons).toEqual([rejection]);
    expect((reasons[0] as Error).message).toMatch(/exceeds the 4 GiB limit; the archive cannot be completed\. Keep each sheet below 4 GiB compressed and uncompressed/);
  });

  it('reports a failing sink through onFailure exactly once, with the first reason', async () => {
    const diskFull = new Error('disk full');
    const zip = new ZipWriter(new FailingSink(diskFull));
    const reasons: unknown[] = [];
    zip.onFailure = (reason) => reasons.push(reason);
    const entry = zip.beginEntry('a');
    await expect(entry.write(encode('x'))).rejects.toBe(diskFull);
    await expect(entry.close()).rejects.toBe(diskFull);
    await expect(zip.finish()).rejects.toBe(diskFull);
    zip.fail(new Error('a later failure'));
    expect(reasons).toEqual([diskFull]);
  });

  it('does not call onFailure when nothing fails', async () => {
    const zip = new ZipWriter(new MemorySink());
    let calls = 0;
    zip.onFailure = () => calls++;
    const entry = zip.beginEntry('a');
    await entry.write(encode('x'));
    await entry.close();
    await zip.finish();
    expect(calls).toBe(0);
  });

  it('propagates sink failures to every operation', async () => {
    const zip = new ZipWriter(new FailingSink(new Error('disk full')));
    const entry = zip.beginEntry('a');
    await expect(entry.write(encode('x'))).rejects.toThrow('disk full');
    await expect(entry.close()).rejects.toThrow('disk full');
    await expect(zip.finish()).rejects.toThrow('disk full');
  });

  it('releases a write blocked on backpressure when the consumer cancels the stream', { timeout: 5000 }, async () => {
    const sink = new StreamSink(64 * 1024);
    const zip = new ZipWriter(sink);
    const entry = zip.beginEntry('big.bin');
    // Incompressible 8 MB: the compressed output far exceeds the sink's high-water mark, so the pump stalls.
    const pendingWrite = entry.write(randomBytes(8 * 1024 * 1024));
    const pendingWriteResult = expect(pendingWrite).rejects.toBe('consumer went away');
    await new Promise((resolve) => setTimeout(resolve, 50));

    await sink.readable.cancel('consumer went away');

    await pendingWriteResult;
    await expect(entry.close()).rejects.toBe('consumer went away');
    await expect(zip.finish()).rejects.toBe('consumer went away');
  });

  it('never lets an entry reach the 0xFFFFFFFF size sentinel', () => {
    expect(new ZipWriter(new MemorySink()).entrySizeLimit).toBe(0xfffffffe);
  });

  it('switches to ZIP64 end records once the entry count reaches the 0xFFFF sentinel', async () => {
    const finishWith = async (count: number): Promise<Uint8Array> => {
      const sink = new MemorySink();
      const zip = new ZipWriter(sink);
      const record: EntryRecord = { name: encode('x'), crc: 0, compressedSize: 0, uncompressedSize: 0, offset: 0, time: 0, date: 0 };
      for (let i = 0; i < count; i++) zip.complete({} as ZipEntry, record); // bookkeeping only, no real entries
      await zip.finish();
      return sink.bytes();
    };
    const signatureAt = (bytes: Uint8Array, offset: number): number => littleEndian(bytes).getUint32(offset, true);

    const below = await finishWith(0xfffe);
    expect(signatureAt(below, below.byteLength - 22)).toBe(0x06054b50);
    expect(littleEndian(below).getUint16(below.byteLength - 22 + 10, true)).toBe(0xfffe);
    expect(signatureAt(below, below.byteLength - 22 - 20)).not.toBe(0x07064b50); // no ZIP64 locator

    const at = await finishWith(0xffff);
    expect(signatureAt(at, at.byteLength - 22 - 20)).toBe(0x07064b50); // ZIP64 locator right before the EOCD
    expect(signatureAt(at, at.byteLength - 22 - 20 - 56)).toBe(0x06064b50); // ZIP64 EOCD record before the locator
    expect(littleEndian(at).getBigUint64(at.byteLength - 22 - 20 - 56 + 24, true)).toBe(0xffffn);
  });

  it('encodes ZIP64 structures when offsets exceed 32 bits', () => {
    const record: EntryRecord = {
      name: encode('x'),
      crc: 1,
      compressedSize: 2,
      uncompressedSize: 3,
      offset: 0x1_0000_0005,
      time: 0,
      date: 0,
    };
    const header = littleEndian(centralDirectoryHeader(record));
    expect(header.byteLength).toBe(46 + 1 + 12);
    expect(header.getUint16(4, true)).toBe(45); // version made by
    expect(header.getUint16(6, true)).toBe(45); // version needed
    expect(header.getUint16(30, true)).toBe(12); // extra length
    expect(header.getUint32(42, true)).toBe(0xffffffff);
    expect(header.getUint16(47, true)).toBe(0x0001); // ZIP64 extra field id
    expect(header.getUint16(49, true)).toBe(8);
    expect(header.getBigUint64(51, true)).toBe(0x1_0000_0005n);

    // 0xFFFFFFFF is the "see the ZIP64 extra field" sentinel, so an offset equal to it must use the ZIP64 form too.
    const sentinel = littleEndian(centralDirectoryHeader({ ...record, offset: 0xffffffff }));
    expect(sentinel.byteLength).toBe(46 + 1 + 12);
    expect(sentinel.getUint16(30, true)).toBe(12);
    expect(sentinel.getUint32(42, true)).toBe(0xffffffff);
    expect(sentinel.getBigUint64(51, true)).toBe(0xffffffffn);
    const belowSentinel = littleEndian(centralDirectoryHeader({ ...record, offset: 0xfffffffe }));
    expect(belowSentinel.byteLength).toBe(47);
    expect(belowSentinel.getUint32(42, true)).toBe(0xfffffffe);

    const plain = littleEndian(centralDirectoryHeader({ ...record, offset: 7 }));
    expect(plain.byteLength).toBe(47);
    expect(plain.getUint16(6, true)).toBe(20);
    expect(plain.getUint32(42, true)).toBe(7);

    const eocd = littleEndian(endOfCentralDirectory(70_000, 123, 0x1_0000_0000, true));
    expect(eocd.getUint16(8, true)).toBe(0xffff);
    expect(eocd.getUint16(10, true)).toBe(0xffff);
    expect(eocd.getUint32(12, true)).toBe(0xffffffff);
    expect(eocd.getUint32(16, true)).toBe(0xffffffff);

    const normal = littleEndian(endOfCentralDirectory(2, 123, 456, false));
    expect(normal.getUint16(10, true)).toBe(2);
    expect(normal.getUint32(12, true)).toBe(123);
    expect(normal.getUint32(16, true)).toBe(456);

    const z64 = littleEndian(zip64EndOfCentralDirectory(70_000, 123, 0x1_0000_0000));
    expect(z64.byteLength).toBe(56);
    expect(z64.getUint32(0, true)).toBe(0x06064b50);
    expect(z64.getBigUint64(4, true)).toBe(44n);
    expect(z64.getBigUint64(24, true)).toBe(70_000n);
    expect(z64.getBigUint64(32, true)).toBe(70_000n);
    expect(z64.getBigUint64(40, true)).toBe(123n);
    expect(z64.getBigUint64(48, true)).toBe(0x1_0000_0000n);

    const locator = littleEndian(zip64Locator(0x1_0000_0100));
    expect(locator.byteLength).toBe(20);
    expect(locator.getUint32(0, true)).toBe(0x07064b50);
    expect(locator.getBigUint64(8, true)).toBe(0x1_0000_0100n);
    expect(locator.getUint32(16, true)).toBe(1);
  });
});
