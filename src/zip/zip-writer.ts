import type { ByteSink } from '../sink.js';
import { crc32 } from './crc32.js';

export interface ZipEntry {
  /** Appends uncompressed bytes. Resolves when the compressor accepted them (backpressure). */
  write(data: Uint8Array): Promise<void>;
  /** Finishes compression and writes the data descriptor. Idempotent. */
  close(): Promise<void>;
}

export interface ZipWriterOptions {
  /** Maximum compressed and uncompressed size of a single entry. Default 0xFFFFFFFF. Tests lower it. */
  entrySizeLimit?: number;
}

/** @internal Bookkeeping for one archive member. */
export interface EntryRecord {
  name: Uint8Array;
  crc: number;
  compressedSize: number;
  uncompressedSize: number;
  offset: number;
  time: number;
  date: number;
}

const SIG_LOCAL_HEADER = 0x04034b50;
const SIG_DATA_DESCRIPTOR = 0x08074b50;
const SIG_CENTRAL_HEADER = 0x02014b50;
const SIG_EOCD = 0x06054b50;
const SIG_ZIP64_EOCD = 0x06064b50;
const SIG_ZIP64_LOCATOR = 0x07064b50;
const FLAG_DATA_DESCRIPTOR = 0x0008;
const FLAG_UTF8_NAMES = 0x0800;
const FLAGS = FLAG_DATA_DESCRIPTOR | FLAG_UTF8_NAMES;
const METHOD_DEFLATE = 8;
const VERSION_DEFLATE = 20;
const VERSION_ZIP64 = 45;
const MAX_UINT16 = 0xffff;
const MAX_UINT32 = 0xffffffff;
const ENCODER = new TextEncoder();
const noop = (): void => {};

/**
 * Writes a ZIP archive to a ByteSink one entry at a time, without knowing sizes in advance
 * (data descriptors). Every operation is sequenced through one promise chain, so callers may
 * begin the next entry right after calling close() on the previous one.
 */
export class ZipWriter {
  private readonly limit: number;
  private offset = 0;
  private readonly entries: EntryRecord[] = [];
  private tail: Promise<void> = Promise.resolve();
  private current: Entry | null = null;
  private failure: { reason: unknown } | null = null;
  private finishPromise: Promise<void> | null = null;

  constructor(
    private readonly sink: ByteSink,
    options: ZipWriterOptions = {},
  ) {
    this.limit = options.entrySizeLimit ?? MAX_UINT32;
  }

  beginEntry(name: string): ZipEntry {
    if (this.failure) throw this.failure.reason;
    if (this.finishPromise) throw new Error('ZipWriter is already finished');
    if (this.current && !this.current.closing) {
      throw new Error(`Cannot start entry "${name}" while the previous entry is still open`);
    }
    const entry = new Entry(this, name);
    this.current = entry;
    return entry;
  }

  /** Writes the central directory and end-of-central-directory records. Idempotent. */
  finish(): Promise<void> {
    if (this.finishPromise) return this.finishPromise;
    this.finishPromise = this.enqueue(async () => {
      if (this.current) throw new Error('Cannot finish the archive while an entry is still open');
      const cdOffset = this.offset;
      for (const entry of this.entries) await this.emit(centralDirectoryHeader(entry));
      const cdSize = this.offset - cdOffset;
      const zip64 = cdOffset > MAX_UINT32 || cdSize > MAX_UINT32 || this.entries.length > MAX_UINT16;
      if (zip64) {
        const zip64EocdOffset = this.offset;
        await this.emit(zip64EndOfCentralDirectory(this.entries.length, cdSize, cdOffset));
        await this.emit(zip64Locator(zip64EocdOffset));
      }
      await this.emit(endOfCentralDirectory(this.entries.length, cdSize, cdOffset, zip64));
    });
    return this.finishPromise;
  }

  /** Marks the archive as failed; every later operation rejects with `reason`. */
  fail(reason: unknown): void {
    if (!this.failure) this.failure = { reason };
  }

  /** @internal Runs `op` after every previously enqueued operation. */
  enqueue<T>(op: () => Promise<T>): Promise<T> {
    const run = this.tail.then(() => {
      if (this.failure) throw this.failure.reason;
      return op();
    });
    this.tail = run.then(noop, (reason) => this.fail(reason));
    return run;
  }

  /** @internal Writes raw bytes to the sink and advances the archive offset. */
  async emit(bytes: Uint8Array): Promise<void> {
    this.offset += bytes.byteLength;
    await this.sink.write(bytes);
  }

  /** @internal */
  get position(): number {
    return this.offset;
  }

  /** @internal */
  get entrySizeLimit(): number {
    return this.limit;
  }

  /**
   * @internal Called by an entry once its data descriptor is written. The next entry may already have begun
   * (an entry can start as soon as the previous one is closing), so only clear `current` if it is still `entry`.
   */
  complete(entry: ZipEntry, record: EntryRecord): void {
    this.entries.push(record);
    if (this.current === entry) this.current = null;
  }
}

class Entry implements ZipEntry {
  closing = false;
  private readonly record: EntryRecord;
  private writer: WritableStreamDefaultWriter<Uint8Array> | null = null;
  private pump: Promise<void> = Promise.resolve();
  private crc = 0;
  private uncompressed = 0;
  private compressed = 0;
  private closePromise: Promise<void> | null = null;

  constructor(
    private readonly zip: ZipWriter,
    private readonly name: string,
  ) {
    this.record = { name: ENCODER.encode(name), crc: 0, compressedSize: 0, uncompressedSize: 0, offset: 0, ...dosDateTime(new Date()) };
    // A failure here resurfaces on the next write()/close(), which run later in the same chain.
    zip.enqueue(() => this.start()).catch(noop);
  }

  write(data: Uint8Array): Promise<void> {
    if (this.closing) return Promise.reject(new Error(`ZIP entry "${this.name}" is already closed`));
    return this.zip.enqueue(async () => {
      this.uncompressed += data.byteLength;
      if (this.uncompressed > this.zip.entrySizeLimit) await this.abort(this.sizeError());
      this.crc = crc32(data, this.crc);
      await this.writer!.write(data);
    });
  }

  close(): Promise<void> {
    if (this.closePromise) return this.closePromise;
    this.closing = true;
    this.closePromise = this.zip.enqueue(async () => {
      await this.writer!.close();
      await this.pump;
      this.record.crc = this.crc;
      this.record.compressedSize = this.compressed;
      this.record.uncompressedSize = this.uncompressed;
      await this.zip.emit(dataDescriptor(this.record));
      this.zip.complete(this, this.record);
    });
    return this.closePromise;
  }

  private async start(): Promise<void> {
    this.record.offset = this.zip.position;
    await this.zip.emit(localFileHeader(this.record));
    const compressor = new CompressionStream('deflate-raw');
    // The DOM lib types the writable side as BufferSource, which rejects Uint8Array<ArrayBufferLike>.
    this.writer = compressor.writable.getWriter() as WritableStreamDefaultWriter<Uint8Array>;
    this.pump = this.drain(compressor.readable);
    this.pump.catch(noop); // observed later in close(); avoid an unhandled rejection meanwhile
  }

  /** Forwards compressed output to the archive, counting compressed bytes. */
  private async drain(readable: ReadableStream<Uint8Array>): Promise<void> {
    const reader = readable.getReader();
    try {
      for (;;) {
        const { done, value } = await reader.read();
        if (done) return;
        this.compressed += value.byteLength;
        if (this.compressed > this.zip.entrySizeLimit) throw this.sizeError();
        await this.zip.emit(value);
      }
    } catch (reason) {
      // The pump has stopped reading, so a write()/close() in flight on the compressor is blocked on
      // backpressure and writer.abort() would wait for it forever. Cancelling the output side errors
      // the transform and releases them before the writer is aborted.
      void reader.cancel(reason).catch(noop);
      await this.abort(reason);
    }
  }

  /** Fails the archive and unblocks any write waiting on the compressor. Always throws. */
  private async abort(reason: unknown): Promise<never> {
    this.zip.fail(reason);
    await this.writer?.abort(reason).catch(noop);
    throw reason;
  }

  private sizeError(): RangeError {
    return new RangeError(`ZIP entry "${this.name}" exceeds the 4 GiB limit; split the data across several sheets`);
  }
}

function dosDateTime(d: Date): { time: number; date: number } {
  const year = Math.max(1980, d.getFullYear());
  return {
    time: (d.getHours() << 11) | (d.getMinutes() << 5) | (d.getSeconds() >> 1),
    date: ((year - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate(),
  };
}

/** @internal */
export function localFileHeader(e: EntryRecord): Uint8Array {
  const bytes = new Uint8Array(30 + e.name.byteLength);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, SIG_LOCAL_HEADER, true);
  view.setUint16(4, VERSION_DEFLATE, true);
  view.setUint16(6, FLAGS, true);
  view.setUint16(8, METHOD_DEFLATE, true);
  view.setUint16(10, e.time, true);
  view.setUint16(12, e.date, true);
  // CRC-32 (14), compressed (18) and uncompressed (22) sizes stay 0: the data descriptor carries them.
  view.setUint16(26, e.name.byteLength, true);
  view.setUint16(28, 0, true);
  bytes.set(e.name, 30);
  return bytes;
}

/** @internal */
export function dataDescriptor(e: EntryRecord): Uint8Array {
  const bytes = new Uint8Array(16);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, SIG_DATA_DESCRIPTOR, true);
  view.setUint32(4, e.crc, true);
  view.setUint32(8, e.compressedSize, true);
  view.setUint32(12, e.uncompressedSize, true);
  return bytes;
}

/** @internal */
export function centralDirectoryHeader(e: EntryRecord): Uint8Array {
  const zip64 = e.offset > MAX_UINT32;
  const extraLength = zip64 ? 12 : 0;
  const bytes = new Uint8Array(46 + e.name.byteLength + extraLength);
  const view = new DataView(bytes.buffer);
  const version = zip64 ? VERSION_ZIP64 : VERSION_DEFLATE;
  view.setUint32(0, SIG_CENTRAL_HEADER, true);
  view.setUint16(4, version, true); // version made by
  view.setUint16(6, version, true); // version needed to extract
  view.setUint16(8, FLAGS, true);
  view.setUint16(10, METHOD_DEFLATE, true);
  view.setUint16(12, e.time, true);
  view.setUint16(14, e.date, true);
  view.setUint32(16, e.crc, true);
  view.setUint32(20, e.compressedSize, true);
  view.setUint32(24, e.uncompressedSize, true);
  view.setUint16(28, e.name.byteLength, true);
  view.setUint16(30, extraLength, true);
  // Comment length (32), disk number (34), internal (36) and external (38) attributes stay 0.
  view.setUint32(42, zip64 ? MAX_UINT32 : e.offset, true);
  bytes.set(e.name, 46);
  if (zip64) {
    const extra = 46 + e.name.byteLength;
    view.setUint16(extra, 0x0001, true); // ZIP64 extended information
    view.setUint16(extra + 2, 8, true);
    view.setBigUint64(extra + 4, BigInt(e.offset), true);
  }
  return bytes;
}

/** @internal */
export function zip64EndOfCentralDirectory(count: number, cdSize: number, cdOffset: number): Uint8Array {
  const bytes = new Uint8Array(56);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, SIG_ZIP64_EOCD, true);
  view.setBigUint64(4, 44n, true); // size of the remaining record
  view.setUint16(12, VERSION_ZIP64, true);
  view.setUint16(14, VERSION_ZIP64, true);
  // Disk number (16) and central directory disk (20) stay 0.
  view.setBigUint64(24, BigInt(count), true);
  view.setBigUint64(32, BigInt(count), true);
  view.setBigUint64(40, BigInt(cdSize), true);
  view.setBigUint64(48, BigInt(cdOffset), true);
  return bytes;
}

/** @internal */
export function zip64Locator(zip64EocdOffset: number): Uint8Array {
  const bytes = new Uint8Array(20);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, SIG_ZIP64_LOCATOR, true);
  view.setUint32(4, 0, true);
  view.setBigUint64(8, BigInt(zip64EocdOffset), true);
  view.setUint32(16, 1, true); // total number of disks
  return bytes;
}

/** @internal */
export function endOfCentralDirectory(count: number, cdSize: number, cdOffset: number, zip64: boolean): Uint8Array {
  const bytes = new Uint8Array(22);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, SIG_EOCD, true);
  // Disk number (4) and central directory disk (6) stay 0.
  view.setUint16(8, zip64 ? MAX_UINT16 : count, true);
  view.setUint16(10, zip64 ? MAX_UINT16 : count, true);
  view.setUint32(12, zip64 ? MAX_UINT32 : cdSize, true);
  view.setUint32(16, zip64 ? MAX_UINT32 : cdOffset, true);
  view.setUint16(20, 0, true);
  return bytes;
}
