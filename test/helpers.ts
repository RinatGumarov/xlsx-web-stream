import { strFromU8, unzipSync } from 'fflate';
import { XMLParser } from 'fast-xml-parser';
import type { ByteSink } from '../src/sink.js';
import type { ZipEntry } from '../src/zip/zip-writer.js';

export const encode = (text: string): Uint8Array => new TextEncoder().encode(text);
export const decode = (bytes: Uint8Array): string => new TextDecoder().decode(bytes);

export function concat(chunks: Uint8Array[]): Uint8Array {
  const total = chunks.reduce((sum, c) => sum + c.byteLength, 0);
  const out = new Uint8Array(total);
  let offset = 0;
  for (const chunk of chunks) {
    out.set(chunk, offset);
    offset += chunk.byteLength;
  }
  return out;
}

/** Collects a whole ReadableStream into one buffer. */
export async function collect(readable: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  return new Uint8Array(await new Response(readable).arrayBuffer());
}

/** Unzips with fflate (an independent implementation) → { 'path/in/zip': text }. */
export function unzipText(bytes: Uint8Array): Record<string, string> {
  const files = unzipSync(bytes);
  return Object.fromEntries(Object.entries(files).map(([name, data]) => [name, strFromU8(data)]));
}

export const xmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '',
  isArray: (name) => ['Override', 'Relationship', 'sheet', 'definedName', 'col', 'row', 'c', 'mergeCell', 'xf', 'font', 'fill', 'border', 'numFmt'].includes(name),
});

export function parseXml(xml: string): any {
  return xmlParser.parse(xml);
}

/** In-memory ByteSink. */
export class MemorySink implements ByteSink {
  chunks: Uint8Array[] = [];
  async write(chunk: Uint8Array): Promise<void> {
    this.chunks.push(chunk);
  }
  bytes(): Uint8Array {
    return concat(this.chunks);
  }
}

/** ByteSink that rejects every write with the given error. */
export class FailingSink implements ByteSink {
  constructor(private readonly error: Error) {}
  write(): Promise<void> {
    return Promise.reject(this.error);
  }
}

/** In-memory ZipEntry for testing the sheet writer in isolation. */
export class MemoryEntry implements ZipEntry {
  chunks: Uint8Array[] = [];
  closed = false;
  async write(data: Uint8Array): Promise<void> {
    this.chunks.push(data);
  }
  async close(): Promise<void> {
    this.closed = true;
  }
  text(): string {
    return decode(concat(this.chunks));
  }
}

export function randomBytes(length: number): Uint8Array {
  const out = new Uint8Array(length);
  for (let i = 0; i < length; i += 65536) crypto.getRandomValues(out.subarray(i, Math.min(i + 65536, length)));
  return out;
}

export const littleEndian = (bytes: Uint8Array): DataView => new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
