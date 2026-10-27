import { describe, expect, it } from 'vitest';
import { createWorkbook } from '../src/index.js';
import { randomBytes } from './helpers.js';

/** Incompressible text so that the 1 MiB output queue fills after a few MB of input. */
function noise(): string {
  return btoa(String.fromCharCode(...randomBytes(768)));
}

const settle = (): Promise<void> => new Promise((resolve) => setTimeout(resolve, 50));

describe('backpressure', () => {
  it('stops accepting rows while nobody reads the stream and resumes when reading starts', async () => {
    const wb = createWorkbook();
    const reader = wb.readable.getReader();
    const sheet = wb.addSheet('S');
    const writes: Promise<void>[] = [];
    for (let i = 0; i < 4000; i++) writes.push(sheet.writeRow([i, noise()]));
    let allSettled = false;
    const all = Promise.all(writes).then(() => {
      allSettled = true;
    });
    await settle();
    expect(allSettled).toBe(false); // the producer is blocked by the full queue

    let bytes = 0;
    const drain = (async () => {
      for (;;) {
        const { done, value } = await reader.read();
        if (done) return;
        bytes += value.byteLength;
      }
    })();
    await all;
    expect(allSettled).toBe(true);
    await wb.close();
    await drain;
    expect(bytes).toBeGreaterThan(2_500_000); // ~4 MB of base64 noise deflates to roughly 75 %
  });

  it('never buffers more than a few MB of output ahead of a slow consumer', async () => {
    const wb = createWorkbook();
    const sheet = wb.addSheet('S');
    let produced = 0;
    let consumed = 0;
    const samples: Array<[produced: number, consumed: number]> = [];
    const output = wb.readable.pipeTo(
      new WritableStream(
        {
          async write(chunk) {
            await new Promise((resolve) => setTimeout(resolve, 1));
            consumed += chunk.byteLength;
          },
        },
        new CountQueuingStrategy({ highWaterMark: 1 }),
      ),
    );
    for (let i = 0; i < 12_000; i++) {
      await sheet.writeRow([i, noise()]);
      produced += 1024;
      samples.push([produced, consumed]);
    }
    await wb.close();
    await output;
    // `consumed` counts deflated bytes and `produced` raw ones. Convert with the overall ratio so that the lag is
    // measured in raw bytes and does not grow with the deflate savings (the naive difference reaches ~4 MB by itself).
    const ratio = consumed / produced;
    const maxLag = samples.reduce((max, [p, c]) => Math.max(max, p - c / ratio), 0);
    // A healthy pipeline peaks at ~1.6 MB: 1 MiB output queue + compressor queues + one 64 KiB flush. The input (12 MB)
    // has to be far above the bound, otherwise a writer that ignores backpressure (6.6 MB lag at the sink, 12 MB if
    // writeRow never waits) would pass as well.
    expect(maxLag).toBeLessThan(3 * 1024 * 1024);
  });
});
