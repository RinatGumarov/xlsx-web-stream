import { describe, expect, it } from 'vitest';
import { StreamSink } from '../src/sink.js';

const tick = (): Promise<void> => new Promise((resolve) => setTimeout(resolve, 0));

describe('StreamSink', () => {
  it('resolves writes immediately while the queue is below the high-water mark', async () => {
    const sink = new StreamSink(100);
    await expect(sink.write(new Uint8Array(10))).resolves.toBeUndefined();
  });

  it('delivers written chunks to the readable in order', async () => {
    const sink = new StreamSink();
    void sink.write(new Uint8Array([1, 2]));
    void sink.write(new Uint8Array([3]));
    sink.close();
    const chunks: number[] = [];
    const reader = sink.readable.getReader();
    for (let result = await reader.read(); !result.done; result = await reader.read()) {
      chunks.push(...result.value);
    }
    expect(chunks).toEqual([1, 2, 3]);
  });

  it('holds the write promise until the consumer drains the queue', async () => {
    const sink = new StreamSink(16);
    const reader = sink.readable.getReader();
    await sink.write(new Uint8Array(8)); // desiredSize 8 > 0 → resolves at once
    let settled = false;
    const pending = sink.write(new Uint8Array(16)).then(() => {
      settled = true;
    });
    await tick();
    expect(settled).toBe(false); // desiredSize is now -8
    await reader.read(); // queue: 16 bytes, desiredSize 0 → still blocked
    await tick();
    expect(settled).toBe(false);
    await reader.read(); // queue empty → pull → release
    await pending;
    expect(settled).toBe(true);
  });

  it('rejects pending and future writes when the consumer cancels', async () => {
    const sink = new StreamSink(8);
    const failures: unknown[] = [];
    sink.onFailure = (reason) => failures.push(reason);
    const reader = sink.readable.getReader();
    const pending = sink.write(new Uint8Array(16));
    await reader.cancel('gone');
    await expect(pending).rejects.toBe('gone');
    await expect(sink.write(new Uint8Array(1))).rejects.toBe('gone');
    expect(failures).toEqual(['gone']);
  });

  it('uses a descriptive error when the consumer cancels without a reason', async () => {
    const sink = new StreamSink(8);
    const reader = sink.readable.getReader();
    await reader.cancel();
    await expect(sink.write(new Uint8Array(1))).rejects.toThrow('cancelled by the consumer');
  });

  it('error() fails the readable and rejects later writes', async () => {
    const sink = new StreamSink();
    const reader = sink.readable.getReader();
    const boom = new Error('boom');
    sink.error(boom);
    await expect(reader.read()).rejects.toBe(boom);
    await expect(sink.write(new Uint8Array(1))).rejects.toBe(boom);
  });

  it('close() ends the readable and rejects later writes', async () => {
    const sink = new StreamSink();
    const reader = sink.readable.getReader();
    sink.close();
    expect(await reader.read()).toEqual({ done: true, value: undefined });
    await expect(sink.write(new Uint8Array(1))).rejects.toThrow('already closed');
  });

  it('settles a write that is waiting for backpressure when close() is called', async () => {
    const sink = new StreamSink(8);
    let outcome = 'pending';
    const pending = sink.write(new Uint8Array(16)).then(
      () => {
        outcome = 'resolved';
      },
      () => {
        outcome = 'rejected';
      },
    );
    sink.close();
    const reader = sink.readable.getReader();
    let total = 0;
    for (let result = await reader.read(); !result.done; result = await reader.read()) {
      total += result.value.length;
    }
    await tick();
    expect(total).toBe(16); // the chunk was accepted before close and is still delivered
    expect(outcome).toBe('resolved');
    await pending;
  });

  it('settles a write waiting for backpressure when the consumer cancels after close()', async () => {
    const sink = new StreamSink(8);
    let outcome = 'pending';
    const pending = sink.write(new Uint8Array(16)).then(
      () => {
        outcome = 'resolved';
      },
      () => {
        outcome = 'rejected';
      },
    );
    sink.close();
    const reader = sink.readable.getReader();
    await reader.cancel('gone');
    await tick();
    expect(outcome).toBe('resolved'); // the chunk was accepted before close
    await pending;
  });
});
