/** Destination for archive bytes. `write` resolves when the producer may continue (backpressure). */
export interface ByteSink {
  write(chunk: Uint8Array): Promise<void>;
}

interface Waiter {
  resolve(): void;
  reject(reason: unknown): void;
}

/**
 * Exposes a push-style sink over a ReadableStream.
 * Writes resolve immediately while the stream's queue is below its high-water mark,
 * otherwise they wait until the consumer pulls. Consumer cancellation fails the sink.
 */
export class StreamSink implements ByteSink {
  readonly readable: ReadableStream<Uint8Array>;
  /** Invoked once when the sink fails (consumer cancelled the stream or `error()` was called). */
  onFailure: ((reason: unknown) => void) | null = null;

  private controller!: ReadableStreamDefaultController<Uint8Array>;
  private waiters: Waiter[] = [];
  private failure: { reason: unknown } | null = null;
  private closed = false;

  constructor(highWaterMark = 1 << 20) {
    this.readable = new ReadableStream<Uint8Array>(
      {
        start: (controller) => {
          this.controller = controller;
        },
        pull: () => {
          this.release();
        },
        cancel: (reason) => {
          this.fail(reason === undefined ? new Error('Stream was cancelled by the consumer') : reason);
        },
      },
      new ByteLengthQueuingStrategy({ highWaterMark }),
    );
  }

  write(chunk: Uint8Array): Promise<void> {
    if (this.failure) return Promise.reject(this.failure.reason);
    if (this.closed) return Promise.reject(new Error('Stream is already closed'));
    this.controller.enqueue(chunk);
    if ((this.controller.desiredSize ?? 0) > 0) return Promise.resolve();
    return new Promise<void>((resolve, reject) => {
      this.waiters.push({ resolve, reject });
    });
  }

  close(): void {
    if (this.failure || this.closed) return;
    this.closed = true;
    this.controller.close();
  }

  error(reason: unknown): void {
    this.fail(reason);
  }

  private fail(reason: unknown): void {
    if (this.failure || this.closed) return;
    this.failure = { reason };
    try {
      this.controller.error(reason);
    } catch {
      // The stream is already cancelled or errored; nothing more to do.
    }
    const waiters = this.waiters;
    this.waiters = [];
    for (const waiter of waiters) waiter.reject(reason);
    this.onFailure?.(reason);
  }

  private release(): void {
    if (this.waiters.length === 0) return;
    const waiters = this.waiters;
    this.waiters = [];
    for (const waiter of waiters) waiter.resolve();
  }
}
