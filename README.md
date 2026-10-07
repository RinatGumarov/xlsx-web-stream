# xlsx-web-stream

Streaming XLSX writer for browsers and Node.js. Constant memory, zero dependencies, built on the Web Streams API.

- **Streams, doesn't buffer.** Rows are serialized, compressed and handed to your destination as you write them. A million rows costs the same memory as a hundred.
- **Zero dependencies.** Compression uses the platform's native `CompressionStream`; the ZIP container is written by the library itself.
- **Runs everywhere Web Streams run.** Browsers, Node.js ≥ 20.12, Deno, Bun, Cloudflare Workers, Web Workers.
- **Formatting that covers real reports.** Fonts, fills, borders, alignment, number and date formats, column widths, frozen headers, auto filters, merged cells, formulas, hyperlinks.

## Install

```bash
npm install xlsx-web-stream
```

## Quick start

```ts
import { createWorkbook } from 'xlsx-web-stream';

const wb = createWorkbook();

// Start consuming the stream first so backpressure can flow while you write.
// (Browser download / File System Access / Node fs variants are in Recipes below.)
const saving = wb.readable.pipeTo(destination);

// Register styles once, reuse them for millions of cells.
const header = wb.style({ font: { bold: true }, fill: '#EEEEEE', border: 'thin' });
const money = wb.style({ numFmt: '#,##0.00' });

const sheet = wb.addSheet('Report', {
  columns: [{ width: 30 }, { width: 14, style: money }, { width: 12 }],
  freeze: { rows: 1 },
});

await sheet.writeRow(['Item', 'Amount', 'Date'], { style: header });
for await (const item of fetchItems()) {
  await sheet.writeRow([item.name, item.total, item.date]);
}
await sheet.writeRow(['Total', { formula: `SUM(B2:B${sheet.nextRow - 1})` }]);
sheet.autoFilter('A1:C1');

await Promise.all([wb.close(), saving]);
```

`wb.readable` is a standard `ReadableStream<Uint8Array>`. Attach its consumer *before* writing rows: the stream has a 1 MiB queue, and once it is full `await sheet.writeRow()` waits until the consumer catches up — that is how memory stays constant. Writing without a consumer attached stalls after a few MB.

## Recipes

### Download in the browser

```ts
const wb = createWorkbook();
const blobReady = new Response(wb.readable).blob(); // collects the stream in the browser's memory
// ... add sheets and rows ...
await wb.close();
const url = URL.createObjectURL(await blobReady);
Object.assign(document.createElement('a'), { href: url, download: 'report.xlsx' }).click();
URL.revokeObjectURL(url);
```

### Save straight to disk in the browser (File System Access API)

```ts
const handle = await showSaveFilePicker({ suggestedName: 'report.xlsx' });
const wb = createWorkbook();
const saving = wb.readable.pipeTo(await handle.createWritable());
// ... add sheets and rows ...
await Promise.all([wb.close(), saving]);
```

Nothing is held in memory: the browser writes chunks to the file as they are produced.

### Node.js file

```ts
import { createWriteStream } from 'node:fs';
import { Writable } from 'node:stream';

const wb = createWorkbook();
const saving = wb.readable.pipeTo(Writable.toWeb(createWriteStream('report.xlsx')));
// ... add sheets and rows ...
await Promise.all([wb.close(), saving]);
```

### HTTP response

Fetch-style runtimes (Cloudflare Workers, Deno, Bun, Next.js route handlers):

```ts
// `writeReport` adds sheets and rows and ends with `await wb.close()`.
// Start it without awaiting, so the response is already being consumed while rows are produced.
const wb = createWorkbook();
writeReport(wb).catch((e) => wb.abort(e));
return new Response(wb.readable, {
  headers: {
    'content-type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'content-disposition': 'attachment; filename="report.xlsx"',
  },
});
```

Node.js `http` / Express:

```ts
import { pipeline } from 'node:stream/promises';
import { Readable } from 'node:stream';
import type { ReadableStream as NodeReadableStream } from 'node:stream/web';

const wb = createWorkbook();
res.setHeader('content-type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
res.setHeader('content-disposition', 'attachment; filename="report.xlsx"');
writeReport(wb).catch((e) => wb.abort(e)); // produce rows while the response streams
// (The cast is only needed under TypeScript's DOM lib, where the two ReadableStream types differ.)
await pipeline(Readable.fromWeb(wb.readable as NodeReadableStream), res); // errors and client disconnects propagate both ways
```

### Web Worker

The library touches no DOM, so generation can run in a Worker. Post `wb.readable` to the main thread with `postMessage(readable, [readable])` *before* writing rows and pipe it there, or build the file inside the worker and post the resulting `Blob`.

## API

### `createWorkbook(options?)`

| Option | Default | Meaning |
|---|---|---|
| `dates` | `'utc'` | Interpret `Date` values by their UTC components (`'utc'`) or local components (`'local'`). Excel has no time zones; `'utc'` produces the same file wherever it is generated, `'local'` shows the time the user sees in their browser. |
| `dateFormat` | `'yyyy-mm-dd'` | Number format for `Date` cells whose style has no `numFmt`. |

Returns a `Workbook`:

- `readable: ReadableStream<Uint8Array>` — the XLSX bytes.
- `style(definition): Style` — registers a style; equal definitions resolve to the same style id.
- `addSheet(name, options?): Sheet` — starts a sheet. Sheets are written one after another: adding a sheet closes the previous one.
- `close(): Promise<void>` — closes the current sheet, writes workbook metadata and ends the stream.
- `abort(reason?)` — fails the stream and every pending operation.

### `Sheet`

- `writeRow(cells, { style?, height? }): Promise<void>` — the promise waits only when the consumer is slower than you; `await` it to keep memory bounded.
- `writeRows(iterable | asyncIterable): Promise<void>`
- `merge('A1:C1')`, `autoFilter('A1:C1')` — may be called any time before the sheet is closed. One `autoFilter` per sheet; `merge` needs at least two cells.
- `nextRow` — 1-based number of the next row, handy for formulas.
- `close(): Promise<void>`

Sheet options: `columns: [{ width?, style? }]` (widths in Excel character units) and `freeze: { rows?, columns? }`. Both must be given up front — they precede the data in the file.

### Cell values

| Value | Written as |
|---|---|
| `string` | Text (inline string) |
| `number` | Number; `NaN`/`Infinity` become the `#NUM!` error cell |
| `boolean` | Boolean |
| `Date` | Date serial with `dateFormat` unless the style has its own `numFmt`; invalid dates and dates outside 1900-01-01 … 9999-12-31 become `#NUM!` |
| `null`, `undefined`, `''` | Empty (a styled empty cell still gets its fill and borders) |
| `{ value, style }` | Value with a cell-level style |
| `{ formula: 'SUM(A1:A9)', value? }` | Formula with an optional cached result; Excel recalculates on open |
| `{ link: 'https://…', value? }` | Hyperlink (`HYPERLINK()` formula) in blue underlined text unless styled |

Style precedence: cell → row → column. The most specific style applies as a whole; styles are not merged property by property (a row style replaces a column style's number format too).

### Styles

```ts
wb.style({
  font: { bold: true, italic: false, underline: false, strike: false, size: 11, color: '#1F4E78', name: 'Calibri' },
  fill: '#DDEBF7',
  border: 'thin', // or per side: { top: 'thin', bottom: { style: 'double', color: '#00FF00' } }
  align: {
    horizontal: 'center', // 'left' | 'center' | 'right' | 'fill' | 'justify' | 'centerContinuous' | 'distributed'
    vertical: 'top', // 'top' | 'center' | 'bottom' | 'justify' | 'distributed'
    wrap: true,
  },
  numFmt: '#,##0.00',
});
```

Every property is optional. Border styles: `thin`, `medium`, `thick`, `dashed`, `dotted`, `double`, `hair`.

## Errors

Validation errors (`TypeError`, `RangeError`, `Error`) are thrown synchronously with a message that says what is wrong: invalid sheet names, more than 1,048,576 rows or 16,384 columns, text over 32,767 characters, malformed ranges, writing to a closed sheet, styles from another workbook. Inside `writeRows()` the same validation errors surface as a rejection of the returned promise. A thrown validation error leaves the workbook open; call `wb.abort(error)` so a consumer waiting on the stream is released.

`NaN`, `Infinity` and unrepresentable dates become `#NUM!` instead of throwing, so one odd value does not fail a long export; structurally invalid cells (text over 32,767 characters, objects that are neither `{ value }`, `{ formula }` nor `{ link }`) still throw.

If the consumer cancels the stream (the user stops a download), the next `writeRow()`/`close()` rejects with the cancellation reason.

## Limits

- A single sheet must stay under 4 GiB compressed and uncompressed (ZIP data-descriptor limit). Larger exports: split across sheets. The archive as a whole may exceed 4 GiB (ZIP64 central directory).
- Strings are written inline rather than through a shared-string table: this keeps memory constant; deflate removes most of the size difference.
- Hyperlinks use the `HYPERLINK()` formula, which needs no per-sheet relationship file.

## Performance

1,000,000 rows × 10 cells: 5.4 s (≈184k rows/s), 52 MB of output, live heap flat at ~5 MB after GC (Apple M3 Max, Node 25). Memory does not grow with row count.

## Supported environments

Anything with `ReadableStream`, `TextEncoder` and `CompressionStream('deflate-raw')`:
Chrome/Edge 103+, Firefox 113+, Safari 16.4+, Node.js 20.12+, Deno, Bun, Cloudflare Workers.

## Roadmap

Streaming reader, optional shared strings, native hyperlinks, conditional formatting, data validation, comments, images, rich text, pluggable compressor for older engines.

## License

MIT
