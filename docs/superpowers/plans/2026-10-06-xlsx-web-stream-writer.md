# xlsx-web-stream v0.1 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Потоковый writer XLSX с постоянным потреблением памяти для браузера и Node.js, опубликованный в npm как `xlsx-web-stream`.

**Architecture:** Строки сериализуются в XML листа и буферизуются (~64 КиБ), байты идут через CRC32 и нативный `CompressionStream('deflate-raw')` в собственный потоковый ZIP-writer, который отдаёт результат в `ReadableStream<Uint8Array>`. Все операции с ZIP выстроены в одну промис-цепочку; служебные части пакета (`styles.xml`, `workbook.xml`, `[Content_Types].xml`) пишутся при закрытии книги.

**Tech Stack:** TypeScript 5.9 (ESM only, сборка `tsc`), Web Streams API, Vitest 4; dev-зависимости для тестов: `fflate` (независимая распаковка), `fast-xml-parser`, `exceljs` (чтение результата сторонней библиотекой), `@types/node`.

**Spec:** `docs/superpowers/specs/2026-10-06-xlsxstream-writer-design.md`

## Global Constraints

- Имя пакета `xlsx-web-stream`, версия `0.1.0`, лицензия MIT, `"type": "module"`, только ESM.
- Рантайм-зависимостей нет. В `src/` запрещены любые Node-специфичные API (`Buffer`, `node:*`), только Web-стандарты: `ReadableStream`, `CompressionStream`, `TextEncoder`, `DataView`.
- `engines.node >= 20.12.0` (первая версия Node 20 с `CompressionStream('deflate-raw')`).
- Код, комментарии, сообщения об ошибках и README — на английском (аудитория npm мировая).
- Строки пишутся inline (`t="inlineStr"`), гиперссылки — формулой `HYPERLINK()`, стили — только через handle из `wb.style()`.
- Даты: по умолчанию UTC-компоненты, опция `dates: 'local'`; формат по умолчанию `yyyy-mm-dd`.
- Ошибки валидации — синхронные `TypeError`/`RangeError`/`Error` с понятным текстом; `NaN`, `Infinity`, невалидные даты и даты до 1900-01-01 — ячейка `#NUM!`, не ошибка.
- Одна ZIP-запись ≤ 4 ГиБ (сжатая и несжатая) → иначе `RangeError`; центральный каталог — ZIP64 при необходимости.
- TDD: каждый модуль сначала получает падающий тест. Коммит после каждой задачи.
- Тесты импортируют исходники как `../src/<module>.js` (расширение `.js` обязательно при `module: NodeNext`; Vite разрешает его в `.ts`).

## File Structure

| Файл | Ответственность |
|---|---|
| `src/types.ts` | Публичные типы (раздел 3.1 спеки) |
| `src/xml.ts` | Экранирование, очистка недопустимых символов, XML-константы (декларация, namespaces) |
| `src/zip/crc32.ts` | Табличный инкрементальный CRC-32 |
| `src/sink.ts` | `ByteSink` и `StreamSink`: `ReadableStream` с backpressure и обработкой отмены |
| `src/zip/zip-writer.ts` | `ZipWriter`/`ZipEntry`: заголовки, data descriptor, центральный каталог, ZIP64, промис-цепочка |
| `src/cell.ts` | Адреса ячеек, диапазоны, серийные даты, `cellXml()` |
| `src/styles.ts` | `StyleHandle`, `StyleRegistry`, генерация `styles.xml` |
| `src/parts.ts` | `[Content_Types].xml`, `_rels/.rels`, `workbook.xml`, `workbook.xml.rels` |
| `src/sheet.ts` | `SheetWriter`: шапка листа, строки, буфер, хвост |
| `src/workbook.ts` | `WorkbookImpl`: оркестрация, валидация имён, закрытие, abort |
| `src/index.ts` | `createWorkbook()` и реэкспорт типов |
| `test/*.test.ts` | По файлу на модуль + `workbook.test.ts` (интеграция), `lifecycle.test.ts`, `backpressure.test.ts` |
| `test/helpers.ts` | `collect()`, `unzip()`, `parseXml()`, `MemorySink`, `MemoryEntry` |
| `bench/million-rows.mjs`, `examples/sample.mjs` | Бенчмарк и образец файла для ручной проверки в Excel |

Порядок задач следует зависимостям: types → xml → crc32 → sink → zip-writer → cell → styles → parts → sheet → workbook → lifecycle → backpressure → bench → docs/publish.

---

### Task 1: Project scaffolding

**Files:**
- Create: `package.json`, `tsconfig.json`, `tsconfig.build.json`, `vitest.config.ts`, `.gitignore`, `LICENSE`, `src/types.ts`, `src/index.ts`

**Interfaces:**
- Produces: все публичные типы из `src/types.ts` (ниже), которые используют все последующие задачи.

- [ ] **Step 1: Create package.json**

```json
{
  "name": "xlsx-web-stream",
  "version": "0.1.0",
  "description": "Streaming XLSX writer for browsers and Node.js: constant memory, zero dependencies, Web Streams API",
  "type": "module",
  "main": "./dist/index.js",
  "types": "./dist/index.d.ts",
  "exports": {
    ".": {
      "types": "./dist/index.d.ts",
      "default": "./dist/index.js"
    }
  },
  "files": ["dist"],
  "sideEffects": false,
  "engines": {
    "node": ">=20.12.0"
  },
  "scripts": {
    "build": "tsc -p tsconfig.build.json",
    "typecheck": "tsc -p tsconfig.json",
    "test": "vitest run",
    "test:watch": "vitest",
    "bench": "node bench/million-rows.mjs",
    "prepublishOnly": "npm run typecheck && npm test && npm run build"
  },
  "keywords": ["xlsx", "excel", "spreadsheet", "stream", "streaming", "web-streams", "browser", "export", "zip", "large-files"],
  "license": "MIT"
}
```

- [ ] **Step 2: Install dev dependencies**

```bash
npm install -D typescript@^5.9 vitest@^4 fflate@^0.8 fast-xml-parser@^5 exceljs@^4.4 @types/node@^22
```

Vitest 4 выбран намеренно: Vitest 5 требует Node ^22.12 || ^24, а Vitest 4 работает на Node 20/22/24+ — это совпадает с `engines` пакета и позволяет гонять CI на Node 20.

- [ ] **Step 3: Create tsconfig.json (typecheck of src + tests)**

```json
{
  "compilerOptions": {
    "target": "ES2022",
    "module": "NodeNext",
    "moduleResolution": "NodeNext",
    "lib": ["ES2022", "DOM", "DOM.Iterable"],
    "types": ["node"],
    "strict": true,
    "noUncheckedIndexedAccess": true,
    "noImplicitOverride": true,
    "verbatimModuleSyntax": true,
    "isolatedModules": true,
    "skipLibCheck": true,
    "noEmit": true
  },
  "include": ["src", "test", "vitest.config.ts"]
}
```

Если `tsc` сообщит о конфликте глобальных типов `ReadableStream`/`CompressionStream` между `DOM` и `@types/node`, убрать `"DOM", "DOM.Iterable"` из `lib` — `@types/node` ≥ 20 объявляет эти глобалы сам.

- [ ] **Step 4: Create tsconfig.build.json (emit dist from src only, no Node types)**

```json
{
  "extends": "./tsconfig.json",
  "compilerOptions": {
    "noEmit": false,
    "outDir": "dist",
    "rootDir": "src",
    "declaration": true,
    "declarationMap": true,
    "sourceMap": true,
    "types": []
  },
  "include": ["src"]
}
```

`"types": []` гарантирует, что код библиотеки не опирается на Node-типы.

- [ ] **Step 5: Create vitest.config.ts, .gitignore, LICENSE**

`vitest.config.ts`:
```ts
import { defineConfig } from 'vitest/config';

export default defineConfig({
  test: {
    include: ['test/**/*.test.ts'],
    passWithNoTests: true,
  },
});
```

`.gitignore`:
```
node_modules/
dist/
coverage/
*.xlsx
```

`LICENSE`:
```
MIT License

Copyright (c) 2026 xlsx-web-stream contributors

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

- [ ] **Step 6: Create src/types.ts (public types, verbatim from spec §3.1)**

```ts
/** Options accepted by `createWorkbook()`. */
export interface WorkbookOptions {
  /** Interpret `Date` values by their UTC components (default) or by local components. */
  dates?: 'utc' | 'local';
  /** Number format applied to `Date` cells whose style has no `numFmt`. Default `'yyyy-mm-dd'`. */
  dateFormat?: string;
}

/** `'#RRGGBB'` or `'RRGGBB'`, case-insensitive. */
export type Color = string;

export type BorderStyle = 'thin' | 'medium' | 'thick' | 'dashed' | 'dotted' | 'double' | 'hair';

export interface BorderSide {
  style: BorderStyle;
  color?: Color;
}

export type BorderSideSpec = BorderStyle | BorderSide;

export type HorizontalAlignment =
  | 'left'
  | 'center'
  | 'right'
  | 'fill'
  | 'justify'
  | 'centerContinuous'
  | 'distributed';

export type VerticalAlignment = 'top' | 'center' | 'bottom' | 'justify' | 'distributed';

export interface FontDefinition {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  /** Font size in points. Default 11. */
  size?: number;
  color?: Color;
  /** Font name. Default Calibri. */
  name?: string;
}

export interface BorderDefinition {
  top?: BorderSideSpec;
  right?: BorderSideSpec;
  bottom?: BorderSideSpec;
  left?: BorderSideSpec;
}

export interface AlignmentDefinition {
  horizontal?: HorizontalAlignment;
  vertical?: VerticalAlignment;
  wrap?: boolean;
}

export interface StyleDefinition {
  font?: FontDefinition;
  /** Solid fill color. */
  fill?: Color;
  /** One style for all four sides, or per-side settings. */
  border?: BorderStyle | BorderDefinition;
  align?: AlignmentDefinition;
  /** Excel number format code, e.g. `'#,##0.00'` or `'yyyy-mm-dd'`. */
  numFmt?: string;
}

/** Opaque handle returned by `workbook.style()`. Valid only within the workbook that created it. */
export interface Style {
  readonly id: number;
}

export type ScalarValue = string | number | boolean | Date | null | undefined;

export type CellValue =
  | ScalarValue
  | { value: ScalarValue; style?: Style }
  | { formula: string; value?: string | number | boolean; style?: Style }
  | { link: string; value?: string; style?: Style };

export interface ColumnOptions {
  /** Width in Excel character units (0–255). */
  width?: number;
  /** Default style for cells in this column. */
  style?: Style;
}

export interface SheetOptions {
  columns?: ColumnOptions[];
  /** Freeze the first `rows` rows and/or `columns` columns. */
  freeze?: { rows?: number; columns?: number };
}

export interface RowOptions {
  /** Default style for cells in this row (overrides column styles). */
  style?: Style;
  /** Row height in points (0–409). */
  height?: number;
}

export interface Sheet {
  readonly name: string;
  /** 1-based number of the next row `writeRow()` will produce. */
  readonly nextRow: number;
  writeRow(cells: readonly CellValue[], options?: RowOptions): Promise<void>;
  writeRows(rows: Iterable<readonly CellValue[]> | AsyncIterable<readonly CellValue[]>): Promise<void>;
  /** Merge a range such as `'A1:C1'`. May be called any time before the sheet is closed. */
  merge(range: string): void;
  /** Add an auto filter to a range such as `'A1:C1'`. At most one per sheet. */
  autoFilter(range: string): void;
  close(): Promise<void>;
}

export interface Workbook {
  /** The XLSX bytes. Pipe it anywhere a Web ReadableStream is accepted. */
  readonly readable: ReadableStream<Uint8Array>;
  style(definition: StyleDefinition): Style;
  /** Starts a new sheet. The previous sheet, if still open, is closed implicitly. */
  addSheet(name: string, options?: SheetOptions): Sheet;
  /** Closes the current sheet, writes workbook metadata and ends the stream. Idempotent. */
  close(): Promise<void>;
  /** Fails the stream and every pending or future operation with `reason`. */
  abort(reason?: unknown): void;
}
```

- [ ] **Step 7: Create a placeholder src/index.ts so the build has an entry point**

```ts
export type * from './types.js';
```

(Функция `createWorkbook` появится в Task 10.)

- [ ] **Step 8: Verify tooling**

Run: `npm run typecheck && npm test && npm run build && ls dist`
Expected: typecheck без ошибок; vitest сообщает «No test files found» и завершается успешно (`passWithNoTests`); в `dist/` лежат `index.js`, `index.d.ts`, `types.js`, `types.d.ts`.

- [ ] **Step 9: Commit**

```bash
git add package.json package-lock.json tsconfig.json tsconfig.build.json vitest.config.ts .gitignore LICENSE src/types.ts src/index.ts
git commit -m "chore: scaffold xlsx-web-stream package with TypeScript and Vitest"
```

---

### Task 2: XML escaping and constants

**Files:**
- Create: `src/xml.ts`
- Test: `test/xml.test.ts`

**Interfaces:**
- Produces:
  - `XML_DECL: string`, `NS_MAIN: string`, `NS_REL: string`
  - `sanitize(text: string): string` — удаляет символы, недопустимые в XML 1.0
  - `escapeText(text: string): string` — экранирует `& < >` (содержимое элементов)
  - `escapeAttr(text: string): string` — экранирует `& < > " '` (значения атрибутов)
  - `xmlText(text: string): string` = `escapeText(sanitize(text))`
  - `xmlAttr(text: string): string` = `escapeAttr(sanitize(text))`

- [ ] **Step 1: Write the failing test**

`test/xml.test.ts`:
```ts
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
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run test/xml.test.ts`
Expected: FAIL — `Failed to load url ../src/xml.js` (модуль не существует).

- [ ] **Step 3: Write the implementation**

`src/xml.ts`:
```ts
export const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
export const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
export const NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

// Characters outside the XML 1.0 Char production (except TAB, LF, CR). Excel refuses files that contain them.
const INVALID_CHARS = /[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/;
const INVALID_CHARS_GLOBAL = /[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g;
const TEXT_SPECIALS = /[&<>]/;
const TEXT_SPECIALS_GLOBAL = /[&<>]/g;
const ATTR_SPECIALS = /[&<>"']/;
const ATTR_SPECIALS_GLOBAL = /[&<>"']/g;

const REPLACEMENTS: Record<string, string> = {
  '&': '&amp;',
  '<': '&lt;',
  '>': '&gt;',
  '"': '&quot;',
  "'": '&apos;',
};

const replace = (ch: string): string => REPLACEMENTS[ch]!;

/** Removes characters that are not allowed in an XML 1.0 document. */
export function sanitize(text: string): string {
  return INVALID_CHARS.test(text) ? text.replace(INVALID_CHARS_GLOBAL, '') : text;
}

/** Escapes a string for use as element content. */
export function escapeText(text: string): string {
  return TEXT_SPECIALS.test(text) ? text.replace(TEXT_SPECIALS_GLOBAL, replace) : text;
}

/** Escapes a string for use inside a double- or single-quoted attribute value. */
export function escapeAttr(text: string): string {
  return ATTR_SPECIALS.test(text) ? text.replace(ATTR_SPECIALS_GLOBAL, replace) : text;
}

/** Sanitizes and escapes user-provided element content. */
export function xmlText(text: string): string {
  return escapeText(sanitize(text));
}

/** Sanitizes and escapes a user-provided attribute value. */
export function xmlAttr(text: string): string {
  return escapeAttr(sanitize(text));
}
```

Пара «тестовый regex без `g` + заменяющий с `g`» намеренная: `.test()` у глобального regex меняет `lastIndex` и даёт ложные результаты при повторных вызовах.

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run test/xml.test.ts`
Expected: PASS (7 tests).

- [ ] **Step 5: Commit**

```bash
git add src/xml.ts test/xml.test.ts
git commit -m "feat: XML escaping and sanitizing helpers"
```

---

### Task 3: CRC-32

**Files:**
- Create: `src/zip/crc32.ts`
- Test: `test/crc32.test.ts`

**Interfaces:**
- Produces: `crc32(data: Uint8Array, seed = 0): number` — беззнаковый CRC-32 (IEEE 802.3); для инкрементального расчёта передавать предыдущий результат как `seed`.

- [ ] **Step 1: Write the failing test**

`test/crc32.test.ts`:
```ts
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
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run test/crc32.test.ts`
Expected: FAIL — модуль `../src/zip/crc32.js` не найден.

- [ ] **Step 3: Write the implementation**

`src/zip/crc32.ts`:
```ts
const TABLE = new Int32Array(256);
for (let n = 0; n < 256; n++) {
  let c = n;
  for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
  TABLE[n] = c;
}

/**
 * CRC-32 (IEEE 802.3) as used by ZIP. Returns an unsigned 32-bit integer.
 * Pass a previous result as `seed` to continue a running checksum.
 */
export function crc32(data: Uint8Array, seed = 0): number {
  let c = ~seed;
  for (let i = 0; i < data.length; i++) {
    c = TABLE[(c ^ data[i]!) & 0xff]! ^ (c >>> 8);
  }
  return ~c >>> 0;
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run test/crc32.test.ts`
Expected: PASS (4 tests).

- [ ] **Step 5: Commit**

```bash
git add src/zip/crc32.ts test/crc32.test.ts
git commit -m "feat: incremental CRC-32"
```

---

### Task 4: StreamSink — ReadableStream with backpressure

**Files:**
- Create: `src/sink.ts`
- Test: `test/sink.test.ts`

**Interfaces:**
- Produces:
  - `interface ByteSink { write(chunk: Uint8Array): Promise<void> }` — промис разрешается, когда продюсеру можно писать дальше.
  - `class StreamSink implements ByteSink` с `readonly readable: ReadableStream<Uint8Array>`, `write(chunk)`, `close(): void`, `error(reason: unknown): void`, `onFailure: ((reason: unknown) => void) | null` (вызывается один раз при отмене потребителем или `error()`).
  - Конструктор: `new StreamSink(highWaterMark = 1 << 20)`.

- [ ] **Step 1: Write the failing test**

`test/sink.test.ts`:
```ts
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
    for await (const chunk of sink.readable) chunks.push(...chunk);
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
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run test/sink.test.ts`
Expected: FAIL — модуль `../src/sink.js` не найден.

- [ ] **Step 3: Write the implementation**

`src/sink.ts`:
```ts
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
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run test/sink.test.ts`
Expected: PASS (7 tests).

- [ ] **Step 5: Commit**

```bash
git add src/sink.ts test/sink.test.ts
git commit -m "feat: StreamSink with backpressure and cancellation"
```

---

### Task 5: Streaming ZIP writer

**Files:**
- Create: `src/zip/zip-writer.ts`
- Test: `test/zip-writer.test.ts`, `test/helpers.ts`

**Interfaces:**
- Consumes: `crc32` (Task 3), `ByteSink` (Task 4).
- Produces:
  - `interface ZipEntry { write(data: Uint8Array): Promise<void>; close(): Promise<void> }`
  - `class ZipWriter { constructor(sink: ByteSink, options?: { entrySizeLimit?: number }); beginEntry(name: string): ZipEntry; finish(): Promise<void>; fail(reason: unknown): void }`
  - Все операции выстроены в одну цепочку: `beginEntry` можно вызывать сразу после `close()` предыдущей записи, не дожидаясь её промиса. После первой ошибки все последующие операции отклоняются той же причиной.
  - Экспортируемые для тестов билдеры: `localFileHeader`, `dataDescriptor`, `centralDirectoryHeader`, `endOfCentralDirectory`, `zip64EndOfCentralDirectory`, `zip64Locator` и тип `EntryRecord`.

- [ ] **Step 1: Create shared test helpers**

`test/helpers.ts`:
```ts
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
```

- [ ] **Step 2: Write the failing test**

`test/zip-writer.test.ts`:
```ts
import { unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { crc32 } from '../src/zip/crc32.js';
import {
  centralDirectoryHeader,
  endOfCentralDirectory,
  zip64EndOfCentralDirectory,
  zip64Locator,
  ZipWriter,
  type EntryRecord,
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

  it('propagates sink failures to every operation', async () => {
    const zip = new ZipWriter(new FailingSink(new Error('disk full')));
    const entry = zip.beginEntry('a');
    await expect(entry.write(encode('x'))).rejects.toThrow('disk full');
    await expect(entry.close()).rejects.toThrow('disk full');
    await expect(zip.finish()).rejects.toThrow('disk full');
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
```

- [ ] **Step 3: Run test to verify it fails**

Run: `npx vitest run test/zip-writer.test.ts`
Expected: FAIL — модуль `../src/zip/zip-writer.js` не найден.

- [ ] **Step 4: Write the implementation**

`src/zip/zip-writer.ts`:
```ts
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

  /** @internal Called by an entry once its data descriptor is written. */
  complete(record: EntryRecord): void {
    this.entries.push(record);
    this.current = null;
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
      this.zip.complete(this.record);
    });
    return this.closePromise;
  }

  private async start(): Promise<void> {
    this.record.offset = this.zip.position;
    await this.zip.emit(localFileHeader(this.record));
    const compressor = new CompressionStream('deflate-raw');
    this.writer = compressor.writable.getWriter();
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
```

- [ ] **Step 5: Run test to verify it passes**

Run: `npx vitest run test/zip-writer.test.ts`
Expected: PASS (8 tests). Если тест «sequences entries» зависает, проверить, что `beginEntry` допускает `current.closing === true`, а `Entry` ставит `start()` в очередь уже в конструкторе.

- [ ] **Step 6: Commit**

```bash
git add src/zip/zip-writer.ts test/zip-writer.test.ts test/helpers.ts
git commit -m "feat: streaming ZIP writer with data descriptors and ZIP64 central directory"
```

---

### Task 6: Cell addressing, dates and cell XML

**Files:**
- Create: `src/cell.ts`
- Test: `test/cell.test.ts`

**Interfaces:**
- Consumes: `CellValue`, `Style` (Task 1); `escapeText`, `sanitize` (Task 2).
- Produces:
  - `MAX_ROWS = 1_048_576`, `MAX_COLS = 16_384`, `MAX_TEXT_LENGTH = 32_767`
  - `type DateMode = 'utc' | 'local'`
  - `columnLetter(col: number): string`, `cellRef(col: number, row: number): string` (1-based)
  - `parseCellRef(ref: string): { col: number; row: number } | null`
  - `interface Range { c1: number; r1: number; c2: number; r2: number }`
  - `parseRange(range: string): Range` (нормализует порядок углов, `TypeError` при ошибке), `formatRange(r: Range): string`, `formatAbsoluteRange(r: Range): string` (`$A$1:$C$3`)
  - `dateToSerial(date: Date, mode: DateMode): number | null`
  - `interface CellContext { readonly dates: DateMode; styleId(style: Style | undefined): number | undefined; dateStyle(styleId: number): number; linkStyle(): number }`
  - `cellXml(col: number, row: number, value: CellValue, inherited: number, ctx: CellContext): string` — `''`, если ячейку писать не нужно.

- [ ] **Step 1: Write the failing test**

`test/cell.test.ts`:
```ts
import { describe, expect, it } from 'vitest';
import {
  cellRef,
  cellXml,
  columnLetter,
  dateToSerial,
  formatAbsoluteRange,
  formatRange,
  parseCellRef,
  parseRange,
  type CellContext,
} from '../src/cell.js';
import type { Style } from '../src/types.js';

const style = (id: number): Style => ({ id });

/** Fake context: handles resolve to their id, the date variant of style 0 is 99, hyperlink style is 7. */
const ctx: CellContext = {
  dates: 'utc',
  styleId: (s) => s?.id,
  dateStyle: (s) => (s === 0 ? 99 : s),
  linkStyle: () => 7,
};

describe('column letters and references', () => {
  it('converts 1-based column indexes to letters', () => {
    expect(columnLetter(1)).toBe('A');
    expect(columnLetter(26)).toBe('Z');
    expect(columnLetter(27)).toBe('AA');
    expect(columnLetter(52)).toBe('AZ');
    expect(columnLetter(53)).toBe('BA');
    expect(columnLetter(702)).toBe('ZZ');
    expect(columnLetter(703)).toBe('AAA');
    expect(columnLetter(16384)).toBe('XFD');
  });

  it('builds and parses A1 references', () => {
    expect(cellRef(3, 12)).toBe('C12');
    expect(parseCellRef('C12')).toEqual({ col: 3, row: 12 });
    expect(parseCellRef('xfd1048576')).toEqual({ col: 16384, row: 1048576 });
    expect(parseCellRef('XFE1')).toBeNull();
    expect(parseCellRef('A0')).toBeNull();
    expect(parseCellRef('A1048577')).toBeNull();
    expect(parseCellRef('1A')).toBeNull();
    expect(parseCellRef('')).toBeNull();
  });
});

describe('ranges', () => {
  it('normalizes corner order', () => {
    expect(parseRange('C3:A1')).toEqual({ c1: 1, r1: 1, c2: 3, r2: 3 });
    expect(formatRange(parseRange('C3:A1'))).toBe('A1:C3');
    expect(formatAbsoluteRange(parseRange('A1:C3'))).toBe('$A$1:$C$3');
  });

  it('rejects malformed ranges', () => {
    for (const bad of ['A1', 'A1:B', 'A1:B2:C3', 'A1-B2', '', 'XFE1:XFE2']) {
      expect(() => parseRange(bad), bad).toThrow(TypeError);
    }
    expect(() => parseRange(42 as unknown as string)).toThrow(TypeError);
  });
});

describe('dateToSerial', () => {
  it('matches Excel serial numbers including the 1900 leap-year bug', () => {
    expect(dateToSerial(new Date(Date.UTC(1900, 0, 1)), 'utc')).toBe(1);
    expect(dateToSerial(new Date(Date.UTC(1900, 1, 28)), 'utc')).toBe(59);
    expect(dateToSerial(new Date(Date.UTC(1900, 2, 1)), 'utc')).toBe(61);
    expect(dateToSerial(new Date(Date.UTC(2000, 0, 1)), 'utc')).toBe(36526);
    expect(dateToSerial(new Date(Date.UTC(2026, 9, 6)), 'utc')).toBe(46301);
    expect(dateToSerial(new Date(Date.UTC(2026, 9, 6, 12)), 'utc')).toBe(46301.5);
  });

  it('uses local components in local mode', () => {
    expect(dateToSerial(new Date(2026, 9, 6, 12), 'local')).toBe(46301.5);
    expect(dateToSerial(new Date(2026, 9, 6, 6), 'local')).toBe(46301.25);
  });

  it('returns null for dates Excel cannot represent', () => {
    expect(dateToSerial(new Date(NaN), 'utc')).toBeNull();
    expect(dateToSerial(new Date(Date.UTC(1899, 11, 31)), 'utc')).toBeNull();
    const ancient = new Date(2000, 0, 1);
    ancient.setFullYear(50);
    expect(dateToSerial(ancient, 'local')).toBeNull();
  });
});

describe('cellXml', () => {
  it('writes strings inline, preserving significant whitespace', () => {
    expect(cellXml(1, 1, 'hi', 0, ctx)).toBe('<c r="A1" t="inlineStr"><is><t>hi</t></is></c>');
    expect(cellXml(2, 3, ' lead', 0, ctx)).toBe('<c r="B3" t="inlineStr"><is><t xml:space="preserve"> lead</t></is></c>');
    expect(cellXml(1, 1, 'a<b>&', 0, ctx)).toBe('<c r="A1" t="inlineStr"><is><t>a&lt;b&gt;&amp;</t></is></c>');
    expect(cellXml(1, 1, 'x\u0000y', 0, ctx)).toBe('<c r="A1" t="inlineStr"><is><t>xy</t></is></c>');
  });

  it('writes numbers, booleans and non-finite numbers', () => {
    expect(cellXml(1, 1, 12.5, 0, ctx)).toBe('<c r="A1"><v>12.5</v></c>');
    expect(cellXml(1, 1, -0, 0, ctx)).toBe('<c r="A1"><v>0</v></c>');
    expect(cellXml(1, 1, 1e21, 0, ctx)).toBe('<c r="A1"><v>1e+21</v></c>');
    expect(cellXml(1, 1, true, 0, ctx)).toBe('<c r="A1" t="b"><v>1</v></c>');
    expect(cellXml(1, 1, false, 0, ctx)).toBe('<c r="A1" t="b"><v>0</v></c>');
    expect(cellXml(1, 1, NaN, 0, ctx)).toBe('<c r="A1" t="e"><v>#NUM!</v></c>');
    expect(cellXml(1, 1, Infinity, 3, ctx)).toBe('<c r="A1" s="3" t="e"><v>#NUM!</v></c>');
  });

  it('writes dates with the date variant of the resolved style', () => {
    expect(cellXml(1, 1, new Date(Date.UTC(2026, 9, 6)), 0, ctx)).toBe('<c r="A1" s="99"><v>46301</v></c>');
    expect(cellXml(1, 1, new Date(Date.UTC(2026, 9, 6)), 5, ctx)).toBe('<c r="A1" s="5"><v>46301</v></c>');
    expect(cellXml(1, 1, new Date(NaN), 0, ctx)).toBe('<c r="A1" t="e"><v>#NUM!</v></c>');
  });

  it('omits empty cells unless they carry a style', () => {
    expect(cellXml(1, 1, null, 0, ctx)).toBe('');
    expect(cellXml(1, 1, undefined, 0, ctx)).toBe('');
    expect(cellXml(1, 1, '', 0, ctx)).toBe('');
    expect(cellXml(1, 1, null, 4, ctx)).toBe('<c r="A1" s="4"/>');
    expect(cellXml(1, 1, { value: null, style: style(6) }, 0, ctx)).toBe('<c r="A1" s="6"/>');
    expect(cellXml(1, 1, { style: style(6) } as never, 0, ctx)).toBe('<c r="A1" s="6"/>');
  });

  it('lets a cell style override the inherited one', () => {
    expect(cellXml(1, 1, { value: 1, style: style(2) }, 9, ctx)).toBe('<c r="A1" s="2"><v>1</v></c>');
    expect(cellXml(1, 1, { value: 'x' }, 9, ctx)).toBe('<c r="A1" s="9" t="inlineStr"><is><t>x</t></is></c>');
  });

  it('writes formulas with optional cached values', () => {
    expect(cellXml(1, 1, { formula: '=SUM(A1:A2)' }, 0, ctx)).toBe('<c r="A1"><f>SUM(A1:A2)</f></c>');
    expect(cellXml(1, 1, { formula: 'A1*2', value: 4 }, 0, ctx)).toBe('<c r="A1"><f>A1*2</f><v>4</v></c>');
    expect(cellXml(1, 1, { formula: 'A1&"x"', value: 'yx' }, 0, ctx)).toBe('<c r="A1" t="str"><f>A1&amp;"x"</f><v>yx</v></c>');
    expect(cellXml(1, 1, { formula: 'A1>0', value: true }, 0, ctx)).toBe('<c r="A1" t="b"><f>A1&gt;0</f><v>1</v></c>');
    expect(cellXml(1, 1, { formula: '1/0', value: NaN }, 0, ctx)).toBe('<c r="A1" t="e"><f>1/0</f><v>#NUM!</v></c>');
  });

  it('writes hyperlinks as HYPERLINK formulas with the built-in link style', () => {
    expect(cellXml(1, 1, { link: 'https://x.com/?a=1&b=2', value: 'Site' }, 0, ctx)).toBe(
      '<c r="A1" s="7" t="str"><f>HYPERLINK("https://x.com/?a=1&amp;b=2","Site")</f><v>Site</v></c>',
    );
    expect(cellXml(1, 1, { link: 'https://x.com' }, 0, ctx)).toBe(
      '<c r="A1" s="7" t="str"><f>HYPERLINK("https://x.com","https://x.com")</f><v>https://x.com</v></c>',
    );
    expect(cellXml(1, 1, { link: 'https://x.com', value: 'say "hi"' }, 0, ctx)).toContain('HYPERLINK("https://x.com","say ""hi""")');
    expect(cellXml(1, 1, { link: 'https://x.com', value: 'x', style: style(3) }, 0, ctx)).toContain(' s="3" ');
    expect(cellXml(1, 1, { link: 'https://x.com', value: 'x' }, 5, ctx)).toContain(' s="5" ');
  });

  it('rejects values it cannot represent', () => {
    expect(() => cellXml(1, 1, 'x'.repeat(32_768), 0, ctx)).toThrow(RangeError);
    expect(() => cellXml(1, 1, {} as never, 0, ctx)).toThrow(TypeError);
    expect(() => cellXml(1, 1, 10n as never, 0, ctx)).toThrow(TypeError);
    expect(() => cellXml(1, 1, { formula: '' } as never, 0, ctx)).toThrow(TypeError);
    expect(() => cellXml(1, 1, { link: '' } as never, 0, ctx)).toThrow(TypeError);
    expect(() => cellXml(1, 1, { formula: 'A1', value: new Date() } as never, 0, ctx)).toThrow(TypeError);
    expect(() => cellXml(1, 1, [1, 2] as never, 0, ctx)).toThrow(TypeError);
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run test/cell.test.ts`
Expected: FAIL — модуль `../src/cell.js` не найден.

- [ ] **Step 3: Write the implementation**

`src/cell.ts`:
```ts
import type { CellValue, Style } from './types.js';
import { escapeText, sanitize } from './xml.js';

export const MAX_ROWS = 1_048_576;
export const MAX_COLS = 16_384;
export const MAX_TEXT_LENGTH = 32_767;

export type DateMode = 'utc' | 'local';

const letterCache: string[] = [];

/** 1-based column index → letters: 1 → "A", 27 → "AA", 16384 → "XFD". */
export function columnLetter(col: number): string {
  const cached = letterCache[col];
  if (cached !== undefined) return cached;
  let n = col;
  let out = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    out = String.fromCharCode(65 + r) + out;
    n = (n - 1 - r) / 26;
  }
  letterCache[col] = out;
  return out;
}

export function cellRef(col: number, row: number): string {
  return columnLetter(col) + row;
}

const CELL_REF = /^([A-Z]{1,3})([0-9]{1,7})$/;

export function parseCellRef(ref: string): { col: number; row: number } | null {
  const match = CELL_REF.exec(ref.toUpperCase());
  if (!match) return null;
  let col = 0;
  for (const ch of match[1]!) col = col * 26 + ch.charCodeAt(0) - 64;
  const row = Number(match[2]);
  if (col > MAX_COLS || row < 1 || row > MAX_ROWS) return null;
  return { col, row };
}

export interface Range {
  c1: number;
  r1: number;
  c2: number;
  r2: number;
}

/** Parses "A1:C3" (corners in any order) into a normalized range. */
export function parseRange(range: string): Range {
  const parts = typeof range === 'string' ? range.split(':') : [];
  const a = parts.length === 2 ? parseCellRef(parts[0]!) : null;
  const b = parts.length === 2 ? parseCellRef(parts[1]!) : null;
  if (!a || !b) throw new TypeError(`Invalid range ${JSON.stringify(range)}: expected the form "A1:C3"`);
  return {
    c1: Math.min(a.col, b.col),
    r1: Math.min(a.row, b.row),
    c2: Math.max(a.col, b.col),
    r2: Math.max(a.row, b.row),
  };
}

export function formatRange(r: Range): string {
  return `${cellRef(r.c1, r.r1)}:${cellRef(r.c2, r.r2)}`;
}

/** "$A$1:$C$3" — the form used in defined names. */
export function formatAbsoluteRange(r: Range): string {
  return `$${columnLetter(r.c1)}$${r.r1}:$${columnLetter(r.c2)}$${r.r2}`;
}

const EXCEL_EPOCH = Date.UTC(1899, 11, 30);
const MS_PER_DAY = 86_400_000;

/** Excel serial number of a date, or null when Excel cannot represent it (invalid or before 1900-01-01). */
export function dateToSerial(date: Date, mode: DateMode): number | null {
  const time = date.getTime();
  if (Number.isNaN(time)) return null;
  let utc = time;
  if (mode === 'local') {
    if (date.getFullYear() < 100) return null; // Date.UTC would map years 0–99 to 1900–1999
    utc = Date.UTC(
      date.getFullYear(),
      date.getMonth(),
      date.getDate(),
      date.getHours(),
      date.getMinutes(),
      date.getSeconds(),
      date.getMilliseconds(),
    );
  }
  let serial = (utc - EXCEL_EPOCH) / MS_PER_DAY;
  if (serial < 2) return null; // before 1900-01-01
  if (serial < 61) serial -= 1; // Excel treats 1900 as a leap year; skip its phantom Feb 29
  return serial;
}

export interface CellContext {
  readonly dates: DateMode;
  /** Resolves a style handle to its xf index; throws if the handle belongs to another workbook. */
  styleId(style: Style | undefined): number | undefined;
  /** xf index to use for a Date value whose resolved style has no number format. */
  dateStyle(styleId: number): number;
  /** xf index of the built-in hyperlink style. */
  linkStyle(): number;
}

interface RichCell {
  value?: unknown;
  style?: Style;
  formula?: unknown;
  link?: unknown;
}

const ERROR_VALUE = '#NUM!';
const LEADING_OR_TRAILING_SPACE = /^\s|\s$/;

function sAttr(s: number): string {
  return s === 0 ? '' : ` s="${s}"`;
}

function errorCell(ref: string, s: number): string {
  return `<c r="${ref}"${sAttr(s)} t="e"><v>${ERROR_VALUE}</v></c>`;
}

function text(value: string, ref: string): string {
  if (value.length > MAX_TEXT_LENGTH) {
    throw new RangeError(`Text in cell ${ref} is ${value.length} characters long; Excel allows at most ${MAX_TEXT_LENGTH}`);
  }
  return escapeText(sanitize(value));
}

/** Quotes a string literal for use inside a formula: "say ""hi""". */
function formulaString(value: string): string {
  return `"${value.replace(/"/g, '""')}"`;
}

/**
 * Serializes one cell. `inherited` is the xf index coming from the row or column (0 = none).
 * Returns '' when the cell is empty and unstyled.
 */
export function cellXml(col: number, row: number, value: CellValue, inherited: number, ctx: CellContext): string {
  const ref = cellRef(col, row);
  let v: unknown = value;
  let s = inherited;
  let ownStyle = false;
  let formula: string | undefined;
  let link: string | undefined;

  if (v !== null && typeof v === 'object' && !(v instanceof Date)) {
    if (Array.isArray(v)) throw new TypeError(`Unsupported value in cell ${ref}: array`);
    const rich = v as RichCell;
    const id = ctx.styleId(rich.style);
    if (id !== undefined) {
      s = id;
      ownStyle = true;
    }
    if (rich.formula !== undefined) {
      if (typeof rich.formula !== 'string' || rich.formula.trim() === '') {
        throw new TypeError(`Formula in cell ${ref} must be a non-empty string`);
      }
      formula = rich.formula;
    } else if (rich.link !== undefined) {
      if (typeof rich.link !== 'string' || rich.link === '') {
        throw new TypeError(`Link in cell ${ref} must be a non-empty string`);
      }
      link = rich.link;
    } else if (!('value' in rich) && !('style' in rich)) {
      throw new TypeError(`Unsupported value in cell ${ref}: expected { value, style }, { formula } or { link }`);
    }
    v = rich.value;
  }

  if (link !== undefined) {
    const label = v === undefined || v === null ? link : String(v);
    if (!ownStyle && s === 0) s = ctx.linkStyle();
    const f = `HYPERLINK(${formulaString(link)},${formulaString(label)})`;
    return `<c r="${ref}"${sAttr(s)} t="str"><f>${text(f, ref)}</f><v>${text(label, ref)}</v></c>`;
  }

  if (formula !== undefined) {
    const f = `<f>${text(formula.startsWith('=') ? formula.slice(1) : formula, ref)}</f>`;
    if (v === undefined || v === null) return `<c r="${ref}"${sAttr(s)}>${f}</c>`;
    if (typeof v === 'number') {
      return Number.isFinite(v)
        ? `<c r="${ref}"${sAttr(s)}>${f}<v>${v}</v></c>`
        : `<c r="${ref}"${sAttr(s)} t="e">${f}<v>${ERROR_VALUE}</v></c>`;
    }
    if (typeof v === 'boolean') return `<c r="${ref}"${sAttr(s)} t="b">${f}<v>${v ? 1 : 0}</v></c>`;
    if (typeof v === 'string') return `<c r="${ref}"${sAttr(s)} t="str">${f}<v>${text(v, ref)}</v></c>`;
    throw new TypeError(`Cached formula value in cell ${ref} must be a string, number or boolean`);
  }

  switch (typeof v) {
    case 'string':
      if (v === '') break;
      return `<c r="${ref}"${sAttr(s)} t="inlineStr"><is><t${LEADING_OR_TRAILING_SPACE.test(v) ? ' xml:space="preserve"' : ''}>${text(v, ref)}</t></is></c>`;
    case 'number':
      return Number.isFinite(v) ? `<c r="${ref}"${sAttr(s)}><v>${v}</v></c>` : errorCell(ref, s);
    case 'boolean':
      return `<c r="${ref}"${sAttr(s)} t="b"><v>${v ? 1 : 0}</v></c>`;
    case 'undefined':
      break;
    case 'object':
      if (v === null) break;
      if (v instanceof Date) {
        const serial = dateToSerial(v, ctx.dates);
        return serial === null ? errorCell(ref, s) : `<c r="${ref}" s="${ctx.dateStyle(s)}"><v>${serial}</v></c>`;
      }
      throw new TypeError(`Unsupported value in cell ${ref}: ${Object.prototype.toString.call(v)}`);
    default:
      throw new TypeError(`Unsupported value in cell ${ref}: ${typeof v}`);
  }
  return s === 0 ? '' : `<c r="${ref}" s="${s}"/>`;
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run test/cell.test.ts`
Expected: PASS (14 tests).

- [ ] **Step 5: Commit**

```bash
git add src/cell.ts test/cell.test.ts
git commit -m "feat: cell references, Excel date serials and cell XML serialization"
```

---

### Task 7: Style registry and styles.xml

**Files:**
- Create: `src/styles.ts`
- Test: `test/styles.test.ts`

**Interfaces:**
- Consumes: `Style`, `StyleDefinition`, `FontDefinition`, `BorderDefinition`, `AlignmentDefinition`, `BorderSideSpec` (Task 1); `escapeAttr`, `xmlAttr`, `XML_DECL`, `NS_MAIN` (Task 2).
- Produces:
  - `class StyleHandle implements Style { constructor(readonly id: number, readonly owner: object) }`
  - `class StyleRegistry { constructor(owner: object, dateFormat: string); register(def: StyleDefinition): StyleHandle; idOf(style: Style | undefined): number | undefined; dateVariant(xfId: number): number; linkStyle(): number; toXml(): string }`
  - `normalizeColor(color: unknown, what: string): string` → `'FFRRGGBB'`

- [ ] **Step 1: Write the failing test**

`test/styles.test.ts`:
```ts
import { describe, expect, it } from 'vitest';
import { normalizeColor, StyleHandle, StyleRegistry } from '../src/styles.js';
import { parseXml } from './helpers.js';

const owner = {};
const registry = (): StyleRegistry => new StyleRegistry(owner, 'yyyy-mm-dd');

describe('normalizeColor', () => {
  it('accepts #RRGGBB and RRGGBB in any case and returns ARGB', () => {
    expect(normalizeColor('#eeeeee', 'fill')).toBe('FFEEEEEE');
    expect(normalizeColor('0563C1', 'fill')).toBe('FF0563C1');
  });

  it('rejects anything else', () => {
    for (const bad of ['#fff', 'red', '#GGGGGG', '', 123, null]) {
      expect(() => normalizeColor(bad, 'fill'), String(bad)).toThrow(TypeError);
    }
  });
});

describe('StyleRegistry.register', () => {
  it('returns handles owned by the registry owner and deduplicates equal definitions', () => {
    const r = registry();
    const a = r.register({ font: { bold: true }, fill: '#EEEEEE' });
    const b = r.register({ fill: '#eeeeee', font: { bold: true } });
    const c = r.register({ font: { bold: true } });
    expect(a).toBeInstanceOf(StyleHandle);
    expect(a.owner).toBe(owner);
    expect(a.id).toBe(b.id);
    expect(a.id).not.toBe(c.id);
    expect(a.id).toBeGreaterThan(0);
  });

  it('maps an empty definition and default-only fonts to xf 0', () => {
    const r = registry();
    expect(r.register({}).id).toBe(0);
    expect(r.register({ font: {} }).id).toBe(0);
    expect(r.register({ numFmt: 'General' }).id).toBe(0);
  });

  it('rejects unknown keys and invalid values', () => {
    const r = registry();
    expect(() => r.register({ numFormat: '0' } as never)).toThrow(TypeError);
    expect(() => r.register({ font: { weight: 'bold' } } as never)).toThrow(TypeError);
    expect(() => r.register({ font: { bold: 'yes' } } as never)).toThrow(TypeError);
    expect(() => r.register({ font: { size: -1 } })).toThrow(TypeError);
    expect(() => r.register({ fill: 'red' })).toThrow(TypeError);
    expect(() => r.register({ border: 'bold' as never })).toThrow(TypeError);
    expect(() => r.register({ border: { top: { style: 'thin', color: 'nope' } } })).toThrow(TypeError);
    expect(() => r.register({ border: { diagonal: 'thin' } } as never)).toThrow(TypeError);
    expect(() => r.register({ align: { horizontal: 'middle' as never } })).toThrow(TypeError);
    expect(() => r.register({ align: { vertical: 'up' as never } })).toThrow(TypeError);
    expect(() => r.register({ numFmt: '' })).toThrow(TypeError);
    expect(() => r.register(null as never)).toThrow(TypeError);
  });
});

describe('StyleRegistry.idOf', () => {
  it('resolves own handles, passes undefined through and rejects foreign handles', () => {
    const r = registry();
    const h = r.register({ font: { italic: true } });
    expect(r.idOf(h)).toBe(h.id);
    expect(r.idOf(undefined)).toBeUndefined();
    expect(() => r.idOf({ id: 1 })).toThrow(TypeError);
    expect(() => r.idOf(new StyleHandle(1, {}))).toThrow(TypeError);
  });
});

describe('date variants and link style', () => {
  it('creates a date variant once per style and keeps styles that already have a numFmt', () => {
    const r = registry();
    const bold = r.register({ font: { bold: true } });
    const money = r.register({ numFmt: '#,##0.00' });
    const dateDefault = r.dateVariant(0);
    const dateBold = r.dateVariant(bold.id);
    expect(dateDefault).not.toBe(0);
    expect(dateBold).not.toBe(bold.id);
    expect(dateBold).not.toBe(dateDefault);
    expect(r.dateVariant(0)).toBe(dateDefault);
    expect(r.dateVariant(money.id)).toBe(money.id);
    expect(r.register({ numFmt: 'yyyy-mm-dd' }).id).toBe(dateDefault);
  });

  it('registers the hyperlink style lazily and only once', () => {
    const r = registry();
    const id = r.linkStyle();
    expect(id).toBeGreaterThan(0);
    expect(r.linkStyle()).toBe(id);
    expect(r.register({ font: { color: '#0563C1', underline: true } }).id).toBe(id);
  });
});

describe('StyleRegistry.toXml', () => {
  it('emits the mandatory default fills, fonts and borders first', () => {
    const xml = registry().toXml();
    const doc = parseXml(xml);
    expect(xml.startsWith('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')).toBe(true);
    expect(doc.styleSheet.fills.count).toBe('2');
    expect(doc.styleSheet.fills.fill[0].patternFill.patternType).toBe('none');
    expect(doc.styleSheet.fills.fill[1].patternFill.patternType).toBe('gray125');
    expect(doc.styleSheet.fonts.count).toBe('1');
    expect(doc.styleSheet.borders.count).toBe('1');
    expect(doc.styleSheet.cellXfs.count).toBe('1');
    expect(doc.styleSheet.cellStyles.cellStyle.name).toBe('Normal');
    expect(xml).not.toContain('<numFmts');
  });

  it('serializes registered components in the order Excel expects', () => {
    const r = registry();
    const h = r.register({
      font: { bold: true, italic: true, underline: true, strike: true, size: 12, color: '#FF0000', name: 'Arial' },
      fill: '#EEEEEE',
      border: { top: 'thin', bottom: { style: 'double', color: '#00FF00' } },
      align: { horizontal: 'center', vertical: 'top', wrap: true },
      numFmt: '#,##0.00',
    });
    const xml = r.toXml();
    expect(xml).toContain('<numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0.00"/></numFmts>');
    expect(xml).toContain('<font><b/><i/><strike/><u/><sz val="12"/><color rgb="FFFF0000"/><name val="Arial"/><family val="2"/></font>');
    expect(xml).toContain('<fill><patternFill patternType="solid"><fgColor rgb="FFEEEEEE"/><bgColor indexed="64"/></patternFill></fill>');
    expect(xml).toContain('<border><left/><right/><top style="thin"/><bottom style="double"><color rgb="FF00FF00"/></bottom><diagonal/></border>');
    expect(xml).toContain(
      `<xf numFmtId="164" fontId="1" fillId="2" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="top" wrapText="1"/></xf>`,
    );
    expect(h.id).toBe(1);
  });

  it('applies a border style string to all four sides and escapes format codes', () => {
    const r = registry();
    r.register({ border: 'medium', numFmt: '"<"0' });
    const xml = r.toXml();
    expect(xml).toContain('<border><left style="medium"/><right style="medium"/><top style="medium"/><bottom style="medium"/><diagonal/></border>');
    expect(xml).toContain('formatCode="&quot;&lt;&quot;0"');
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run test/styles.test.ts`
Expected: FAIL — модуль `../src/styles.js` не найден.

- [ ] **Step 3: Write the implementation**

`src/styles.ts`:
```ts
import type {
  AlignmentDefinition,
  BorderDefinition,
  BorderSideSpec,
  FontDefinition,
  Style,
  StyleDefinition,
} from './types.js';
import { NS_MAIN, XML_DECL, xmlAttr } from './xml.js';

const BORDER_STYLES: ReadonlySet<string> = new Set(['thin', 'medium', 'thick', 'dashed', 'dotted', 'double', 'hair']);
const HORIZONTAL: ReadonlySet<string> = new Set(['left', 'center', 'right', 'fill', 'justify', 'centerContinuous', 'distributed']);
const VERTICAL: ReadonlySet<string> = new Set(['top', 'center', 'bottom', 'justify', 'distributed']);
const STYLE_KEYS: ReadonlySet<string> = new Set(['font', 'fill', 'border', 'align', 'numFmt']);
const FONT_KEYS: ReadonlySet<string> = new Set(['bold', 'italic', 'underline', 'strike', 'size', 'color', 'name']);
const BORDER_KEYS: ReadonlySet<string> = new Set(['top', 'right', 'bottom', 'left']);
const ALIGN_KEYS: ReadonlySet<string> = new Set(['horizontal', 'vertical', 'wrap']);
const SIDES = ['left', 'right', 'top', 'bottom'] as const;

const FIRST_CUSTOM_NUMFMT = 164;
const DEFAULT_FONT = '<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>';
const DEFAULT_FILLS = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
const DEFAULT_BORDER = '<border><left/><right/><top/><bottom/><diagonal/></border>';
const LINK_COLOR = '#0563C1';

interface Xf {
  numFmtId: number;
  fontId: number;
  fillId: number;
  borderId: number;
  alignment: string;
}

const DEFAULT_XF: Xf = { numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, alignment: '' };

/** Converts '#RRGGBB' / 'RRGGBB' to the ARGB form used in styles.xml. */
export function normalizeColor(color: unknown, what: string): string {
  if (typeof color !== 'string' || !/^#?[0-9a-fA-F]{6}$/.test(color)) {
    throw new TypeError(`Invalid ${what} ${JSON.stringify(color)}: expected "#RRGGBB"`);
  }
  return `FF${(color.startsWith('#') ? color.slice(1) : color).toUpperCase()}`;
}

/** Opaque handle returned by `workbook.style()`. */
export class StyleHandle implements Style {
  constructor(
    readonly id: number,
    readonly owner: object,
  ) {}
}

function assertObject(value: unknown, what: string): asserts value is Record<string, unknown> {
  if (value === null || typeof value !== 'object' || Array.isArray(value)) throw new TypeError(`${what} must be an object`);
}

function assertKeys(value: object, allowed: ReadonlySet<string>, what: string): void {
  for (const key of Object.keys(value)) {
    if (!allowed.has(key)) throw new TypeError(`Unknown ${what} property "${key}"`);
  }
}

function assertOptionalBoolean(value: unknown, what: string): void {
  if (value !== undefined && typeof value !== 'boolean') throw new TypeError(`${what} must be a boolean`);
}

function fontXml(font: FontDefinition): string {
  assertObject(font, 'font');
  assertKeys(font, FONT_KEYS, 'font');
  for (const key of ['bold', 'italic', 'underline', 'strike'] as const) assertOptionalBoolean(font[key], `font.${key}`);
  if (font.size !== undefined && !(typeof font.size === 'number' && font.size > 0 && font.size <= 409)) {
    throw new TypeError('font.size must be a number between 0 and 409');
  }
  if (font.name !== undefined && (typeof font.name !== 'string' || font.name === '')) {
    throw new TypeError('font.name must be a non-empty string');
  }
  let xml = '<font>';
  if (font.bold) xml += '<b/>';
  if (font.italic) xml += '<i/>';
  if (font.strike) xml += '<strike/>';
  if (font.underline) xml += '<u/>';
  xml += `<sz val="${font.size ?? 11}"/>`;
  xml += font.color === undefined ? '<color theme="1"/>' : `<color rgb="${normalizeColor(font.color, 'font.color')}"/>`;
  xml += `<name val="${xmlAttr(font.name ?? 'Calibri')}"/><family val="2"/>`;
  if (font.name === undefined) xml += '<scheme val="minor"/>';
  return `${xml}</font>`;
}

function fillXml(color: string): string {
  return `<fill><patternFill patternType="solid"><fgColor rgb="${normalizeColor(color, 'fill')}"/><bgColor indexed="64"/></patternFill></fill>`;
}

function borderXml(border: BorderDefinition | string): string {
  let sides: BorderDefinition;
  if (typeof border === 'string') {
    sides = { left: border as BorderSideSpec, right: border as BorderSideSpec, top: border as BorderSideSpec, bottom: border as BorderSideSpec };
  } else {
    assertObject(border, 'border');
    assertKeys(border, BORDER_KEYS, 'border');
    sides = border;
  }
  let xml = '<border>';
  for (const side of SIDES) {
    const spec = sides[side];
    if (spec === undefined) {
      xml += `<${side}/>`;
      continue;
    }
    let style: unknown;
    let color: unknown;
    if (typeof spec === 'string') {
      style = spec;
    } else {
      assertObject(spec, `border.${side}`);
      style = spec.style;
      color = spec.color;
    }
    if (typeof style !== 'string' || !BORDER_STYLES.has(style)) {
      throw new TypeError(`Invalid border style ${JSON.stringify(style)} for border.${side}`);
    }
    xml +=
      color === undefined
        ? `<${side} style="${style}"/>`
        : `<${side} style="${style}"><color rgb="${normalizeColor(color, `border.${side}.color`)}"/></${side}>`;
  }
  return `${xml}<diagonal/></border>`;
}

function alignmentXml(align: AlignmentDefinition): string {
  assertObject(align, 'align');
  assertKeys(align, ALIGN_KEYS, 'align');
  assertOptionalBoolean(align.wrap, 'align.wrap');
  if (align.horizontal !== undefined && !HORIZONTAL.has(align.horizontal)) {
    throw new TypeError(`Invalid align.horizontal ${JSON.stringify(align.horizontal)}`);
  }
  if (align.vertical !== undefined && !VERTICAL.has(align.vertical)) {
    throw new TypeError(`Invalid align.vertical ${JSON.stringify(align.vertical)}`);
  }
  let attrs = '';
  if (align.horizontal) attrs += ` horizontal="${align.horizontal}"`;
  if (align.vertical) attrs += ` vertical="${align.vertical}"`;
  if (align.wrap) attrs += ' wrapText="1"';
  return attrs === '' ? '' : `<alignment${attrs}/>`;
}

function xfXml(xf: Xf): string {
  const apply =
    `${xf.numFmtId ? ' applyNumberFormat="1"' : ''}${xf.fontId ? ' applyFont="1"' : ''}` +
    `${xf.fillId ? ' applyFill="1"' : ''}${xf.borderId ? ' applyBorder="1"' : ''}${xf.alignment ? ' applyAlignment="1"' : ''}`;
  const open = `<xf numFmtId="${xf.numFmtId}" fontId="${xf.fontId}" fillId="${xf.fillId}" borderId="${xf.borderId}" xfId="0"${apply}`;
  return xf.alignment ? `${open}>${xf.alignment}</xf>` : `${open}/>`;
}

function xfKey(xf: Xf): string {
  return `${xf.numFmtId}/${xf.fontId}/${xf.fillId}/${xf.borderId}/${xf.alignment}`;
}

/** Returns the index of `xml` in `list`, appending it when new. */
function intern(list: string[], ids: Map<string, number>, xml: string): number {
  let id = ids.get(xml);
  if (id === undefined) {
    id = list.length;
    list.push(xml);
    ids.set(xml, id);
  }
  return id;
}

/** Deduplicating registry of cell formats; produces styles.xml. */
export class StyleRegistry {
  private readonly numFmts: string[] = [];
  private readonly numFmtIds = new Map<string, number>();
  private readonly fonts = [DEFAULT_FONT];
  private readonly fontIds = new Map([[DEFAULT_FONT, 0]]);
  private readonly fills = [...DEFAULT_FILLS];
  private readonly fillIds = new Map(DEFAULT_FILLS.map((xml, i) => [xml, i] as const));
  private readonly borders = [DEFAULT_BORDER];
  private readonly borderIds = new Map([[DEFAULT_BORDER, 0]]);
  private readonly xfs: Xf[] = [DEFAULT_XF];
  private readonly xfIds = new Map([[xfKey(DEFAULT_XF), 0]]);
  private readonly dateVariants = new Map<number, number>();
  private linkId: number | undefined;

  constructor(
    private readonly owner: object,
    private readonly dateFormat: string,
  ) {}

  register(definition: StyleDefinition): StyleHandle {
    assertObject(definition, 'Style definition');
    assertKeys(definition, STYLE_KEYS, 'style');
    const xf: Xf = {
      numFmtId: this.numFmtId(definition.numFmt),
      fontId: definition.font === undefined ? 0 : intern(this.fonts, this.fontIds, fontXml(definition.font)),
      fillId: definition.fill === undefined ? 0 : intern(this.fills, this.fillIds, fillXml(definition.fill)),
      borderId: definition.border === undefined ? 0 : intern(this.borders, this.borderIds, borderXml(definition.border)),
      alignment: definition.align === undefined ? '' : alignmentXml(definition.align),
    };
    return new StyleHandle(this.xfId(xf), this.owner);
  }

  /** Resolves a handle to its xf index. Throws for handles made by another workbook. */
  idOf(style: Style | undefined): number | undefined {
    if (style === undefined) return undefined;
    if (!(style instanceof StyleHandle) || style.owner !== this.owner) {
      throw new TypeError('Style must be created with style() of the same workbook');
    }
    return style.id;
  }

  /** The same style with the default date format applied, for Date cells whose style has no numFmt. */
  dateVariant(xfId: number): number {
    const xf = this.xfs[xfId]!;
    if (xf.numFmtId !== 0) return xfId;
    let id = this.dateVariants.get(xfId);
    if (id === undefined) {
      id = this.xfId({ ...xf, numFmtId: this.numFmtId(this.dateFormat) });
      this.dateVariants.set(xfId, id);
    }
    return id;
  }

  /** Blue underlined text, applied to hyperlink cells without an explicit style. */
  linkStyle(): number {
    this.linkId ??= this.register({ font: { color: LINK_COLOR, underline: true } }).id;
    return this.linkId;
  }

  toXml(): string {
    const numFmts =
      this.numFmts.length === 0
        ? ''
        : `<numFmts count="${this.numFmts.length}">${this.numFmts
            .map((code, i) => `<numFmt numFmtId="${FIRST_CUSTOM_NUMFMT + i}" formatCode="${xmlAttr(code)}"/>`)
            .join('')}</numFmts>`;
    return (
      `${XML_DECL}<styleSheet xmlns="${NS_MAIN}">${numFmts}` +
      `<fonts count="${this.fonts.length}">${this.fonts.join('')}</fonts>` +
      `<fills count="${this.fills.length}">${this.fills.join('')}</fills>` +
      `<borders count="${this.borders.length}">${this.borders.join('')}</borders>` +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      `<cellXfs count="${this.xfs.length}">${this.xfs.map(xfXml).join('')}</cellXfs>` +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
      '</styleSheet>'
    );
  }

  private numFmtId(code: string | undefined): number {
    if (code === undefined || code === 'General') return 0;
    if (typeof code !== 'string' || code === '') throw new TypeError('numFmt must be a non-empty string');
    let id = this.numFmtIds.get(code);
    if (id === undefined) {
      id = FIRST_CUSTOM_NUMFMT + this.numFmts.length;
      this.numFmts.push(code);
      this.numFmtIds.set(code, id);
    }
    return id;
  }

  private xfId(xf: Xf): number {
    const key = xfKey(xf);
    let id = this.xfIds.get(key);
    if (id === undefined) {
      id = this.xfs.length;
      this.xfs.push(xf);
      this.xfIds.set(key, id);
    }
    return id;
  }
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run test/styles.test.ts`
Expected: PASS (11 tests).

- [ ] **Step 5: Commit**

```bash
git add src/styles.ts test/styles.test.ts
git commit -m "feat: deduplicating style registry and styles.xml generation"
```

---

### Task 8: Package parts (content types, relationships, workbook.xml)

**Files:**
- Create: `src/parts.ts`
- Test: `test/parts.test.ts`

**Interfaces:**
- Consumes: `XML_DECL`, `NS_MAIN`, `NS_REL`, `xmlAttr`, `xmlText` (Task 2); `Range`, `formatAbsoluteRange` (Task 6).
- Produces:
  - `interface SheetInfo { name: string; autoFilter: Range | undefined }`
  - `contentTypesXml(sheetCount: number): string`
  - `rootRelsXml(): string`
  - `workbookRelsXml(sheetCount: number): string` — `rId1..rIdN` листы, `rId(N+1)` стили
  - `workbookXml(sheets: readonly SheetInfo[]): string`

- [ ] **Step 1: Write the failing test**

`test/parts.test.ts`:
```ts
import { describe, expect, it } from 'vitest';
import { parseRange } from '../src/cell.js';
import { contentTypesXml, rootRelsXml, workbookRelsXml, workbookXml } from '../src/parts.js';
import { parseXml } from './helpers.js';

describe('contentTypesXml', () => {
  it('lists defaults, the workbook, every sheet and the styles part', () => {
    const doc = parseXml(contentTypesXml(2));
    const overrides = doc.Types.Override.map((o: any) => o.PartName);
    expect(overrides).toEqual(['/xl/workbook.xml', '/xl/worksheets/sheet1.xml', '/xl/worksheets/sheet2.xml', '/xl/styles.xml']);
    expect(doc.Types.Default).toHaveLength(2);
    expect(doc.Types.xmlns).toBe('http://schemas.openxmlformats.org/package/2006/content-types');
  });
});

describe('relationships', () => {
  it('points the package at xl/workbook.xml', () => {
    const doc = parseXml(rootRelsXml());
    expect(doc.Relationships.Relationship[0]).toMatchObject({
      Id: 'rId1',
      Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
      Target: 'xl/workbook.xml',
    });
  });

  it('relates sheets by index and styles last', () => {
    const rels = parseXml(workbookRelsXml(2)).Relationships.Relationship;
    expect(rels.map((r: any) => [r.Id, r.Target])).toEqual([
      ['rId1', 'worksheets/sheet1.xml'],
      ['rId2', 'worksheets/sheet2.xml'],
      ['rId3', 'styles.xml'],
    ]);
    expect(rels[2].Type).toBe('http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles');
  });
});

describe('workbookXml', () => {
  it('declares sheets, auto-filter defined names and full recalculation on load', () => {
    const xml = workbookXml([
      { name: 'Data & "more"', autoFilter: parseRange('A1:C1') },
      { name: "O'Brien", autoFilter: undefined },
      { name: 'Third', autoFilter: parseRange('B2:D9') },
    ]);
    const doc = parseXml(xml);
    expect(doc.workbook.sheets.sheet.map((s: any) => [s.name, s.sheetId, s['r:id']])).toEqual([
      ['Data & "more"', '1', 'rId1'],
      ["O'Brien", '2', 'rId2'],
      ['Third', '3', 'rId3'],
    ]);
    expect(xml).toContain(`<definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">'Data &amp; "more"'!$A$1:$C$1</definedName>`);
    expect(xml).toContain(`<definedName name="_xlnm._FilterDatabase" localSheetId="2" hidden="1">'Third'!$B$2:$D$9</definedName>`);
    expect(doc.workbook.definedNames.definedName).toHaveLength(2);
    expect(doc.workbook.calcPr.fullCalcOnLoad).toBe('1');
    expect(xml.indexOf('<definedNames>')).toBeLessThan(xml.indexOf('<calcPr'));
  });

  it('doubles apostrophes in sheet names inside defined names and omits definedNames without filters', () => {
    const xml = workbookXml([{ name: "O'Brien", autoFilter: parseRange('A1:A1') }]);
    expect(xml).toContain(`'O''Brien'!$A$1:$A$1`);
    expect(workbookXml([{ name: 'Plain', autoFilter: undefined }])).not.toContain('definedNames');
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run test/parts.test.ts`
Expected: FAIL — модуль `../src/parts.js` не найден.

- [ ] **Step 3: Write the implementation**

`src/parts.ts`:
```ts
import { formatAbsoluteRange, type Range } from './cell.js';
import { NS_MAIN, NS_REL, XML_DECL, xmlAttr, xmlText } from './xml.js';

const NS_CONTENT_TYPES = 'http://schemas.openxmlformats.org/package/2006/content-types';
const NS_PACKAGE_RELS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CT_RELS = 'application/vnd.openxmlformats-package.relationships+xml';
const CT_WORKBOOK = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml';
const CT_WORKSHEET = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml';
const CT_STYLES = 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml';
const REL_OFFICE_DOCUMENT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
const REL_WORKSHEET = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';
const REL_STYLES = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles';

export interface SheetInfo {
  name: string;
  autoFilter: Range | undefined;
}

export function contentTypesXml(sheetCount: number): string {
  let xml =
    `${XML_DECL}<Types xmlns="${NS_CONTENT_TYPES}">` +
    `<Default Extension="rels" ContentType="${CT_RELS}"/>` +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    `<Override PartName="/xl/workbook.xml" ContentType="${CT_WORKBOOK}"/>`;
  for (let i = 1; i <= sheetCount; i++) {
    xml += `<Override PartName="/xl/worksheets/sheet${i}.xml" ContentType="${CT_WORKSHEET}"/>`;
  }
  return `${xml}<Override PartName="/xl/styles.xml" ContentType="${CT_STYLES}"/></Types>`;
}

export function rootRelsXml(): string {
  return (
    `${XML_DECL}<Relationships xmlns="${NS_PACKAGE_RELS}">` +
    `<Relationship Id="rId1" Type="${REL_OFFICE_DOCUMENT}" Target="xl/workbook.xml"/>` +
    '</Relationships>'
  );
}

export function workbookRelsXml(sheetCount: number): string {
  let xml = `${XML_DECL}<Relationships xmlns="${NS_PACKAGE_RELS}">`;
  for (let i = 1; i <= sheetCount; i++) {
    xml += `<Relationship Id="rId${i}" Type="${REL_WORKSHEET}" Target="worksheets/sheet${i}.xml"/>`;
  }
  return `${xml}<Relationship Id="rId${sheetCount + 1}" Type="${REL_STYLES}" Target="styles.xml"/></Relationships>`;
}

/** Sheet names inside formulas are single-quoted with apostrophes doubled. */
function quoteSheetName(name: string): string {
  return `'${name.replace(/'/g, "''")}'`;
}

export function workbookXml(sheets: readonly SheetInfo[]): string {
  let xml = `${XML_DECL}<workbook xmlns="${NS_MAIN}" xmlns:r="${NS_REL}"><workbookPr/><bookViews><workbookView/></bookViews><sheets>`;
  sheets.forEach((sheet, i) => {
    xml += `<sheet name="${xmlAttr(sheet.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`;
  });
  xml += '</sheets>';
  let definedNames = '';
  sheets.forEach((sheet, i) => {
    if (!sheet.autoFilter) return;
    definedNames +=
      `<definedName name="_xlnm._FilterDatabase" localSheetId="${i}" hidden="1">` +
      `${xmlText(quoteSheetName(sheet.name))}!${formatAbsoluteRange(sheet.autoFilter)}</definedName>`;
  });
  if (definedNames) xml += `<definedNames>${definedNames}</definedNames>`;
  return `${xml}<calcPr fullCalcOnLoad="1"/></workbook>`;
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run test/parts.test.ts`
Expected: PASS (5 tests).

- [ ] **Step 5: Commit**

```bash
git add src/parts.ts test/parts.test.ts
git commit -m "feat: package parts — content types, relationships, workbook.xml"
```

---

### Task 9: SheetWriter

**Files:**
- Create: `src/sheet.ts`
- Test: `test/sheet.test.ts`

**Interfaces:**
- Consumes: `Sheet`, `SheetOptions`, `RowOptions`, `CellValue` (Task 1); `XML_DECL`, `NS_MAIN`, `NS_REL` (Task 2); `ZipEntry` (Task 5); `cellRef`, `cellXml`, `parseRange`, `formatRange`, `MAX_ROWS`, `MAX_COLS`, `CellContext`, `Range` (Task 6).
- Produces:
  - `interface SheetDeps { entry: ZipEntry; ctx: CellContext; first: boolean }`
  - `class SheetWriter implements Sheet { constructor(name: string, options: SheetOptions, deps: SheetDeps); readonly name; get nextRow(); get autoFilterRange(): Range | undefined; writeRow(); writeRows(); merge(); autoFilter(); close(); fail(reason: unknown): void }`
  - `fail(reason)` — после него `writeRow`/`close` возвращают отклонённый промис, `merge`/`autoFilter` бросают `reason`.

- [ ] **Step 1: Write the failing test**

`test/sheet.test.ts`:
```ts
import { describe, expect, it } from 'vitest';
import type { CellContext } from '../src/cell.js';
import { MAX_ROWS } from '../src/cell.js';
import { SheetWriter } from '../src/sheet.js';
import type { Style } from '../src/types.js';
import { MemoryEntry } from './helpers.js';

const style = (id: number): Style => ({ id });
const ctx: CellContext = { dates: 'utc', styleId: (s) => s?.id, dateStyle: (s) => (s === 0 ? 99 : s), linkStyle: () => 7 };

const HEADER_START = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';

function sheet(options = {}, first = true) {
  const entry = new MemoryEntry();
  return { entry, sheet: new SheetWriter('S', options, { entry, ctx, first }) };
}

describe('SheetWriter header', () => {
  it('writes a minimal header for the first sheet', async () => {
    const { entry, sheet: s } = sheet();
    await s.close();
    expect(entry.text()).toBe(
      `${HEADER_START}<sheetViews><sheetView workbookViewId="0" tabSelected="1"></sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData></sheetData></worksheet>`,
    );
    expect(entry.closed).toBe(true);
  });

  it('marks only the first sheet as selected', async () => {
    const { entry, sheet: s } = sheet({}, false);
    await s.close();
    expect(entry.text()).toContain('<sheetView workbookViewId="0"></sheetView>');
  });

  it('freezes rows and/or columns with the right active pane', async () => {
    const both = sheet({ freeze: { rows: 1, columns: 2 } });
    await both.sheet.close();
    expect(both.entry.text()).toContain('<pane xSplit="2" ySplit="1" topLeftCell="C2" activePane="bottomRight" state="frozen"/>');
    const rows = sheet({ freeze: { rows: 3 } });
    await rows.sheet.close();
    expect(rows.entry.text()).toContain('<pane ySplit="3" topLeftCell="A4" activePane="bottomLeft" state="frozen"/>');
    const cols = sheet({ freeze: { columns: 1 } });
    await cols.sheet.close();
    expect(cols.entry.text()).toContain('<pane xSplit="1" topLeftCell="B1" activePane="topRight" state="frozen"/>');
    const none = sheet({ freeze: { rows: 0 } });
    await none.sheet.close();
    expect(none.entry.text()).not.toContain('<pane');
  });

  it('writes column widths and styles', async () => {
    const { entry, sheet: s } = sheet({ columns: [{ width: 30 }, {}, { style: style(4) }, { width: 12.5, style: style(2) }] });
    await s.close();
    expect(entry.text()).toContain(
      '<cols><col min="1" max="1" width="30" customWidth="1"/><col min="3" max="3" width="8.43" style="4"/><col min="4" max="4" width="12.5" customWidth="1" style="2"/></cols>',
    );
  });

  it('validates options', () => {
    expect(() => sheet({ columns: 'wide' })).toThrow(TypeError);
    expect(() => sheet({ columns: [{ width: 300 }] })).toThrow(RangeError);
    expect(() => sheet({ columns: [{ width: -1 }] })).toThrow(RangeError);
    expect(() => sheet({ freeze: { rows: 1.5 } })).toThrow(RangeError);
    expect(() => sheet({ freeze: { columns: -1 } })).toThrow(RangeError);
    expect(() => sheet({ freeze: 'top' })).toThrow(TypeError);
  });
});

describe('SheetWriter rows', () => {
  it('writes rows with increasing numbers and tracks nextRow', async () => {
    const { entry, sheet: s } = sheet();
    expect(s.nextRow).toBe(1);
    await s.writeRow(['a', 1]);
    await s.writeRow([null, true]);
    expect(s.nextRow).toBe(3);
    await s.close();
    expect(entry.text()).toContain(
      '<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>a</t></is></c><c r="B1"><v>1</v></c></row><row r="2"><c r="B2" t="b"><v>1</v></c></row></sheetData>',
    );
  });

  it('applies row style, column style and cell style with the right precedence', async () => {
    const { entry, sheet: s } = sheet({ columns: [{ style: style(3) }] });
    await s.writeRow([1, 2]); // column style on A only
    await s.writeRow([1, 2, { value: 3, style: style(8) }], { style: style(5) }); // row style beats column style, cell beats row
    await s.close();
    const xml = entry.text();
    expect(xml).toContain('<row r="1"><c r="A1" s="3"><v>1</v></c><c r="B1"><v>2</v></c></row>');
    expect(xml).toContain('<row r="2" s="5" customFormat="1"><c r="A2" s="5"><v>1</v></c><c r="B2" s="5"><v>2</v></c><c r="C2" s="8"><v>3</v></c></row>');
  });

  it('writes custom row heights and validates them', async () => {
    const { entry, sheet: s } = sheet();
    await s.writeRow(['x'], { height: 30 });
    expect(() => s.writeRow(['x'], { height: 500 })).toThrow(RangeError);
    expect(() => s.writeRow(['x'], { height: 'tall' as never })).toThrow(RangeError);
    await s.close();
    expect(entry.text()).toContain('<row r="1" ht="30" customHeight="1">');
  });

  it('writeRows consumes sync and async iterables', async () => {
    const { entry, sheet: s } = sheet();
    await s.writeRows([[1], [2]]);
    async function* more() {
      yield [3];
      yield [4];
    }
    await s.writeRows(more());
    await s.close();
    expect(entry.text()).toMatch(/<row r="1">.*<row r="2">.*<row r="3">.*<row r="4">/);
  });

  it('flushes to the entry in chunks before close', async () => {
    const { entry, sheet: s } = sheet();
    const wide = Array.from({ length: 50 }, (_, i) => `cell number ${i} with some padding text`);
    for (let i = 0; i < 200; i++) await s.writeRow(wide);
    expect(entry.chunks.length).toBeGreaterThan(1);
    await s.close();
    expect(entry.text()).toContain('<row r="200">');
  });

  it('rejects non-array rows, too many cells and too many rows', () => {
    const { sheet: s } = sheet();
    expect(() => s.writeRow('abc' as never)).toThrow(TypeError);
    expect(() => s.writeRow(new Array(16_385).fill(1))).toThrow(RangeError);
    for (let i = 0; i < MAX_ROWS; i++) void s.writeRow([]);
    expect(() => s.writeRow([1])).toThrow(RangeError);
  });
});

describe('SheetWriter merges, filters and lifecycle', () => {
  it('writes autoFilter before mergeCells after sheetData', async () => {
    const { entry, sheet: s } = sheet();
    await s.writeRow(['a', 'b', 'c']);
    s.merge('C3:A3');
    s.autoFilter('A1:C1');
    s.merge('A5:A6');
    await s.close();
    expect(entry.text()).toContain(
      '</sheetData><autoFilter ref="A1:C1"/><mergeCells count="2"><mergeCell ref="A3:C3"/><mergeCell ref="A5:A6"/></mergeCells></worksheet>',
    );
    expect(s.autoFilterRange).toEqual({ c1: 1, r1: 1, c2: 3, r2: 1 });
  });

  it('rejects single-cell merges, malformed ranges and a second auto filter', () => {
    const { sheet: s } = sheet();
    expect(() => s.merge('A1:A1')).toThrow(TypeError);
    expect(() => s.merge('A1')).toThrow(TypeError);
    s.autoFilter('A1:B1');
    expect(() => s.autoFilter('A1:C1')).toThrow('already has an auto filter');
  });

  it('refuses every operation after close and keeps close idempotent', async () => {
    const { sheet: s } = sheet();
    const first = s.close();
    expect(s.close()).toBe(first);
    await first;
    expect(() => s.writeRow([1])).toThrow('already closed');
    expect(() => s.merge('A1:B1')).toThrow('already closed');
    expect(() => s.autoFilter('A1:B1')).toThrow('already closed');
  });

  it('fail() turns later operations into rejections with the given reason', async () => {
    const { sheet: s } = sheet();
    const reason = new Error('stream gone');
    s.fail(reason);
    await expect(s.writeRow([1])).rejects.toBe(reason);
    await expect(s.close()).rejects.toBe(reason);
    expect(() => s.merge('A1:B1')).toThrow(reason);
  });

  it('remembers a failed flush and rejects subsequent writes', async () => {
    const entry = new MemoryEntry();
    const boom = new Error('entry failed');
    entry.write = () => Promise.reject(boom);
    const s = new SheetWriter('S', {}, { entry, ctx, first: true });
    await expect(s.close()).rejects.toBe(boom);
    await expect(s.writeRow([1])).rejects.toBe(boom);
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run test/sheet.test.ts`
Expected: FAIL — модуль `../src/sheet.js` не найден.

- [ ] **Step 3: Write the implementation**

`src/sheet.ts`:
```ts
import { cellRef, cellXml, formatRange, MAX_COLS, MAX_ROWS, parseRange, type CellContext, type Range } from './cell.js';
import type { CellValue, RowOptions, Sheet, SheetOptions } from './types.js';
import { NS_MAIN, NS_REL, XML_DECL } from './xml.js';
import type { ZipEntry } from './zip/zip-writer.js';

/** Buffered XML (UTF-16 code units) that triggers a flush to the compressor. */
const FLUSH_THRESHOLD = 64 * 1024;
const DEFAULT_COLUMN_WIDTH = 8.43;
const MAX_COLUMN_WIDTH = 255;
const MAX_ROW_HEIGHT = 409;
const ENCODER = new TextEncoder();
const RESOLVED = Promise.resolve();

export interface SheetDeps {
  entry: ZipEntry;
  ctx: CellContext;
  /** Only the first sheet of a workbook is marked as selected. */
  first: boolean;
}

function assertCount(value: unknown, what: string, max: number): asserts value is number {
  if (typeof value !== 'number' || !Number.isInteger(value) || value < 0 || value > max) {
    throw new RangeError(`${what} must be an integer between 0 and ${max}`);
  }
}

/** Streams one worksheet's XML into a ZIP entry. */
export class SheetWriter implements Sheet {
  readonly name: string;
  private readonly entry: ZipEntry;
  private readonly ctx: CellContext;
  private readonly columnStyles: number[] = [];
  private readonly merges: string[] = [];
  private filter: Range | undefined;
  private buffer: string;
  private row = 1;
  private closePromise: Promise<void> | null = null;
  private failure: { reason: unknown } | null = null;

  constructor(name: string, options: SheetOptions, deps: SheetDeps) {
    this.name = name;
    this.entry = deps.entry;
    this.ctx = deps.ctx;
    this.buffer = this.header(options, deps.first);
  }

  get nextRow(): number {
    return this.row;
  }

  /** @internal The auto-filter range, needed by workbook.xml's defined names. */
  get autoFilterRange(): Range | undefined {
    return this.filter;
  }

  writeRow(cells: readonly CellValue[], options?: RowOptions): Promise<void> {
    if (this.failure) return Promise.reject(this.failure.reason);
    this.assertOpen('writeRow');
    if (!Array.isArray(cells)) throw new TypeError('writeRow() expects an array of cell values');
    const row = this.row;
    if (row > MAX_ROWS) throw new RangeError(`Sheet "${this.name}" already has ${MAX_ROWS} rows, the maximum Excel supports`);
    if (cells.length > MAX_COLS) throw new RangeError(`Row ${row} has ${cells.length} cells; Excel supports at most ${MAX_COLS} columns`);

    const rowStyle = this.ctx.styleId(options?.style) ?? 0;
    let xml = `<row r="${row}"`;
    if (rowStyle !== 0) xml += ` s="${rowStyle}" customFormat="1"`;
    if (options?.height !== undefined) {
      const height = options.height;
      if (typeof height !== 'number' || !(height >= 0 && height <= MAX_ROW_HEIGHT)) {
        throw new RangeError(`Row height must be a number between 0 and ${MAX_ROW_HEIGHT} points`);
      }
      xml += ` ht="${height}" customHeight="1"`;
    }
    xml += '>';
    for (let i = 0; i < cells.length; i++) {
      const inherited = rowStyle !== 0 ? rowStyle : (this.columnStyles[i] ?? 0);
      xml += cellXml(i + 1, row, cells[i], inherited, this.ctx);
    }
    xml += '</row>';

    this.row = row + 1;
    this.buffer += xml;
    return this.buffer.length >= FLUSH_THRESHOLD ? this.flush() : RESOLVED;
  }

  async writeRows(rows: Iterable<readonly CellValue[]> | AsyncIterable<readonly CellValue[]>): Promise<void> {
    for await (const row of rows) await this.writeRow(row);
  }

  merge(range: string): void {
    this.assertUsable('merge');
    const r = parseRange(range);
    if (r.c1 === r.c2 && r.r1 === r.r2) throw new TypeError(`Cannot merge the single cell ${JSON.stringify(range)}`);
    this.merges.push(formatRange(r));
  }

  autoFilter(range: string): void {
    this.assertUsable('autoFilter');
    if (this.filter) throw new Error(`Sheet "${this.name}" already has an auto filter`);
    this.filter = parseRange(range);
  }

  close(): Promise<void> {
    if (this.closePromise) return this.closePromise;
    if (this.failure) return (this.closePromise = Promise.reject(this.failure.reason));
    let tail = '</sheetData>';
    if (this.filter) tail += `<autoFilter ref="${formatRange(this.filter)}"/>`;
    if (this.merges.length > 0) {
      tail += `<mergeCells count="${this.merges.length}">${this.merges.map((m) => `<mergeCell ref="${m}"/>`).join('')}</mergeCells>`;
    }
    this.buffer += `${tail}</worksheet>`;
    this.closePromise = this.flush().then(() => this.entry.close());
    return this.closePromise;
  }

  /** @internal Marks the sheet as failed; later operations reject/throw with `reason`. */
  fail(reason: unknown): void {
    if (!this.failure) this.failure = { reason };
  }

  private assertOpen(method: string): void {
    if (this.closePromise) throw new Error(`Cannot call ${method}() on sheet "${this.name}": it is already closed`);
  }

  private assertUsable(method: string): void {
    if (this.failure) throw this.failure.reason;
    this.assertOpen(method);
  }

  private flush(): Promise<void> {
    const text = this.buffer;
    this.buffer = '';
    return this.entry.write(ENCODER.encode(text)).catch((reason: unknown) => {
      this.fail(reason);
      throw reason;
    });
  }

  private header(options: SheetOptions, first: boolean): string {
    const { columns = [], freeze } = options;
    if (!Array.isArray(columns)) throw new TypeError('columns must be an array');
    if (columns.length > MAX_COLS) throw new RangeError(`columns has ${columns.length} entries; Excel supports at most ${MAX_COLS}`);

    let xml = `${XML_DECL}<worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_REL}"><sheetViews><sheetView workbookViewId="0"${first ? ' tabSelected="1"' : ''}>`;
    if (freeze !== undefined) {
      if (freeze === null || typeof freeze !== 'object') throw new TypeError('freeze must be an object like { rows: 1 }');
      const rows = freeze.rows ?? 0;
      const cols = freeze.columns ?? 0;
      assertCount(rows, 'freeze.rows', MAX_ROWS - 1);
      assertCount(cols, 'freeze.columns', MAX_COLS - 1);
      if (rows > 0 || cols > 0) {
        const pane = rows > 0 && cols > 0 ? 'bottomRight' : rows > 0 ? 'bottomLeft' : 'topRight';
        xml += `<pane${cols > 0 ? ` xSplit="${cols}"` : ''}${rows > 0 ? ` ySplit="${rows}"` : ''} topLeftCell="${cellRef(cols + 1, rows + 1)}" activePane="${pane}" state="frozen"/>`;
      }
    }
    xml += '</sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/>';

    let cols = '';
    columns.forEach((column, i) => {
      if (column === null || typeof column !== 'object') throw new TypeError(`columns[${i}] must be an object`);
      const styleId = this.ctx.styleId(column.style) ?? 0;
      this.columnStyles[i] = styleId;
      const { width } = column;
      if (width !== undefined && !(typeof width === 'number' && width >= 0 && width <= MAX_COLUMN_WIDTH)) {
        throw new RangeError(`columns[${i}].width must be a number between 0 and ${MAX_COLUMN_WIDTH}`);
      }
      if (width === undefined && styleId === 0) return;
      cols += `<col min="${i + 1}" max="${i + 1}" width="${width ?? DEFAULT_COLUMN_WIDTH}"${width !== undefined ? ' customWidth="1"' : ''}${styleId !== 0 ? ` style="${styleId}"` : ''}/>`;
    });
    if (cols) xml += `<cols>${cols}</cols>`;
    return `${xml}<sheetData>`;
  }
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run test/sheet.test.ts`
Expected: PASS (15 tests). Тест с `MAX_ROWS` пустых строк идёт около секунды — это нормально.

- [ ] **Step 5: Commit**

```bash
git add src/sheet.ts test/sheet.test.ts
git commit -m "feat: SheetWriter — buffered worksheet XML streaming"
```

---

### Task 10: Workbook, public entry point and integration tests

**Files:**
- Create: `src/workbook.ts`
- Modify: `src/index.ts`
- Test: `test/workbook.test.ts`

**Interfaces:**
- Consumes: всё из задач 1–9.
- Produces:
  - `class WorkbookImpl implements Workbook` (не экспортируется из `index.ts`)
  - `createWorkbook(options?: WorkbookOptions): Workbook` — единственная функция публичного API.
  - `index.ts` реэкспортирует все типы из `types.ts`.

- [ ] **Step 1: Write the failing integration test**

`test/workbook.test.ts`:
```ts
import * as ExcelJS from 'exceljs';
import { describe, expect, it } from 'vitest';
import { createWorkbook } from '../src/index.js';
import { collect, parseXml, unzipText } from './helpers.js';

async function readBack(bytes: Uint8Array): Promise<ExcelJS.Workbook> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(Buffer.from(bytes)); // exceljs wants a Node Buffer; tests may use Node APIs, src/ may not
  return workbook;
}

describe('createWorkbook end to end', () => {
  it('produces a package with every required part, readable by an independent library', async () => {
    const wb = createWorkbook();
    const header = wb.style({ font: { bold: true }, fill: '#EEEEEE', border: 'thin' });
    const money = wb.style({ numFmt: '#,##0.00', align: { horizontal: 'right' } });
    const output = collect(wb.readable);

    const report = wb.addSheet('Report', { columns: [{ width: 30 }, { width: 14, style: money }, { width: 12 }], freeze: { rows: 1 } });
    await report.writeRow(['Item', 'Amount', 'Date'], { style: header });
    await report.writeRow(['Coffee & tea', 12.5, new Date(Date.UTC(2026, 9, 6))]);
    await report.writeRow(['Books', 1000, new Date(Date.UTC(2026, 0, 1, 12))]);
    await report.writeRow(['Total', { formula: 'SUM(B2:B3)' }, { link: 'https://example.com', value: 'Details' }]);
    await report.writeRow([true, false, null]);
    report.autoFilter('A1:C1');
    report.merge('A6:C6');
    await report.writeRow(['Merged title']);

    const second = wb.addSheet('Second');
    await second.writeRow([NaN, 'plain']);
    await wb.close();

    const bytes = await output;
    const files = unzipText(bytes);
    expect(Object.keys(files).sort()).toEqual(
      ['[Content_Types].xml', '_rels/.rels', 'xl/_rels/workbook.xml.rels', 'xl/styles.xml', 'xl/workbook.xml', 'xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml'].sort(),
    );
    const contentTypes = parseXml(files['[Content_Types].xml']!);
    expect(contentTypes.Types.Override.map((o: any) => o.PartName)).toContain('/xl/worksheets/sheet2.xml');
    const workbookXml = parseXml(files['xl/workbook.xml']!);
    expect(workbookXml.workbook.sheets.sheet.map((s: any) => s.name)).toEqual(['Report', 'Second']);
    expect(files['xl/workbook.xml']).toContain(`'Report'!$A$1:$C$1`);
    expect(files['xl/worksheets/sheet1.xml']).toContain('<autoFilter ref="A1:C1"/><mergeCells count="1"><mergeCell ref="A6:C6"/></mergeCells>');
    expect(files['xl/worksheets/sheet2.xml']).toContain('<c r="A1" t="e"><v>#NUM!</v></c>');

    const excel = await readBack(bytes);
    expect(excel.worksheets.map((ws) => ws.name)).toEqual(['Report', 'Second']);
    const ws = excel.getWorksheet('Report')!;
    expect(ws.getCell('A1').value).toBe('Item');
    expect(ws.getCell('A1').font.bold).toBe(true);
    expect(ws.getCell('A1').fill).toMatchObject({ type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEEEEEE' } });
    expect(ws.getCell('A1').border.top?.style).toBe('thin');
    expect(ws.getCell('A2').value).toBe('Coffee & tea');
    expect(ws.getCell('B2').value).toBe(12.5);
    expect(ws.getCell('B2').numFmt).toBe('#,##0.00');
    expect(ws.getCell('B2').alignment.horizontal).toBe('right');
    expect(ws.getCell('C2').value).toBeInstanceOf(Date);
    expect((ws.getCell('C2').value as Date).getTime()).toBe(Date.UTC(2026, 9, 6));
    expect((ws.getCell('C3').value as Date).getTime()).toBe(Date.UTC(2026, 0, 1, 12));
    expect(ws.getCell('C2').numFmt).toBe('yyyy-mm-dd');
    expect(ws.getCell('B4').value).toMatchObject({ formula: 'SUM(B2:B3)' });
    expect(ws.getCell('C4').value).toMatchObject({ formula: 'HYPERLINK("https://example.com","Details")', result: 'Details' });
    expect(ws.getCell('C4').font.underline).toBeTruthy();
    expect(ws.getCell('A5').value).toBe(true);
    expect(ws.getCell('B5').value).toBe(false);
    expect(ws.getCell('C5').value).toBeNull();
    expect(ws.getColumn(1).width).toBe(30);
    expect(ws.getColumn(2).width).toBe(14);
    expect(ws.views[0]).toMatchObject({ state: 'frozen', ySplit: 1 });
    expect(ws.autoFilter).toBe('A1:C1');
    expect(ws.model.merges).toContain('A6:C6');
    expect(excel.getWorksheet('Second')!.getCell('A1').value).toMatchObject({ error: '#NUM!' });
    expect(excel.getWorksheet('Second')!.getCell('B1').value).toBe('plain');
  });

  it('supports local date mode and a custom default date format', async () => {
    const wb = createWorkbook({ dates: 'local', dateFormat: 'dd.mm.yyyy hh:mm' });
    const output = collect(wb.readable);
    const sheet = wb.addSheet('Dates');
    await sheet.writeRow([new Date(2026, 9, 6, 15, 30)]);
    await wb.close();
    const excel = await readBack(await output);
    const cell = excel.getWorksheet('Dates')!.getCell('A1');
    expect((cell.value as Date).getTime()).toBe(Date.UTC(2026, 9, 6, 15, 30));
    expect(cell.numFmt).toBe('dd.mm.yyyy hh:mm');
  });

  it('closes the previous sheet implicitly when a new one is added', async () => {
    const wb = createWorkbook();
    const output = collect(wb.readable);
    const a = wb.addSheet('A');
    void a.writeRow([1]);
    const b = wb.addSheet('B');
    await b.writeRow([2]);
    await wb.close();
    const files = unzipText(await output);
    expect(files['xl/worksheets/sheet1.xml']).toContain('<c r="A1"><v>1</v></c>');
    expect(files['xl/worksheets/sheet2.xml']).toContain('<c r="A1"><v>2</v></c>');
    expect(() => a.writeRow([3])).toThrow('already closed');
  });

  it('handles 100k rows with constant-size chunks and a valid result', async () => {
    const wb = createWorkbook();
    let chunks = 0;
    let largest = 0;
    const output = wb.readable.pipeTo(
      new WritableStream({
        write(chunk) {
          chunks++;
          largest = Math.max(largest, chunk.byteLength);
        },
      }),
    );
    const sheet = wb.addSheet('Big');
    for (let i = 1; i <= 100_000; i++) await sheet.writeRow([i, `row ${i}`, i % 7 === 0]);
    expect(sheet.nextRow).toBe(100_001);
    await wb.close();
    await output;
    expect(chunks).toBeGreaterThan(10);
    expect(largest).toBeLessThan(2 * 1024 * 1024);
  });
});
```

Если `import * as ExcelJS from 'exceljs'` даёт в рантайме `ExcelJS.Workbook === undefined` (exceljs — CommonJS с `module.exports = ExcelJS`), заменить на:
```ts
import { createRequire } from 'node:module';
const ExcelJS = createRequire(import.meta.url)('exceljs') as typeof import('exceljs');
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npx vitest run test/workbook.test.ts`
Expected: FAIL — `createWorkbook` не экспортируется из `../src/index.js`.

- [ ] **Step 3: Write the implementation**

`src/workbook.ts`:
```ts
import { MAX_COLS, type CellContext } from './cell.js';
import { contentTypesXml, rootRelsXml, workbookRelsXml, workbookXml } from './parts.js';
import { SheetWriter } from './sheet.js';
import { StreamSink } from './sink.js';
import { StyleRegistry } from './styles.js';
import type { Sheet, SheetOptions, Style, StyleDefinition, Workbook, WorkbookOptions } from './types.js';
import { ZipWriter } from './zip/zip-writer.js';

const ENCODER = new TextEncoder();
const MAX_SHEET_NAME_LENGTH = 31;
const FORBIDDEN_SHEET_NAME_CHARS = /[[\]:*?/\\]/;
const noop = (): void => {};

function validateSheetName(name: unknown, taken: ReadonlySet<string>): asserts name is string {
  if (typeof name !== 'string' || name.length === 0) throw new TypeError('Sheet name must be a non-empty string');
  if (name.length > MAX_SHEET_NAME_LENGTH) {
    throw new TypeError(`Sheet name ${JSON.stringify(name)} is longer than ${MAX_SHEET_NAME_LENGTH} characters`);
  }
  if (FORBIDDEN_SHEET_NAME_CHARS.test(name)) {
    throw new TypeError(`Sheet name ${JSON.stringify(name)} contains a forbidden character ([ ] : * ? / \\)`);
  }
  if (name.startsWith("'") || name.endsWith("'")) {
    throw new TypeError(`Sheet name ${JSON.stringify(name)} must not start or end with an apostrophe`);
  }
  if (taken.has(name.toLowerCase())) {
    throw new Error(`Sheet name ${JSON.stringify(name)} is already used (sheet names are case-insensitive)`);
  }
}

export class WorkbookImpl implements Workbook {
  readonly readable: ReadableStream<Uint8Array>;
  private readonly sink: StreamSink;
  private readonly zip: ZipWriter;
  private readonly styles: StyleRegistry;
  private readonly ctx: CellContext;
  private readonly sheets: SheetWriter[] = [];
  private readonly names = new Set<string>();
  private current: SheetWriter | null = null;
  private closePromise: Promise<void> | null = null;
  private closed = false;

  constructor(options: WorkbookOptions = {}) {
    if (options === null || typeof options !== 'object') throw new TypeError('Workbook options must be an object');
    const dates = options.dates ?? 'utc';
    if (dates !== 'utc' && dates !== 'local') throw new TypeError(`dates must be "utc" or "local", got ${JSON.stringify(dates)}`);
    const dateFormat = options.dateFormat ?? 'yyyy-mm-dd';
    if (typeof dateFormat !== 'string' || dateFormat === '') throw new TypeError('dateFormat must be a non-empty string');

    this.sink = new StreamSink();
    this.sink.onFailure = (reason) => this.fail(reason);
    this.readable = this.sink.readable;
    this.zip = new ZipWriter(this.sink);
    this.styles = new StyleRegistry(this, dateFormat);
    this.ctx = {
      dates,
      styleId: (style) => this.styles.idOf(style),
      dateStyle: (id) => this.styles.dateVariant(id),
      linkStyle: () => this.styles.linkStyle(),
    };
  }

  style(definition: StyleDefinition): Style {
    this.assertOpen('style');
    return this.styles.register(definition);
  }

  addSheet(name: string, options: SheetOptions = {}): Sheet {
    this.assertOpen('addSheet');
    validateSheetName(name, this.names);
    if (options === null || typeof options !== 'object') throw new TypeError('Sheet options must be an object');
    if (options.columns !== undefined && Array.isArray(options.columns) && options.columns.length > MAX_COLS) {
      throw new RangeError(`columns has ${options.columns.length} entries; Excel supports at most ${MAX_COLS}`);
    }
    // Close the previous sheet implicitly; its failure, if any, resurfaces through the ZIP chain.
    if (this.current) this.current.close().catch(noop);
    const index = this.sheets.length;
    const entry = this.zip.beginEntry(`xl/worksheets/sheet${index + 1}.xml`);
    const sheet = new SheetWriter(name, options, { entry, ctx: this.ctx, first: index === 0 });
    this.sheets.push(sheet);
    this.names.add(name.toLowerCase());
    this.current = sheet;
    return sheet;
  }

  close(): Promise<void> {
    if (this.closePromise) return this.closePromise;
    this.closed = true;
    if (this.sheets.length === 0) {
      const error = new Error('Workbook must contain at least one sheet');
      this.fail(error);
      this.sink.error(error);
      return (this.closePromise = Promise.reject(error));
    }
    this.closePromise = this.finish().catch((reason: unknown) => {
      this.fail(reason);
      this.sink.error(reason);
      throw reason;
    });
    return this.closePromise;
  }

  abort(reason: unknown = new Error('Workbook was aborted')): void {
    this.closed = true;
    this.fail(reason);
    this.sink.error(reason);
  }

  private async finish(): Promise<void> {
    if (this.current) await this.current.close();
    await this.writePart('xl/styles.xml', this.styles.toXml());
    await this.writePart('xl/workbook.xml', workbookXml(this.sheets.map((s) => ({ name: s.name, autoFilter: s.autoFilterRange }))));
    await this.writePart('xl/_rels/workbook.xml.rels', workbookRelsXml(this.sheets.length));
    await this.writePart('_rels/.rels', rootRelsXml());
    await this.writePart('[Content_Types].xml', contentTypesXml(this.sheets.length));
    await this.zip.finish();
    this.sink.close();
  }

  private async writePart(name: string, xml: string): Promise<void> {
    const entry = this.zip.beginEntry(name);
    await entry.write(ENCODER.encode(xml));
    await entry.close();
  }

  private fail(reason: unknown): void {
    this.zip.fail(reason);
    for (const sheet of this.sheets) sheet.fail(reason);
  }

  private assertOpen(method: string): void {
    if (this.closed) throw new Error(`Cannot call ${method}() after the workbook was closed`);
  }
}
```

`src/index.ts` (замена placeholder):
```ts
import type { Workbook, WorkbookOptions } from './types.js';
import { WorkbookImpl } from './workbook.js';

export type * from './types.js';

/**
 * Creates a streaming XLSX workbook. Pipe `workbook.readable` to its destination,
 * add sheets and rows, then `await workbook.close()`.
 */
export function createWorkbook(options?: WorkbookOptions): Workbook {
  return new WorkbookImpl(options);
}
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npx vitest run test/workbook.test.ts`
Expected: PASS (4 tests). Если exceljs не читает дату как `Date` в `C2`, проверить, что в `styles.xml` ячейка ссылается на xf с `numFmtId="164"` и формат `yyyy-mm-dd` объявлен в `numFmts`.

- [ ] **Step 5: Run the whole suite and the typecheck**

Run: `npm run typecheck && npm test`
Expected: всё зелёное.

- [ ] **Step 6: Commit**

```bash
git add src/workbook.ts src/index.ts test/workbook.test.ts
git commit -m "feat: createWorkbook — orchestrates sheets, metadata parts and the output stream"
```

---

### Task 11: Lifecycle and error-handling tests

**Files:**
- Test: `test/lifecycle.test.ts`
- Modify (only if a test exposes a gap): `src/workbook.ts`, `src/sheet.ts`

**Interfaces:**
- Consumes: `createWorkbook` (Task 10).

- [ ] **Step 1: Write the tests**

`test/lifecycle.test.ts`:
```ts
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
```

- [ ] **Step 2: Run the tests**

Run: `npx vitest run test/lifecycle.test.ts`
Expected: PASS (10 tests). Если что-то падает — это дефект в `workbook.ts`/`sheet.ts`, а не в тесте: исправить реализацию так, чтобы поведение соответствовало разделу 6 спеки, и перегнать весь набор.

- [ ] **Step 3: Commit**

```bash
git add test/lifecycle.test.ts src
git commit -m "test: workbook lifecycle, validation, abort and cancellation"
```

---

### Task 12: Backpressure test

**Files:**
- Test: `test/backpressure.test.ts`

**Interfaces:**
- Consumes: `createWorkbook` (Task 10).

- [ ] **Step 1: Write the test**

`test/backpressure.test.ts`:
```ts
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
    let maxQueued = 0;
    let produced = 0;
    let consumed = 0;
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
    for (let i = 0; i < 3000; i++) {
      await sheet.writeRow([i, noise()]);
      produced += 1024;
      maxQueued = Math.max(maxQueued, produced - consumed);
    }
    await wb.close();
    await output;
    // 1 MiB stream queue + compressor queues + one 64 KiB flush: comfortably under 4 MB of raw input.
    expect(maxQueued).toBeLessThan(4 * 1024 * 1024);
  });
});
```

- [ ] **Step 2: Run the test**

Run: `npx vitest run test/backpressure.test.ts`
Expected: PASS (2 tests) за несколько секунд. Если первый тест видит `allSettled === true`, значит backpressure не работает: проверить, что `SheetWriter.flush()` возвращает промис `entry.write()`, а `Entry.write()` ждёт `writer.write()` компрессора.

- [ ] **Step 3: Commit**

```bash
git add test/backpressure.test.ts
git commit -m "test: backpressure with a slow or absent consumer"
```

---

### Task 13: Benchmark and sample file

**Files:**
- Create: `bench/million-rows.mjs`, `examples/sample.mjs`

**Interfaces:**
- Consumes: собранный `dist/index.js` (Task 10 + `npm run build`).

- [ ] **Step 1: Write the benchmark**

`bench/million-rows.mjs`:
```js
// Usage: npm run build && node bench/million-rows.mjs [rows]
// Writes `rows` × 10 cells to a discarding sink and reports throughput and heap usage.
import { createWorkbook } from '../dist/index.js';

const ROWS = Number(process.argv[2] ?? 1_000_000);
const COLUMNS = 10;

const wb = createWorkbook();
let outputBytes = 0;
const done = wb.readable.pipeTo(
  new WritableStream({
    write(chunk) {
      outputBytes += chunk.byteLength;
    },
  }),
);

const bold = wb.style({ font: { bold: true } });
const money = wb.style({ numFmt: '#,##0.00' });
const sheet = wb.addSheet('Bench', {
  columns: Array.from({ length: COLUMNS }, () => ({ width: 16 })),
  freeze: { rows: 1 },
});
await sheet.writeRow(Array.from({ length: COLUMNS }, (_, i) => `Column ${i + 1}`), { style: bold });

const heapSamples = [];
const started = performance.now();
for (let i = 1; i <= ROWS; i++) {
  await sheet.writeRow([
    i,
    `Item ${i}`,
    { value: i * 1.5, style: money },
    i % 2 === 0,
    new Date(Date.UTC(2026, 0, 1 + (i % 365))),
    'Lorem ipsum dolor sit amet',
    i * 3,
    `Code-${i % 1000}`,
    i / 7,
    i % 10 === 0 ? null : 'end',
  ]);
  if (i % 100_000 === 0) {
    const heapMb = process.memoryUsage().heapUsed / 1024 / 1024;
    heapSamples.push(heapMb);
    console.log(`${i.toLocaleString()} rows, heap ${heapMb.toFixed(1)} MB`);
  }
}
await wb.close();
await done;
const seconds = (performance.now() - started) / 1000;

console.log('---');
console.log(`rows: ${ROWS.toLocaleString()} × ${COLUMNS} cells`);
console.log(`time: ${seconds.toFixed(1)} s (${Math.round(ROWS / seconds).toLocaleString()} rows/s)`);
console.log(`output: ${(outputBytes / 1024 / 1024).toFixed(1)} MB`);
console.log(`heap: min ${Math.min(...heapSamples).toFixed(1)} MB, max ${Math.max(...heapSamples).toFixed(1)} MB`);
```

- [ ] **Step 2: Write the sample generator for manual checks in Excel / Numbers / Google Sheets**

`examples/sample.mjs`:
```js
// Usage: npm run build && node examples/sample.mjs [path]
// Produces a small workbook that exercises every v0.1 feature; open it in Excel, Numbers and Google Sheets.
import { createWriteStream } from 'node:fs';
import { Writable } from 'node:stream';
import { createWorkbook } from '../dist/index.js';

const path = process.argv[2] ?? 'sample.xlsx';
const wb = createWorkbook();
const saved = wb.readable.pipeTo(Writable.toWeb(createWriteStream(path)));

const header = wb.style({ font: { bold: true, color: '#FFFFFF' }, fill: '#4472C4', border: 'thin', align: { horizontal: 'center' } });
const money = wb.style({ numFmt: '#,##0.00 €' });
const percent = wb.style({ numFmt: '0.0%', font: { italic: true } });
const wrapped = wb.style({ align: { wrap: true, vertical: 'top' } });
const title = wb.style({ font: { size: 16, bold: true } });

const sales = wb.addSheet('Sales', {
  columns: [{ width: 28 }, { width: 14, style: money }, { width: 10, style: percent }, { width: 14 }, { width: 40, style: wrapped }],
  freeze: { rows: 2 },
});
await sales.writeRow(['Quarterly sales report'], { style: title, height: 28 });
sales.merge('A1:E1');
await sales.writeRow(['Product', 'Revenue', 'Share', 'Updated', 'Notes'], { style: header });
const products = [
  ['Coffee & tea', 12500.5, 0.42, new Date(Date.UTC(2026, 9, 1)), 'Line one\nLine two — wrapped text'],
  ['Books "Classics"', 8300, 0.28, new Date(Date.UTC(2026, 9, 2, 14, 30)), '  leading and trailing spaces  '],
  ['Stationery <misc>', 4200.25, 0.14, new Date(Date.UTC(2026, 9, 3)), 'Unicode: Привет, 你好, 😀'],
  ['Returns', -1200, -0.04, new Date(Date.UTC(2026, 9, 4)), { link: 'https://example.com/returns', value: 'Return policy' }],
];
for (const row of products) await sales.writeRow(row);
const last = sales.nextRow - 1;
await sales.writeRow(['Total', { formula: `SUM(B3:B${last})` }, { formula: `SUM(C3:C${last})` }, null, { formula: `COUNTA(A3:A${last})&" products"` }], { style: header });
sales.autoFilter(`A2:E${last}`);

const misc = wb.addSheet('Edge cases', { columns: [{ width: 24 }, { width: 24 }] });
await misc.writeRow(['NaN becomes #NUM!', NaN]);
await misc.writeRow(['Infinity becomes #NUM!', Infinity]);
await misc.writeRow(['Invalid date becomes #NUM!', new Date(NaN)]);
await misc.writeRow(['Date before 1900 becomes #NUM!', new Date(Date.UTC(1899, 0, 1)), 'x']);
await misc.writeRow(['Booleans', true, false]);
await misc.writeRow(['Large number', 1e21]);
await misc.writeRow(['Empty styled cell →', { value: null, style: header }]);
await misc.writeRow(['Control chars stripped', 'a\u0000b\u0001c']);

await wb.close();
await saved;
console.log(`wrote ${path}`);
```

- [ ] **Step 3: Run both scripts**

```bash
npm run build && node bench/million-rows.mjs 200000 && node examples/sample.mjs
```

Expected: бенчмарк печатает heap-семплы примерно одинаковой величины (не растут с числом строк) и итог; `sample.xlsx` появляется в корне (он в `.gitignore`). Затем прогнать полный `node bench/million-rows.mjs` и записать `rows/s`, объём и heap — цифры нужны для README.

Проверить `sample.xlsx` вручную: открыть в Excel (должно открыться без диалога «восстановить»), посмотреть закрепление, фильтр, объединение, форматы, ссылку и лист «Edge cases». Эту проверку делает пользователь.

- [ ] **Step 4: Commit**

```bash
git add bench/million-rows.mjs examples/sample.mjs
git commit -m "chore: benchmark script and sample workbook generator"
```

---

### Task 14: Documentation, CI and publish preparation

**Files:**
- Create: `README.md`, `CHANGELOG.md`, `.github/workflows/ci.yml`
- Modify: `package.json` (`repository`/`homepage`, если к этому моменту есть GitHub-remote)

- [ ] **Step 1: Write README.md**

`README.md` (цифры в разделе Performance — из прогона бенчмарка в Task 13):
````markdown
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

await wb.close();
```

`wb.readable` is a standard `ReadableStream<Uint8Array>`. Pipe it wherever the bytes should go — ideally *before* writing rows, so backpressure can flow:

```ts
const saving = wb.readable.pipeTo(destination); // start consuming
// ... write rows ...
await wb.close();
await saving;
```

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
await wb.close();
await saving;
```

Nothing is held in memory: the browser writes chunks to the file as they are produced.

### Node.js file

```ts
import { createWriteStream } from 'node:fs';
import { Writable } from 'node:stream';

const wb = createWorkbook();
const saving = wb.readable.pipeTo(Writable.toWeb(createWriteStream('report.xlsx')));
// ... add sheets and rows ...
await wb.close();
await saving;
```

### HTTP response

```ts
// Fetch-style runtimes (Cloudflare Workers, Deno, Bun, Next.js route handlers):
return new Response(wb.readable, {
  headers: {
    'content-type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'content-disposition': 'attachment; filename="report.xlsx"',
  },
});

// Node.js http / Express:
import { Readable } from 'node:stream';
Readable.fromWeb(wb.readable).pipe(res);
```

### Web Worker

The library touches no DOM, so generation can run in a Worker. Post `wb.readable` to the main thread with `postMessage(readable, [readable])` and pipe it there, or build the file inside the worker and post the resulting `Blob`.

## API

### `createWorkbook(options?)`

| Option | Default | Meaning |
|---|---|---|
| `dates` | `'utc'` | Interpret `Date` values by their UTC components (`'utc'`) or local components (`'local'`). Excel has no time zones; `'utc'` produces the same file wherever it is generated, `'local'` shows the time the user sees in their browser. |
| `dateFormat` | `'yyyy-mm-dd'` | Number format for `Date` cells whose style has no `numFmt`. |

Returns a `Workbook`:

- `readable: ReadableStream<Uint8Array>` — the XLSX bytes.
- `style(definition): Style` — registers a style; equal definitions share one handle.
- `addSheet(name, options?): Sheet` — starts a sheet. Sheets are written one after another: adding a sheet closes the previous one.
- `close(): Promise<void>` — closes the current sheet, writes workbook metadata and ends the stream.
- `abort(reason?)` — fails the stream and every pending operation.

### `Sheet`

- `writeRow(cells, { style?, height? }): Promise<void>` — the promise waits only when the consumer is slower than you; `await` it to keep memory bounded.
- `writeRows(iterable | asyncIterable): Promise<void>`
- `merge('A1:C1')`, `autoFilter('A1:C1')` — may be called any time before the sheet is closed.
- `nextRow` — 1-based number of the next row, handy for formulas.
- `close(): Promise<void>`

Sheet options: `columns: [{ width?, style? }]` (widths in Excel character units) and `freeze: { rows?, columns? }`. Both must be given up front — they precede the data in the file.

### Cell values

| Value | Written as |
|---|---|
| `string` | Text (inline string) |
| `number` | Number; `NaN`/`Infinity` become the `#NUM!` error cell |
| `boolean` | Boolean |
| `Date` | Date serial with `dateFormat` unless the style has its own `numFmt`; invalid dates and dates before 1900 become `#NUM!` |
| `null`, `undefined`, `''` | Empty (a styled empty cell still gets its fill and borders) |
| `{ value, style }` | Value with a cell-level style |
| `{ formula: 'SUM(A1:A9)', value? }` | Formula with an optional cached result; Excel recalculates on open |
| `{ link: 'https://…', value? }` | Hyperlink (`HYPERLINK()` formula) in blue underlined text unless styled |

Style precedence: cell → row → column.

### Styles

```ts
wb.style({
  font: { bold, italic, underline, strike, size, color: '#RRGGBB', name },
  fill: '#RRGGBB',
  border: 'thin' | { top: 'thin', bottom: { style: 'double', color: '#00FF00' } },
  align: { horizontal: 'left' | 'center' | 'right' | …, vertical: 'top' | 'center' | 'bottom' | …, wrap: true },
  numFmt: '#,##0.00',
});
```

Border styles: `thin`, `medium`, `thick`, `dashed`, `dotted`, `double`, `hair`.

## Errors

Validation errors (`TypeError`, `RangeError`, `Error`) are thrown synchronously with a message that says what is wrong: invalid sheet names, more than 1,048,576 rows or 16,384 columns, text over 32,767 characters, malformed ranges, writing to a closed sheet, styles from another workbook. Bad *data* is not an error: `NaN`, `Infinity` and unrepresentable dates become `#NUM!` so that a long export never fails over one value.

If the consumer cancels the stream (the user stops a download), the next `writeRow()`/`close()` rejects with the cancellation reason.

## Limits

- A single sheet must stay under 4 GiB compressed and uncompressed (ZIP data-descriptor limit). Larger exports: split across sheets. The archive as a whole may exceed 4 GiB (ZIP64 central directory).
- Strings are written inline rather than through a shared-string table: this keeps memory constant; deflate removes most of the size difference.
- Hyperlinks use the `HYPERLINK()` formula, which needs no per-sheet relationship file.

## Performance

<!-- Fill in from `npm run bench` on the development machine: -->
1,000,000 rows × 10 cells: **N s** (**N rows/s**), **N MB** of output, heap flat at **N MB**.

## Supported environments

Anything with `ReadableStream`, `TextEncoder` and `CompressionStream('deflate-raw')`:
Chrome/Edge 103+, Firefox 113+, Safari 16.4+, Node.js 20.12+, Deno, Bun, Cloudflare Workers.

## Roadmap

Streaming reader, optional shared strings, native hyperlinks, conditional formatting, data validation, comments, images, rich text, pluggable compressor for older engines.

## License

MIT
````

- [ ] **Step 2: Write CHANGELOG.md**

```markdown
# Changelog

## 0.1.0

Initial release.

- Streaming XLSX writer on the Web Streams API with constant memory usage.
- Zero dependencies: native `CompressionStream('deflate-raw')` and a built-in ZIP writer (ZIP64 central directory).
- Multiple sheets; strings, numbers, booleans, dates, formulas, hyperlinks.
- Fonts, fills, borders, alignment, number formats; column widths, frozen panes, auto filters, merged cells.
- Backpressure through `writeRow()` promises; `abort()` and consumer cancellation.
```

- [ ] **Step 3: Write the CI workflow**

`.github/workflows/ci.yml`:
```yaml
name: CI

on:
  push:
    branches: [main]
  pull_request:

jobs:
  test:
    runs-on: ubuntu-latest
    strategy:
      matrix:
        node: [20, 22, 24]
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-node@v4
        with:
          node-version: ${{ matrix.node }}
          cache: npm
      - run: npm ci
      - run: npm run typecheck
      - run: npm test
      - run: npm run build
```

- [ ] **Step 4: Fill in the benchmark numbers in README**

Run: `npm run bench` (полный миллион строк), подставить результаты в раздел Performance и удалить HTML-комментарий.

- [ ] **Step 5: Verify the package contents**

Run: `npm run build && npm pack --dry-run`
Expected: в списке только `dist/**` (`.js`, `.d.ts`, `.map`), `package.json`, `README.md`, `LICENSE`; ни `src/`, ни `test/`, ни `bench/`. Размер tarball — порядка десятков килобайт.

Run: `node -e "import('xlsx-web-stream').then(m => console.log(typeof m.createWorkbook))"` из временной папки после `npm link` или проверить `node --input-type=module -e "import('./dist/index.js').then(m => console.log(typeof m.createWorkbook))"` → `function`.

- [ ] **Step 6: Commit**

```bash
git add README.md CHANGELOG.md .github/workflows/ci.yml package.json
git commit -m "docs: README, changelog and CI workflow"
```

- [ ] **Step 7: Publish (requires the user)**

1. Пользователь выполняет `npm login` (библиотека не вводит учётные данные) и при желании создаёт GitHub-репозиторий, добавляет remote и `repository`/`homepage` в `package.json`.
2. После явного подтверждения пользователя: `npm publish` (`prepublishOnly` прогонит typecheck, тесты и сборку).
3. Проверить: `npm view xlsx-web-stream version` → `0.1.0`; установить в чистую папку и выполнить пример из README.
4. `git tag v0.1.0`.

---

## Self-review (done while writing the plan)

**Spec coverage.** Раздел 3 (API) → Tasks 1, 9, 10; §3.3 (приоритет стилей, датовые варианты, стиль ссылок) → Tasks 6, 7, 9; §4 (сопоставление значений) → Task 6; §4.1 (XML листа) → Task 9; §4.2 (служебные части, `_xlnm._FilterDatabase`, `fullCalcOnLoad`) → Tasks 7, 8, 10; §5 (ZIP, ZIP64, лимит 4 ГиБ) → Task 5; §5.1 (backpressure) → Tasks 4, 12; §6 (валидация и ошибки) → Tasks 6, 7, 9, 10, 11; §7 (структура репозитория, dev-зависимости) → Tasks 1, 14; §8 (тесты) → Tasks 2–12; §9 (публикация) → Task 14. Пробелов не найдено.

**Отступления от спеки, зафиксированные в плане:** (1) `escapeText` экранирует только `& < >` в содержимом элементов — `"` и `'` экранируются в атрибутах; это корректный XML и меньший файл. (2) `{ style }` без `value` допускается как пустая стилизованная ячейка. (3) Объединение одной ячейки (`A1:A1`) отклоняется — Excel считает такой файл повреждённым. (4) Вторая загрузка dev-зависимостей: Vitest 4, а не 5 (Node 20 в CI). Спека обновлена соответствующими правками.

**Type consistency.** `CellContext` (`dates`, `styleId`, `dateStyle`, `linkStyle`) одинаков в Tasks 6, 9, 10. `ZipEntry { write, close }` одинаков в Tasks 5, 9, 10. `StyleRegistry.idOf/dateVariant/linkStyle/register/toXml` — Tasks 7, 10. `SheetWriter.autoFilterRange: Range | undefined` соответствует `SheetInfo.autoFilter: Range | undefined` (Tasks 8, 9, 10). `StreamSink.onFailure` — Tasks 4, 10. `ZipWriter.fail(reason)` — Tasks 5, 10.

