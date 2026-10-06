# @riguran/xlsxstream — потоковый writer XLSX. Дизайн v0.1

Дата: 2026-10-06
Статус: согласован, ожидает реализации

## 1. Цель

Библиотека для JavaScript/TypeScript, которая генерирует XLSX-файлы любого размера
с постоянным потреблением памяти: строки сериализуются, сжимаются и отдаются
потребителю по мере поступления, ничего не накапливается. Работает в браузере
(основная цель), Node.js и других средах с Web Streams API. Публикуется в npm
как `@riguran/xlsxstream`.

Вне объёма v0.1: чтение XLSX, shared strings, «настоящие» гиперссылки через
файл связей, условное форматирование, валидация данных, комментарии, картинки,
rich text, подключаемый компрессор. Все эти вещи можно добавить позже без
поломки публичного API.

## 2. Ключевые решения

| Решение | Выбор | Почему |
|---|---|---|
| Среда | Браузер + Node 20.12+ (Deno, Bun, Workers — где есть Web Streams и `CompressionStream('deflate-raw')`) | Один код на Web Streams покрывает всё; основная цель — браузер |
| Сжатие | Нативный `CompressionStream('deflate-raw')`, собственный потоковый ZIP-writer | Ноль зависимостей, маленький бандл, сжатие в нативном коде |
| Строки | Inline strings (`t="inlineStr"`) | Таблица shared strings росла бы в памяти с файлом; deflate хорошо сжимает повторы |
| Гиперссылки | Формула `HYPERLINK("url","text")` | Механизм `<hyperlinks>` требует файл связей, растущий с каждой ссылкой |
| Стили | Регистрируются заранее через `wb.style()` и передаются как handle | Нет хеширования объекта на каждую ячейку |
| Даты | По умолчанию UTC-компоненты, опция `dates: 'local'` | Детерминированный файл независимо от места генерации; `local` — для «как видит пользователь» |
| Служебные части (`styles.xml`, `workbook.xml`, `[Content_Types].xml`) | Пишутся при закрытии книги, после листов | Только тогда известны все стили и листы; порядок записей в ZIP читателям не важен |
| Формат пакета | TypeScript, только ESM, сборка `tsc` | Зависимостей нет — бандлить нечего |
| Тесты | Vitest в Node; независимый распаковщик, XML-парсер и `exceljs` как dev-зависимости | Проверяем и структуру, и читаемость сторонним читателем |

## 3. Публичный API

```ts
import { createWorkbook } from '@riguran/xlsxstream';

const wb = createWorkbook({ dates: 'utc', dateFormat: 'yyyy-mm-dd' }); // оба поля необязательны

// Выход — стандартный Web ReadableStream<Uint8Array>
const saved = wb.readable.pipeTo(destination);
// Браузер: const blob = await new Response(wb.readable).blob();

const header = wb.style({ font: { bold: true }, fill: '#EEEEEE', border: 'thin' });
const money  = wb.style({ numFmt: '#,##0.00', align: { horizontal: 'right' } });

const sheet = wb.addSheet('Отчёт', {
  columns: [{ width: 30 }, { width: 14, style: money }, { width: 12 }],
  freeze: { rows: 1 },
});

await sheet.writeRow(['Товар', 'Сумма', 'Дата'], { style: header });
for await (const item of items) {
  await sheet.writeRow([item.name, item.total, item.date]);
}
await sheet.writeRow(['Итого', { formula: 'SUM(B2:B1000001)' }]);
sheet.autoFilter('A1:C1');
sheet.merge('A1000003:C1000003');

await wb.close();
await saved;
```

### 3.1. Типы

```ts
export interface WorkbookOptions {
  /** Как интерпретировать Date: по UTC-компонентам или по локальным. По умолчанию 'utc'. */
  dates?: 'utc' | 'local';
  /** Формат чисел по умолчанию для Date-ячеек без numFmt. По умолчанию 'yyyy-mm-dd'. */
  dateFormat?: string;
}

export type Color = string; // '#RRGGBB' или 'RRGGBB', регистр не важен

export type BorderStyle =
  | 'thin' | 'medium' | 'thick' | 'dashed' | 'dotted' | 'double' | 'hair';

export interface BorderSide { style: BorderStyle; color?: Color }

export interface StyleDefinition {
  font?: {
    bold?: boolean; italic?: boolean; underline?: boolean; strike?: boolean;
    size?: number; color?: Color; name?: string;
  };
  fill?: Color;
  border?: BorderStyle | {
    top?: BorderStyle | BorderSide; right?: BorderStyle | BorderSide;
    bottom?: BorderStyle | BorderSide; left?: BorderStyle | BorderSide;
  };
  align?: {
    horizontal?: 'left' | 'center' | 'right' | 'fill' | 'justify' | 'centerContinuous' | 'distributed';
    vertical?: 'top' | 'center' | 'bottom' | 'justify' | 'distributed';
    wrap?: boolean;
  };
  numFmt?: string;
}

/** Непрозрачный handle, возвращаемый wb.style(). Принадлежит конкретной книге. */
export interface Style { readonly id: number }

export type CellValue =
  | string | number | boolean | Date | null | undefined
  | { value: string | number | boolean | Date | null | undefined; style?: Style }
  | { formula: string; value?: string | number | boolean; style?: Style }
  | { link: string; value?: string; style?: Style };

export interface ColumnOptions { width?: number; style?: Style }

export interface SheetOptions {
  columns?: ColumnOptions[];
  freeze?: { rows?: number; columns?: number };
}

export interface RowOptions { style?: Style; height?: number }

export interface Sheet {
  readonly name: string;
  /** Номер следующей строки (1-based). Полезно для формул и диапазонов. */
  readonly nextRow: number;
  writeRow(cells: readonly CellValue[], options?: RowOptions): Promise<void>;
  writeRows(rows: Iterable<readonly CellValue[]> | AsyncIterable<readonly CellValue[]>): Promise<void>;
  merge(range: string): void;        // 'A1:C1'
  autoFilter(range: string): void;   // один раз на лист
  close(): Promise<void>;
}

export interface Workbook {
  readonly readable: ReadableStream<Uint8Array>;
  style(definition: StyleDefinition): Style;
  addSheet(name: string, options?: SheetOptions): Sheet;
  close(): Promise<void>;
  abort(reason?: unknown): void;
}

export function createWorkbook(options?: WorkbookOptions): Workbook;
```

### 3.2. Семантика

- **`readable`** создаётся сразу; первые байты появляются при первом `addSheet`.
  Это обычный `ReadableStream`, его можно передавать в `pipeTo`, `Response`,
  `Readable.fromWeb` и т. д. Библиотека не предоставляет функций скачивания —
  рецепты для браузера и Node живут в README.
- **`style()`** дедуплицирует: два одинаковых определения дают один `id`.
  Handle действителен только в своей книге; чужой handle → `TypeError`.
  `style()` после `close()` → `Error`.
- **`addSheet()`** синхронный. Если предыдущий лист не закрыт, он закрывается
  неявно; все операции с ZIP выстроены во внутреннюю промис-цепочку, поэтому
  запись нового листа начнётся строго после завершения предыдущего.
- **`writeRow()`** сериализует строку синхронно во внутренний буфер. Промис
  ждёт только тогда, когда буфер сброшен в компрессор и получатель не успевает.
  Если `await` забыт, порядок строк сохраняется — растёт лишь буфер.
- **`writeRows()`** — обёртка: `for await (row of rows) await writeRow(row)`.
- **`merge()` и `autoFilter()`** можно вызывать в любой момент до закрытия
  листа; они записываются в хвост XML листа.
- **`close()`** книги: закрывает текущий лист, пишет `styles.xml`,
  `workbook.xml`, связи, `[Content_Types].xml`, центральный каталог ZIP и
  закрывает `readable`. Повторный вызов возвращает тот же промис. Книга без
  листов → `Error` (Excel требует хотя бы один лист).
- **`abort(reason)`**: переводит `readable` в состояние ошибки, все ожидающие и
  последующие операции отклоняются с `reason`.
- Если потребитель отменил `readable` (например, пользователь прервал
  скачивание), следующие `writeRow`/`close` отклоняются с причиной отмены.

### 3.3. Приоритет стилей

Для каждой записанной ячейки итоговый стиль = стиль ячейки ?? стиль строки ??
стиль колонки ?? стиль по умолчанию. В OOXML стиль строки/колонки применяется
только к пустым ячейкам, поэтому writer проставляет разрешённый `s` явно на
каждую ячейку. Если итоговый стиль не задаёт `numFmt`, а значение — `Date`,
используется «датовый вариант» этого стиля (те же шрифт/заливка/границы/
выравнивание + `dateFormat` книги); варианты кэшируются в реестре.

Ячейка `{ link }` без стиля получает встроенный стиль гиперссылки
(синий `#0563C1`, подчёркивание).

## 4. Сопоставление значений и XML

| Значение | XML ячейки |
|---|---|
| `string` | `<c r="A1" t="inlineStr"><is><t>…</t></is></c>`; `xml:space="preserve"` добавляется, если строка начинается или заканчивается пробельным символом |
| `number` конечное | `<c r="A1"><v>12.5</v></c>`; текст — `String(n)` (экспоненциальная запись допустима) |
| `NaN`, `±Infinity` | `<c r="A1" t="e"><v>#NUM!</v></c>` |
| `boolean` | `<c r="A1" t="b"><v>1</v></c>` |
| `Date` | `<c r="A1" s="…"><v>серийный номер</v></c>`; невалидная дата или дата раньше 1900-01-01 → `#NUM!` |
| `null`, `undefined`, `''` | Ячейка пропускается; если у неё есть стиль (явный, строки или колонки) — пишется пустая `<c r="A1" s="…"/>`, чтобы заливка и границы отобразились |
| `{ formula }` | `<c r="A1"><f>SUM(A1:A2)</f></c>`; ведущий `=` отбрасывается. Кэшированное `value`: число → `<v>`, строка → `t="str"` + `<v>`, boolean → `t="b"` |
| `{ link, value }` | `<c r="A1" t="str" s="…"><f>HYPERLINK("url","text")</f><v>text</v></c>`; кавычки внутри экранируются удвоением; `value` по умолчанию = `link` |

Серийный номер даты: `(Date.UTC(y, m, d, h, mi, s, ms) − Date.UTC(1899, 11, 30)) / 86 400 000`,
где компоненты берутся по UTC или локально согласно `dates`. Для дат с
1900-01-01 по 1900-02-28 из результата вычитается 1, чтобы совпасть с «ошибкой
1900 года» в Excel.

Во всех текстах (значения, формулы, имена листов) экранируются `& < > " '`, а
недопустимые для XML 1.0 управляющие символы (кроме `\t \n \r`) удаляются.

### 4.1. Структура XML листа (порядок элементов по схеме)

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetViews>
    <sheetView workbookViewId="0" tabSelected="1">            <!-- tabSelected только у первого листа -->
      <pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/>
    </sheetView>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <cols><col min="1" max="1" width="30" customWidth="1" style="2"/></cols>
  <sheetData>
    <row r="1" s="1" customFormat="1" ht="20" customHeight="1">…</row>
  </sheetData>
  <autoFilter ref="A1:C1"/>
  <mergeCells count="1"><mergeCell ref="A5:C5"/></mergeCells>
</worksheet>
```

`activePane`: обе оси → `bottomRight`, только строки → `bottomLeft`, только
колонки → `topRight`. Элемент `<dimension>` не пишется (необязателен, а
размер заранее неизвестен).

### 4.2. Служебные части пакета

- `[Content_Types].xml` — Default для `rels` и `xml`; Override для
  `xl/workbook.xml`, каждого `xl/worksheets/sheetN.xml`, `xl/styles.xml`.
- `_rels/.rels` — связь officeDocument → `xl/workbook.xml`.
- `xl/workbook.xml` — `<workbookPr/>`, `<bookViews><workbookView/></bookViews>`,
  `<sheets>` с `name`/`sheetId`/`r:id`, `<definedNames>` с
  `_xlnm._FilterDatabase` (`localSheetId`, `hidden="1"`) для каждого листа с
  автофильтром, `<calcPr fullCalcOnLoad="1"/>` — чтобы формулы без кэша
  пересчитались при открытии.
- `xl/_rels/workbook.xml.rels` — `rId` на каждый лист и на `styles.xml`.
- `xl/styles.xml` — `numFmts` (пользовательские id от 164, `General` не
  регистрируется), `fonts` (индекс 0 — Calibri 11), `fills` (индексы 0 и 1 —
  обязательные `none` и `gray125`), `borders` (индекс 0 — пустая),
  `cellStyleXfs` (одна запись), `cellXfs` (индекс 0 — по умолчанию, далее
  зарегистрированные), `cellStyles` (Normal). Каждая коллекция дедуплицируется
  по ключу.
- `docProps/*` не пишутся — необязательны.

Порядок записей в ZIP: `xl/worksheets/sheet1.xml`, …, `sheetN.xml`,
`xl/styles.xml`, `xl/workbook.xml`, `xl/_rels/workbook.xml.rels`, `_rels/.rels`,
`[Content_Types].xml`.

## 5. Потоковый ZIP

- Каждая запись: local file header (сигнатура `0x04034b50`, version needed 20,
  флаги: бит 3 — data descriptor, бит 11 — имя в UTF-8; метод 8 — deflate;
  CRC и размеры в заголовке — нули), затем сжатые данные, затем data descriptor
  (`0x08074b50`, CRC32, сжатый размер, несжатый размер, по 4 байта).
- Сжатие: на запись создаётся `new CompressionStream('deflate-raw')`. Байты
  записи идут в его `writable`, параллельная «накачка» читает `readable` и
  передаёт в выход, считая сжатый размер. CRC32 считается по несжатым байтам
  до компрессора.
- Центральный каталог (`0x02014b50` на запись) и EOCD (`0x06054b50`) пишутся при
  закрытии книги. Если смещение или размер каталога превышают `0xFFFFFFFF`
  либо записей больше 65 535, используются ZIP64: extra field `0x0001` в
  записях каталога, ZIP64 EOCD record (`0x06064b50`) и locator (`0x07064b50`).
- **Ограничение:** одна запись (один лист) не может превышать 4 ГиБ ни в сжатом,
  ни в несжатом виде — размер data descriptor фиксируется до того, как размер
  известен. При превышении writer отклоняет запись с `RangeError` и
  рекомендацией разбить данные на несколько листов. Форсированный ZIP64 для
  записей — возможное расширение.
- Временная метка записей — текущее время в формате DOS.

### 5.1. Backpressure

`readable` создаётся с `ByteLengthQueuingStrategy({ highWaterMark: 1 MiB })`.
Накачка компрессора ждёт, пока `controller.desiredSize > 0`. Когда очередь
компрессора заполнена, `writer.write()` в его `writable` не разрешается, и
именно этот промис возвращает `writeRow()` при сбросе буфера. Цепочка
давления: потребитель → очередь `readable` → очередь `CompressionStream` →
`writeRow`.

Буфер XML листа — строка; сбрасывается в компрессор при достижении ~64 КиБ
(после `TextEncoder.encode`), а также при закрытии листа.

## 6. Валидация и ошибки

Все проверки — синхронные, с понятным сообщением:

| Ситуация | Ошибка |
|---|---|
| Имя листа пустое, длиннее 31 символа, содержит `[ ] : * ? / \`, начинается или заканчивается `'`, или уже занято (без учёта регистра) | `TypeError` / `Error` (дубликат) |
| Строка > 1 048 576 или колонка > 16 384 | `RangeError` |
| Текст длиннее 32 767 символов | `RangeError` (молчаливое усечение данных хуже) |
| Некорректный диапазон в `merge`/`autoFilter` (формат не `A1:B2`, выходит за пределы, второй `autoFilter`) | `TypeError` / `Error` |
| `width` вне `0..255`, `height` вне `0..409`, `freeze.rows/columns` не целое ≥ 0 | `RangeError` |
| Цвет не соответствует `#?[0-9A-Fa-f]{6}` | `TypeError` |
| Неизвестное значение `BorderStyle`, `horizontal`, `vertical` | `TypeError` |
| Handle стиля из другой книги | `TypeError` |
| `writeRow`/`merge`/`autoFilter` в закрытый лист, `addSheet`/`style` после `close`, `close` без листов | `Error` |
| Размер записи ZIP > 4 ГиБ | `RangeError` (асинхронно, через промис `writeRow`/`close`) |
| Отмена `readable` потребителем или `abort()` | Промисы отклоняются с причиной |

`NaN`, `Infinity`, невалидные даты и даты до 1900 года ошибками не считаются и
записываются как `#NUM!` — экспорт на миллион строк не должен падать из-за
одного значения.

## 7. Структура репозитория

```
package.json          name @riguran/xlsxstream, version 0.1.0, type module, license MIT,
                      exports { ".": { types, default } }, files ["dist"], sideEffects false,
                      engines.node ">=20.12.0", publishConfig.access "public",
                      prepublishOnly = typecheck + test + build
tsconfig.json         strict, target ES2022, module NodeNext, lib ES2022 + DOM (для типов Web Streams)
vitest.config.ts
LICENSE               MIT
README.md             на английском: установка, пример, рецепты (скачивание в браузере,
                      File System Access, Node fs, HTTP-ответ, Web Worker), таблица поддержки
                      сред, ограничения, roadmap
CHANGELOG.md
.github/workflows/ci.yml   typecheck + test на Node 20/22/24
src/
  index.ts            публичные экспорты
  types.ts            публичные типы (раздел 3.1)
  workbook.ts         createWorkbook, оркестрация листов и служебных частей
  sheet.ts            SheetWriter: буфер, строки, хвост XML
  styles.ts           StyleRegistry + генерация styles.xml
  cell.ts             сериализация ячеек, адреса, даты
  xml.ts              экранирование и очистка
  parts.ts            шаблоны служебных частей пакета
  zip/crc32.ts
  zip/zip-writer.ts   ZipWriter: записи, descriptor, каталог, ZIP64
test/                 по файлу на модуль + интеграционные
bench/                скрипт: 1 млн строк × 10 колонок, heap и throughput
docs/superpowers/specs/   этот документ
```

Dev-зависимости: `typescript`, `vitest`, `fflate` (независимая распаковка в
тестах), `fast-xml-parser` (разбор XML в тестах), `exceljs` (чтение
результата сторонней библиотекой). Рантайм-зависимостей нет.

## 8. Тестирование

- **`crc32`** — известные векторы (`"123456789"` → `0xCBF43926`, пустой ввод →
  `0`), инкрементальный расчёт равен расчёту за один вызов.
- **`zip-writer`** — несколько записей со случайными данными распаковываются
  `fflate.unzipSync` в исходные байты; флаги и data descriptor корректны;
  кодирование записей каталога и EOCD в ZIP64 проверяется на синтетических
  смещениях > 4 ГиБ без реальной генерации таких объёмов; превышение 4 ГиБ
  на запись отклоняется (через подмену счётчика).
- **`cell`** — адреса (`1→A`, `26→Z`, `27→AA`, `16384→XFD`), серийные даты
  (`1900-01-01→1`, `1900-02-28→59`, `1900-03-01→61`, известная современная дата),
  режимы `utc`/`local`, `NaN`/`Infinity`/невалидная дата → `#NUM!`,
  экранирование и `xml:space`, формулы с кэшем и без, экранирование кавычек в
  `HYPERLINK`.
- **`styles`** — дедупликация, стабильность `id`, обязательные `fills` 0/1,
  датовые варианты, порядок элементов `<font>`.
- **`workbook` (интеграция)** — книга с двумя листами, стилями, объединениями,
  автофильтром, закреплением: распаковка → проверка `[Content_Types].xml`,
  связей, `workbook.xml` парсером → чтение `exceljs`: значения, типы, имена
  листов, ширины колонок, закреплённые строки, объединения совпадают.
- **Ошибки** — каждая строка таблицы раздела 6; поведение при `abort()` и при
  отмене `readable` потребителем; операции после `close`.
- **Backpressure** — медленный потребитель: промис `writeRow` действительно
  ожидает, объём данных в очередях ограничен.
- **Бенчмарк** (`bench/`, не входит в `npm test`) — 1 млн строк × 10 колонок:
  heap после прогрева не растёт с числом строк; throughput фиксируется в README.
- **Ручная проверка** — открытие сгенерированных файлов в Excel, Numbers и
  Google Sheets выполняет пользователь.

## 9. Публикация

1. `npm login` выполняет пользователь (библиотека не вводит учётные данные).
2. `npm publish --access public` — после явного подтверждения пользователя,
   `prepublishOnly` прогоняет проверку типов, тесты и сборку.
3. Версия `0.1.0`; дальнейшие релизы по semver, изменения — в `CHANGELOG.md`.
4. CI-workflow готов в репозитории; подключается при появлении GitHub-remote.

## 10. Возможные расширения (вне v0.1)

Потоковое чтение XLSX; опция shared strings; гиперссылки через `<hyperlinks>`;
условное форматирование; валидация данных; комментарии; картинки; rich text;
подключаемый компрессор (fallback без `CompressionStream`); форсированный ZIP64
для записей > 4 ГиБ; `docProps`.
