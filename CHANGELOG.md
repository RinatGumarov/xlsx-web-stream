# Changelog

## 0.1.1

- `merge()` no longer rescans every earlier merge when the new range starts below all of them, so one merge per row stays linear: 100,000 row-by-row merges take about 40 ms instead of 6 s, and 1,000,000 about 0.3 s. Overlapping ranges are still rejected with the same error.
- README: the `HYPERLINK()` length, formula length and overlapping-merge validation errors are listed; the File System Access and Node.js file recipes call `wb.abort(error)` on failure; a sheet over 4 GiB compressed is documented as rejected with a `RangeError`; dynamic-array formulas (`FILTER`, `SORT`, `UNIQUE`, `SEQUENCE`, …) are documented as unsupported in 0.x.

## 0.1.0

Initial release.

- Streaming XLSX writer on the Web Streams API with constant memory usage.
- Zero dependencies: native `CompressionStream('deflate-raw')` and a built-in ZIP writer (ZIP64 central directory).
- Multiple sheets; strings, numbers, booleans, dates (`dates: 'utc' | 'local'`, `dateFormat` options), formulas, hyperlinks.
- Fonts, fills, borders, alignment, number formats; column widths, row heights, frozen panes, auto filters, merged cells.
- Backpressure through `writeRow()` promises; `abort()` and consumer cancellation.
- Validation: HYPERLINK() arguments ≤ 255 chars, formulas ≤ 8192 chars, overlapping merges rejected.
