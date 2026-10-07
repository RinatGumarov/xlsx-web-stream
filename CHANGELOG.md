# Changelog

## 0.1.0

Initial release.

- Streaming XLSX writer on the Web Streams API with constant memory usage.
- Zero dependencies: native `CompressionStream('deflate-raw')` and a built-in ZIP writer (ZIP64 central directory).
- Multiple sheets; strings, numbers, booleans, dates (`dates: 'utc' | 'local'`, `dateFormat` options), formulas, hyperlinks.
- Fonts, fills, borders, alignment, number formats; column widths, row heights, frozen panes, auto filters, merged cells.
- Backpressure through `writeRow()` promises; `abort()` and consumer cancellation.
