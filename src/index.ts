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
