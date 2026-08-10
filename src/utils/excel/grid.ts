import * as XLSX from 'xlsx';
import type { CellValue, SheetGrid } from '../../types/excel';

/**
 * 시트를 2차원 배열로 바꾸는 공용 헬퍼.
 * 신장분사 파서와 연차 파서가 함께 사용한다.
 */

/**
 * 시트를 2차원 배열로 변환한다.
 *
 * sheet_to_json 을 쓰지 않는 이유:
 * 빈 행/빈 열이 있으면 행·열 인덱스가 밀려서 좌표 기반 탐색이 전부 어긋난다.
 * !ref 범위를 직접 순회하면 grid[row][col] 이 원본 시트 좌표와 정확히 일치한다.
 */
export function buildGrid(sheet: XLSX.WorkSheet): SheetGrid {
  const ref = sheet['!ref'];
  if (ref === undefined) {
    return [];
  }

  const range = XLSX.utils.decode_range(ref);
  const grid: CellValue[][] = [];

  for (let row = range.s.r; row <= range.e.r; row += 1) {
    const rowValues: CellValue[] = [];
    for (let col = range.s.c; col <= range.e.c; col += 1) {
      const address = XLSX.utils.encode_cell({ r: row, c: col });
      const cell = sheet[address] as XLSX.CellObject | undefined;
      rowValues.push(cell?.v ?? null);
    }
    grid.push(rowValues);
  }

  return grid;
}
