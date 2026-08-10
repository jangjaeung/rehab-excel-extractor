import type { CellValue, SheetGrid } from '../../types/excel';

/**
 * 셀 값 / 그리드 접근을 담당하는 저수준 헬퍼 모음.
 * parser.ts 가 셀 타입 방어 로직으로 지저분해지지 않도록 분리한다.
 */

/** 천 단위 구분자, 공백 등 숫자 판별을 방해하는 문자 */
const NUMBER_NOISE_PATTERN = /[,\s]/g;

/**
 * 문자열 정규화 시 제거할 공백류.
 * JS 의 \s 는 전각 공백(U+3000), non-breaking space(U+00A0) 까지 포함하므로
 * 엑셀에서 흔히 섞여 들어오는 공백을 모두 걸러 낸다.
 */
const WHITESPACE_PATTERN = /\s+/g;

/**
 * 그리드 범위를 벗어나거나 비어 있으면 null 을 돌려주는 안전한 셀 접근자.
 * (tsconfig 의 noUncheckedIndexedAccess 대응)
 */
export function getCell(grid: SheetGrid, row: number, col: number): CellValue {
  const targetRow = grid[row];
  if (targetRow === undefined) {
    return null;
  }
  return targetRow[col] ?? null;
}

/** 해당 행의 길이 (없으면 0) */
export function getRowLength(grid: SheetGrid, row: number): number {
  return grid[row]?.length ?? 0;
}

/**
 * 셀 값을 사람이 읽는 문자열로 변환한다.
 * 빈 셀 / boolean / Date 등은 비교 대상이 아니므로 빈 문자열로 취급한다.
 */
export function toText(value: CellValue): string {
  if (value === null || typeof value === 'boolean' || value instanceof Date) {
    return '';
  }
  return String(value).trim();
}

/**
 * 비교용으로 정규화한 문자열을 만든다.
 * 엑셀에는 '신장분사 C20', '합계 건수' 처럼 공백이 섞여 들어오는 경우가 잦아
 * 모든 공백을 제거한 뒤 비교한다.
 */
export function normalizeText(value: CellValue): string {
  return toText(value).replace(WHITESPACE_PATTERN, '');
}

/**
 * 셀을 숫자로 변환한다. 숫자로 해석할 수 없으면 null.
 * '숫자 없음(null)' 과 '값이 0' 을 구분해야 하므로 0 을 반환하지 않는다.
 */
export function toNumberOrNull(value: CellValue): number | null {
  if (typeof value === 'number') {
    return Number.isFinite(value) ? value : null;
  }
  if (typeof value === 'string') {
    const cleaned = value.replace(NUMBER_NOISE_PATTERN, '');
    if (cleaned === '') {
      return null;
    }
    const parsed = Number(cleaned);
    return Number.isFinite(parsed) ? parsed : null;
  }
  return null;
}

/** 셀이 비어 있는지 여부 (공백만 있는 셀도 빈 셀로 간주) */
export function isBlank(value: CellValue): boolean {
  return normalizeText(value) === '';
}
