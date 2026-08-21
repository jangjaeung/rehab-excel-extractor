import type { CellPosition, CellValue, SheetGrid } from '../../types/excel';
import {
  BLOCK_LABEL_WORDS,
  MAX_NAME_LOOKUP_ROWS,
  MAX_PT_LOOKUP_ROWS,
  PT_NUMBER_PATTERN,
  SPRAY_ITEM_PREFIX,
  TOTAL_COUNT_HEADER,
} from '../constants';
import { getCell, getRowLength, isBlank, normalizeText, toNumberOrNull, toText } from './cell';

/**
 * 치료사 블록을 찾는 공용 로직.
 *
 * 「주간 환자치료 타임 현황관리판」의 블록 구조는 신장분사든 감염치료건수든 동일하다.
 *
 *      B       C            D            E~H       I
 *  5 │ PT팀장                7월        1일~4일  합계건수   ← 블록 머리글
 *  6 │ 치료사  풀타임치료유무
 *  7 │ 허정훈  감염치료건수     4    5    7        16       ← 항목 행
 *  8 │ PT288   도수치료건수  신장분사A20
 *
 * 블록은 위아래로만이 아니라 좌우로도 나란히 반복되므로,
 * '합계건수' 헤더를 기준으로 각 블록의 열 범위를 먼저 확정해야
 * 옆 블록의 값이나 PT번호를 잘못 가져오지 않는다.
 */

/** 치료사 블록의 가로 범위 */
export interface BlockBounds {
  /** 합계건수 열. 헤더를 찾지 못하면 null */
  readonly totalColumn: number | null;
  /** 합계건수 헤더가 있는 행 (= 날짜가 적힌 머리글 행). 못 찾으면 null */
  readonly headerRow: number | null;
  /** 블록의 왼쪽 경계 (inclusive) */
  readonly startCol: number;
  /** 블록의 오른쪽 경계 (inclusive) */
  readonly endCol: number;
}

/** 블록에서 찾아낸 치료사 정보 */
export interface TherapistIdentity {
  readonly therapist: string;
  /** PT 번호. 시트에 적혀 있지 않으면 빈 문자열 */
  readonly pt: string;
}

/** 오른쪽 경계를 알 수 없을 때 사용하는 값 (해당 행 끝까지) */
export const UNBOUNDED_COLUMN = Number.MAX_SAFE_INTEGER;

/**
 * '합계건수' 헤더 셀 위치를 모두 찾는다.
 * 블록이 세로/가로로 반복되므로 목록으로 보관한다.
 */
export function findTotalHeaders(grid: SheetGrid): CellPosition[] {
  const headers: CellPosition[] = [];

  for (let row = 0; row < grid.length; row += 1) {
    const length = getRowLength(grid, row);
    for (let col = 0; col < length; col += 1) {
      if (normalizeText(getCell(grid, row, col)).includes(TOTAL_COUNT_HEADER)) {
        headers.push({ row, col });
      }
    }
  }

  return headers;
}

/**
 * 셀이 속한 블록의 가로 범위와 합계 열을 확정한다.
 *
 * 좌우 블록의 시작 행이 서로 어긋나는 양식이 있어서(한쪽이 몇 행 밀려 시작한다)
 * "같은 행에 있는 헤더" 로 경계를 잡으면 안 된다. 열을 기준으로 잡는다.
 *
 * 1. 오른쪽 경계 = 항목보다 오른쪽에서 **가장 가까운 열**의 합계건수 헤더.
 *    그 열이 이 블록의 합계 열이다.
 * 2. 왼쪽 경계 = 항목보다 왼쪽에 있는 합계건수 헤더 중 **가장 오른쪽 열**의 다음 열.
 *    그 헤더가 왼쪽 블록의 끝이기 때문이다.
 *
 * 두 경우 모두 항목보다 위(같은 행 포함)에 있는 헤더만 본다.
 * 그 아래 헤더는 다른 블록의 것이다.
 */
export function resolveBlockBounds(headers: readonly CellPosition[], item: CellPosition): BlockBounds {
  // 항목보다 위쪽 헤더를 우선 사용하고, 없으면 전체 헤더를 대상으로 한다.
  const above = headers.filter((header) => header.row <= item.row);
  const candidates = above.length > 0 ? above : headers;

  let right: CellPosition | null = null;
  for (const header of candidates) {
    if (header.col <= item.col) {
      continue;
    }
    // 열이 가까울수록, 같은 열이면 행이 가까울수록(아래일수록) 우선한다.
    const isCloser = right === null || header.col < right.col || (header.col === right.col && header.row > right.row);
    if (isCloser) {
      right = header;
    }
  }

  if (right === null) {
    // 오른쪽에서 헤더를 찾지 못한 경우 (헤더가 아예 없는 양식 등)
    return { totalColumn: null, headerRow: null, startCol: 0, endCol: UNBOUNDED_COLUMN };
  }

  // 항목 왼쪽에서 가장 오른쪽에 있는 헤더 = 왼쪽 블록의 오른쪽 끝
  let leftBoundary = 0;
  for (const header of candidates) {
    if (header.col < item.col) {
      leftBoundary = Math.max(leftBoundary, header.col + 1);
    }
  }

  return { totalColumn: right.col, headerRow: right.row, startCol: leftBoundary, endCol: right.col };
}

/**
 * 항목 셀에서 위로 거슬러 올라가며 PT 번호를 찾고, 그 위에서 이름을 찾는다.
 *
 * 탐색은 블록의 열 범위 안에서, **블록 머리글 아래까지만** 한다.
 * PT 번호는 항상 머리글과 항목 사이에 적히기 때문이다.
 * 머리글을 넘어가면 PT 번호를 빠뜨린 블록에서 위 블록의 번호를 집어
 * 남의 실적으로 붙어 버린다.
 */
export function findTherapistIdentity(
  grid: SheetGrid,
  item: CellPosition,
  bounds: BlockBounds,
  totalHeaders: readonly CellPosition[],
): TherapistIdentity | null {
  // 블록 머리글보다 위는 다른 블록이다.
  const blockTop = bounds.headerRow === null ? 0 : bounds.headerRow + 1;
  const lowestRow = Math.max(0, blockTop, item.row - MAX_PT_LOOKUP_ROWS);
  const searchEndCol = Math.min(bounds.endCol, item.col);

  for (let row = item.row; row >= lowestRow; row -= 1) {
    const found = findPtNumberInRow(grid, row, bounds.startCol, searchEndCol);
    if (found === null) {
      continue;
    }

    return { therapist: findNameAbove(grid, found.position, bounds), pt: found.pt };
  }

  // PT 번호를 빠뜨린 블록이 있어(주차마다 적는 방식이 조금씩 다르다) 이름만으로도 찾아 본다.
  const therapist = findNameWithoutPt(grid, item, bounds, totalHeaders);
  return therapist === '' ? null : { therapist, pt: '' };
}

/**
 * 한 행의 [startCol, endCol] 구간에서 PT 번호 셀을 찾는다.
 * 항목에서 가장 가까운 PT 번호를 쓰기 위해 오른쪽부터 탐색한다.
 */
function findPtNumberInRow(
  grid: SheetGrid,
  row: number,
  startCol: number,
  endCol: number,
): { pt: string; position: CellPosition } | null {
  const lastCol = Math.min(endCol, getRowLength(grid, row) - 1);

  for (let col = lastCol; col >= startCol; col -= 1) {
    const pt = extractPtNumber(getCell(grid, row, col));
    if (pt !== null) {
      return { pt, position: { row, col } };
    }
  }

  return null;
}

/**
 * 셀에서 PT 번호를 추출하여 'PT288' 형태로 정규화한다.
 * 'pt 288', 'PT-288' 등도 같은 값으로 취급하여 중복 집계를 막는다.
 * 'PT팀장', 'PT부팀장' 처럼 숫자가 없는 셀은 PT 번호가 아니다.
 */
export function extractPtNumber(value: CellValue): string | null {
  const match = PT_NUMBER_PATTERN.exec(normalizeText(value));
  if (match === null) {
    return null;
  }
  const digits = match[1];
  return digits === undefined ? null : `PT${digits}`;
}

/**
 * PT 번호 셀의 바로 위 행부터 위로 올라가며 이름 셀을 찾는다.
 *
 * 이름은 PT 번호와 **같은 열**에 적히므로 그 열을 끝까지 먼저 훑는다.
 * 한 행씩 내려가며 같은 열과 블록 전체를 번갈아 보면,
 * 이름과 PT 번호 사이에 다른 행이 끼어들었을 때
 * 옆 칸의 고정 라벨(감염치료 상세 등)을 이름으로 잘못 집는다.
 * (주차마다 블록에 행이 하나씩 늘거나 줄어드는 양식이 있다)
 *
 * 같은 열에서 못 찾으면 그때 블록 범위를 훑는다. 두 경우 모두 고정 라벨은 건너뛴다.
 */
function findNameAbove(grid: SheetGrid, ptPosition: CellPosition, bounds: BlockBounds): string {
  const highestRow = Math.max(0, ptPosition.row - MAX_NAME_LOOKUP_ROWS);

  for (let row = ptPosition.row - 1; row >= highestRow; row -= 1) {
    const value = getCell(grid, row, ptPosition.col);
    if (isPersonName(value)) {
      return toText(value);
    }
  }

  for (let row = ptPosition.row - 1; row >= highestRow; row -= 1) {
    const lastCol = Math.min(bounds.endCol, getRowLength(grid, row) - 1);
    for (let col = bounds.startCol; col <= lastCol; col += 1) {
      const value = getCell(grid, row, col);
      if (isPersonName(value)) {
        return toText(value);
      }
    }
  }

  return '';
}

/** 이름 후보이면서 블록의 고정 라벨(치료사·감염치료건수·감염치료 상세 …)이 아닌 셀 */
function isPersonName(value: CellValue): boolean {
  return isNameCandidate(value) && !isBlockLabel(value);
}

/** 블록마다 반복되는 고정 라벨인지 (사람 이름이 아니다) */
function isBlockLabel(value: CellValue): boolean {
  const text = normalizeText(value);
  return text.startsWith(SPRAY_ITEM_PREFIX) || BLOCK_LABEL_WORDS.some((word) => text.includes(word));
}

/**
 * PT 번호가 없는 블록에서 이름만 찾는다.
 *
 * 블록의 첫 열(이름·PT번호가 적히는 열)을 위로 훑되,
 * 그 블록의 머리글('합계건수' 가 있는 행)을 만나면 멈춰 위 블록을 넘겨다보지 않는다.
 * 항목명이나 '치료사' 같은 고정 라벨은 이름이 아니므로 건너뛴다.
 */
function findNameWithoutPt(
  grid: SheetGrid,
  item: CellPosition,
  bounds: BlockBounds,
  totalHeaders: readonly CellPosition[],
): string {
  for (let row = item.row - 1; row >= 0; row -= 1) {
    const isBlockHeader = totalHeaders.some(
      (header) => header.row === row && header.col >= bounds.startCol && header.col <= bounds.endCol,
    );
    if (isBlockHeader) {
      return '';
    }

    const value = getCell(grid, row, bounds.startCol);

    if (isBlank(value) || extractPtNumber(value) !== null || isBlockLabel(value)) {
      continue;
    }

    return toText(value);
  }

  return '';
}

/** 이름 후보 셀인지 판단한다. 빈 셀 / 숫자 / PT 번호는 이름이 아니다. */
export function isNameCandidate(value: CellValue): boolean {
  if (isBlank(value)) {
    return false;
  }
  if (toNumberOrNull(value) !== null) {
    return false;
  }
  return extractPtNumber(value) === null;
}
