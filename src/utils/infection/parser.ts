import * as XLSX from 'xlsx';
import type { CellPosition, ParseResult, SheetGrid, TherapistRecord } from '../../types/excel';
import {
  BLOCK_LABEL_WORDS,
  DAY_LABEL_PATTERN,
  DEFAULT_ITEM_COUNT,
  INFECTION_RESULT_SHEET_NAME,
  INFECTION_ROW_LABEL,
  INFECTION_TOTAL_COLUMN,
  MAX_PT_LOOKUP_BELOW_ROWS,
} from '../constants';
import {
  extractPtNumber,
  findTotalHeaders,
  isNameCandidate,
  resolveBlockBounds,
  type BlockBounds,
  type TherapistIdentity,
} from '../excel/blocks';
import { getCell, getRowLength, normalizeText, toNumberOrNull, toText } from '../excel/cell';
import { buildGrid } from '../excel/grid';

/**
 * 감염치료건수 파서.
 *
 * 입력 파일은 신장분사와 같은 「주간 환자치료 타임 현황관리판」이고,
 * 시트 하나가 한 주차다. 다만 결과는 주차별로 나누지 않고
 * **모든 시트를 합쳐 1일~말일의 하루치 건수**로 만든다.
 *
 *      B       C            D        E     F     G     H     I
 *  5 │ PT팀장                7월    1일   2일   3일   4일  합계건수   ← 머리글(날짜)
 *  6 │ 치료사  풀타임치료유무
 *  7 │ 허정훈  감염치료건수          4          5     2     11      ← 이 행을 읽는다
 *  8 │ PT288   도수치료건수
 *
 * 설계 원칙
 * - 셀 주소를 쓰지 않는다. '감염치료건수' 라고 적힌 셀을 찾고,
 *   그 셀이 속한 블록의 머리글 행에서 날짜(1일, 2일 …)를 읽어 열 ↔ 날짜를 맞춘다.
 * - 블록 탐색(좌우로 반복되는 블록 나누기, 치료사 찾기)은 신장분사와 같은 로직을 쓴다.
 */

/** 한 사람의 하루치 건수를 모으는 내부 타입 */
interface InfectionAccumulator {
  therapist: string;
  pt: string;
  /** 날짜(1~31) → 건수 */
  counts: Map<number, number>;
}

// ---------------------------------------------------------------------------
// 공개 API
// ---------------------------------------------------------------------------

/** 화면에서 호출하는 진입점 */
export async function parseInfectionExcel(file: File): Promise<ParseResult> {
  const buffer = await file.arrayBuffer();
  return parseInfectionExcelBuffer(buffer);
}

/** 테스트용 동기 진입점 */
export function parseInfectionExcelBuffer(buffer: ArrayBuffer): ParseResult {
  const workbook = XLSX.read(new Uint8Array(buffer), { type: 'array' });

  if (workbook.SheetNames.length === 0) {
    throw new Error('엑셀 파일에 시트가 없습니다.');
  }

  const warnings: string[] = [];
  const people = new Map<string, InfectionAccumulator>();
  const usedSheets: string[] = [];
  let lastDay = 0;

  // 시트(주차)를 모두 훑어 하루치 건수를 한 곳에 모은다.
  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    if (sheet === undefined) {
      warnings.push(`시트 '${sheetName}' 를 읽지 못해 건너뛰었습니다.`);
      continue;
    }

    const collected = collectFromSheet(buildGrid(sheet), sheetName, people, warnings);
    if (collected.rowCount > 0) {
      usedSheets.push(sheetName);
      lastDay = Math.max(lastDay, collected.lastDay);
    }
  }

  if (usedSheets.length === 0) {
    throw new Error(`'${INFECTION_ROW_LABEL}' 이 적힌 시트를 찾지 못했습니다.`);
  }

  const columns = buildDayColumns(lastDay);
  const rows = [...people.values()].map((person) => finalizeRecord(person, lastDay));

  return {
    // 주차로 나누지 않으므로 결과는 하나다. 시트 이름은 저장 시 시트명으로 쓰인다.
    weeks: [{ sheetName: INFECTION_RESULT_SHEET_NAME, label: `1일~${String(lastDay)}일`, rows, warnings: [] }],
    columns,
    warnings: [`읽은 시트 ${String(usedSheets.length)}개: ${usedSheets.join(', ')}`, ...warnings],
  };
}

// ---------------------------------------------------------------------------
// 1. 시트 → 하루치 건수
// ---------------------------------------------------------------------------

/** 시트 하나에서 감염치료건수 행을 모두 읽어 people 에 누적한다. */
function collectFromSheet(
  grid: SheetGrid,
  sheetName: string,
  people: Map<string, InfectionAccumulator>,
  warnings: string[],
): { rowCount: number; lastDay: number } {
  const totalHeaders = findTotalHeaders(grid);
  const labelCells = findLabelCells(grid);
  let lastDay = 0;

  for (const position of labelCells) {
    const bounds = resolveBlockBounds(totalHeaders, position);

    const identity = findPersonForLabel(grid, position, bounds);
    if (identity === null) {
      warnings.push(`${sheetName} ${String(position.row + 1)}행: 치료사를 찾지 못해 제외했습니다.`);
      continue;
    }

    const dayColumns = readDayColumns(grid, bounds);
    if (dayColumns.size === 0) {
      warnings.push(`${sheetName} ${String(position.row + 1)}행 ${identity.therapist}: 날짜 칸을 찾지 못했습니다.`);
      continue;
    }

    // 주차마다 PT번호를 빠뜨린 시트가 있어(2주차 허정훈) PT번호를 key 로 쓰면 한 사람이 두 줄로 갈린다.
    // 이름으로 묶고, 어느 시트든 PT번호가 적혀 있으면 그것을 채워 둔다.
    const person = people.get(identity.therapist) ?? {
      therapist: identity.therapist,
      pt: identity.pt,
      counts: new Map<number, number>(),
    };
    if (person.pt === '') {
      person.pt = identity.pt;
    }

    for (const [day, column] of dayColumns) {
      lastDay = Math.max(lastDay, day);

      const count = toNumberOrNull(getCell(grid, position.row, column));
      if (count === null) {
        continue;
      }

      // 같은 날짜가 여러 시트에 겹쳐 나와도 더해지지 않도록 한 번만 기록한다.
      if (person.counts.has(day)) {
        warnings.push(`${identity.therapist}: ${String(day)}일 건수가 여러 번 나와 처음 값만 사용했습니다.`);
        continue;
      }
      person.counts.set(day, count);
    }

    people.set(identity.therapist, person);
  }

  return { rowCount: labelCells.length, lastDay };
}

/**
 * 감염치료건수 행의 치료사를 찾는다.
 *
 * 신장분사와 달리 이 행은 PT번호보다 **위**에 있다.
 *
 *   B7 허정훈   C7 감염치료건수   ← 라벨 행
 *   B8 PT288    C8 도수치료건수
 *
 * 그래서 신장분사처럼 위로 거슬러 올라가면 윗 블록의 PT번호를 잘못 집는다.
 * 라벨 왼쪽에서 이름을 찾고, 그 이름 아래에서 PT번호를 찾는다.
 */
function findPersonForLabel(grid: SheetGrid, label: CellPosition, bounds: BlockBounds): TherapistIdentity | null {
  for (let col = label.col - 1; col >= bounds.startCol; col -= 1) {
    const value = getCell(grid, label.row, col);
    if (!isNameCandidate(value)) {
      continue;
    }

    const text = normalizeText(value);
    if (BLOCK_LABEL_WORDS.some((word) => text.includes(word))) {
      continue;
    }

    return { therapist: toText(value), pt: findPtBelow(grid, label.row, col) };
  }

  return null;
}

/** 이름 셀 아래에서 PT 번호를 찾는다. 없으면 빈 문자열. */
function findPtBelow(grid: SheetGrid, nameRow: number, col: number): string {
  const lastRow = Math.min(nameRow + MAX_PT_LOOKUP_BELOW_ROWS, grid.length - 1);

  for (let row = nameRow + 1; row <= lastRow; row += 1) {
    const pt = extractPtNumber(getCell(grid, row, col));
    if (pt !== null) {
      return pt;
    }
  }

  return '';
}

/** '감염치료건수' 라고 적힌 셀을 모두 찾는다. */
function findLabelCells(grid: SheetGrid): CellPosition[] {
  const positions: CellPosition[] = [];

  for (let row = 0; row < grid.length; row += 1) {
    const length = getRowLength(grid, row);
    for (let col = 0; col < length; col += 1) {
      if (normalizeText(getCell(grid, row, col)) === INFECTION_ROW_LABEL) {
        positions.push({ row, col });
      }
    }
  }

  return positions;
}

/**
 * 블록 머리글 행에서 날짜 칸을 읽어 '날짜 → 열' 로 만든다.
 * 머리글에는 '7월' 처럼 날짜가 아닌 칸도 섞여 있으므로 '숫자+일' 형태만 인정한다.
 */
function readDayColumns(grid: SheetGrid, bounds: BlockBounds): Map<number, number> {
  const columns = new Map<number, number>();
  if (bounds.headerRow === null) {
    return columns;
  }

  const lastCol = Math.min(
    bounds.totalColumn === null ? bounds.endCol : bounds.totalColumn - 1,
    getRowLength(grid, bounds.headerRow) - 1,
  );

  for (let col = bounds.startCol; col <= lastCol; col += 1) {
    const matched = DAY_LABEL_PATTERN.exec(normalizeText(getCell(grid, bounds.headerRow, col)));
    const day = matched?.[1];
    if (day !== undefined) {
      columns.set(Number(day), col);
    }
  }

  return columns;
}

// ---------------------------------------------------------------------------
// 2. 결과 조립
// ---------------------------------------------------------------------------

/** 1일 ~ 말일 컬럼 + 합계 */
function buildDayColumns(lastDay: number): string[] {
  const days: string[] = [];
  for (let day = 1; day <= lastDay; day += 1) {
    days.push(`${String(day)}일`);
  }
  return [...days, INFECTION_TOTAL_COLUMN];
}

/** 빠진 날짜를 0 으로 채우고 합계를 붙인다. */
function finalizeRecord(person: InfectionAccumulator, lastDay: number): TherapistRecord {
  const items: Record<string, number> = {};
  let total = DEFAULT_ITEM_COUNT;

  for (let day = 1; day <= lastDay; day += 1) {
    const count = person.counts.get(day) ?? DEFAULT_ITEM_COUNT;
    items[`${String(day)}일`] = count;
    total += count;
  }

  items[INFECTION_TOTAL_COLUMN] = total;
  return { therapist: person.therapist, pt: person.pt, items };
}
