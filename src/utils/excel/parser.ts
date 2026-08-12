import * as XLSX from 'xlsx';
import type { CellPosition, ParseResult, SheetGrid, TherapistRecord } from '../../types/excel';
import { DEFAULT_ITEM_COUNT, SPRAY_ITEM_PREFIX, TOTAL_COUNT_HEADER, WEEK_LABEL_SUFFIX } from '../constants';
import {
  findTherapistIdentity,
  findTotalHeaders,
  resolveBlockBounds,
  type BlockBounds,
  type TherapistIdentity,
} from './blocks';
import { getCell, getRowLength, isBlank, normalizeText, toNumberOrNull } from './cell';
import { buildGrid } from './grid';

/**
 * 병원 치료 실적 엑셀 파서.
 *
 * 설계 원칙
 * - 셀 주소(A1, B3) 나 고정 컬럼 번호를 절대 사용하지 않는다.
 * - 오직 셀의 "내용"으로 위치를 탐색한다.
 *   · 항목      : '신장분사' 로 시작하는 셀
 *   · 날짜 구간 : 항목명 셀 오른쪽 ~ '합계건수' 헤더 열 바로 앞
 *   · 건수      : 그 날짜 구간 값의 합 (시트의 합계 셀은 수식 범위가 어긋난 사례가 있어 쓰지 않는다)
 *   · 치료사    : 항목과 같은 블록 안에서 위로 올라가며 만나는 PT 번호, 그 위 행의 이름
 * - 따라서 날짜 개수 / 날짜 컬럼 위치 / 전체 컬럼 수가 달라져도 그대로 동작한다.
 *
 * 블록 배치
 *   치료사 블록은 위아래로만 반복되는 것이 아니라 좌우로도 나란히 놓인다.
 *
 *      [허정훈 PT288 ... 합계건수]   [강지은 PT287 ... 합계건수]
 *      [권문옥 PT183 ... 합계건수]   [김미정 PT300 ... 합계건수]
 *
 *   그래서 '합계건수' 헤더를 각 블록의 오른쪽 경계로 사용하여
 *   항목이 속한 블록의 열 범위를 먼저 확정한 뒤, 그 범위 안에서만 PT 번호와 이름을 찾는다.
 *   (행 전체를 훑으면 오른쪽 블록의 항목이 왼쪽 블록 치료사에게 잘못 붙는다.)
 */

/** 치료사 블록을 조립하는 동안 사용하는 내부 누적 타입 */
interface TherapistAccumulator extends TherapistRecord {
  /** 중복 항목 감지를 위해 이미 기록한 항목명 집합 */
  readonly recordedItems: Set<string>;
}

// ---------------------------------------------------------------------------
// 공개 API
// ---------------------------------------------------------------------------

/**
 * 엑셀 파일을 읽어 치료사별 신장분사 합계건수를 추출한다.
 * UI 는 이 함수만 호출하면 되며, 파싱 관련 지식을 갖지 않는다.
 *
 * @param file 사용자가 선택하거나 드롭한 엑셀 파일
 */
export async function parseExcel(file: File): Promise<ParseResult> {
  const buffer = await file.arrayBuffer();
  return parseExcelBuffer(buffer);
}

/**
 * 엑셀 바이너리를 직접 파싱한다. (테스트 및 File 이 아닌 입력을 위한 진입점)
 */
export function parseExcelBuffer(buffer: ArrayBuffer): ParseResult {
  const workbook = XLSX.read(new Uint8Array(buffer), { type: 'array' });

  if (workbook.SheetNames.length === 0) {
    throw new Error('엑셀 파일에 시트가 없습니다.');
  }

  const warnings: string[] = [];
  const collected: { sheetName: string; extracted: SheetExtraction }[] = [];

  // 시트 하나가 한 주차다. 빈 시트('시트1' 등)와 신장분사 항목이 없는 시트는 건너뛴다.
  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    if (sheet === undefined) {
      warnings.push(`시트 '${sheetName}' 를 읽지 못해 건너뛰었습니다.`);
      continue;
    }

    const extracted = extractFromGrid(buildGrid(sheet));
    if (extracted.records.length === 0 && extracted.itemNames.size === 0) {
      continue;
    }
    collected.push({ sheetName, extracted });
  }

  if (collected.length === 0) {
    throw new Error(`'${SPRAY_ITEM_PREFIX}' 항목이 있는 시트를 찾지 못했습니다.`);
  }

  // 주차마다 등장하는 항목이 달라도 표 모양은 같아야 비교할 수 있으므로 컬럼을 하나로 합친다.
  const columns = sortItemColumns(new Set(collected.flatMap((item) => [...item.extracted.itemNames])));

  const weeks = collected.map((item, index) => ({
    sheetName: item.sheetName,
    label: `${String(index + 1)}${WEEK_LABEL_SUFFIX}`,
    rows: item.extracted.records.map((accumulator) => finalizeRecord(accumulator, columns)),
    warnings: item.extracted.warnings,
  }));

  return { weeks, columns, warnings };
}

// ---------------------------------------------------------------------------
// 1. 그리드 → 결과
// ---------------------------------------------------------------------------

/** 시트 하나에서 뽑아낸 것 (컬럼 확정 전 단계) */
interface SheetExtraction {
  records: TherapistAccumulator[];
  itemNames: Set<string>;
  warnings: string[];
}

/**
 * 그리드 전체를 순회하며 신장분사 항목을 수집하고 치료사별로 묶는다.
 */
function extractFromGrid(grid: SheetGrid): SheetExtraction {
  const warnings: string[] = [];

  // '합계건수' 헤더 위치를 먼저 모두 찾아 둔다.
  // 세로/가로로 반복되는 블록마다 하나씩 존재하므로 블록 경계로도 사용한다.
  const totalHeaders = findTotalHeaders(grid);
  if (totalHeaders.length === 0) {
    warnings.push(`'${TOTAL_COUNT_HEADER}' 헤더를 찾지 못해 항목 오른쪽의 숫자를 끝까지 더했습니다.`);
  }

  // PT 번호를 key 로 사용하여 같은 치료사가 여러 블록에 나와도 하나로 합친다.
  const therapists = new Map<string, TherapistAccumulator>();
  const itemNames = new Set<string>();
  /** PT번호 없이 이름으로만 처리한 치료사 (경고를 한 번만 남기기 위함) */
  const namedOnly = new Set<string>();

  for (const position of findSprayItemCells(grid)) {
    const itemName = normalizeText(getCell(grid, position.row, position.col));
    const bounds = resolveBlockBounds(totalHeaders, position);

    const identity = findTherapistIdentity(grid, position, bounds, totalHeaders);
    if (identity === null) {
      warnings.push(`${String(position.row + 1)}행 '${itemName}' 위쪽에서 치료사를 찾지 못해 제외했습니다.`);
      continue;
    }
    // 같은 사람의 항목마다 반복되지 않도록 치료사당 한 번만 알린다.
    if (identity.pt === '' && !namedOnly.has(identity.therapist)) {
      namedOnly.add(identity.therapist);
      warnings.push(`${identity.therapist}: 시트에 PT번호가 없어 이름으로 묶었습니다.`);
    }

    const count = readItemCount(grid, position, bounds);
    checkTotalMismatch(grid, position, bounds, { pt: identity.pt, itemName, count }, warnings);

    itemNames.add(itemName);
    addItem(therapists, identity, itemName, count, warnings);
  }

  return { records: [...therapists.values()], itemNames, warnings };
}

/**
 * '신장분사' 로 시작하는 모든 셀 위치를 위→아래, 왼→오른쪽 순서로 수집한다.
 * 항목명을 하드코딩하지 않으므로 신장분사D20 같은 항목이 추가되어도 자동 인식된다.
 */
function findSprayItemCells(grid: SheetGrid): CellPosition[] {
  const positions: CellPosition[] = [];

  for (let row = 0; row < grid.length; row += 1) {
    const length = getRowLength(grid, row);
    for (let col = 0; col < length; col += 1) {
      const text = normalizeText(getCell(grid, row, col));
      // 접두사만 있고 뒤에 항목 구분자가 없는 셀(제목 등)은 제외한다.
      if (text.startsWith(SPRAY_ITEM_PREFIX) && text.length > SPRAY_ITEM_PREFIX.length) {
        positions.push({ row, col });
      }
    }
  }

  return positions;
}

// ---------------------------------------------------------------------------
// 3. 건수 계산
// ---------------------------------------------------------------------------

/**
 * 항목 행의 건수를 **날짜 칸 값의 합**으로 계산한다.
 *
 * 원본 시트의 '합계건수' 수식 범위가 실제 항목 행과 어긋나 있는 경우가 확인되어
 * (예: 신장분사B20 의 합계 셀이 두 행 아래 C20 의 날짜를 더하고 있음)
 * 합계 셀 값을 그대로 쓰지 않고 날짜 칸을 직접 더한다.
 *
 * '합계건수' 열은 날짜 구간의 오른쪽 끝을 정하는 기준과 교차 검증용으로만 사용한다.
 */
function readItemCount(grid: SheetGrid, item: CellPosition, bounds: BlockBounds): number {
  // 날짜 구간 = 항목명 셀 오른쪽 ~ 합계건수 열 바로 앞
  const lastDateColumn = bounds.totalColumn !== null ? bounds.totalColumn - 1 : bounds.endCol;
  return sumNumericCells(grid, item.row, item.col + 1, lastDateColumn);
}

/**
 * 지정한 행의 [startCol, endCol] 구간에 있는 숫자를 모두 더한다.
 * 빈 칸은 건너뛰고, 글자가 들어 있는 셀을 만나면 다른 영역이 시작된 것으로 보고 멈춘다.
 * (합계건수 헤더가 없는 양식에서 옆 블록까지 더해 버리는 것을 막는 안전장치)
 */
function sumNumericCells(grid: SheetGrid, row: number, startCol: number, endCol: number): number {
  const lastCol = Math.min(endCol, getRowLength(grid, row) - 1);
  let sum = DEFAULT_ITEM_COUNT;

  for (let col = startCol; col <= lastCol; col += 1) {
    const value = getCell(grid, row, col);
    if (isBlank(value)) {
      continue;
    }

    const numeric = toNumberOrNull(value);
    if (numeric === null) {
      break;
    }
    sum += numeric;
  }

  return sum;
}

/**
 * 시트에 적힌 합계건수와 날짜 합계를 비교해 어긋나면 경고를 남긴다.
 * 원본 수식이 잘못된 위치를 참조하고 있는지 사용자가 확인할 수 있게 하기 위함이다.
 */
function checkTotalMismatch(
  grid: SheetGrid,
  item: CellPosition,
  bounds: BlockBounds,
  context: { pt: string; itemName: string; count: number },
  warnings: string[],
): void {
  if (bounds.totalColumn === null) {
    return;
  }

  const declared = toNumberOrNull(getCell(grid, item.row, bounds.totalColumn));
  if (declared === null || declared === context.count) {
    return;
  }

  warnings.push(
    `${context.pt} '${context.itemName}': 시트의 합계건수(${String(declared)})와 ` +
      `날짜 합계(${String(context.count)})가 다릅니다. 날짜 합계를 사용했습니다.`,
  );
}

// ---------------------------------------------------------------------------
// 4. 결과 조립
// ---------------------------------------------------------------------------

/**
 * 치료사 누적 맵에 항목 1건을 기록한다.
 * 같은 치료사의 같은 항목이 다시 나오면 최초 값을 유지하여 중복 집계를 방지한다.
 */
function addItem(
  therapists: Map<string, TherapistAccumulator>,
  identity: TherapistIdentity,
  itemName: string,
  count: number,
  warnings: string[],
): void {
  // PT 번호가 없으면 이름을 key 로 쓴다.
  const key = identity.pt === '' ? identity.therapist : identity.pt;
  let accumulator = therapists.get(key);

  if (accumulator === undefined) {
    accumulator = {
      therapist: identity.therapist,
      pt: identity.pt,
      items: {},
      recordedItems: new Set<string>(),
    };
    therapists.set(key, accumulator);
  }

  if (accumulator.recordedItems.has(itemName)) {
    warnings.push(`${identity.pt} 의 '${itemName}' 항목이 중복되어 처음 값만 사용했습니다.`);
    return;
  }

  accumulator.recordedItems.add(itemName);
  accumulator.items[itemName] = count;
}

/** 누적 데이터를 최종 레코드로 변환하고 누락된 컬럼을 0 으로 채운다. */
function finalizeRecord(accumulator: TherapistAccumulator, columns: readonly string[]): TherapistRecord {
  const items: Record<string, number> = {};
  for (const column of columns) {
    items[column] = accumulator.items[column] ?? DEFAULT_ITEM_COUNT;
  }

  return { therapist: accumulator.therapist, pt: accumulator.pt, items };
}

/**
 * 수집된 항목명을 오름차순 정렬하여 컬럼 목록을 만든다.
 * numeric 옵션 덕분에 신장분사A9 < 신장분사A20 처럼 숫자도 자연스럽게 정렬된다.
 */
function sortItemColumns(itemNames: ReadonlySet<string>): string[] {
  return [...itemNames].sort((left, right) => left.localeCompare(right, 'ko', { numeric: true }));
}
