import * as XLSX from 'xlsx';
import type { CellValue, SheetGrid } from '../../types/excel';
import type { HalfDayPeriod, LeaveEntry, LeaveKind, LeaveParseResult, LeavePersonSummary } from '../../types/leave';
import {
  AFTERNOON_KEYWORD,
  ANNUAL_LEAVE_KEYWORD,
  DEDUCTED_KINDS,
  FAMILY_EVENT_KEYWORD,
  FULL_DAY_VALUE,
  HALF_DAY_KEYWORD,
  HALF_DAY_VALUE,
  LEAVE_COUNTER_PATTERN,
  LEAVE_KIND_KEYWORDS,
  LEAVE_QUALIFIERS,
  MAX_DATE_SERIAL,
  MAX_WEEK_BLOCK_ROWS,
  MIN_DATE_CELLS_IN_ROW,
  MIN_DATE_SERIAL,
  MORNING_KEYWORD,
  PUBLIC_LEAVE_KEYWORD,
  SPECIAL_DUTY_BRACKET_PATTERN,
  SPECIAL_DUTY_KEYWORD,
  WEEKDAY_LABELS,
} from '../constants';
import { getCell, getRowLength, toText } from '../excel/cell';
import { buildGrid } from '../excel/grid';

/**
 * 연차 계획표(월별 달력 시트) 파서. OT/PT 두 파일 모두 같은 형식이다.
 *
 * 시트 구조
 *   시트 하나가 한 달이고, 안에 달력이 그려져 있다.
 *
 *        A(일)   C(월)   E(화)   G(수)   I(목)   K(금)   M(토)
 *    3 │  28      29      30      31       1       2       3     ← 날짜 행
 *    4 │                                  신 정                   ┐
 *    5 │                                                          │ 내용 영역
 *    6 │                                                          │ (병합 셀)
 *    7 │                                                          ┘
 *    8 │   4       5       6       7       8       9      10     ← 다음 날짜 행
 *
 * 설계 원칙
 * - 셀 주소나 고정 행 번호를 쓰지 않는다. 날짜 행은 '날짜 셀이 여러 개 있는 행'으로 찾고,
 *   요일 열은 그 행에서 날짜가 실제로 놓인 열로 정한다.
 * - 날짜는 엑셀 시리얼 값에서 직접 계산한다. (toDateFromSerial 주석 참고)
 * - **이름은 근무표 명단(roster)에 있는 것만 인정한다.**
 *   달력 칸에는 공휴일·행사·검진·퇴사 메모가 사람 이름과 뒤섞여 적히고
 *   ('지방선거(특근)', '직원검진 임재민(오후반차)', '김우진(14.5) 이호근알바 퇴사')
 *   글자 모양만으로는 사람과 행사를 구분할 수 없기 때문이다.
 */

/** 엑셀 날짜 시리얼의 기준일 (1899-12-30) */
const EXCEL_EPOCH_YEAR = 1899;
const EXCEL_EPOCH_MONTH_INDEX = 11;
const EXCEL_EPOCH_DAY = 30;
const EXCEL_EPOCH_UTC = Date.UTC(EXCEL_EPOCH_YEAR, EXCEL_EPOCH_MONTH_INDEX, EXCEL_EPOCH_DAY);

/** 하루를 밀리초로 */
const MS_PER_DAY = 86_400_000;

/** 0.5(반차)를 더하다 생기는 부동소수점 오차를 소수 한 자리에서 정리하기 위한 값 */
const DAYS_ROUNDING_FACTOR = 10;

/** 줄바꿈·연속 공백을 공백 하나로 정리하기 위한 패턴 */
const WHITESPACE_RUN = /\s+/g;

/** 한글이 아닌 문자 (숫자·기호·영문) */
const NON_KOREAN_PATTERN = /[^가-힣]/g;

/** 괄호 안 내용 */
const PAREN_CONTENT_PATTERN = /\(([^)]*)\)/;

/** 파싱에 필요한 바깥 정보 */
export interface LeaveParseOptions {
  /** 근무표에서 읽은 이름 목록. 이 목록에 있는 이름만 사람으로 인정한다. */
  roster: readonly string[];
  /** 결과에 표시할 구분 (OT / PT) */
  department: string;
}

// ---------------------------------------------------------------------------
// 공개 API
// ---------------------------------------------------------------------------

/** 화면에서 호출하는 진입점 */
export async function parseLeaveExcel(file: File, options: LeaveParseOptions): Promise<LeaveParseResult> {
  const buffer = await file.arrayBuffer();
  return parseLeaveExcelBuffer(buffer, options);
}

/** 테스트용 동기 진입점 */
export function parseLeaveExcelBuffer(buffer: ArrayBuffer, options: LeaveParseOptions): LeaveParseResult {
  const workbook = XLSX.read(new Uint8Array(buffer), { type: 'array' });

  if (workbook.SheetNames.length === 0) {
    throw new Error('엑셀 파일에 시트가 없습니다.');
  }

  const roster = [...options.roster].sort((a, b) => b.length - a.length);
  const entries: LeaveEntry[] = [];
  const warnings: string[] = [];
  // 달력에는 앞뒤 달의 날짜가 겹쳐 나온다(1월 시트 끝에 2월 1일 등).
  // 같은 날짜가 두 시트에 모두 적혀 있어도 한 번만 세도록 중복을 걸러 낸다.
  const seen = new Set<string>();

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    if (sheet === undefined) {
      warnings.push(`시트 '${sheetName}' 를 읽지 못해 건너뛰었습니다.`);
      continue;
    }
    collectFromSheet(buildGrid(sheet), sheetName, roster, options.department, entries, warnings, seen);
  }

  if (entries.length === 0) {
    warnings.push('근무표 명단에 있는 이름으로 적힌 연차 기록을 찾지 못했습니다.');
  }

  entries.sort(compareEntries);

  return {
    entries,
    people: summarizeLeaveEntries(entries),
    sheetNames: [...workbook.SheetNames],
    warnings,
  };
}

// ---------------------------------------------------------------------------
// 1. 시트 → 연차 기록
// ---------------------------------------------------------------------------

/** 시트 하나를 훑어 연차 기록을 entries 에 채운다. */
function collectFromSheet(
  grid: SheetGrid,
  sheetName: string,
  roster: readonly string[],
  department: string,
  entries: LeaveEntry[],
  warnings: string[],
  seen: Set<string>,
): void {
  const dateRows = findDateRows(grid);

  dateRows.forEach((dateRow, index) => {
    const dayColumns = findDateColumns(grid, dateRow);
    const nextDateRow = dateRows[index + 1];
    // 마지막 주는 아래에 안내문·서명란이 이어지므로 달력 칸 높이만큼만 본다.
    const lastRow = Math.min(nextDateRow === undefined ? grid.length - 1 : nextDateRow - 1, dateRow + MAX_WEEK_BLOCK_ROWS);
    // 요일 열 간격 (보통 2열). 마지막 요일의 오른쪽 끝을 정하는 데 쓴다.
    const columnSpan = dayColumns.length > 1 ? (dayColumns[1] ?? 0) - (dayColumns[0] ?? 0) : 1;

    dayColumns.forEach((column, columnIndex) => {
      const date = toDateFromSerial(getCell(grid, dateRow, column));
      if (date === null) {
        return;
      }

      const lastColumn = (dayColumns[columnIndex + 1] ?? column + columnSpan) - 1;

      for (let row = dateRow + 1; row <= lastRow; row += 1) {
        for (let col = column; col <= lastColumn; col += 1) {
          readDayCell(grid, row, col, date, sheetName, roster, department, entries, warnings, seen);
        }
      }
    });
  });
}

/** 셀 하나에서 연차 표기를 읽어 entries 에 추가한다. */
function readDayCell(
  grid: SheetGrid,
  row: number,
  col: number,
  date: Date,
  sheetName: string,
  roster: readonly string[],
  department: string,
  entries: LeaveEntry[],
  warnings: string[],
  seen: Set<string>,
): void {
  const value = getCell(grid, row, col);
  if (isDateSerial(value)) {
    return;
  }

  const text = flatten(toText(value));
  if (text === '') {
    return;
  }

  const isoDate = formatIsoDate(date);
  const parsed = parseLeaveText(text, roster);

  // 괄호가 붙었는데 아는 이름이 하나도 없으면 확인이 필요하다. ('지방선거(특근)')
  // 공휴일·행사처럼 괄호도 없는 메모는 조용히 무시한다.
  if (parsed.length === 0 && text.includes('(')) {
    warnings.push(`${sheetName} ${isoDate}: 근무표 명단에 없는 이름이라 건너뛰었습니다. → '${text}'`);
    return;
  }

  for (const item of parsed) {
    // 같은 날짜·같은 사람이 두 시트에 겹쳐 적힌 경우 한 번만 센다.
    const key = `${isoDate}|${item.name}`;
    if (seen.has(key)) {
      continue;
    }
    seen.add(key);

    // 괄호를 열어 두고 숫자를 안 적은 경우만 알려 준다. ('김우진()')
    // 괄호 없이 이름만 적는 달도 있어서, 그때까지 경고하면 확인할 것이 묻힌다.
    if (item.hasParen && item.counter === null && item.kind === ANNUAL_LEAVE_KEYWORD) {
      warnings.push(`${sheetName} ${isoDate} ${item.name}: 괄호 안 숫자가 비어 있습니다. → '${text}'`);
    }

    if (item.half && item.halfPeriod === null) {
      warnings.push(`${sheetName} ${isoDate} ${item.name}: 반차인데 오전/오후 표기가 없습니다. → '${text}'`);
    }

    entries.push({
      name: item.name,
      department,
      isoDate,
      monthDay: `${String(date.getUTCMonth() + 1)}/${String(date.getUTCDate())}`,
      weekday: WEEKDAY_LABELS[date.getUTCDay()] ?? '',
      kind: item.kind,
      half: item.half,
      halfPeriod: item.halfPeriod,
      days: item.kind === SPECIAL_DUTY_KEYWORD ? 0 : item.half ? HALF_DAY_VALUE : FULL_DAY_VALUE,
      counter: item.counter,
      sheetName,
      raw: text,
    });
  }
}

// ---------------------------------------------------------------------------
// 2. 달력 좌표 찾기
// ---------------------------------------------------------------------------

/** 날짜 셀이 여러 개 들어 있는 행 = 달력의 한 주 머리글 */
function findDateRows(grid: SheetGrid): number[] {
  const rows: number[] = [];

  for (let row = 0; row < grid.length; row += 1) {
    let count = 0;
    const length = getRowLength(grid, row);
    for (let col = 0; col < length; col += 1) {
      if (isDateSerial(getCell(grid, row, col))) {
        count += 1;
      }
    }
    if (count >= MIN_DATE_CELLS_IN_ROW) {
      rows.push(row);
    }
  }

  return rows;
}

/** 날짜 행에서 실제로 날짜가 놓인 열 목록 (= 요일 열) */
function findDateColumns(grid: SheetGrid, row: number): number[] {
  const columns: number[] = [];
  const length = getRowLength(grid, row);

  for (let col = 0; col < length; col += 1) {
    if (isDateSerial(getCell(grid, row, col))) {
      columns.push(col);
    }
  }

  return columns;
}

// ---------------------------------------------------------------------------
// 3. 셀 텍스트 → 연차 표기
// ---------------------------------------------------------------------------

/** 셀에서 뽑아낸 표기 하나 */
interface ParsedLeave {
  name: string;
  kind: LeaveKind;
  half: boolean;
  halfPeriod: HalfDayPeriod | null;
  counter: string | null;
  /** 괄호가 붙은 표기였는지 (누적 일수를 적을 자리가 있었는지) */
  hasParen: boolean;
}

/**
 * 셀 텍스트에서 연차/반차/공가/경조 표기를 모두 뽑는다.
 *
 * 한 칸에 여러 명이 적히고 표기 방식도 제각각이라
 * ('공학중(연차11) 김우진(연차8)', '나명승오후반차 (3.5)', '허정훈 경조휴가1', '오지석 노미경(오후반차)')
 * 형태를 규칙화하는 대신 **명단에 있는 이름의 위치를 기준으로 문장을 나눈다.**
 * 이름이 나온 자리부터 다음 이름 직전까지가 그 사람의 기록이다.
 */
function parseLeaveText(text: string, roster: readonly string[]): ParsedLeave[] {
  // 공백을 없앤 뒤 훑는다. '윤 송' 처럼 이름 안에 공백이 들어가거나
  // '오후 반차' / '오후반차' 가 섞여 있어도 같은 규칙으로 읽기 위함이다.
  const scanText = text.replace(WHITESPACE_RUN, '');
  const hits = findRosterNames(scanText, roster);
  const brackets = findBracketRanges(scanText);

  return hits.flatMap((hit, index) => {
    const scope = scanText.slice(hit.start, hits[index + 1]?.start ?? scanText.length);
    const paren = PAREN_CONTENT_PATTERN.exec(scope)?.[1] ?? null;

    // 대괄호로 묶인 사람은 쉬는 것이 아니라 특근이다. ('[홍길동]')
    if (brackets.some((range) => hit.start >= range.start && hit.start < range.end)) {
      return [{ name: hit.name, kind: SPECIAL_DUTY_KEYWORD, half: false, halfPeriod: null, counter: null, hasParen: false }];
    }

    // 괄호가 없는데 이름 말고 다른 낱말이 남으면 연차 기록이 아니다.
    // ('이찬규교육', '이호근알바 퇴사', '오혜원 퇴사')
    if (paren === null && leftoverText(scope, hit.name) !== '') {
      return [];
    }

    // 첫 기록 앞에 붙은 설명은 그 사람의 것으로 본다. ('직원검진 임재민(오후반차)')
    const reasonText = index === 0 ? scanText.slice(0, hit.start) + scope : scope;
    const half = scope.includes(HALF_DAY_KEYWORD);

    return [
      {
        name: hit.name,
        kind: resolveKind(reasonText),
        half,
        halfPeriod: half ? findHalfPeriod(scope) : null,
        counter: paren === null ? null : (LEAVE_COUNTER_PATTERN.exec(paren)?.[0] ?? null),
        hasParen: paren !== null,
      },
    ];
  });
}

/** 대괄호 구간 (특근 표기) */
function findBracketRanges(text: string): { start: number; end: number }[] {
  const pattern = new RegExp(SPECIAL_DUTY_BRACKET_PATTERN.source, 'g');
  const ranges: { start: number; end: number }[] = [];

  let match = pattern.exec(text);
  while (match !== null) {
    ranges.push({ start: match.index, end: pattern.lastIndex });
    match = pattern.exec(text);
  }

  return ranges;
}

/**
 * 문장에서 명단에 있는 이름이 나오는 위치를 앞에서부터 찾는다.
 * 같은 자리에서는 긴 이름을 우선하고, 찾은 이름끼리 겹치지 않게 건너뛴다.
 */
function findRosterNames(text: string, roster: readonly string[]): { name: string; start: number }[] {
  const hits: { name: string; start: number }[] = [];

  let index = 0;
  while (index < text.length) {
    const found = roster.find((name) => text.startsWith(name, index));
    if (found === undefined) {
      index += 1;
      continue;
    }
    hits.push({ name: found, start: index });
    index += found.length;
  }

  return hits;
}

/** 이름과 수식어를 빼고 남는 한글 (남는 게 있으면 연차 기록이 아니다) */
function leftoverText(scope: string, name: string): string {
  let rest = scope.replace(name, '');
  for (const qualifier of LEAVE_QUALIFIERS) {
    rest = rest.split(qualifier).join('');
  }
  return rest.replace(NON_KOREAN_PATTERN, '');
}

/**
 * 종류 판정.
 * 키워드 표에서 먼저 맞는 것을 쓰고, 아무것도 없으면 연차로 본다.
 */
function resolveKind(scope: string): LeaveKind {
  return LEAVE_KIND_KEYWORDS.find((candidate) => scope.includes(candidate.keyword))?.kind ?? ANNUAL_LEAVE_KEYWORD;
}

/**
 * 반차를 오전에 썼는지 오후에 썼는지 찾는다.
 * 시트 안내문에 '반차 사용시 오전/오후 인지 기록할 것' 이라고 적혀 있지만
 * 빠뜨린 표기도 있을 수 있으므로 못 찾으면 null 을 돌려준다.
 */
function findHalfPeriod(scope: string): HalfDayPeriod | null {
  const morning = scope.indexOf(MORNING_KEYWORD);
  const afternoon = scope.indexOf(AFTERNOON_KEYWORD);

  if (morning === -1 && afternoon === -1) {
    return null;
  }
  if (morning === -1) {
    return AFTERNOON_KEYWORD;
  }
  if (afternoon === -1) {
    return MORNING_KEYWORD;
  }
  return morning < afternoon ? MORNING_KEYWORD : AFTERNOON_KEYWORD;
}

// ---------------------------------------------------------------------------
// 4. 날짜 변환
// ---------------------------------------------------------------------------

/** 날짜로 볼 수 있는 엑셀 시리얼 값인지 (달력의 '6' 같은 일반 숫자를 걸러 낸다) */
function isDateSerial(value: CellValue): boolean {
  return typeof value === 'number' && Number.isFinite(value) && value >= MIN_DATE_SERIAL && value <= MAX_DATE_SERIAL;
}

/**
 * 엑셀 날짜 시리얼을 날짜로 바꾼다.
 *
 * xlsx 의 cellDates 옵션을 쓰지 않는 이유:
 * 이 옵션은 2026-01-01 을 '2025-12-31 23:59:08' 같은 값으로 만들어 놓기 때문에
 * 로컬 시간대로 읽으면 날짜가 하루씩 밀린다.
 * 시리얼 값에서 UTC 기준으로 직접 계산하면 시간대와 무관하게 정확하다.
 */
function toDateFromSerial(value: CellValue): Date | null {
  if (!isDateSerial(value) || typeof value !== 'number') {
    return null;
  }
  return new Date(EXCEL_EPOCH_UTC + Math.round(value) * MS_PER_DAY);
}

/** 2026-02-09 형태로 */
function formatIsoDate(date: Date): string {
  const year = String(date.getUTCFullYear());
  const month = String(date.getUTCMonth() + 1).padStart(2, '0');
  const day = String(date.getUTCDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

// ---------------------------------------------------------------------------
// 5. 정리
// ---------------------------------------------------------------------------

/** 줄바꿈·전각 공백 등을 공백 하나로 정리한다. */
function flatten(text: string): string {
  return text.replace(WHITESPACE_RUN, ' ').trim();
}

/** 날짜 오름차순, 같은 날이면 이름 가나다순 */
function compareEntries(a: LeaveEntry, b: LeaveEntry): number {
  if (a.isoDate !== b.isoDate) {
    return a.isoDate < b.isoDate ? -1 : 1;
  }
  return a.name.localeCompare(b.name, 'ko');
}

/** 연차에서 차감되는 종류인지 (공가·경조는 차감되지 않는다) */
function isDeducted(kind: LeaveKind): boolean {
  return DEDUCTED_KINDS.includes(kind);
}

/** 이름별로 묶어 합계를 낸다. (사용 일수 많은 순 → 이름순) */
export function summarizeLeaveEntries(entries: readonly LeaveEntry[]): LeavePersonSummary[] {
  const byName = new Map<string, LeavePersonSummary>();

  for (const entry of entries) {
    const found = byName.get(entry.name) ?? {
      name: entry.name,
      department: entry.department,
      fullCount: 0,
      halfCount: 0,
      morningCount: 0,
      afternoonCount: 0,
      publicCount: 0,
      familyEventCount: 0,
      specialDutyCount: 0,
      totalDays: 0,
      dates: [],
    };

    if (!isDeducted(entry.kind)) {
      // 공가·경조·특근은 연차에서 차감되지 않으므로 합계 일수에 넣지 않는다.
      if (entry.kind === PUBLIC_LEAVE_KEYWORD) {
        found.publicCount += 1;
      } else if (entry.kind === FAMILY_EVENT_KEYWORD) {
        found.familyEventCount += 1;
      } else {
        found.specialDutyCount += 1;
      }
      found.dates.push(`${entry.monthDay}(${entry.kind})`);
    } else {
      if (entry.half) {
        found.halfCount += 1;
        if (entry.halfPeriod === MORNING_KEYWORD) {
          found.morningCount += 1;
        } else if (entry.halfPeriod === AFTERNOON_KEYWORD) {
          found.afternoonCount += 1;
        }
      } else {
        found.fullCount += 1;
      }
      // 부동소수점 오차(0.5 누적)를 피하려고 소수 한 자리에서 정리한다.
      found.totalDays = Math.round((found.totalDays + entry.days) * DAYS_ROUNDING_FACTOR) / DAYS_ROUNDING_FACTOR;
      found.dates.push(entry.monthDay);
    }

    byName.set(entry.name, found);
  }

  return [...byName.values()].sort(
    (a, b) => b.totalDays - a.totalDays || a.name.localeCompare(b.name, 'ko'),
  );
}
