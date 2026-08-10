import ExcelJS from 'exceljs';
import {
  MAX_DAY_OF_MONTH,
  MIN_DAY_OF_MONTH,
  SCHEDULE_DATE_HEADER,
  SCHEDULE_HEADER_SCAN_COLS,
  SCHEDULE_HEADER_SCAN_ROWS,
  SCHEDULE_MONTH_PATTERN,
  SCHEDULE_NAME_GAP_LIMIT,
  SCHEDULE_NON_NAME_LABELS,
  SCHEDULE_YEAR_PATTERN,
} from '../constants';

/** 성명 열에서 이름으로 인정할 글자 모양 */
const PERSON_NAME_PATTERN = /^[가-힣]{2,5}$/;

/** 공백류 (전각 공백·줄바꿈 포함) */
const WHITESPACE_RUN = /\s+/g;

/**
 * 근무표 엑셀을 읽어 '어느 시트의 어느 칸이 누구의 며칠인지' 를 알아낸다.
 *
 * 시트 구조
 *      B      C        D    E    F   ...  AH
 *   2 │ 2026년 1월 재활치료부 근무표
 *   4 │      날짜      1    2    3   ...  31     ← 날짜 행
 *   5 │   요일 성명    목   금   토   ...  토
 *   6 │ PT  나명승     ·    D    ·   ...  ·      ← 사람마다 한 행
 *
 * 설계 원칙
 * - 셀 주소를 쓰지 않는다. '날짜' 라고 적힌 셀을 찾아 그 행을 날짜 행,
 *   그 열을 성명 열로 삼는다. (근무표 양식이 위아래로 밀려도 그대로 동작한다)
 * - 연도까지 확인한다. 이 파일에는 작년(2025년) 9·10월 시트가 남아 있어서
 *   월만 맞춰 쓰면 올해 연차가 작년 시트에 들어간다.
 * - 부서(PT/OT) 는 믿지 않는다. 같은 사람이 달마다 PT/OT 로 다르게 적힌 경우가 있어
 *   이름만으로 찾는다.
 */

/** 근무표 시트 하나의 구조 */
export interface ScheduleSheet {
  worksheet: ExcelJS.Worksheet;
  sheetName: string;
  year: number | null;
  month: number | null;
  /** 날짜(1~31) → 열 번호 (1-based) */
  dayColumns: Map<number, number>;
  /** 이름 → 행 번호 (1-based) */
  nameRows: Map<string, number>;
}

/** 근무표 파일 전체 */
export interface ScheduleWorkbook {
  workbook: ExcelJS.Workbook;
  sheets: ScheduleSheet[];
  /** 모든 시트에 등장한 이름 (연차표에서 사람을 가려내는 기준) */
  roster: string[];
  warnings: string[];
}

/** 근무표 파일을 읽어 구조를 분석한다. */
export async function readSchedule(file: File): Promise<ScheduleWorkbook> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await file.arrayBuffer());

  const sheets: ScheduleSheet[] = [];
  const warnings: string[] = [];
  const roster = new Set<string>();

  for (const worksheet of workbook.worksheets) {
    const sheet = analyzeSheet(worksheet);
    if (sheet === null) {
      warnings.push(`'${worksheet.name}' 시트에서 '${SCHEDULE_DATE_HEADER}' 머리글을 찾지 못해 건너뛰었습니다.`);
      continue;
    }

    if (sheet.dayColumns.size === 0) {
      warnings.push(`'${worksheet.name}' 시트에 날짜 칸이 없어 건너뛰었습니다.`);
    }

    sheets.push(sheet);
    for (const name of sheet.nameRows.keys()) {
      roster.add(name);
    }
  }

  if (roster.size === 0) {
    throw new Error('근무표에서 이름 명단을 찾지 못했습니다. 근무표 파일이 맞는지 확인해 주세요.');
  }

  return { workbook, sheets, roster: [...roster], warnings };
}

/** 시트 하나의 구조를 분석한다. 머리글을 못 찾으면 null. */
function analyzeSheet(worksheet: ExcelJS.Worksheet): ScheduleSheet | null {
  const header = findDateHeader(worksheet);
  if (header === null) {
    return null;
  }

  return {
    worksheet,
    sheetName: worksheet.name,
    year: findYear(worksheet, header.row),
    month: findMonth(worksheet, header.row),
    dayColumns: findDayColumns(worksheet, header),
    nameRows: findNameRows(worksheet, header),
  };
}

/** '날짜' 라고 적힌 셀 찾기 */
function findDateHeader(worksheet: ExcelJS.Worksheet): { row: number; col: number } | null {
  for (let row = 1; row <= SCHEDULE_HEADER_SCAN_ROWS; row += 1) {
    for (let col = 1; col <= SCHEDULE_HEADER_SCAN_COLS; col += 1) {
      if (compact(cellText(worksheet, row, col)) === SCHEDULE_DATE_HEADER) {
        return { row, col };
      }
    }
  }
  return null;
}

/** 머리글 위쪽 제목에서 연도를 읽는다. */
function findYear(worksheet: ExcelJS.Worksheet, headerRow: number): number | null {
  for (let row = 1; row < headerRow; row += 1) {
    for (let col = 1; col <= SCHEDULE_HEADER_SCAN_COLS; col += 1) {
      const matched = SCHEDULE_YEAR_PATTERN.exec(cellText(worksheet, row, col));
      if (matched?.[1] !== undefined) {
        return Number(matched[1]);
      }
    }
  }
  return null;
}

/** 시트 이름에서 월을 읽고, 없으면 위쪽 제목에서 찾는다. */
function findMonth(worksheet: ExcelJS.Worksheet, headerRow: number): number | null {
  const fromName = SCHEDULE_MONTH_PATTERN.exec(worksheet.name);
  if (fromName?.[1] !== undefined) {
    return Number(fromName[1]);
  }

  for (let row = 1; row < headerRow; row += 1) {
    for (let col = 1; col <= SCHEDULE_HEADER_SCAN_COLS; col += 1) {
      const matched = SCHEDULE_MONTH_PATTERN.exec(cellText(worksheet, row, col));
      if (matched?.[1] !== undefined) {
        return Number(matched[1]);
      }
    }
  }
  return null;
}

/** 날짜 행에서 1~31 이 적힌 열을 모은다. */
function findDayColumns(worksheet: ExcelJS.Worksheet, header: { row: number; col: number }): Map<number, number> {
  const columns = new Map<number, number>();

  for (let col = header.col + 1; col <= worksheet.columnCount; col += 1) {
    const value = worksheet.getCell(header.row, col).value;
    if (typeof value !== 'number' || !Number.isInteger(value)) {
      continue;
    }
    if (value < MIN_DAY_OF_MONTH || value > MAX_DAY_OF_MONTH || columns.has(value)) {
      continue;
    }
    columns.set(value, col);
  }

  return columns;
}

/** 성명 열을 아래로 훑어 사람마다의 행을 모은다. */
function findNameRows(worksheet: ExcelJS.Worksheet, header: { row: number; col: number }): Map<string, number> {
  const rows = new Map<string, number>();
  let gap = 0;

  for (let row = header.row + 1; row <= worksheet.rowCount; row += 1) {
    const name = compact(cellText(worksheet, row, header.col));

    // 머리글 바로 아래 '요일 성명' 같은 안내 칸은 이름이 아니다.
    if (name === '' || !isPersonName(name)) {
      if (rows.size > 0) {
        gap += 1;
        if (gap >= SCHEDULE_NAME_GAP_LIMIT) {
          break;
        }
      }
      continue;
    }

    gap = 0;
    if (!rows.has(name)) {
      rows.set(name, row);
    }
  }

  return rows;
}

/**
 * 사람 이름으로 볼 수 있는 한글 낱말인지.
 * 성명 열에는 '요일 성명' 안내나 '비고' 같은 라벨도 섞여 있어 함께 걸러 낸다.
 */
function isPersonName(text: string): boolean {
  return PERSON_NAME_PATTERN.test(text) && !SCHEDULE_NON_NAME_LABELS.some((label) => text.includes(label));
}

/** 셀 값을 문자열로 (수식 셀은 계산된 값을 쓴다) */
function cellText(worksheet: ExcelJS.Worksheet, row: number, col: number): string {
  const value = worksheet.getCell(row, col).value;

  if (value === null || value === undefined) {
    return '';
  }
  if (typeof value === 'object') {
    // 서식 있는 글자 / 수식 결과 / 하이퍼링크는 표시되는 글자만 꺼낸다.
    if ('richText' in value) {
      return value.richText.map((part) => part.text).join('');
    }
    if ('result' in value) {
      const { result } = value;
      return typeof result === 'object' ? '' : String(result ?? '');
    }
    if ('text' in value) {
      return value.text;
    }
    return '';
  }
  return String(value);
}

/** 공백을 모두 없앤 비교용 문자열 */
export function compact(text: string): string {
  return text.replace(WHITESPACE_RUN, '');
}

export { cellText };
