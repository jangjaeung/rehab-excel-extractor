import ExcelJS from 'exceljs';
import { HOLIDAY_RED_DOMINANCE, HOLIDAY_RED_MIN_RED, HOLIDAY_TITLE_SCAN_COLS } from '../constants';

/**
 * 연차표에서 공휴일·휴업일을 읽어 낸다.
 *
 * 연차표 달력의 날짜는 글자색으로 성격을 구분한다.
 *   붉은색  일요일 · 공휴일 · 휴업일
 *   파란색  토요일
 *   검은색  평일
 *
 * 그래서 '붉은 글자로 적힌 날짜' 를 쉬는 날로 본다.
 *
 * 색은 xlsx(SheetJS) 로는 읽기 어려워 exceljs 로 따로 한 번 더 읽는다.
 * (근무표 쓰기에 이미 exceljs 를 쓰고 있으므로 추가 의존성은 없다)
 */

/** 연차표에서 붉게 칠해진 날짜를 모두 모은다. (YYYY-MM-DD) */
export async function readHolidayDates(file: File): Promise<Set<string>> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await file.arrayBuffer());

  const holidays = new Set<string>();

  for (const worksheet of workbook.worksheets) {
    const own = findSheetMonth(worksheet);
    if (own === null) {
      continue;
    }

    worksheet.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        const date = toDate(cell.value);
        if (date === null || !isRed(fontColor(cell))) {
          return;
        }

        // 달력 위아래 줄에는 앞뒤 달의 날짜도 붉게 칠해져 있다.
        // 그대로 받으면 6월 1일(월) 같은 평범한 평일이 휴일로 잡히므로 자기 달만 인정한다.
        if (date.getUTCFullYear() !== own.year || date.getUTCMonth() + 1 !== own.month) {
          return;
        }

        holidays.add(formatIsoDate(date));
      });
    });
  }

  return holidays;
}

/** 시트가 나타내는 연·월 (제목 칸에 들어 있는 날짜에서 읽는다) */
function findSheetMonth(worksheet: ExcelJS.Worksheet): { year: number; month: number } | null {
  for (let col = 1; col <= HOLIDAY_TITLE_SCAN_COLS; col += 1) {
    const title = toDate(worksheet.getCell(1, col).value);
    if (title !== null) {
      return { year: title.getUTCFullYear(), month: title.getUTCMonth() + 1 };
    }
  }
  return null;
}

/** 셀 값에서 날짜를 꺼낸다. 날짜 칸은 대부분 수식이라 계산된 값도 함께 본다. */
function toDate(value: ExcelJS.CellValue): Date | null {
  if (value instanceof Date) {
    return value;
  }
  if (value !== null && typeof value === 'object' && 'result' in value && value.result instanceof Date) {
    return value.result;
  }
  return null;
}

/**
 * 셀의 글자색을 꺼낸다.
 * 타입 정의상 font 는 항상 있다고 되어 있지만, 서식이 지정되지 않은 셀에서는
 * 실제로 undefined 가 오므로 방어적으로 읽는다.
 */
function fontColor(cell: ExcelJS.Cell): Partial<ExcelJS.Color> | undefined {
  const font = cell.font as Partial<ExcelJS.Font> | undefined;
  return font?.color;
}

/** ARGB 문자열에서 각 색이 차지하는 자리 (예: FFC00000 → 알파 FF, 빨강 C0, 초록 00, 파랑 00) */
const ARGB_LENGTH = 'AARRGGBB'.length;
const CHANNEL_WIDTH = 'FF'.length;
const RED_OFFSET = 'AA'.length;
const GREEN_OFFSET = 'AARR'.length;
const BLUE_OFFSET = 'AARRGG'.length;

/** 붉은 계열 글자색인지 (빨강이 충분히 크고 다른 색보다 뚜렷하게 높은지) */
function isRed(color: Partial<ExcelJS.Color> | undefined): boolean {
  const argb = color?.argb;
  if (argb === undefined || argb.length !== ARGB_LENGTH) {
    return false;
  }

  const channel = (offset: number): number => parseInt(argb.slice(offset, offset + CHANNEL_WIDTH), 16);
  const red = channel(RED_OFFSET);

  return (
    red > HOLIDAY_RED_MIN_RED &&
    channel(GREEN_OFFSET) < red * HOLIDAY_RED_DOMINANCE &&
    channel(BLUE_OFFSET) < red * HOLIDAY_RED_DOMINANCE
  );
}

/** 2026-02-09 형태로 */
function formatIsoDate(date: Date): string {
  const year = String(date.getUTCFullYear());
  const month = String(date.getUTCMonth() + 1).padStart(2, '0');
  const day = String(date.getUTCDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}
