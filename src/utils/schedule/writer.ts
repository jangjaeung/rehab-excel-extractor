import type { LeaveEntry } from '../../types/leave';
import type { AppliedCell, ApplyStatus, ScheduleApplyResult, ScheduleSheetInfo } from '../../types/schedule';
import {
  FAMILY_EVENT_KEYWORD,
  PUBLIC_LEAVE_KEYWORD,
  SATURDAY_LABEL,
  SCHEDULE_FULL_UNIT,
  SCHEDULE_HALF_UNIT,
  SCHEDULE_OFF_MARKER,
  SCHEDULE_PLAIN_VALUES,
  SCHEDULE_REST_MARKER,
  SCHEDULE_WORKING_WEEKDAYS,
  SCHEDULE_WORK_MARKER,
  SPECIAL_DUTY_KEYWORD,
  WEEKDAY_LABELS,
} from '../constants';
import { cellText, compact, type ScheduleSheet, type ScheduleWorkbook } from './reader';

/** 'YYYY-MM-DD' 에서 연도가 차지하는 글자 수 */
const YEAR_LENGTH = 4;

/**
 * 연차 기록을 근무표에 반영한다.
 *
 * 순서가 중요하다.
 *   ① 날짜·요일 행을 해당 월에 맞게 고친다      (요일을 알아야 평일을 판단할 수 있다)
 *   ② 평일 칸을 D 로 채운다                     ('·' 로 적힌 휴무·공휴일은 그대로 둔다)
 *   ③ 연차 기록을 off 표기로 덮어쓴다
 *   ④ 근무표에 실제로 적힌 값을 세어 집계 열을 채운다
 *
 * 쓰기에 exceljs 를 쓰는 이유:
 * xlsx(SheetJS) 로 다시 쓰면 셀 배경색 등 서식이 모두 사라진다.
 * 근무표는 그대로 인쇄해서 쓰는 문서이므로 원본 서식을 유지해야 한다.
 */
export async function applyLeaveToSchedule(
  schedule: ScheduleWorkbook,
  entries: readonly LeaveEntry[],
): Promise<ScheduleApplyResult> {
  const warnings = [...schedule.warnings];

  // 연차표에 없는 연도의 시트(지난해 잔여 시트 등)는 건드리지 않는다.
  // 그대로 손대면 복구할 근거가 없는 지난 기록을 지우게 된다.
  const years = new Set(entries.map((entry) => Number(entry.isoDate.slice(0, YEAR_LENGTH))));
  const targets = schedule.sheets.filter((sheet) => sheet.year !== null && years.has(sheet.year));

  for (const sheet of schedule.sheets) {
    if (!targets.includes(sheet)) {
      warnings.push(`'${sheet.sheetName}' 는 연차표에 없는 연도(${String(sheet.year ?? 0)})라 손대지 않았습니다.`);
    }
  }

  for (const sheet of targets) {
    fixCalendar(sheet, warnings);
    fillWorkingDays(sheet, warnings);
  }

  const applied = entries.map((entry) => applyEntry(schedule, entry));

  for (const sheet of targets) {
    fillCountColumns(sheet, warnings);
  }

  const buffer = await schedule.workbook.xlsx.writeBuffer();

  return {
    data: new Uint8Array(buffer),
    applied,
    sheets: summarizeSheets(schedule),
    warnings,
  };
}

// ---------------------------------------------------------------------------
// ① 날짜·요일 맞추기
// ---------------------------------------------------------------------------

/**
 * 날짜 행을 1일~말일로, 요일 행을 실제 요일로 다시 쓴다.
 * 남는 칸은 비워서 지난달 흔적이 남지 않게 한다.
 */
function fixCalendar(sheet: ScheduleSheet, warnings: string[]): void {
  const { year, month, dateBand } = sheet;
  if (year === null || month === null) {
    warnings.push(`'${sheet.sheetName}' 는 연도/월을 읽지 못해 날짜를 맞추지 않았습니다.`);
    return;
  }
  if (dateBand.length === 0) {
    warnings.push(`'${sheet.sheetName}' 에 날짜를 넣을 열이 없어 건너뛰었습니다.`);
    return;
  }

  const lastDay = lastDayOfMonth(year, month);
  if (dateBand.length < lastDay) {
    warnings.push(
      `'${sheet.sheetName}' 의 날짜 칸이 ${String(dateBand.length)}개뿐이라 ${String(lastDay)}일까지 넣지 못했습니다.`,
    );
  }

  sheet.dayColumns.clear();

  dateBand.forEach((column, index) => {
    const day = index + 1;

    if (day > lastDay) {
      sheet.worksheet.getCell(sheet.headerRow, column).value = null;
      sheet.worksheet.getCell(sheet.weekdayRow, column).value = null;
      return;
    }

    sheet.worksheet.getCell(sheet.headerRow, column).value = day;
    sheet.worksheet.getCell(sheet.weekdayRow, column).value = weekdayLabel(year, month, day);
    sheet.dayColumns.set(day, column);
  });
}

/** 그 달의 마지막 날 */
function lastDayOfMonth(year: number, month: number): number {
  return new Date(Date.UTC(year, month, 0)).getUTCDate();
}

/** 요일 한 글자 */
function weekdayLabel(year: number, month: number, day: number): string {
  return WEEKDAY_LABELS[weekdayIndex(year, month, day)] ?? '';
}

/** 0=일 … 6=토 */
function weekdayIndex(year: number, month: number, day: number): number {
  return new Date(Date.UTC(year, month - 1, day)).getUTCDay();
}

// ---------------------------------------------------------------------------
// ② 평일 D 채우기
// ---------------------------------------------------------------------------

/**
 * 연차를 넣기 전에 근무 칸을 정리한다. (연차표를 항상 우선하기 위한 초기화)
 *
 * 평일(월~금)
 *   '·' 는 공휴일이거나 개인 휴무이므로 그대로 두고, 나머지는 D 로 채운다.
 *
 * 토·일
 *   근무 배정(D)은 근무표에서만 정하는 것이라 건드리지 않고,
 *   **off 표기만 D 로 되돌린다.**
 *   주말에 off 가 적혀 있다는 것은 원래 근무 배정(D)이 있었는데 쉬었다는 뜻이므로,
 *   연차표에 그 기록이 없다면 근무한 것으로 돌려놓는 것이 맞다.
 *   (이 정리를 안 하면 지난달 근무표의 off 가 그대로 남는다)
 */
function fillWorkingDays(sheet: ScheduleSheet, warnings: string[]): void {
  const { year, month } = sheet;
  if (year === null || month === null) {
    return;
  }

  let cleared = 0;

  for (const [day, column] of sheet.dayColumns) {
    const isWeekday = SCHEDULE_WORKING_WEEKDAYS.includes(weekdayIndex(year, month, day));

    for (const row of sheet.nameRows.values()) {
      const current = compact(cellText(sheet.worksheet, row, column));
      const hasOff = current.includes(SCHEDULE_OFF_MARKER);

      // 평일은 '·' 만 남기고 모두 D, 주말은 off 만 D 로 되돌린다.
      if (isWeekday ? current === SCHEDULE_REST_MARKER : !hasOff) {
        continue;
      }
      if (hasOff) {
        cleared += 1;
      }
      sheet.worksheet.getCell(row, column).value = SCHEDULE_WORK_MARKER;
    }
  }

  if (cleared > 0) {
    warnings.push(
      `'${sheet.sheetName}': 원래 적혀 있던 off 표기 ${String(cleared)}건을 D 로 되돌린 뒤 연차표 기준으로 다시 넣었습니다. 연차표에 없는 off 는 사라집니다.`,
    );
  }
}

// ---------------------------------------------------------------------------
// ③ 연차 기록 반영
// ---------------------------------------------------------------------------

/** 기록 하나를 근무표에 반영한다. */
function applyEntry(schedule: ScheduleWorkbook, entry: LeaveEntry): AppliedCell {
  const marker = markerFor(entry);
  const base = {
    isoDate: entry.isoDate,
    weekday: entry.weekday,
    name: entry.name,
    department: entry.department,
    marker,
  };

  const [year, month, day] = entry.isoDate.split('-').map(Number);
  const sheet = schedule.sheets.find((candidate) => candidate.year === year && candidate.month === month);

  if (sheet === undefined) {
    return {
      ...base,
      sheetName: null,
      address: null,
      status: '미반영',
      previous: null,
      reason: `${String(year)}년 ${String(month)}월 근무표 시트가 없습니다.`,
    };
  }

  const column = day === undefined ? undefined : sheet.dayColumns.get(day);
  if (column === undefined) {
    return {
      ...base,
      sheetName: sheet.sheetName,
      address: null,
      status: '미반영',
      previous: null,
      reason: `'${sheet.sheetName}' 에 ${String(day)}일 칸이 없습니다.`,
    };
  }

  const row = sheet.nameRows.get(entry.name);
  if (row === undefined) {
    return {
      ...base,
      sheetName: sheet.sheetName,
      address: null,
      status: '미반영',
      previous: null,
      reason: `'${sheet.sheetName}' 명단에 ${entry.name} 이(가) 없습니다.`,
    };
  }

  const cell = sheet.worksheet.getCell(row, column);
  const previous = cellText(sheet.worksheet, row, column);
  const status = resolveStatus(previous, marker);

  // 이미 같은 내용이면 원본 표기를 그대로 둔다. ('오후off' 처럼 띄어쓰기만 다른 경우)
  if (status !== '동일') {
    cell.value = marker;
  }

  return {
    ...base,
    sheetName: sheet.sheetName,
    address: cell.address,
    status,
    previous: status === '덮어씀' ? previous : null,
    reason: null,
  };
}

/** 원래 값과 넣을 값을 견주어 결과를 정한다. */
function resolveStatus(previous: string, marker: string): ApplyStatus {
  if (compact(previous) === compact(marker)) {
    return '동일';
  }
  return SCHEDULE_PLAIN_VALUES.includes(compact(previous)) ? '입력' : '덮어씀';
}

/**
 * 기록 → 근무표에 넣을 표기.
 *   연차 종일 → off        연차 오후 → 오후 off
 *   공가 종일 → 공가 off    공가 오후 → 오후 공가 off
 *   특근      → D          (쉬는 것이 아니라 근무이므로 근무 표시를 넣는다)
 */
function markerFor(entry: LeaveEntry): string {
  if (entry.kind === SPECIAL_DUTY_KEYWORD) {
    return SCHEDULE_WORK_MARKER;
  }

  // 오전/오후가 안 적힌 반차는 종일과 구분할 수 없으므로 종일로 넣고 경고로 남긴다.
  const period = entry.halfPeriod === null ? '' : `${entry.halfPeriod} `;
  const reason = entry.kind === PUBLIC_LEAVE_KEYWORD || entry.kind === FAMILY_EVENT_KEYWORD ? `${entry.kind} ` : '';

  return `${period}${reason}${SCHEDULE_OFF_MARKER}`;
}

// ---------------------------------------------------------------------------
// ④ 집계 열 채우기
// ---------------------------------------------------------------------------

/**
 * 근무표에 실제로 적힌 값을 세어 '오프' 와 '월차연차' 열을 채운다.
 *
 * 연차표가 아니라 근무표 칸을 세는 이유:
 * 손으로 적어 둔 표기까지 포함해야 실제 근무표와 숫자가 맞는다.
 *
 *   오프    = 모든 off (공가·경조 포함).  오전/오후는 0.5, 그 외는 1
 *   월차연차 = 공가·경조를 뺀 off 만.       특근은 어느 쪽에도 넣지 않는다
 */
function fillCountColumns(sheet: ScheduleSheet, warnings: string[]): void {
  if (sheet.offColumn === null && sheet.leaveColumn === null && sheet.saturdayColumn === null) {
    warnings.push(`'${sheet.sheetName}' 에 집계 열('오프' 등)이 없어 집계를 채우지 않았습니다.`);
    return;
  }

  const saturday = WEEKDAY_LABELS.indexOf(SATURDAY_LABEL);

  for (const row of sheet.nameRows.values()) {
    let offTotal = 0;
    let leaveTotal = 0;
    let saturdayTotal = 0;

    for (const [day, column] of sheet.dayColumns) {
      const value = compact(cellText(sheet.worksheet, row, column));

      // 토요일 근무는 토요일 칸에 근무 표시(D)가 있는 것만 센다.
      const isSaturday =
        sheet.year !== null && sheet.month !== null && weekdayIndex(sheet.year, sheet.month, day) === saturday;
      if (isSaturday && value === SCHEDULE_WORK_MARKER) {
        saturdayTotal += SCHEDULE_FULL_UNIT;
      }

      if (!value.includes(SCHEDULE_OFF_MARKER)) {
        continue;
      }

      const unit = value.includes('오전') || value.includes('오후') ? SCHEDULE_HALF_UNIT : SCHEDULE_FULL_UNIT;
      offTotal += unit;

      if (!value.includes(PUBLIC_LEAVE_KEYWORD) && !value.includes(FAMILY_EVENT_KEYWORD)) {
        leaveTotal += unit;
      }
    }

    if (sheet.offColumn !== null) {
      sheet.worksheet.getCell(row, sheet.offColumn).value = offTotal;
    }
    if (sheet.leaveColumn !== null) {
      sheet.worksheet.getCell(row, sheet.leaveColumn).value = leaveTotal;
    }
    // 근무가 없으면 '·' 대신 0 이 들어간다.
    if (sheet.saturdayColumn !== null) {
      sheet.worksheet.getCell(row, sheet.saturdayColumn).value = saturdayTotal;
    }
  }
}

/** 화면에 보여 줄 시트 구조 요약 */
function summarizeSheets(schedule: ScheduleWorkbook): ScheduleSheetInfo[] {
  return schedule.sheets.map((sheet: ScheduleSheet) => ({
    sheetName: sheet.sheetName,
    year: sheet.year,
    month: sheet.month,
    dayCount: sheet.dayColumns.size,
    peopleCount: sheet.nameRows.size,
  }));
}
