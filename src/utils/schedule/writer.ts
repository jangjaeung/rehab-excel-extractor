import type ExcelJS from 'exceljs';
import type { LeaveEntry } from '../../types/leave';
import type { AppliedCell, ApplyStatus, ScheduleApplyResult, ScheduleSheetInfo } from '../../types/schedule';
import {
  AFTERNOON_KEYWORD,
  FAMILY_EVENT_KEYWORD,
  MORNING_KEYWORD,
  PUBLIC_LEAVE_KEYWORD,
  SATURDAY_LABEL,
  SCHEDULE_DEPARTMENT_FONT_SIZE,
  SCHEDULE_FULL_UNIT,
  SCHEDULE_HALF_OFF_FONT_SIZE,
  SCHEDULE_HALF_UNIT,
  SCHEDULE_HOLIDAY_FILL_ARGB,
  SCHEDULE_MARKER_LINE_BREAK,
  SCHEDULE_OFF_FONT_SIZE,
  SCHEDULE_OFF_MARKER,
  SCHEDULE_REASON_OFF_FONT_SIZE,
  SCHEDULE_PLAIN_VALUES,
  SCHEDULE_REST_MARKER,
  SCHEDULE_WORKING_WEEKDAYS,
  SCHEDULE_WORK_FONT_NAME,
  SCHEDULE_WORK_FONT_SIZE,
  SCHEDULE_WORK_MARKER,
  SPECIAL_DUTY_KEYWORD,
  WEEKDAY_LABELS,
} from '../constants';
import { cellText, compact, type ScheduleSheet, type ScheduleWorkbook } from './reader';

/**
 * 값·서식을 적을 칸.
 *
 * 병합된 칸은 대표 칸의 값만 화면에 나온다.
 * 종속 칸에 적으면 아무 일도 일어나지 않은 것처럼 옛 값이 그대로 보이므로
 * 언제나 대표 칸을 거쳐서 쓴다.
 */
function writableCell(worksheet: ExcelJS.Worksheet, row: number, column: number): ExcelJS.Cell {
  const cell = worksheet.getCell(row, column);
  return cell.isMerged ? cell.master : cell;
}

/** 'YYYY-MM-DD' 에서 연도가 차지하는 글자 수 */
const YEAR_LENGTH = 4;

/**
 * 연차 기록을 근무표에 반영한다.
 *
 * 근무표는 틀(템플릿)일 뿐이고 내용은 모두 연차표에서 나온다.
 * 그래서 순서가 중요하다.
 *   ① 날짜·요일 행을 해당 월에 맞게 고친다   (요일을 알아야 평일/주말을 가른다)
 *   ② 평일은 D, 토·일은 · 로 되돌린다        (이전 달 데이터를 모두 지운다)
 *     이름이 없는 빈 줄은 함께 비운다
 *   ③ 연차 기록을 off / 특근(D) 표기로 넣는다
 *   ④ 휴일 칸에 색을 입히고 D 글자 서식을 맞춘다
 *   ⑤ 근무표에 적힌 값을 세어 집계 열을 채운다
 *   ⑥ 이름 왼쪽 소속(OT/PT) 을 연차표 기준으로 맞춘다
 *
 * 쓰기에 exceljs 를 쓰는 이유:
 * xlsx(SheetJS) 로 다시 쓰면 셀 배경색 등 서식이 모두 사라진다.
 * 근무표는 그대로 인쇄해서 쓰는 문서이므로 원본 서식을 유지해야 한다.
 */
export async function applyLeaveToSchedule(
  schedule: ScheduleWorkbook,
  entries: readonly LeaveEntry[],
  holidays: ReadonlySet<string>,
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
    resetToTemplate(sheet, holidays, warnings);
    clearEmptyRosterRows(sheet);
  }

  const applied = entries.map((entry) => applyEntry(schedule, entry));

  const departments = collectDepartments(entries);

  for (const sheet of targets) {
    applyCellStyles(sheet, holidays);
    fillCountColumns(sheet, warnings);
    applyDepartmentLabels(sheet, departments, warnings);
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
      clearDayColumn(sheet, column);
      return;
    }

    writableCell(sheet.worksheet, sheet.headerRow, column).value = day;
    writableCell(sheet.worksheet, sheet.weekdayRow, column).value = weekdayLabel(year, month, day);
    sheet.dayColumns.set(day, column);
  });
}

/**
 * 그 달에 없는 날짜 열(30일까지인 달의 31일 칸)을 통째로 비운다.
 *
 * 날짜·요일만 지우면 그 아래 근무 칸에 D·off 가 그대로 남아
 * 있지도 않은 날에 근무가 잡힌 것처럼 보인다. 배경색도 함께 지운다.
 */
function clearDayColumn(sheet: ScheduleSheet, column: number): void {
  // 이름이 없는 빈 줄에도 지난 달 D 가 남아 있다.
  const rows = [sheet.headerRow, sheet.weekdayRow, ...sheet.nameRows.values(), ...findEmptyRosterRows(sheet)];

  for (const row of rows) {
    const cell = writableCell(sheet.worksheet, row, column);
    const style = { ...cell.style };
    style.fill = { type: 'pattern', pattern: 'none' };
    cell.style = style;
    cell.value = null;
  }
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
 * 근무 칸을 템플릿 상태로 되돌린다.
 *
 *   평일(월~금)        → D
 *   토·일              → ·
 *   공휴일·휴업일      → ·   (연차표에 붉은 글자로 적힌 날)
 *
 * 근무표는 틀일 뿐이고 내용은 모두 연차표에서 나오므로,
 * 이전 달 데이터가 남아 있어도 전부 지우고 처음부터 다시 채운다.
 * 토요일 근무는 연차표의 특근 표기('[홍길동]')가 D 로 들어가면서 생긴다.
 */
function resetToTemplate(sheet: ScheduleSheet, holidays: ReadonlySet<string>, warnings: string[]): void {
  const { year, month } = sheet;
  if (year === null || month === null) {
    return;
  }

  let cleared = 0;
  let restDays = 0;

  for (const [day, column] of sheet.dayColumns) {
    const isHoliday = holidays.has(formatIsoDate(year, month, day));
    const isWeekday = !isHoliday && SCHEDULE_WORKING_WEEKDAYS.includes(weekdayIndex(year, month, day));
    const template = isWeekday ? SCHEDULE_WORK_MARKER : SCHEDULE_REST_MARKER;

    if (isHoliday) {
      restDays += 1;
    }

    for (const row of sheet.nameRows.values()) {
      if (compact(cellText(sheet.worksheet, row, column)).includes(SCHEDULE_OFF_MARKER)) {
        cleared += 1;
      }
      writableCell(sheet.worksheet, row, column).value = template;
    }
  }

  if (cleared > 0) {
    warnings.push(
      `'${sheet.sheetName}': 원래 적혀 있던 off 표기 ${String(cleared)}건을 지우고 연차표 기준으로 다시 채웠습니다.`,
    );
  }
  if (restDays > 0) {
    warnings.push(`'${sheet.sheetName}': 연차표에 붉게 표시된 공휴일·휴업일 ${String(restDays)}일을 · 로 두었습니다.`);
  }
}

/**
 * 명단에 자리만 있고 이름이 없는 줄을 찾는다.
 *
 * 근무표에는 사람이 빠진 뒤에도 소속(OT/PT)만 남은 빈 줄이 있다.
 * 명단 첫 줄부터 아래로 훑되, 소속·성명이 모두 빈 줄이나
 * '비고' 같은 명단 밖 줄을 만나면 거기서 멈춘다.
 */
function findEmptyRosterRows(sheet: ScheduleSheet): number[] {
  const named = new Set(sheet.nameRows.values());
  if (named.size === 0) {
    return [];
  }

  const empty: number[] = [];

  for (let row = Math.min(...named); row <= sheet.worksheet.rowCount; row += 1) {
    if (named.has(row)) {
      continue;
    }

    const name = compact(cellText(sheet.worksheet, row, sheet.nameColumn));
    const department = sheet.nameColumn > 1 ? compact(cellText(sheet.worksheet, row, sheet.nameColumn - 1)) : '';

    // 이름이 적혀 있는데 명단에 없는 줄('비고' 등)이면 명단이 끝난 것이다.
    if (name !== '') {
      break;
    }
    if (department === '') {
      break;
    }

    empty.push(row);
  }

  return empty;
}

/**
 * 이름이 없는 빈 줄의 근무 칸과 집계 칸을 비운다.
 * 사람이 없는 줄에 D 나 집계 숫자가 남아 있으면 근무한 것처럼 보인다.
 */
function clearEmptyRosterRows(sheet: ScheduleSheet): void {
  const countColumns = [sheet.offColumn, sheet.leaveColumn, sheet.saturdayColumn];

  for (const row of findEmptyRosterRows(sheet)) {
    for (const column of sheet.dayColumns.values()) {
      writableCell(sheet.worksheet, row, column).value = null;
    }
    for (const column of countColumns) {
      if (column !== null) {
        writableCell(sheet.worksheet, row, column).value = null;
      }
    }
  }
}

/** 2026-02-09 형태로 */
function formatIsoDate(year: number, month: number, day: number): string {
  return `${String(year)}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
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

  const cell = writableCell(sheet.worksheet, row, column);
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
 *   연차 종일 → off         연차 오후 → 오후 + 줄바꿈 + off
 *   공가 종일 → 공가+off     공가 오후 → 오후 공가 + 줄바꿈 + off
 *   특근      → D           (쉬는 것이 아니라 근무이므로 근무 표시를 넣는다)
 *
 * 앞에 말이 붙을 때는 off 앞에서 줄을 바꾼다. 날짜 칸이 좁아 한 줄로는 잘려 보인다.
 */
function markerFor(entry: LeaveEntry): string {
  if (entry.kind === SPECIAL_DUTY_KEYWORD) {
    return SCHEDULE_WORK_MARKER;
  }

  // 오전/오후가 안 적힌 반차는 종일과 구분할 수 없으므로 종일로 넣고 경고로 남긴다.
  const period = entry.halfPeriod ?? '';
  const reason = entry.kind === PUBLIC_LEAVE_KEYWORD || entry.kind === FAMILY_EVENT_KEYWORD ? entry.kind : '';
  const prefix = [period, reason].filter((part) => part !== '').join(' ');

  return prefix === ''
    ? SCHEDULE_OFF_MARKER
    : `${prefix}${SCHEDULE_MARKER_LINE_BREAK}${SCHEDULE_OFF_MARKER}`;
}

// ---------------------------------------------------------------------------
// ④ 휴일 색칠 · D 글자 서식
// ---------------------------------------------------------------------------

/**
 * 날짜 칸의 서식을 다시 입힌다.
 *
 * - 평일 칸은 배경색을 지운다. 지난달 근무표에서 쉬는 날이었던 칸에 색이 남아 있으면
 *   이번 달 달력과 어긋나 보인다.
 * - 토·일·공휴일·대체휴일 칸은 하늘색으로 칠한다. 쉬는 날 판단 기준은
 *   값을 채울 때와 똑같아서(주말이거나 연차표에 붉게 적힌 날) 색과 내용이 항상 맞는다.
 * - 근무(D) 글자는 맑은 고딕 10pt 굵게로 맞춘다.
 * - off 는 D 와 같은 10pt, 오전/오후 off 는 8pt, 공가·경조 off 는 7pt 로 맞춘다.
 *   앞에 말이 붙은 것은 두 줄이 되므로 자동 줄바꿈도 함께 켠다.
 *
 * 날짜 행·요일 행까지 함께 칠해 세로로 한 칸이 통째로 구분되게 한다.
 *
 * 주의: exceljs 는 서식이 같은 셀끼리 style 객체를 공유한다.
 * cell.fill = ... 처럼 바로 바꾸면 같은 서식을 쓰던 다른 칸까지 함께 바뀌어,
 * 평일 칸을 지우면 휴일 색까지 지워지고 다시 칠하면 평일까지 칠해진다.
 * 그래서 칸마다 style 을 새 객체로 복사한 뒤 바꾼다.
 */
function applyCellStyles(sheet: ScheduleSheet, holidays: ReadonlySet<string>): void {
  const { year, month } = sheet;
  if (year === null || month === null) {
    return;
  }

  // 빈 줄도 같은 배경색을 써야 열이 끊겨 보이지 않는다.
  const rows = [sheet.headerRow, sheet.weekdayRow, ...sheet.nameRows.values(), ...findEmptyRosterRows(sheet)];

  for (const [day, column] of sheet.dayColumns) {
    const isRestDay =
      holidays.has(formatIsoDate(year, month, day)) ||
      !SCHEDULE_WORKING_WEEKDAYS.includes(weekdayIndex(year, month, day));

    for (const row of rows) {
      const cell = writableCell(sheet.worksheet, row, column);
      const style = { ...cell.style };

      // 손으로 적힌 칸까지 글꼴을 맞춘다. 크기·굵기는 아래에서 표기별로 덮어쓴다.
      style.font = { ...style.font, name: SCHEDULE_WORK_FONT_NAME };

      style.fill = isRestDay
        ? { type: 'pattern', pattern: 'solid', fgColor: { argb: SCHEDULE_HOLIDAY_FILL_ARGB } }
        : { type: 'pattern', pattern: 'none' };

      const text = compact(cellText(sheet.worksheet, row, column));

      if (text === SCHEDULE_WORK_MARKER) {
        // 글자색 등 원래 서식은 두고 글꼴만 바꾼다.
        style.font = { ...style.font, name: SCHEDULE_WORK_FONT_NAME, size: SCHEDULE_WORK_FONT_SIZE, bold: true };
      } else if (text.includes(SCHEDULE_OFF_MARKER)) {
        // off 앞에 말이 붙으면 두 줄이 되므로 글자를 줄이고 자동 줄바꿈을 켠다.
        const prefixed = text !== SCHEDULE_OFF_MARKER;
        style.font = { ...style.font, name: SCHEDULE_WORK_FONT_NAME, size: offFontSize(text), bold: true };
        if (prefixed) {
          style.alignment = { ...style.alignment, wrapText: true };
        }
      }

      cell.style = style;
    }
  }
}

/**
 * off 표기의 글자 크기.
 * 글자가 길수록 줄인다. 셋이 겹치는 표기(오후 공가 off)는 가장 작은 것을 쓴다.
 */
function offFontSize(text: string): number {
  if (text.includes(PUBLIC_LEAVE_KEYWORD) || text.includes(FAMILY_EVENT_KEYWORD)) {
    return SCHEDULE_REASON_OFF_FONT_SIZE;
  }
  if (text.includes(MORNING_KEYWORD) || text.includes(AFTERNOON_KEYWORD)) {
    return SCHEDULE_HALF_OFF_FONT_SIZE;
  }
  return SCHEDULE_OFF_FONT_SIZE;
}

// ---------------------------------------------------------------------------
// ⑤ 집계 열 채우기
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

      const unit =
        value.includes(MORNING_KEYWORD) || value.includes(AFTERNOON_KEYWORD) ? SCHEDULE_HALF_UNIT : SCHEDULE_FULL_UNIT;
      offTotal += unit;

      if (!value.includes(PUBLIC_LEAVE_KEYWORD) && !value.includes(FAMILY_EVENT_KEYWORD)) {
        leaveTotal += unit;
      }
    }

    writeCount(sheet, row, sheet.offColumn, offTotal);
    writeCount(sheet, row, sheet.leaveColumn, leaveTotal);
    // 근무가 없으면 '·' 대신 0 이 들어간다.
    writeCount(sheet, row, sheet.saturdayColumn, saturdayTotal);
  }
}

/**
 * 집계 칸에 숫자를 적고 글꼴을 맞춘다.
 * 열이 없는 시트도 있으므로 null 이면 아무것도 하지 않는다.
 */
function writeCount(sheet: ScheduleSheet, row: number, column: number | null, total: number): void {
  if (column === null) {
    return;
  }
  const cell = writableCell(sheet.worksheet, row, column);
  const style = { ...cell.style };
  // 서식이 아예 없던 칸(토요일근무 등)은 크기까지 채워야 글꼴이 저장된다.
  style.font = { size: SCHEDULE_WORK_FONT_SIZE, ...style.font, name: SCHEDULE_WORK_FONT_NAME };
  cell.style = style;
  cell.value = total;
}

// ---------------------------------------------------------------------------
// ⑥ 소속(OT/PT) 맞추기
// ---------------------------------------------------------------------------

/**
 * 사람마다 어느 연차표에 이름이 있었는지 모은다.
 * 두 파일 모두에 있으면 어느 쪽이 맞는지 알 수 없으므로 둘 다 담아 둔다.
 */
function collectDepartments(entries: readonly LeaveEntry[]): Map<string, Set<string>> {
  const departments = new Map<string, Set<string>>();

  for (const entry of entries) {
    const found = departments.get(entry.name) ?? new Set<string>();
    found.add(entry.department);
    departments.set(entry.name, found);
  }

  return departments;
}

/**
 * 이름 왼쪽 칸의 소속(OT/PT) 을 연차표 기준으로 고친다.
 *
 * 기본은 근무표에 적힌 값이다.
 * 그 사람이 PT 연차표에 있으면 PT, OT 연차표에 있으면 OT 로 바꾼다.
 * 두 파일 모두에 있거나 어느 쪽에도 없으면 근무표 값을 그대로 둔다.
 * 값을 바꾸지 않는 사람도 글꼴은 함께 맞춘다.
 */
function applyDepartmentLabels(
  sheet: ScheduleSheet,
  departments: ReadonlyMap<string, Set<string>>,
  warnings: string[],
): void {
  if (sheet.nameColumn <= 1) {
    warnings.push(`'${sheet.sheetName}' 는 성명 열 왼쪽에 칸이 없어 소속을 적지 않았습니다.`);
    return;
  }

  const column = sheet.nameColumn - 1;
  const ambiguous: string[] = [];

  for (const [name, row] of sheet.nameRows) {
    const cell = writableCell(sheet.worksheet, row, column);
    const style = { ...cell.style };
    style.font = {
      ...style.font,
      name: SCHEDULE_WORK_FONT_NAME,
      size: SCHEDULE_DEPARTMENT_FONT_SIZE,
      bold: true,
    };
    cell.style = style;

    const found = departments.get(name);
    if (found === undefined) {
      continue;
    }
    if (found.size > 1) {
      ambiguous.push(name);
      continue;
    }

    const [department] = [...found];
    if (department !== undefined) {
      cell.value = department;
    }
  }

  if (ambiguous.length > 0) {
    warnings.push(
      `'${sheet.sheetName}': ${ambiguous.join(', ')} 은(는) OT·PT 연차표에 모두 있어 소속을 근무표 값으로 두었습니다.`,
    );
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
