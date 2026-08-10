import type { LeaveEntry } from '../../types/leave';
import type { AppliedCell, ApplyStatus, ScheduleApplyResult, ScheduleSheetInfo } from '../../types/schedule';
import {
  AFTERNOON_KEYWORD,
  FAMILY_EVENT_KEYWORD,
  HALF_DAY_KEYWORD,
  MORNING_KEYWORD,
  PUBLIC_LEAVE_KEYWORD,
  SCHEDULE_MARKERS,
  SCHEDULE_PLAIN_VALUES,
} from '../constants';
import { cellText, compact, type ScheduleSheet, type ScheduleWorkbook } from './reader';

/**
 * 연차 기록을 근무표의 해당 칸에 써 넣는다.
 *
 * 쓰기에 exceljs 를 쓰는 이유:
 * xlsx(SheetJS) 로 다시 쓰면 셀 배경색 등 서식이 모두 사라진다.
 * 근무표는 그대로 인쇄해서 쓰는 문서이므로 원본 서식을 유지해야 한다.
 */
export async function applyLeaveToSchedule(
  schedule: ScheduleWorkbook,
  entries: readonly LeaveEntry[],
): Promise<ScheduleApplyResult> {
  const applied: AppliedCell[] = [];
  const warnings = [...schedule.warnings];

  for (const entry of entries) {
    applied.push(applyEntry(schedule, entry));
  }

  const buffer = await schedule.workbook.xlsx.writeBuffer();

  return {
    data: new Uint8Array(buffer),
    applied,
    sheets: summarizeSheets(schedule),
    warnings,
  };
}

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

/** 기록 종류 → 근무표에 넣을 표기 */
function markerFor(entry: LeaveEntry): string {
  if (entry.kind === PUBLIC_LEAVE_KEYWORD) {
    return SCHEDULE_MARKERS.public;
  }
  if (entry.kind === FAMILY_EVENT_KEYWORD) {
    return SCHEDULE_MARKERS.familyEvent;
  }
  if (entry.kind === HALF_DAY_KEYWORD) {
    if (entry.halfPeriod === MORNING_KEYWORD) {
      return SCHEDULE_MARKERS.morningHalf;
    }
    if (entry.halfPeriod === AFTERNOON_KEYWORD) {
      return SCHEDULE_MARKERS.afternoonHalf;
    }
    // 오전/오후가 안 적힌 반차는 종일과 구분할 수 없으므로 off 로 넣고 경고로 남긴다.
    return SCHEDULE_MARKERS.annual;
  }
  return SCHEDULE_MARKERS.annual;
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
