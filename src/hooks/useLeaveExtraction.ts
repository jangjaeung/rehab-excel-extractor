import { useCallback, useState } from 'react';
import type { LeaveEntry, LeavePersonSummary } from '../types/leave';
import type { ScheduleApplyResult } from '../types/schedule';
import { readHolidayDates } from '../utils/leave/holidays';
import { parseLeaveExcel, summarizeLeaveEntries } from '../utils/leave/parser';
import { readSchedule } from '../utils/schedule/reader';
import { applyLeaveToSchedule } from '../utils/schedule/writer';
import { saveWorkbookData } from '../utils/excel/save';
import { isExcelFile, toErrorMessage } from '../utils/file';
import {
  ACCEPTED_EXTENSIONS,
  LEAVE_SLOTS,
  SCHEDULE_RESULT_SUFFIX,
  WORK_SCHEDULE_LABEL,
  type LeaveSlotId,
} from '../utils/constants';

/** 화면이 표현해야 하는 진행 상태 */
export type LeaveStatus = 'idle' | 'parsing' | 'saving';

/** 추출 결과 (연차 기록 + 근무표 반영 결과) */
export interface LeaveResult {
  entries: LeaveEntry[];
  people: LeavePersonSummary[];
  schedule: ScheduleApplyResult;
}

/** 훅이 화면에 제공하는 상태와 동작 */
export interface LeaveExtraction {
  files: Record<LeaveSlotId, File | null>;
  result: LeaveResult | null;
  status: LeaveStatus;
  error: string | null;
  notice: string | null;
  selectFile: (slot: LeaveSlotId, file: File) => void;
  extract: () => Promise<void>;
  save: () => Promise<void>;
}

/** 업로드 슬롯 초기 상태 */
const EMPTY_FILES: Record<LeaveSlotId, File | null> = { ot: null, pt: null, schedule: null };

/**
 * 연차 추출기 화면의 상태를 관리하는 훅.
 *
 * 흐름
 *   ① 근무표를 먼저 읽어 명단을 얻는다 (연차표에서 사람을 가려내는 기준이 된다)
 *   ② OT/PT 연차표를 그 명단으로 파싱한다
 *   ③ 기록을 근무표 칸에 써 넣고, 결과를 화면에 보여 준다
 *   ④ [근무표 저장] 으로 서식이 유지된 xlsx 를 내려받는다
 */
export function useLeaveExtraction(): LeaveExtraction {
  const [files, setFiles] = useState<Record<LeaveSlotId, File | null>>(EMPTY_FILES);
  const [result, setResult] = useState<LeaveResult | null>(null);
  const [status, setStatus] = useState<LeaveStatus>('idle');
  const [error, setError] = useState<string | null>(null);
  const [notice, setNotice] = useState<string | null>(null);

  /** 확장자를 확인해 통과하면 보관하고, 아니면 어떤 업로드가 잘못됐는지 알린다. */
  const selectFile = useCallback((slot: LeaveSlotId, file: File): void => {
    // 파일이 바뀌면 이전 결과는 낡은 것이 되므로 비운다.
    setResult(null);
    setNotice(null);

    const label = LEAVE_SLOTS.find((candidate) => candidate.id === slot)?.title ?? slot;

    if (!isExcelFile(file)) {
      setFiles((current) => ({ ...current, [slot]: null }));
      setError(`${label}: 지원하지 않는 파일 형식입니다. (${ACCEPTED_EXTENSIONS.join(', ')} 만 가능)`);
      return;
    }

    setError(null);
    setFiles((current) => ({ ...current, [slot]: file }));
  }, []);

  const extract = useCallback(async (): Promise<void> => {
    const scheduleFile = files.schedule;
    if (scheduleFile === null) {
      setError(`${WORK_SCHEDULE_LABEL}을 선택하세요. 근무표의 명단을 기준으로 연차표를 읽습니다.`);
      return;
    }
    if (files.ot === null && files.pt === null) {
      setError('OT 연차 또는 PT 연차 중 하나 이상을 선택하세요.');
      return;
    }

    setStatus('parsing');
    setError(null);
    setNotice(null);

    try {
      const schedule = await readSchedule(scheduleFile);

      const entries: LeaveEntry[] = [];
      const warnings: string[] = [];
      // 공휴일·휴업일은 연차표 날짜의 붉은 글자로 표시되어 있다.
      const holidays = new Set<string>();

      for (const slot of LEAVE_SLOTS) {
        const file = files[slot.id];
        if (slot.id === 'schedule' || file === null) {
          continue;
        }
        const parsed = await parseLeaveExcel(file, { roster: schedule.roster, department: slot.department });
        entries.push(...parsed.entries);
        warnings.push(...parsed.warnings.map((warning) => `[${slot.department}] ${warning}`));

        for (const holiday of await readHolidayDates(file)) {
          holidays.add(holiday);
        }
      }

      entries.sort((a, b) => (a.isoDate === b.isoDate ? a.name.localeCompare(b.name, 'ko') : a.isoDate < b.isoDate ? -1 : 1));

      const applyResult = await applyLeaveToSchedule(schedule, entries, holidays);
      applyResult.warnings.unshift(...warnings);

      setResult({ entries, people: summarizeLeaveEntries(entries), schedule: applyResult });

      if (entries.length === 0) {
        setNotice('연차 기록을 찾지 못했습니다. 파일 내용을 확인해 주세요.');
      }
    } catch (caught) {
      setResult(null);
      setError(toErrorMessage(caught));
    } finally {
      setStatus('idle');
    }
  }, [files]);

  const save = useCallback(async (): Promise<void> => {
    const scheduleFile = files.schedule;
    if (result === null || scheduleFile === null) {
      setError('저장할 결과가 없습니다. 먼저 추출하세요.');
      return;
    }

    setStatus('saving');
    setError(null);
    setNotice(null);

    try {
      const savedPath = await saveWorkbookData(result.schedule.data, toResultFileName(scheduleFile.name));
      setNotice(savedPath === null ? '저장을 취소했습니다.' : `저장 완료: ${savedPath}`);
    } catch (caught) {
      setError(toErrorMessage(caught));
    } finally {
      setStatus('idle');
    }
  }, [files.schedule, result]);

  return { files, result, status, error, notice, selectFile, extract, save };
}

/** 원본 근무표 파일명 뒤에 '_연차반영' 을 붙인다. */
function toResultFileName(originalName: string): string {
  const dot = originalName.lastIndexOf('.');
  const base = dot === -1 ? originalName : originalName.slice(0, dot);
  return `${base.trim()}${SCHEDULE_RESULT_SUFFIX}.xlsx`;
}
