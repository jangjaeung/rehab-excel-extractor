import { useMemo, type JSX } from 'react';
import { FileDropZone } from '../components/FileDropZone';
import { LeaveEntryTable } from '../components/LeaveEntryTable';
import { LeavePersonTable } from '../components/LeavePersonTable';
import { MessageBar } from '../components/MessageBar';
import { ScheduleApplyTable } from '../components/ScheduleApplyTable';
import { WarningList } from '../components/WarningList';
import { useLeaveExtraction } from '../hooks/useLeaveExtraction';
import { LEAVE_SLOTS } from '../utils/constants';

/**
 * 연차 추출기 탭.
 * OT/PT 연차표와 근무표를 올려 추출하면, 근무표의 날짜·이름 칸에 off 표기를 채워 준다.
 */
export function LeaveTab(): JSX.Element {
  const { files, result, status, error, notice, selectFile, extract, save } = useLeaveExtraction();

  const isBusy = status !== 'idle';
  const canExtract = files.schedule !== null && (files.ot !== null || files.pt !== null) && !isBusy;
  const canSave = result !== null && !isBusy;

  const counts = useMemo(() => {
    const applied = result?.schedule.applied ?? [];
    return {
      written: applied.filter((cell) => cell.status === '입력' || cell.status === '덮어씀').length,
      same: applied.filter((cell) => cell.status === '동일').length,
      failed: applied.filter((cell) => cell.status === '미반영').length,
    };
  }, [result]);

  return (
    <>
      <div className="upload-grid">
        {LEAVE_SLOTS.map((slot) => (
          <FileDropZone
            key={slot.id}
            file={files[slot.id]}
            onSelect={(file) => {
              selectFile(slot.id, file);
            }}
            disabled={isBusy}
            title={slot.title}
            buttonLabel={`${slot.title} 선택`}
            hint={slot.hint}
          />
        ))}
      </div>

      <div className="actions">
        <button type="button" className="button button--primary" onClick={() => void extract()} disabled={!canExtract}>
          {status === 'parsing' ? '추출 중...' : '추출하기'}
        </button>
        <button type="button" className="button" onClick={() => void save()} disabled={!canSave}>
          {status === 'saving' ? '저장 중...' : '근무표 저장'}
        </button>
      </div>

      <MessageBar message={error} tone="error" />
      <MessageBar message={notice} tone="info" />

      {result !== null && (
        <section className="result">
          <div className="result__summary">
            <span>연차 기록 {result.entries.length}건</span>
            <span>인원 {result.people.length}명</span>
            <span>근무표 반영 {counts.written}건</span>
            <span>이미 동일 {counts.same}건</span>
            <span>미반영 {counts.failed}건</span>
          </div>

          <WarningList warnings={result.schedule.warnings} />

          <h2 className="result__title">근무표 반영 결과</h2>
          <ScheduleApplyTable applied={result.schedule.applied} />

          <h2 className="result__title">인원별 요약</h2>
          <LeavePersonTable people={result.people} />

          <h2 className="result__title">연차 사용 내역</h2>
          <LeaveEntryTable entries={result.entries} />
        </section>
      )}
    </>
  );
}
