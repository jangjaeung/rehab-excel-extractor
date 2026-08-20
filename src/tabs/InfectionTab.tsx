import { useMemo, type JSX } from 'react';
import { CopyValuesButton } from '../components/CopyValuesButton';
import { FileDropZone } from '../components/FileDropZone';
import { MessageBar } from '../components/MessageBar';
import { NameOrderInput } from '../components/NameOrderInput';
import { ResultTable } from '../components/ResultTable';
import { WarningList } from '../components/WarningList';
import { useExcelExtraction } from '../hooks/useExcelExtraction';
import { parseInfectionExcel } from '../utils/infection/parser';
import { INFECTION_RESULT_FILE_NAME, INFECTION_TOTAL_COLUMN } from '../utils/constants';

/**
 * 감염치료건수 탭.
 * 신장분사와 같은 파일을 올리지만, 주차로 나누지 않고
 * 모든 시트를 합쳐 1일~말일의 하루치 건수로 보여 준다.
 */
export function InfectionTab(): JSX.Element {
  const {
    file,
    result,
    status,
    error,
    notice,
    selectFile,
    extract,
    save,
    nameOrderText,
    names,
    missingNames,
    setNameOrderText,
  } = useExcelExtraction({
    parse: parseInfectionExcel,
    fileName: INFECTION_RESULT_FILE_NAME,
  });

  const isBusy = status !== 'idle';
  const canExtract = file !== null && !isBusy;
  const canSave = result !== null && !isBusy;

  const week = useMemo(() => result?.weeks[0] ?? null, [result]);

  /**
   * 복사할 컬럼 — 합계는 뺀 일자만.
   * 합계는 엑셀 쪽에 이미 수식이 있어서 덮어쓰면 안 된다.
   */
  const copyColumns = useMemo(
    () => (result?.columns ?? []).filter((column) => column !== INFECTION_TOTAL_COLUMN),
    [result],
  );

  return (
    <>
      <FileDropZone file={file} onSelect={selectFile} disabled={isBusy} />

      <NameOrderInput
        value={nameOrderText}
        names={names}
        missing={missingNames}
        onChange={setNameOrderText}
        disabled={isBusy}
      />

      <div className="actions">
        <button type="button" className="button button--primary" onClick={() => void extract()} disabled={!canExtract}>
          {status === 'parsing' ? '추출 중...' : '추출하기'}
        </button>
        <button type="button" className="button" onClick={() => void save()} disabled={!canSave}>
          {status === 'saving' ? '저장 중...' : '엑셀 저장'}
        </button>
      </div>

      <MessageBar message={error} tone="error" />
      <MessageBar message={notice} tone="info" />

      {result !== null && week !== null && (
        <section className="result">
          <div className="result__head">
            <div className="result__summary">
              <span>기간 {week.label}</span>
              <span>인원 {week.rows.length}명</span>
            </div>
            <CopyValuesButton columns={copyColumns} rows={week.rows} />
          </div>

          <WarningList warnings={result.warnings} />
          <ResultTable columns={result.columns} rows={week.rows} />
        </section>
      )}
    </>
  );
}
