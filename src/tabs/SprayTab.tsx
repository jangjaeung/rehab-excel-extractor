import type { JSX } from 'react';
import { FileDropZone } from '../components/FileDropZone';
import { MessageBar } from '../components/MessageBar';
import { NameOrderInput } from '../components/NameOrderInput';
import { ResultTable } from '../components/ResultTable';
import { WarningList } from '../components/WarningList';
import { useExcelExtraction } from '../hooks/useExcelExtraction';
import { parseExcel } from '../utils/excel/parser';
import { RESULT_FILE_NAME } from '../utils/constants';

/**
 * 신장분사 추출기 탭.
 * 파일 선택 → 추출 → 저장 흐름은 useExcelExtraction 훅이 담당하고,
 * 여기서는 배치와 사용자 상호작용만 다룬다.
 */
export function SprayTab(): JSX.Element {
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
    parse: parseExcel,
    fileName: RESULT_FILE_NAME,
  });

  const isBusy = status !== 'idle';
  const canExtract = file !== null && !isBusy;
  const canSave = result !== null && result.weeks.length > 0 && !isBusy;

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

      {result !== null && (
        <section className="result">
          <div className="result__summary">
            <span>주차 {result.weeks.length}개</span>
            <span>항목 {result.columns.length}개</span>
          </div>

          <WarningList warnings={result.warnings} />

          {result.weeks.map((week) => (
            <div key={week.sheetName}>
              <h2 className="result__title">
                {week.label} <span className="result__subtitle">{week.sheetName} · 치료사 {week.rows.length}명</span>
              </h2>
              <WarningList warnings={week.warnings} />
              <ResultTable columns={result.columns} rows={week.rows} />
            </div>
          ))}
        </section>
      )}
    </>
  );
}
