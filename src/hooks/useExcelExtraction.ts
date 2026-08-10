import { useCallback, useState } from 'react';
import type { ParseResult } from '../types/excel';
import { parseExcel } from '../utils/excel/parser';
import { saveResultWorkbook } from '../utils/excel/exporter';
import { isExcelFile, toErrorMessage } from '../utils/file';
import { ACCEPTED_EXTENSIONS } from '../utils/constants';

/** 화면이 표현해야 하는 진행 상태 */
export type ExtractionStatus = 'idle' | 'parsing' | 'saving';

/** 훅이 화면에 제공하는 상태와 동작 */
export interface ExcelExtraction {
  file: File | null;
  result: ParseResult | null;
  status: ExtractionStatus;
  error: string | null;
  notice: string | null;
  selectFile: (file: File) => void;
  extract: () => Promise<void>;
  save: () => Promise<void>;
}

/**
 * 파일 선택 → 추출 → 저장까지의 화면 상태를 관리하는 훅.
 * 엑셀 관련 실제 처리는 parser / exporter 에 위임하고, 여기서는 흐름과 상태만 다룬다.
 */
export function useExcelExtraction(): ExcelExtraction {
  const [file, setFile] = useState<File | null>(null);
  const [result, setResult] = useState<ParseResult | null>(null);
  const [status, setStatus] = useState<ExtractionStatus>('idle');
  const [error, setError] = useState<string | null>(null);
  const [notice, setNotice] = useState<string | null>(null);

  /** 새 파일을 선택하면 이전 결과를 모두 비운다. */
  const selectFile = useCallback((nextFile: File): void => {
    setResult(null);
    setNotice(null);

    if (!isExcelFile(nextFile)) {
      setFile(null);
      setError(`지원하지 않는 파일 형식입니다. (${ACCEPTED_EXTENSIONS.join(', ')} 만 가능)`);
      return;
    }

    setError(null);
    setFile(nextFile);
  }, []);

  const extract = useCallback(async (): Promise<void> => {
    if (file === null) {
      setError('먼저 엑셀 파일을 선택하세요.');
      return;
    }

    setStatus('parsing');
    setError(null);
    setNotice(null);

    try {
      const parsed = await parseExcel(file);
      setResult(parsed);

      if (parsed.rows.length === 0) {
        setNotice('신장분사 항목을 찾지 못했습니다. 파일 내용을 확인해 주세요.');
      }
    } catch (caught) {
      setResult(null);
      setError(toErrorMessage(caught));
    } finally {
      setStatus('idle');
    }
  }, [file]);

  const save = useCallback(async (): Promise<void> => {
    if (result === null || result.rows.length === 0) {
      setError('저장할 결과가 없습니다. 먼저 추출하세요.');
      return;
    }

    setStatus('saving');
    setError(null);
    setNotice(null);

    try {
      const savedPath = await saveResultWorkbook(result);
      setNotice(savedPath === null ? '저장을 취소했습니다.' : `저장 완료: ${savedPath}`);
    } catch (caught) {
      setError(toErrorMessage(caught));
    } finally {
      setStatus('idle');
    }
  }, [result]);

  return { file, result, status, error, notice, selectFile, extract, save };
}
