import { useCallback, useEffect, useRef, useState, type JSX } from 'react';
import type { TherapistRecord } from '../types/excel';
import { copyText, toClipboardTable } from '../utils/clipboard';
import { COPY_BUTTON_DONE, COPY_BUTTON_FAILED, COPY_BUTTON_LABEL, COPY_FEEDBACK_MS } from '../utils/constants';

interface CopyValuesButtonProps {
  /** 복사할 항목 컬럼 (이름·PT번호는 포함하지 않는다) */
  columns: readonly string[];
  rows: readonly TherapistRecord[];
}

/** 버튼이 보여 줄 상태 */
type CopyState = 'idle' | 'done' | 'failed';

/**
 * 표의 숫자만 클립보드에 복사하는 버튼.
 * 이름·PT번호를 뺀 항목 값만 담으므로 엑셀의 원하는 칸에 바로 붙여넣을 수 있다.
 */
export function CopyValuesButton({ columns, rows }: CopyValuesButtonProps): JSX.Element {
  const [state, setState] = useState<CopyState>('idle');
  const timer = useRef<number | null>(null);

  // 다른 탭으로 옮기는 등으로 사라질 때 남은 타이머를 정리한다.
  useEffect(
    () => () => {
      if (timer.current !== null) {
        window.clearTimeout(timer.current);
      }
    },
    [],
  );

  const handleCopy = useCallback(async (): Promise<void> => {
    const copied = await copyText(toClipboardTable(rows, columns));

    setState(copied ? 'done' : 'failed');
    if (timer.current !== null) {
      window.clearTimeout(timer.current);
    }
    timer.current = window.setTimeout(() => {
      setState('idle');
    }, COPY_FEEDBACK_MS);
  }, [columns, rows]);

  const label = state === 'done' ? COPY_BUTTON_DONE : state === 'failed' ? COPY_BUTTON_FAILED : COPY_BUTTON_LABEL;

  return (
    <button
      type="button"
      className={state === 'idle' ? 'copy-button' : `copy-button copy-button--${state}`}
      onClick={() => void handleCopy()}
      disabled={rows.length === 0}
      title={`${String(rows.length)}명 × ${String(columns.length)}개 항목의 숫자만 복사합니다`}
    >
      {label}
    </button>
  );
}
