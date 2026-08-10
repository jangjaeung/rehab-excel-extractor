import type { JSX } from 'react';

interface WarningListProps {
  warnings: readonly string[];
}

/**
 * 파싱 중 발생한 비치명적 경고 목록.
 * (PT번호를 못 찾은 행, 중복 항목 등 확인이 필요한 내용)
 */
export function WarningList({ warnings }: WarningListProps): JSX.Element | null {
  if (warnings.length === 0) {
    return null;
  }

  return (
    <details className="warnings">
      <summary>확인이 필요한 항목 {warnings.length}건</summary>
      <ul>
        {warnings.map((warning, index) => (
          <li key={`${String(index)}:${warning}`}>{warning}</li>
        ))}
      </ul>
    </details>
  );
}
