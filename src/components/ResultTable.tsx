import type { JSX } from 'react';
import type { ParseResult } from '../types/excel';
import { COLUMN_LABEL_PT, COLUMN_LABEL_THERAPIST, DEFAULT_ITEM_COUNT } from '../utils/constants';

interface ResultTableProps {
  result: ParseResult;
}

/**
 * 추출 결과 테이블.
 * 컬럼은 파서가 만들어 준 순서(오름차순 정렬)를 그대로 사용하므로
 * 새로운 신장분사 항목이 생겨도 컴포넌트 수정이 필요 없다.
 */
export function ResultTable({ result }: ResultTableProps): JSX.Element {
  if (result.rows.length === 0) {
    return <p className="empty">추출된 데이터가 없습니다.</p>;
  }

  return (
    <div className="table-wrapper">
      <table className="result-table">
        <thead>
          <tr>
            <th>{COLUMN_LABEL_THERAPIST}</th>
            <th>{COLUMN_LABEL_PT}</th>
            {result.columns.map((column) => (
              <th key={column}>{column}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {result.rows.map((row) => (
            <tr key={row.pt}>
              <td>{row.therapist}</td>
              <td>{row.pt}</td>
              {result.columns.map((column) => (
                <td key={column} className="numeric">
                  {row.items[column] ?? DEFAULT_ITEM_COUNT}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}
