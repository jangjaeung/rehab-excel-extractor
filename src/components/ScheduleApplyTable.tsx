import type { JSX } from 'react';
import type { AppliedCell } from '../types/schedule';
import { APPLY_COLUMN_LABELS } from '../utils/constants';

interface ScheduleApplyTableProps {
  applied: readonly AppliedCell[];
}

/** 결과에 따라 색을 달리해 한눈에 보이게 한다. */
function statusClassName(cell: AppliedCell): string {
  if (cell.status === '미반영') {
    return 'status status--error';
  }
  if (cell.status === '덮어씀') {
    return 'status status--warn';
  }
  return 'status';
}

/** 근무표에 무엇을 어디에 넣었는지 보여 주는 표 */
export function ScheduleApplyTable({ applied }: ScheduleApplyTableProps): JSX.Element {
  if (applied.length === 0) {
    return <p className="empty">근무표에 반영할 기록이 없습니다.</p>;
  }

  return (
    <div className="table-wrapper">
      <table className="result-table">
        <thead>
          <tr>
            <th>{APPLY_COLUMN_LABELS.date}</th>
            <th>{APPLY_COLUMN_LABELS.weekday}</th>
            <th>{APPLY_COLUMN_LABELS.department}</th>
            <th>{APPLY_COLUMN_LABELS.name}</th>
            <th>{APPLY_COLUMN_LABELS.marker}</th>
            <th>{APPLY_COLUMN_LABELS.sheet}</th>
            <th>{APPLY_COLUMN_LABELS.address}</th>
            <th>{APPLY_COLUMN_LABELS.status}</th>
            <th>{APPLY_COLUMN_LABELS.detail}</th>
          </tr>
        </thead>
        <tbody>
          {applied.map((cell) => (
            <tr key={`${cell.isoDate}:${cell.name}`}>
              <td>{cell.isoDate}</td>
              <td>{cell.weekday}</td>
              <td>{cell.department}</td>
              <td>{cell.name}</td>
              <td>{cell.marker}</td>
              <td>{cell.sheetName ?? '-'}</td>
              <td>{cell.address ?? '-'}</td>
              <td className={statusClassName(cell)}>{cell.status}</td>
              <td className="wrap muted">
                {cell.reason ?? (cell.previous === null ? '' : `원래 값: ${cell.previous}`)}
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}
