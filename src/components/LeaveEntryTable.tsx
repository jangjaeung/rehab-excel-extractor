import type { JSX } from 'react';
import type { LeaveEntry } from '../types/leave';
import { ANNUAL_LEAVE_KEYWORD, HALF_DAY_KEYWORD, LEAVE_COLUMN_LABELS } from '../utils/constants';

interface LeaveEntryTableProps {
  entries: readonly LeaveEntry[];
}

/** '연차' / '오후 반차' / '오후 공가' / '경조' / '특근' */
function formatKind(entry: LeaveEntry): string {
  const period = entry.halfPeriod === null ? '' : `${entry.halfPeriod} `;

  if (entry.kind === ANNUAL_LEAVE_KEYWORD && entry.half) {
    return `${period}${HALF_DAY_KEYWORD}`;
  }
  return `${period}${entry.kind}`;
}

/**
 * 연차 사용 내역 전체 (날짜 오름차순).
 * 원본 표기를 함께 보여 주어 시트와 대조할 수 있게 한다.
 */
export function LeaveEntryTable({ entries }: LeaveEntryTableProps): JSX.Element {
  if (entries.length === 0) {
    return <p className="empty">추출된 연차가 없습니다.</p>;
  }

  return (
    <div className="table-wrapper">
      <table className="result-table">
        <thead>
          <tr>
            <th>{LEAVE_COLUMN_LABELS.date}</th>
            <th>{LEAVE_COLUMN_LABELS.weekday}</th>
            <th>{LEAVE_COLUMN_LABELS.department}</th>
            <th>{LEAVE_COLUMN_LABELS.name}</th>
            <th>{LEAVE_COLUMN_LABELS.kind}</th>
            <th>{LEAVE_COLUMN_LABELS.counter}</th>
            <th>{LEAVE_COLUMN_LABELS.sheet}</th>
            <th>{LEAVE_COLUMN_LABELS.raw}</th>
          </tr>
        </thead>
        <tbody>
          {entries.map((entry) => (
            <tr key={`${entry.isoDate}:${entry.name}`}>
              <td>{entry.isoDate}</td>
              <td>{entry.weekday}</td>
              <td>{entry.department}</td>
              <td>{entry.name}</td>
              <td>{formatKind(entry)}</td>
              <td className="numeric">{entry.counter ?? '-'}</td>
              <td>{entry.sheetName}</td>
              <td className="wrap muted">{entry.raw}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}
