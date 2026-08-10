import type { JSX } from 'react';
import type { LeavePersonSummary } from '../types/leave';
import { LEAVE_COLUMN_LABELS } from '../utils/constants';

interface LeavePersonTableProps {
  people: readonly LeavePersonSummary[];
}

/** '2회 (오전 1, 오후 1)' 처럼 반차 건수와 시간대 내역을 함께 보여 준다. */
function formatHalfCount(person: LeavePersonSummary): string {
  if (person.halfCount === 0) {
    return '0회';
  }

  const breakdown: string[] = [];
  if (person.morningCount > 0) {
    breakdown.push(`오전 ${String(person.morningCount)}`);
  }
  if (person.afternoonCount > 0) {
    breakdown.push(`오후 ${String(person.afternoonCount)}`);
  }

  const unknown = person.halfCount - person.morningCount - person.afternoonCount;
  if (unknown > 0) {
    breakdown.push(`미표기 ${String(unknown)}`);
  }

  return `${String(person.halfCount)}회 (${breakdown.join(', ')})`;
}

/** 이름별 연차 사용 요약 (누가 며칠 썼고 언제 썼는지) */
export function LeavePersonTable({ people }: LeavePersonTableProps): JSX.Element {
  if (people.length === 0) {
    return <p className="empty">추출된 연차가 없습니다.</p>;
  }

  return (
    <div className="table-wrapper">
      <table className="result-table">
        <thead>
          <tr>
            <th>{LEAVE_COLUMN_LABELS.department}</th>
            <th>{LEAVE_COLUMN_LABELS.name}</th>
            <th>{LEAVE_COLUMN_LABELS.fullCount}</th>
            <th>{LEAVE_COLUMN_LABELS.halfCount}</th>
            <th>{LEAVE_COLUMN_LABELS.publicCount}</th>
            <th>{LEAVE_COLUMN_LABELS.familyEventCount}</th>
            <th>{LEAVE_COLUMN_LABELS.totalDays}</th>
            <th>{LEAVE_COLUMN_LABELS.dates}</th>
          </tr>
        </thead>
        <tbody>
          {people.map((person) => (
            <tr key={person.name}>
              <td>{person.department}</td>
              <td>{person.name}</td>
              <td className="numeric">{person.fullCount}일</td>
              <td className="numeric">{formatHalfCount(person)}</td>
              <td className="numeric">{person.publicCount}일</td>
              <td className="numeric">{person.familyEventCount}일</td>
              <td className="numeric">{person.totalDays}일</td>
              <td className="wrap">{person.dates.join(', ')}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}
