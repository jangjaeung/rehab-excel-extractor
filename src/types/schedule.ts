/** 근무표 읽기/쓰기에 사용하는 타입 정의 */

/** 근무표 셀 한 칸에 값을 넣은 결과 */
export type ApplyStatus = '입력' | '동일' | '덮어씀' | '미반영';

/** 연차 기록 하나를 근무표에 반영한 결과 */
export interface AppliedCell {
  isoDate: string;
  weekday: string;
  name: string;
  department: string;
  /** 근무표에 넣은 값 (off / 오전 off / 공가 off …) */
  marker: string;
  /** 반영된 시트 이름. 반영하지 못했으면 null */
  sheetName: string | null;
  /** 반영된 셀 주소 (예: Y6). 반영하지 못했으면 null */
  address: string | null;
  status: ApplyStatus;
  /** 원래 들어 있던 값 (덮어쓴 경우) */
  previous: string | null;
  /** 미반영 사유 */
  reason: string | null;
}

/** 근무표 시트 하나의 구조 */
export interface ScheduleSheetInfo {
  sheetName: string;
  /** 시트가 나타내는 연도 (제목에서 읽는다) */
  year: number | null;
  /** 시트가 나타내는 월 */
  month: number | null;
  /** 날짜 칸 수 */
  dayCount: number;
  /** 명단 인원 수 */
  peopleCount: number;
}

/** 근무표에 연차를 반영한 최종 결과 */
export interface ScheduleApplyResult {
  /** 저장할 수 있는 xlsx 바이너리 (원본 서식 그대로) */
  data: Uint8Array<ArrayBuffer>;
  /** 기록별 반영 결과 (날짜 오름차순) */
  applied: AppliedCell[];
  /** 근무표 시트 구조 요약 */
  sheets: ScheduleSheetInfo[];
  warnings: string[];
}
