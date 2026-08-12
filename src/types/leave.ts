/** 연차 시트 파싱 결과에 사용하는 타입 정의 */

/** 반차를 쓴 시간대 */
export type HalfDayPeriod = '오전' | '오후';

/**
 * 기록의 종류(= 쉬는 이유).
 * 공가(예비군·교육·검진 포함)와 경조는 연차에서 차감되지 않으므로 따로 센다.
 * 특근은 쉬는 것이 아니라 추가 근무다.
 *
 * 반차는 종류가 아니라 '길이'(half)로 따로 둔다.
 * '검진 오후반차' 처럼 이유와 길이가 함께 붙는 표기가 있기 때문이다.
 */
export type LeaveKind = '연차' | '공가' | '경조' | '특근';

/** 연차/반차 표기 1건 (= 한 사람이 하루 쓴 기록) */
export interface LeaveEntry {
  /** 사용자 이름 (근무표 명단과 같은 표기) */
  name: string;
  /** 어느 연차표에서 나온 기록인지 (OT / PT) */
  department: string;
  /** 사용 날짜 (2026-02-09) */
  isoDate: string;
  /** 화면 표시용 짧은 날짜 (2/9) */
  monthDay: string;
  /** 요일 한 글자 (월) */
  weekday: string;
  /** 쉬는 이유 (연차 / 공가 / 경조 / 특근) */
  kind: LeaveKind;
  /** 반차 여부 */
  half: boolean;
  /** 반차를 쓴 시간대. 반차가 아니거나 시트에 오전/오후가 적혀 있지 않으면 null */
  halfPeriod: HalfDayPeriod | null;
  /** 사용 일수 (종일 1, 반차 0.5, 특근 0) */
  days: number;
  /** 괄호 안에 적힌 누적 사용 표기 ('연차9' → '9'). 비어 있으면 null */
  counter: string | null;
  /** 이 기록이 있던 시트 (1월 …) */
  sheetName: string;
  /** 원본 셀 문자열 (줄바꿈은 공백으로 정리) */
  raw: string;
}

/** 사람 1명의 연차 사용 요약 */
export interface LeavePersonSummary {
  name: string;
  /** OT / PT */
  department: string;
  /** 종일 연차 건수 */
  fullCount: number;
  /** 반차 건수 */
  halfCount: number;
  /** 그중 오전 반차 건수 */
  morningCount: number;
  /** 그중 오후 반차 건수 */
  afternoonCount: number;
  /** 공가 건수 (예비군·교육·검진 포함. 연차에서 차감되지 않는다) */
  publicCount: number;
  /** 경조 건수 (연차에서 차감되지 않는다) */
  familyEventCount: number;
  /** 특근 건수 */
  specialDutyCount: number;
  /** 연차 합계 일수 (반차는 0.5로 계산, 공가·경조·특근은 제외) */
  totalDays: number;
  /** 사용한 날짜 목록 (2/9 형태, 날짜 오름차순). 공가·경조는 '7/20(공가)' 로 표시한다. */
  dates: string[];
}

/** parseLeaveExcel 의 최종 반환값 */
export interface LeaveParseResult {
  /** 날짜 오름차순으로 정렬된 전체 기록 */
  entries: LeaveEntry[];
  /** 이름별 요약 (사용 일수 많은 순) */
  people: LeavePersonSummary[];
  /** 읽은 시트 이름 목록 */
  sheetNames: string[];
  /** 확인이 필요한 비치명적 문제 */
  warnings: string[];
}
