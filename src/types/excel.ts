/** 엑셀 파싱과 결과 표현에 사용하는 공용 타입 정의 */

/** 시트 셀 하나가 가질 수 있는 원시 값 (빈 셀은 null) */
export type CellValue = string | number | boolean | Date | null;

/** 행 우선(row-major) 2차원 시트 데이터. grid[row][col] 형태로 접근한다. */
export type SheetGrid = readonly (readonly CellValue[])[];

/** 시트 상의 좌표 (0-based) */
export interface CellPosition {
  readonly row: number;
  readonly col: number;
}

/** 치료사 1명의 추출 결과 */
export interface TherapistRecord {
  /** 치료사 이름 (예: 허정훈) */
  therapist: string;
  /** PT 번호 (예: PT288) */
  pt: string;
  /** 신장분사 항목명 → 합계건수 */
  items: Record<string, number>;
}

/** 주차 1개(시트 1개)의 추출 결과 */
export interface WeekResult {
  /** 원본 시트 이름 (예: 26년 7월 1일~4일) */
  sheetName: string;
  /** 화면 표시용 순번 (예: 1주차) */
  label: string;
  /** 치료사별 결과 (시트에 등장한 순서 유지) */
  rows: TherapistRecord[];
  /** 이 주차에서 발견한 비치명적 문제 */
  warnings: string[];
}

/** parseExcel 의 최종 반환값 */
export interface ParseResult {
  /** 주차별 결과 (시트 순서 유지) */
  weeks: WeekResult[];
  /**
   * 모든 주차를 합친 항목 컬럼 목록 (오름차순).
   * 주차마다 등장하는 항목이 달라도 표 모양이 같아야 비교할 수 있으므로 하나로 통일한다.
   */
  columns: string[];
  /** 파일 전체에 해당하는 비치명적 문제 */
  warnings: string[];
}
