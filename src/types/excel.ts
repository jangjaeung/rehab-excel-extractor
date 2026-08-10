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

/** parseExcel 의 최종 반환값 */
export interface ParseResult {
  /** 실제로 파싱한 시트 이름 */
  sheetName: string;
  /** 오름차순 정렬된 신장분사 항목 컬럼 목록 */
  columns: string[];
  /** 치료사별 결과 (시트에 등장한 순서 유지) */
  rows: TherapistRecord[];
  /** 파싱 도중 발견한 비치명적 문제 (예: PT번호를 찾지 못한 행) */
  warnings: string[];
}
