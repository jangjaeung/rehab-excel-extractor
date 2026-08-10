/**
 * 메인 프로세스와 렌더러가 공유하는 IPC 채널 이름 및 페이로드 타입.
 * 문자열 채널명을 여러 곳에 흩뿌리지 않기 위해 한 곳에서 정의한다.
 */

/** 결과 엑셀 저장 채널 */
export const IPC_SAVE_EXCEL = 'excel:save' as const;

/** 렌더러 → 메인: 저장할 파일 이름과 xlsx 바이너리 */
export interface SaveExcelRequest {
  /** 저장 다이얼로그에 표시할 기본 파일명 (예: 결과.xlsx) */
  defaultFileName: string;
  /** xlsx 바이너리 데이터 */
  data: Uint8Array;
}

/** 메인 → 렌더러: 저장 결과 */
export interface SaveExcelResponse {
  /** 실제로 파일이 저장되었는지 여부 (사용자가 취소하면 false) */
  saved: boolean;
  /** 저장된 전체 경로 */
  filePath?: string;
  /** 저장 실패 사유 */
  error?: string;
}
