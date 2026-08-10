/** preload 가 contextBridge 로 노출하는 API 의 렌더러 측 타입 선언 */

interface SaveExcelRequest {
  /** 저장 다이얼로그의 기본 파일명 */
  defaultFileName: string;
  /** xlsx 바이너리 */
  data: Uint8Array;
}

interface SaveExcelResponse {
  /** 실제 저장 여부 (사용자가 취소하면 false) */
  saved: boolean;
  /** 저장된 전체 경로 */
  filePath?: string;
  /** 실패 사유 */
  error?: string;
}

interface ElectronAPI {
  saveExcel: (request: SaveExcelRequest) => Promise<SaveExcelResponse>;
}

declare global {
  interface Window {
    /** Electron 이 아닌 환경(브라우저 개발 모드)에서는 undefined */
    electronAPI?: ElectronAPI;
  }
}

export {};
