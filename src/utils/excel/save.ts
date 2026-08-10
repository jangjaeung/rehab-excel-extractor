/**
 * 만들어 둔 엑셀 바이너리를 파일로 저장하는 공용 헬퍼.
 * 렌더러에는 fs 접근 권한이 없으므로 Electron 에서는 IPC 로 메인 프로세스에 맡긴다.
 */

/**
 * 엑셀 바이너리를 저장한다.
 * Electron 환경에서는 네이티브 저장 다이얼로그를, 그 외(브라우저 개발 모드)에서는
 * 다운로드 방식을 사용한다.
 *
 * @returns 저장된 경로. 사용자가 취소하면 null.
 */
export async function saveWorkbookData(data: Uint8Array<ArrayBuffer>, defaultFileName: string): Promise<string | null> {
  const api = window.electronAPI;

  if (api === undefined) {
    downloadInBrowser(data, defaultFileName);
    return defaultFileName;
  }

  const response = await api.saveExcel({ defaultFileName, data });
  if (!response.saved) {
    if (response.error !== undefined) {
      throw new Error(`엑셀 저장에 실패했습니다: ${response.error}`);
    }
    return null;
  }

  return response.filePath ?? defaultFileName;
}

/** Electron 없이 브라우저에서 실행할 때의 저장 방식 */
function downloadInBrowser(data: Uint8Array<ArrayBuffer>, fileName: string): void {
  const blob = new Blob([data], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = fileName;
  anchor.click();
  URL.revokeObjectURL(url);
}
