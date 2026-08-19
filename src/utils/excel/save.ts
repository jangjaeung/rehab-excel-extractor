/**
 * 만들어 둔 엑셀 바이너리를 파일로 내려받는 헬퍼.
 * 브라우저에서 동작하므로 Blob 을 만들어 다운로드시킨다.
 */

/** 엑셀 MIME 타입 */
const XLSX_MIME_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

/**
 * 엑셀 바이너리를 내려받는다.
 *
 * 브라우저는 저장 위치를 알려 주지 않으므로(실제 경로는 다운로드 폴더),
 * 화면에 보여 줄 이름으로 파일명을 그대로 돌려준다.
 */
export function saveWorkbookData(data: Uint8Array<ArrayBuffer>, fileName: string): Promise<string | null> {
  const blob = new Blob([data], { type: XLSX_MIME_TYPE });
  const url = URL.createObjectURL(blob);

  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = fileName;
  anchor.click();

  URL.revokeObjectURL(url);
  return Promise.resolve(fileName);
}
