import { ACCEPTED_EXTENSIONS } from './constants';

/**
 * 파일이 지원하는 엑셀 확장자인지 검사한다.
 * 드래그 앤 드롭으로 아무 파일이나 들어올 수 있으므로 UI 진입 시점에서 걸러 낸다.
 */
export function isExcelFile(file: File): boolean {
  const lowerName = file.name.toLowerCase();
  return ACCEPTED_EXTENSIONS.some((extension) => lowerName.endsWith(extension));
}

/** 오류 객체에서 사용자에게 보여줄 메시지를 뽑아낸다. */
export function toErrorMessage(error: unknown): string {
  if (error instanceof Error) {
    return error.message;
  }
  return String(error);
}
