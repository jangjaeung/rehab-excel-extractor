import type { TherapistRecord } from '../types/excel';
import { CLIPBOARD_COLUMN_SEPARATOR, CLIPBOARD_ROW_SEPARATOR, DEFAULT_ITEM_COUNT } from './constants';

/**
 * 표 값을 엑셀에 그대로 붙여넣을 수 있는 형태로 만든다.
 *
 * 엑셀은 붙여넣기 글자를 이렇게 해석한다.
 *   탭    → 오른쪽 옆 칸
 *   줄바꿈 → 아래 칸
 * 그래서 값을 탭으로 잇고 줄바꿈으로 나누면 셀 하나하나에 알아서 들어간다.
 *
 * 줄바꿈은 \r\n 을 쓴다. 윈도우 엑셀이 가장 확실하게 행으로 인식한다.
 */
export function toClipboardTable(rows: readonly TherapistRecord[], columns: readonly string[]): string {
  return rows
    .map((row) =>
      columns.map((column) => String(row.items[column] ?? DEFAULT_ITEM_COUNT)).join(CLIPBOARD_COLUMN_SEPARATOR),
    )
    .join(CLIPBOARD_ROW_SEPARATOR);
}

/**
 * 글자를 클립보드에 넣는다.
 *
 * navigator.clipboard 는 https 나 localhost 에서만 동작하므로,
 * 사내에서 http 로 열었을 때를 대비해 예전 방식으로 한 번 더 시도한다.
 */
export async function copyText(text: string): Promise<boolean> {
  try {
    await navigator.clipboard.writeText(text);
    return true;
  } catch {
    return copyWithTextarea(text);
  }
}

/** navigator.clipboard 를 쓸 수 없을 때의 대체 방법 */
function copyWithTextarea(text: string): boolean {
  const textarea = document.createElement('textarea');
  textarea.value = text;
  // 화면에 보이지 않으면서도 선택은 가능해야 한다.
  textarea.setAttribute('readonly', '');
  textarea.style.position = 'fixed';
  textarea.style.opacity = '0';

  document.body.appendChild(textarea);
  textarea.select();

  try {
    // execCommand 는 폐기 예정이지만, https 가 아닌 곳에서 쓸 수 있는 유일한 방법이라 남겨 둔다.
    // eslint-disable-next-line @typescript-eslint/no-deprecated
    return document.execCommand('copy');
  } catch {
    return false;
  } finally {
    document.body.removeChild(textarea);
  }
}
