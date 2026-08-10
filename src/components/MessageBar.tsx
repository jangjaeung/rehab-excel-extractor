import type { JSX } from 'react';

interface MessageBarProps {
  /** 표시할 메시지 (null 이면 렌더링하지 않는다) */
  message: string | null;
  /** 메시지 성격에 따른 색상 구분 */
  tone: 'error' | 'info';
}

/** 오류 / 안내 메시지를 한 줄로 보여 주는 공용 컴포넌트 */
export function MessageBar({ message, tone }: MessageBarProps): JSX.Element | null {
  if (message === null) {
    return null;
  }
  return <p className={`message message--${tone}`}>{message}</p>;
}
