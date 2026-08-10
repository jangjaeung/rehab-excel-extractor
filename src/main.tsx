import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import App from './App';
import './styles.css';

/**
 * 드롭존 밖에 파일을 떨어뜨리면 Electron 창이 해당 파일로 이동해 버린다.
 * 창 전체에서 기본 동작을 막아 앱이 사라지는 것을 방지한다.
 */
function preventWindowFileDrop(): void {
  const cancel = (event: Event): void => {
    event.preventDefault();
  };
  window.addEventListener('dragover', cancel);
  window.addEventListener('drop', cancel);
}

preventWindowFileDrop();

const container = document.getElementById('root');
if (container === null) {
  throw new Error('#root 엘리먼트를 찾을 수 없습니다.');
}

createRoot(container).render(
  <StrictMode>
    <App />
  </StrictMode>,
);
