import { useCallback, useRef, useState, type DragEvent, type ChangeEvent, type JSX } from 'react';
import { ACCEPT_ATTRIBUTE, DROPZONE_BUTTON_LABEL, DROPZONE_HINT } from '../utils/constants';

interface FileDropZoneProps {
  /** 현재 선택된 파일 (없으면 안내 문구 표시) */
  file: File | null;
  /** 파일이 선택되었을 때 호출된다. */
  onSelect: (file: File) => void;
  /** 추출/저장 중에는 파일 변경을 막는다. */
  disabled: boolean;
  /** 드롭존 위에 표시할 제목. 업로드가 여러 개일 때 무엇을 올리는지 구분한다. */
  title?: string;
  /** 버튼 문구 (기본: '엑셀 파일 선택') */
  buttonLabel?: string;
  /** 드래그 안내 문구 */
  hint?: string;
}

/**
 * 엑셀 파일 선택 영역.
 * 버튼 클릭(숨겨진 file input)과 드래그 앤 드롭을 모두 지원한다.
 * 문구만 바꿔 여러 종류의 파일 업로드에 재사용한다.
 */
export function FileDropZone({
  file,
  onSelect,
  disabled,
  title,
  buttonLabel = DROPZONE_BUTTON_LABEL,
  hint = DROPZONE_HINT,
}: FileDropZoneProps): JSX.Element {
  const inputRef = useRef<HTMLInputElement>(null);
  const [isDragging, setIsDragging] = useState(false);

  const openFileDialog = useCallback((): void => {
    inputRef.current?.click();
  }, []);

  const handleInputChange = useCallback(
    (event: ChangeEvent<HTMLInputElement>): void => {
      const selected = event.target.files?.[0];
      if (selected !== undefined) {
        onSelect(selected);
      }
      // 같은 파일을 다시 선택해도 change 가 발생하도록 값을 초기화한다.
      event.target.value = '';
    },
    [onSelect],
  );

  const handleDragOver = useCallback(
    (event: DragEvent<HTMLDivElement>): void => {
      event.preventDefault();
      if (!disabled) {
        setIsDragging(true);
      }
    },
    [disabled],
  );

  const handleDragLeave = useCallback((event: DragEvent<HTMLDivElement>): void => {
    event.preventDefault();
    setIsDragging(false);
  }, []);

  const handleDrop = useCallback(
    (event: DragEvent<HTMLDivElement>): void => {
      event.preventDefault();
      setIsDragging(false);

      if (disabled) {
        return;
      }

      const dropped = event.dataTransfer.files[0];
      if (dropped !== undefined) {
        onSelect(dropped);
      }
    },
    [disabled, onSelect],
  );

  const zoneClassName = ['dropzone', isDragging ? 'dropzone--active' : '', disabled ? 'dropzone--disabled' : '']
    .filter((token) => token !== '')
    .join(' ');

  return (
    <div className={zoneClassName} onDragOver={handleDragOver} onDragLeave={handleDragLeave} onDrop={handleDrop}>
      <input
        ref={inputRef}
        type="file"
        accept={ACCEPT_ATTRIBUTE}
        onChange={handleInputChange}
        className="dropzone__input"
        tabIndex={-1}
      />

      {title !== undefined && <p className="dropzone__title">{title}</p>}

      <button type="button" className="button button--primary" onClick={openFileDialog} disabled={disabled}>
        {buttonLabel}
      </button>

      <p className="dropzone__hint">{hint}</p>
      <p className="dropzone__filename">{file === null ? '선택된 파일 없음' : file.name}</p>
    </div>
  );
}
