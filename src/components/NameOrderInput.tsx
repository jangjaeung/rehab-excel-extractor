import type { ChangeEvent, JSX } from 'react';
import { NAME_ORDER_HINT, NAME_ORDER_LABEL, NAME_ORDER_ROWS } from '../utils/constants';

interface NameOrderInputProps {
  /** 붙여넣은 원본 글자 */
  value: string;
  /** 인식된 인원 (개수 표시용) */
  names: readonly string[];
  /** 붙여넣은 이름 중 결과에 없는 것 */
  missing: readonly string[];
  onChange: (value: string) => void;
  disabled: boolean;
}

/**
 * 결과 표의 인원 순서를 정하는 입력.
 *
 * 엑셀에서 이름 세로줄을 복사해 붙여넣는 용도라 줄바꿈이 그대로 남아야 하므로
 * input 이 아니라 textarea 를 쓴다. (input 은 붙여넣을 때 줄바꿈이 공백으로 뭉개진다)
 */
export function NameOrderInput({ value, names, missing, onChange, disabled }: NameOrderInputProps): JSX.Element {
  const handleChange = (event: ChangeEvent<HTMLTextAreaElement>): void => {
    onChange(event.target.value);
  };

  return (
    <div className="name-order">
      <div className="name-order__head">
        <label className="name-order__label" htmlFor="name-order">
          {NAME_ORDER_LABEL}
        </label>
        {names.length > 0 && <span className="name-order__count">{names.length}명 인식</span>}
        {value !== '' && (
          <button
            type="button"
            className="name-order__clear"
            onClick={() => {
              onChange('');
            }}
            disabled={disabled}
          >
            지우기
          </button>
        )}
      </div>

      <textarea
        id="name-order"
        className="name-order__input"
        rows={NAME_ORDER_ROWS}
        value={value}
        onChange={handleChange}
        disabled={disabled}
        placeholder={NAME_ORDER_HINT}
        spellCheck={false}
      />

      {missing.length > 0 && (
        <p className="name-order__missing">
          결과에 없는 이름 {missing.length}명: {missing.join(', ')}
        </p>
      )}
    </div>
  );
}
