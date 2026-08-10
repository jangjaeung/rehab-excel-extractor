import type { JSX } from 'react';
import { TAB_ITEMS, type TabId } from '../utils/constants';

interface TabsProps {
  /** 현재 열려 있는 탭 */
  active: TabId;
  /** 탭을 클릭했을 때 호출된다. */
  onChange: (id: TabId) => void;
}

/**
 * 화면 상단 탭 바.
 * 탭 목록은 constants 의 TAB_ITEMS 하나만 보고 그린다.
 */
export function Tabs({ active, onChange }: TabsProps): JSX.Element {
  return (
    <nav className="tabs" role="tablist">
      {TAB_ITEMS.map((tab) => {
        const isActive = tab.id === active;

        return (
          <button
            key={tab.id}
            type="button"
            role="tab"
            aria-selected={isActive}
            className={isActive ? 'tabs__item tabs__item--active' : 'tabs__item'}
            onClick={() => {
              onChange(tab.id);
            }}
          >
            {tab.label}
          </button>
        );
      })}
    </nav>
  );
}
