import { useState, type JSX } from 'react';
import { Tabs } from './components/Tabs';
import { InfectionTab } from './tabs/InfectionTab';
import { LeaveTab } from './tabs/LeaveTab';
import { SprayTab } from './tabs/SprayTab';
import { DEFAULT_TAB_ID, TAB_ITEMS, type TabId } from './utils/constants';

/**
 * 화면 전체 레이아웃.
 * 탭 전환과 헤더 표시만 담당하고, 각 탭의 동작은 tabs/ 아래 컴포넌트가 가진다.
 *
 * 두 탭을 모두 마운트한 채 감추기만 하는 이유는,
 * 탭을 옮겼다 돌아왔을 때 추출 결과나 선택한 파일이 사라지지 않게 하기 위함이다.
 */
export default function App(): JSX.Element {
  const [activeTab, setActiveTab] = useState<TabId>(DEFAULT_TAB_ID);

  const current = TAB_ITEMS.find((tab) => tab.id === activeTab) ?? TAB_ITEMS[0];

  return (
    <div className="app">
      <Tabs active={activeTab} onChange={setActiveTab} />

      <header className="app__header">
        <h1>{current.title}</h1>
        <p className="app__subtitle">{current.description}</p>
      </header>

      <div hidden={activeTab !== 'spray'}>
        <SprayTab />
      </div>
      <div hidden={activeTab !== 'infection'}>
        <InfectionTab />
      </div>
      <div hidden={activeTab !== 'leave'}>
        <LeaveTab />
      </div>
    </div>
  );
}
