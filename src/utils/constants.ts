/**
 * 파싱 규칙에 사용되는 상수 모음.
 * 매직 넘버/매직 스트링을 코드에 직접 쓰지 않고 여기에서만 관리한다.
 */

/** 추출 대상 항목의 접두사. 이 문자열로 시작하는 셀을 모두 항목으로 인식한다. */
export const SPRAY_ITEM_PREFIX = '신장분사';

/** 합계 값이 들어 있는 열을 찾기 위한 헤더 문자열 */
export const TOTAL_COUNT_HEADER = '합계건수';

/** 감염치료건수 행을 찾기 위한 라벨 */
export const INFECTION_ROW_LABEL = '감염치료건수';

/** 이름 셀 아래로 PT 번호를 찾을 때 몇 행까지 볼지 */
export const MAX_PT_LOOKUP_BELOW_ROWS = 3;

/** 블록 머리글의 날짜 칸에서 일자를 뽑는 패턴 ('1일' → 1) */
export const DAY_LABEL_PATTERN = /^(\d{1,2})\s*일$/;

/** 감염치료 결과 테이블의 합계 컬럼 이름 */
export const INFECTION_TOTAL_COLUMN = '합계';

/** 감염치료 결과 저장 파일명 */
export const INFECTION_RESULT_FILE_NAME = '감염치료건수.xlsx';

/** 감염치료 결과 시트 이름 */
export const INFECTION_RESULT_SHEET_NAME = '감염치료건수';

/** PT 번호 패턴. 'PT288', 'pt 288', 'PT-288' 등을 허용한다. */
export const PT_NUMBER_PATTERN = /PT\s*-?\s*(\d+)/i;

/**
 * 시트 하나가 한 주차다. ('26년 7월 1일~4일' 처럼 기간이 이름에 들어간다)
 * 시트 개수가 4주든 5주든 그대로 따라가므로 주차 수를 고정하지 않는다.
 */
export const WEEK_LABEL_SUFFIX = '주차';

/** 숫자를 찾지 못했을 때 사용할 기본 건수 */
export const DEFAULT_ITEM_COUNT = 0;

/** 블록 안에서 이름이 아닌 고정 라벨 (PT번호 없이 이름을 찾을 때 걸러 낸다) */
export const BLOCK_LABEL_WORDS = ['치료사', '팀장', '건수', '유무', '타임', '월'];

/** PT 번호 행 위로 이름을 찾을 때 최대 몇 행까지 거슬러 올라갈지 (빈 행 대비) */
export const MAX_NAME_LOOKUP_ROWS = 5;

/** 신장분사 행에서 위로 PT 번호를 찾을 때 최대 탐색 행 수 */
export const MAX_PT_LOOKUP_ROWS = 200;

/** 결과 저장 기본 파일명 */
export const RESULT_FILE_NAME = '결과.xlsx';

/** 결과 테이블 고정 컬럼 라벨 */
export const COLUMN_LABEL_THERAPIST = '이름';
export const COLUMN_LABEL_PT = 'PT번호';

/** 결과 엑셀 시트 이름 */
export const RESULT_SHEET_NAME = '신장분사';

/** 선택 가능한 엑셀 확장자 */
export const ACCEPTED_EXTENSIONS = ['.xlsx', '.xls', '.xlsm'] as const;

/** file input 의 accept 속성 값 */
export const ACCEPT_ATTRIBUTE = ACCEPTED_EXTENSIONS.join(',');

/* ----- 연차 추출 ----- */

/**
 * 날짜 셀로 인정할 엑셀 시리얼 값 범위 (2000-01-01 ~ 2100-01-01).
 * 달력에는 '6' 같은 일반 숫자도 섞여 있어서, 범위로 진짜 날짜 셀만 걸러 낸다.
 */
export const MIN_DATE_SERIAL = 36526;
export const MAX_DATE_SERIAL = 73050;

/** 한 행에 날짜 셀이 이 개수 이상이면 '날짜 행'(달력의 한 주 머리글)으로 본다. */
export const MIN_DATE_CELLS_IN_ROW = 3;

/** 날짜 행 아래 몇 행까지를 그 주의 내용 영역으로 볼지 (달력 칸 높이) */
export const MAX_WEEK_BLOCK_ROWS = 4;

/**
 * '이름(내용)' 형태의 연차 표기.
 * 이름에 공백이 섞인 경우('윤 송')까지 잡기 위해 한글 토큰을 여러 개 허용한다.
 */
export const LEAVE_ENTRY_PATTERN = /([가-힣]+(?:\s+[가-힣]+)*)\s*\(([^)]*)\)/;

/** 괄호 안에서 누적 사용일수를 뽑는 패턴 ('연차9' → 9, '오후반차 9.5' → 9.5) */
export const LEAVE_COUNTER_PATTERN = /\d+(?:\.\d+)?/;

/**
 * 이름 앞뒤(때로는 이름에 붙어서) 나오는 수식어. 이름으로 오인하지 않도록 걸러 낸다.
 * 긴 낱말이 먼저 지워지도록 '오전반차' 를 '오전'·'반차' 보다 앞에 둔다.
 */
export const LEAVE_QUALIFIERS = [
  '오전반차',
  '오후반차',
  '경조휴가',
  '경조사',
  '예비군',
  '오전',
  '오후',
  '반차',
  '연차',
  '휴가',
  '공가',
  '경조',
  '검진',
  '훈련',
  '교육',
];

/** 기록 종류 (= 쉬는 이유) */
export const ANNUAL_LEAVE_KEYWORD = '연차';
export const PUBLIC_LEAVE_KEYWORD = '공가';
export const FAMILY_EVENT_KEYWORD = '경조';
export const SPECIAL_DUTY_KEYWORD = '특근';

/** 길이(반차) 판정 키워드. 종류와 별개로 붙는다. ('검진 오후반차') */
export const HALF_DAY_KEYWORD = '반차';

/**
 * 셀에 적힌 낱말 → 쉬는 이유.
 * 위에서부터 먼저 맞는 것을 쓰므로, 좁은 낱말('경조사')을 넓은 낱말('경조')보다 앞에 둔다.
 * 종류를 늘리려면 이 표에만 한 줄 추가하면 된다.
 */
export const LEAVE_KIND_KEYWORDS = [
  { keyword: '경조휴가', kind: FAMILY_EVENT_KEYWORD },
  { keyword: '경조사', kind: FAMILY_EVENT_KEYWORD },
  { keyword: '경조', kind: FAMILY_EVENT_KEYWORD },
  { keyword: '예비군', kind: PUBLIC_LEAVE_KEYWORD },
  { keyword: '교육', kind: PUBLIC_LEAVE_KEYWORD },
  { keyword: '검진', kind: PUBLIC_LEAVE_KEYWORD },
  { keyword: '공가', kind: PUBLIC_LEAVE_KEYWORD },
] as const;

/** 연차에서 차감되는 종류 (공가·경조·특근은 차감되지 않는다) */
export const DEDUCTED_KINDS = [ANNUAL_LEAVE_KEYWORD];

/** 대괄호로 묶인 인원은 특근으로 본다. ('[홍길동]') */
export const SPECIAL_DUTY_BRACKET_PATTERN = /\[([^\]]*)\]/g;

/** 반차 시간대 판정 키워드 */
export const MORNING_KEYWORD = '오전';
export const AFTERNOON_KEYWORD = '오후';

/** 사용 일수 */
export const FULL_DAY_VALUE = 1;
export const HALF_DAY_VALUE = 0.5;

/* ----- 공휴일·휴업일 (연차표 날짜 글자색) ----- */

/** 붉은 글자로 인정할 최소 빨강 값 */
export const HOLIDAY_RED_MIN_RED = 100;

/** 초록/파랑이 빨강의 몇 배 미만이어야 '붉은색' 으로 볼지 */
export const HOLIDAY_RED_DOMINANCE = 0.5;

/** 시트 제목(연·월)을 찾을 때 훑어볼 열 수 */
export const HOLIDAY_TITLE_SCAN_COLS = 14;

/** 요일 라벨 (Date 의 getUTCDay 순서) */
export const WEEKDAY_LABELS = ['일', '월', '화', '수', '목', '금', '토'];

/** 연차 결과 테이블 컬럼 라벨 */
export const LEAVE_COLUMN_LABELS = {
  name: '이름',
  department: '구분',
  date: '날짜',
  weekday: '요일',
  kind: '구분',
  counter: '표기',
  sheet: '시트',
  raw: '원본',
  fullCount: '연차',
  halfCount: '반차 (오전/오후)',
  publicCount: '공가 (예비군·교육·검진)',
  familyEventCount: '경조',
  specialDutyCount: '특근',
  totalDays: '연차 합계',
  dates: '사용 날짜',
} as const;

/* ----- 근무표 ----- */

/** 근무표에서 날짜 행을 찾는 기준이 되는 머리글. 이 셀이 있는 열이 성명 열이기도 하다. */
export const SCHEDULE_DATE_HEADER = '날짜';

/** 머리글을 찾을 때 훑어볼 범위 (근무표 상단 제목·결재란을 지나칠 정도면 충분하다) */
export const SCHEDULE_HEADER_SCAN_ROWS = 20;
export const SCHEDULE_HEADER_SCAN_COLS = 12;

/** 명단이 끝났다고 볼 빈 행 연속 개수 */
export const SCHEDULE_NAME_GAP_LIMIT = 5;

/** 성명 열에 섞여 있는, 사람이 아닌 라벨 */
export const SCHEDULE_NON_NAME_LABELS = ['비고', '성명', '요일', '합계', '구분', '근무', '소계', '기타'];

/** 근무표 제목에서 연도를 읽는 패턴 */
export const SCHEDULE_YEAR_PATTERN = /(\d{4})\s*년/;
/** 시트 이름/제목에서 월을 읽는 패턴 */
export const SCHEDULE_MONTH_PATTERN = /(\d{1,2})\s*월/;

/** 날짜 칸으로 인정할 값의 범위 */
export const MIN_DAY_OF_MONTH = 1;
export const MAX_DAY_OF_MONTH = 31;

/**
 * 근무표에 넣을 표기.
 * '이유' 앞에 '오전/오후' 가 붙는 구조라 조각으로 두고 조합한다.
 *   연차 종일 → off        연차 오후 → 오후 off
 *   공가 종일 → 공가 off    공가 오후 → 오후 공가 off
 */
export const SCHEDULE_OFF_MARKER = 'off';
export const SCHEDULE_WORK_MARKER = 'D';
export const SCHEDULE_REST_MARKER = '·';

/** 원래 들어 있어도 덮어썼다고 보지 않는 값 (근무 표시·빈칸) */
export const SCHEDULE_PLAIN_VALUES = ['', 'D', '·', '.', '-'];

/** 집계 컬럼 머리글 (공백을 없앤 형태로 비교한다) */
export const SCHEDULE_OFF_COUNT_HEADER = '오프';
export const SCHEDULE_LEAVE_COUNT_HEADER = '월차연차';
export const SCHEDULE_SATURDAY_COUNT_HEADER = '토요일근무';

/** 토요일 요일 라벨 (요일 번호는 WEEKDAY_LABELS 에서 찾아 쓴다) */
export const SATURDAY_LABEL = '토';

/** 반차 한 건이 차지하는 일수 */
export const SCHEDULE_HALF_UNIT = 0.5;
export const SCHEDULE_FULL_UNIT = 1;

/** 주말 요일 라벨. 토·일은 사람마다 근무가 달라 자동으로 채우지 않는다. */
export const WEEKEND_LABELS = ['일', '토'];

/** 평일을 D 로 채울 요일 번호 (Date 의 getUTCDay 기준) */
export const SCHEDULE_WORKING_WEEKDAYS = WEEKDAY_LABELS.map((label, index) => ({ label, index }))
  .filter((weekday) => !WEEKEND_LABELS.includes(weekday.label))
  .map((weekday) => weekday.index);

/** 저장 파일명 뒤에 붙일 말 */
export const SCHEDULE_RESULT_SUFFIX = '_연차반영';

/** 근무표 반영 결과 테이블 컬럼 라벨 */
export const APPLY_COLUMN_LABELS = {
  date: '날짜',
  weekday: '요일',
  name: '이름',
  department: '구분',
  marker: '넣은 값',
  sheet: '근무표 시트',
  address: '셀',
  status: '결과',
  detail: '비고',
};

/* ----- 화면 탭 ----- */

/** 화면 상단 탭 목록. 탭을 추가하려면 여기에 항목을 넣고 App 에서 분기만 추가하면 된다. */
export const TAB_ITEMS = [
  {
    id: 'spray',
    label: '신장분사 추출기',
    title: '신장분사 실적 추출기',
    description: '치료사별 신장분사 항목의 합계건수를 자동으로 추출합니다.',
  },
  {
    id: 'infection',
    label: '감염치료건수',
    title: '감염치료건수 추출기',
    description: '주간 현황관리판의 모든 시트를 합쳐 1일부터 말일까지 인원별 감염치료건수를 추출합니다.',
  },
  {
    id: 'leave',
    label: '연차 추출기',
    title: '연차 추출기',
    description: '연차 엑셀 시트와 근무표를 올려 연차 사용 내역을 추출합니다.',
  },
] as const;

/** 탭 식별자 (TAB_ITEMS 에서 자동 도출) */
export type TabId = (typeof TAB_ITEMS)[number]['id'];

/** 최초 진입 시 열려 있는 탭 */
export const DEFAULT_TAB_ID: TabId = 'spray';

/** 연차 추출기에서 업로드받는 파일 라벨 (오류 메시지에도 사용) */
export const WORK_SCHEDULE_LABEL = '근무표';

/** 연차 추출기의 업로드 항목 3개 */
export const LEAVE_SLOTS = [
  { id: 'ot', title: 'OT 연차', department: 'OT', hint: 'OT 인원의 연차 계획표' },
  { id: 'pt', title: 'PT 연차', department: 'PT', hint: 'PT 인원의 연차 계획표' },
  { id: 'schedule', title: WORK_SCHEDULE_LABEL, department: '', hint: '연차를 써 넣을 근무표 (필수)' },
] as const;

/** 업로드 항목 식별자 */
export type LeaveSlotId = (typeof LEAVE_SLOTS)[number]['id'];

/** 드롭존 기본 문구 */
export const DROPZONE_BUTTON_LABEL = '엑셀 파일 선택';
export const DROPZONE_HINT = '또는 이곳에 엑셀 파일을 끌어다 놓으세요';
