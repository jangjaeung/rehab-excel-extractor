# 신장분사 실적 추출기

병원 치료 실적 엑셀에서 **치료사별 「신장분사」 항목의 합계건수**를 자동으로 추출하는 Electron 데스크톱 프로그램입니다.

매달 양식에서 **날짜 개수 · 날짜 컬럼 위치 · 전체 컬럼 수**가 달라져도 동작하도록,
셀 주소(A1, B3)나 고정 컬럼 번호를 전혀 사용하지 않고 **셀의 내용만으로 탐색**합니다.

---

## 1. 프로젝트 구조

```
d:\toy
├─ electron/                     # 메인 프로세스 (Node 측)
│  ├─ main.ts                    # 윈도우 생성, 저장 다이얼로그 + 파일 쓰기 IPC 처리
│  ├─ preload.ts                 # contextBridge 로 window.electronAPI 노출
│  └─ ipc.ts                     # IPC 채널명/페이로드 타입 공유 정의
│
├─ src/                          # 렌더러 (React 측)
│  ├─ components/                # 순수 UI 컴포넌트 (로직 없음)
│  │  ├─ FileDropZone.tsx        # [엑셀 파일 선택] 버튼 + Drag & Drop + 파일명 표시
│  │  ├─ ResultTable.tsx         # 결과 테이블 (컬럼은 파서가 준 순서를 그대로 사용)
│  │  ├─ MessageBar.tsx          # 오류/안내 메시지 한 줄 표시
│  │  └─ WarningList.tsx         # 파싱 중 발생한 비치명적 경고 목록
│  │
│  ├─ hooks/
│  │  └─ useExcelExtraction.ts   # 파일선택 → 추출 → 저장 흐름과 화면 상태 관리
│  │
│  ├─ utils/
│  │  ├─ constants.ts            # '신장분사', '합계건수', PT 정규식 등 모든 상수
│  │  ├─ file.ts                 # 확장자 검증, 에러 메시지 변환
│  │  └─ excel/
│  │     ├─ parser.ts            # ★ 엑셀 파싱 전담 (핵심 로직)
│  │     ├─ cell.ts              # 셀 값/그리드 접근 저수준 헬퍼
│  │     └─ exporter.ts          # 결과 → 결과.xlsx 생성 및 저장
│  │
│  ├─ types/
│  │  ├─ excel.ts                # CellValue, SheetGrid, TherapistRecord, ParseResult
│  │  └─ electron.d.ts           # window.electronAPI 타입 선언
│  │
│  ├─ App.tsx                    # 화면 배치
│  ├─ main.tsx                   # React 진입점
│  └─ styles.css
│
├─ index.html
├─ vite.config.ts
├─ tsconfig.json                 # 렌더러용 (strict + noUncheckedIndexedAccess)
├─ tsconfig.electron.json        # 메인 프로세스용
├─ eslint.config.mjs             # ESLint flat config (typescript-eslint strictTypeChecked)
└─ package.json
```

**레이어 분리 원칙**

| 레이어 | 책임 | 엑셀을 아는가 |
|---|---|---|
| `components/` | 화면 표시, 사용자 입력 | ✕ |
| `hooks/` | 상태 전이(로딩/에러/결과) | ✕ (parser 호출만) |
| `utils/excel/` | 파싱·변환·저장 | ○ |

UI 는 `parseExcel(file)` 과 `saveResultWorkbook(result)` 두 함수만 호출합니다.

---

## 2. 실행 방법

### 설치

```bash
npm install
```

### 개발 모드 (HMR)

```bash
npm run dev
```

Vite 개발 서버(5173)가 뜬 뒤 Electron 창이 자동으로 열립니다.

### 프로덕션 실행

```bash
npm start          # 빌드 후 Electron 실행
```

### 설치 파일 생성

```bash
npm run package    # release/ 폴더에 Windows 설치 파일 생성
```

생성물: `release\StretchSprayExtractor Setup 1.0.0.exe` (약 85MB, 사용자 단위 설치)
설치 없이 바로 쓰려면 `release\win-unpacked\StretchSprayExtractor.exe` 를 실행하면 됩니다.
(무설치판은 **폴더째** 옮겨야 합니다. exe 하나만 복사하면 실행되지 않습니다.)

> **코드 서명 도구(winCodeSign) 심볼릭 링크 문제는 자동으로 처리됩니다**
>
> electron-builder 는 설치 파일을 만들기 직전에 `winCodeSign` 압축을 푸는데,
> 그 안에 **macOS 전용 심볼릭 링크**(`darwin/.../libcrypto.dylib`)가 들어 있습니다.
> Windows 에서 심볼릭 링크를 만들려면 관리자 권한이나 개발자 모드가 필요해서 아래 오류로 멈춥니다.
>
> ```
> ⨯ cannot execute  cause=exit status 2
> ERROR: Cannot create symbolic link : 클라이언트가 필요한 권한을 가지고 있지 않습니다.
> ```
>
> 이때 `release\win-unpacked` 까지만 만들어지고 **설치 파일은 나오지 않습니다.**
>
> 그래서 `package` 스크립트가 [scripts/prepare-wincodesign.mjs](scripts/prepare-wincodesign.mjs) 를 먼저 실행합니다.
> Windows 빌드에 `darwin` 폴더는 필요 없으므로 그 폴더만 빼고 캐시에 미리 풀어 두고,
> electron-builder 는 '이미 있다' 고 보고 압축 해제 단계를 건너뜁니다.
>
> * 관리자 권한이 필요 없습니다.
> * 캐시가 이미 준비돼 있으면 아무 일도 하지 않으므로 매번 실행해도 됩니다.
> * 새 PC 에서도 `pnpm install` 후 `pnpm package` 만 하면 됩니다.
> * Windows 가 아니면 그냥 건너뜁니다.

### 기타 스크립트

| 명령 | 설명 |
|---|---|
| `npm run typecheck` | 타입 검사만 수행 (strict) |
| `npm run lint` | ESLint 검사 (경고 0건 강제) |
| `npm run build` | typecheck → 렌더러 빌드 → 메인 프로세스 빌드 |

### 사용 순서

1. **[엑셀 파일 선택]** 버튼을 누르거나, 점선 영역에 파일을 **끌어다 놓습니다.**
2. **[추출하기]** 를 누르면 아래에 결과 테이블이 표시됩니다.
3. **[엑셀 저장]** 을 누르면 저장 위치를 물어본 뒤 `결과.xlsx` 로 저장합니다.

> **문제 해결** — Electron 이 즉시 종료되며 `Cannot read properties of undefined (reading 'whenReady')`
> 오류가 나면, 터미널에 `ELECTRON_RUN_AS_NODE` 환경변수가 설정되어 있는 경우입니다(일부 IDE 내장 터미널).
> 해당 변수를 제거한 뒤 다시 실행하세요.

---

## 3. 주요 로직 설명

### 3.1 왜 셀 주소를 쓰지 않는가

매달 날짜 개수가 달라지면 `합계건수` 열의 위치가 통째로 이동합니다.
따라서 모든 위치 정보를 **셀에 적힌 글자**로부터 역산합니다.

| 찾는 대상 | 판단 기준 |
|---|---|
| 항목 | 셀 내용이 `신장분사` 로 **시작**하고 뒤에 식별자가 더 있음 |
| 합계 열 | 셀 내용에 `합계건수` 가 포함된 헤더의 **열 번호** |
| 블록 경계 | `합계건수` 헤더 열 = 그 블록의 **오른쪽 끝** |
| 날짜 구간 | 항목명 셀 오른쪽 ~ `합계건수` 열 바로 앞 |
| PT 번호 | 셀 내용이 `PT\s*-?\s*숫자` 패턴과 일치 (`PT팀장` 처럼 숫자가 없으면 제외) |
| 치료사 이름 | PT 번호 셀의 **바로 위 행**의 이름 셀 |

### 3.2 블록이 좌우로도 반복되는 구조

실제 양식은 치료사 블록이 위아래로만이 아니라 **좌우로 나란히** 배치됩니다.

```
      B       C            D            E~H       I       J    K       L       M            N~Q     R
 5 │ PT팀장                7월        1일~4일  합계건수  │  PT부팀장  7월                 1일~4일 합계건수
 6 │ 치료사  풀타임치료유무                        0     │  치료사   풀타임치료유무                     0
 7 │ 허정훈  감염치료건수                          0     │  강지은   감염치료건수                       0
 8 │ PT288   도수치료건수  신장분사A20             0     │  PT287    도수치료  신장분사A20              0
 9 │                       신장분사A25             0     │                     신장분사A25              0
```

여기서 핵심은 **PT번호·이름·항목이 같은 행에 걸쳐 있다**는 점입니다.
8행에는 왼쪽 블록의 `PT288`·`신장분사A20` 과 오른쪽 블록의 `PT287`·`신장분사A20` 이 함께 존재합니다.

따라서 항목 행 **전체**를 훑어 PT번호를 찾으면 오른쪽 블록 항목까지 왼쪽 치료사에게 붙고,
같은 항목명이 중복으로 판정되어 **오른쪽 블록이 통째로 누락**됩니다.

이를 막기 위해 `resolveBlockBounds()` 가 항목이 속한 블록의 **열 범위**를 먼저 확정합니다.

* **오른쪽 경계** = 항목보다 위(같은 행 포함)의 `합계건수` 헤더 중 항목 오른쪽에서 가장 가까운 열
* **왼쪽 경계** = 같은 헤더 행에서 항목 왼쪽에 있는 `합계건수` 헤더의 바로 다음 열 (없으면 0)

예) `M8(신장분사A20)` → 오른쪽 경계 `R`, 왼쪽 경계 `J`
→ PT번호를 `J~M` 범위에서만 찾으므로 `K8(PT287)` 이 정확히 선택됩니다.

PT번호·이름 탐색과 합계 대체값 탐색이 모두 이 범위 안에서만 이뤄집니다.
블록이 좌우로 3개, 4개로 늘어나도 헤더 개수만큼 자동으로 나뉩니다.

### 3.3 전체 흐름

```
File
 └─ arrayBuffer()
     └─ XLSX.read                       시트 하나 = 한 주차 (빈 시트는 건너뜀)
         └─ buildGrid()                 시트 → 인덱스가 원본과 일치하는 2차원 배열
             ├─ findTotalHeaders()      '합계건수' 헤더 좌표 전부 수집
             ├─ findSprayItemCells()    '신장분사*' 셀 좌표 전부 수집
             └─ 각 항목 셀마다
                 ├─ resolveBlockBounds()      항목이 속한 블록의 열 범위 + 합계 열 확정
                 ├─ findTherapistIdentity()   블록 범위 안에서 PT번호 → 그 위 이름
                 ├─ readItemCount()           날짜 구간 값을 직접 합산
                 ├─ checkTotalMismatch()      시트의 합계건수와 다르면 경고 기록
                 └─ addItem()                 PT번호 기준으로 누적 (중복 무시)
                     └─ sortItemColumns()     항목명 오름차순 → 컬럼 확정
                         └─ finalizeRecord()  빠진 항목은 0 으로 채움
```

### 3.4 상태 관리 (`useExcelExtraction`)

* `file` / `result` / `status` / `error` / `notice` 5개 상태만 관리합니다.
* 새 파일을 선택하면 이전 결과를 즉시 비워 **낡은 결과가 남아 있는 상황**을 막습니다.
* `status` 가 `idle` 이 아닌 동안 모든 버튼과 드롭존이 비활성화됩니다.

### 3.5 저장 (`exporter.ts` + IPC)

렌더러에는 `fs` 접근 권한이 없습니다(`contextIsolation: true`, `nodeIntegration: false`).

```
렌더러 : buildResultWorkbook() → Uint8Array
       → window.electronAPI.saveExcel({ defaultFileName: '결과.xlsx', data })
메인   : dialog.showSaveDialog() → fs.writeFile()
       → { saved, filePath } 반환
```

* 저장되는 엑셀의 컬럼 순서는 화면 테이블과 **완전히 동일**합니다(같은 `result.columns` 사용).
* xlsx 는 UTF-8 XML 을 사용하므로 한글이 깨지지 않습니다.
* Electron 없이 브라우저에서 열었을 때는 자동으로 Blob 다운로드로 대체됩니다.

---

## 4. `parser.ts` 상세 설명

### 4.1 공개 API

```ts
parseExcel(file: File): Promise<ParseResult>      // UI 가 호출하는 유일한 함수
parseExcelBuffer(buffer: ArrayBuffer): ParseResult // 테스트용 동기 진입점
```

```ts
interface ParseResult {
  weeks: WeekResult[];    // 주차별 결과 (시트 순서 유지)
  columns: string[];      // 모든 주차를 합친 항목 목록 (오름차순)
  warnings: string[];     // 파일 전체에 해당하는 문제
}

interface WeekResult {
  sheetName: string;      // 원본 시트 이름 (26년 7월 1일~4일)
  label: string;          // 1주차
  rows: TherapistRecord[];// 치료사별 결과 (시트 등장 순서 유지)
  warnings: string[];     // 확인이 필요한 비치명적 문제
}

interface TherapistRecord {
  therapist: string;               // 허정훈
  pt: string;                      // PT288
  items: Record<string, number>;   // { 신장분사C20: 1, 신장분사C25: 26, ... }
}
```

### 4.2 내부 함수

#### `buildGrid(sheet)`
`sheet_to_json` 대신 `!ref` 범위를 직접 순회해 2차원 배열을 만듭니다.
빈 행·빈 열 때문에 **행/열 인덱스가 밀리는 문제**를 원천 차단하기 위함이며,
이후 모든 탐색은 이 그리드의 (row, col) 좌표만 사용합니다.

#### `findSprayItemCells(grid)`
전체 셀을 위→아래, 왼→오른쪽으로 순회하며
정규화된 텍스트가 `신장분사` 로 시작하고 **접두사보다 긴** 셀을 모읍니다.
항목명을 하드코딩하지 않으므로 `신장분사D20`, `신장분사E25` 가 추가되어도
**코드 수정 없이** 자동 인식·자동 컬럼 생성됩니다.

#### `findTotalHeaders(grid)`
`합계건수` 헤더는 블록마다 하나씩 반복되므로 **좌표 목록 전체**를 모아 둡니다.
합계 값을 읽는 용도이자, 블록의 가로 경계를 나누는 기준으로도 쓰입니다.

#### `resolveBlockBounds(headers, item)`  ★ 좌우 반복 블록 대응의 핵심
항목 셀 하나에 대해 `{ totalColumn, startCol, endCol }` 을 계산합니다.

* `totalColumn` / `endCol` — 항목보다 위(같은 행 포함)의 헤더 중,
  항목 **오른쪽**에서 가장 가까운 헤더의 열 (행이 가까울수록 → 열이 가까울수록 우선)
* `startCol` — 같은 헤더 행에서 항목 **왼쪽**에 있는 헤더의 다음 열 (없으면 0)

날짜 개수가 달라 합계 열 위치가 블록마다 달라도, 블록이 좌우로 몇 개가 놓여도
각 항목이 자기 블록의 열 범위만 보게 됩니다.

#### `readItemCount(grid, item, bounds)` / `sumNumericCells(...)`
건수는 시트의 합계 셀을 읽지 않고 **날짜 칸 값을 직접 더해서** 계산합니다.

원본 시트의 `합계건수` 수식 범위가 실제 항목 행과 어긋난 사례가 확인되었기 때문입니다.
예를 들어 어떤 블록에서는 `신장분사B20` 의 합계 셀이 **두 행 아래 `C20` 의 날짜**를 더하고 있어서,
합계 셀을 그대로 믿으면 값이 엉뚱한 항목에 붙습니다.

* 합산 구간 = `항목명 셀 오른쪽` ~ `합계건수 열 바로 앞`
* 빈 칸은 건너뛰고, **글자가 든 셀을 만나면 중단**합니다.
  (합계건수 헤더가 없는 양식에서 옆 블록까지 더해 버리는 것을 막는 안전장치)
* 숫자가 하나도 없으면 → `0`

#### `checkTotalMismatch(...)`
계산한 날짜 합계와 시트에 적힌 `합계건수` 값이 다르면 경고를 남깁니다.
값은 **날짜 합계를 사용**하며, 경고는 원본 수식이 깨진 위치를 찾는 용도입니다.
결과 화면의 `확인이 필요한 항목 N건` 을 펼치면 볼 수 있습니다.

```
PT211 '신장분사B20': 시트의 합계건수(3)와 날짜 합계(0)가 다릅니다. 날짜 합계를 사용했습니다.
```

#### `findTherapistIdentity(grid, item, bounds)`
항목 셀에서 위로 올라가며 PT 번호를 찾고(`findPtNumberInRow`),
그 셀 **바로 위 행**부터 이름을 찾습니다(`findNameAbove`).

* 탐색 열 범위는 `[bounds.startCol, min(bounds.endCol, 항목 열)]` 로 제한합니다.
  PT 번호는 항상 항목의 왼쪽에 있고, 다른 블록을 넘겨다보면 안 되기 때문입니다.
* 한 행 안에서는 **오른쪽부터** 탐색하여 항목에서 가장 가까운 PT 번호를 고릅니다.
* 이름은 PT 번호와 **같은 열**을 우선 확인하고, 없으면 블록 범위 안의 첫 이름 후보를 씁니다.
* 빈 행이 끼어 있을 수 있어 최대 `MAX_NAME_LOOKUP_ROWS` 행까지 거슬러 올라갑니다.
* PT 번호를 못 찾은 항목은 결과에서 제외하고 `warnings` 에 행 번호를 남깁니다.

#### `extractPtNumber(value)`
`PT288`, `pt 288`, `PT-288` 을 모두 `PT288` 로 **정규화**합니다.
표기 흔들림 때문에 같은 사람이 두 줄로 나뉘는 것을 막습니다.

#### `addItem(...)`
`PT번호`를 key 로 하는 `Map` 에 누적합니다.

* 같은 치료사가 여러 블록에 등장해도 **한 행으로 합쳐집니다.**
* 같은 치료사의 **같은 항목이 중복**되면 최초 값만 유지하고 경고를 남깁니다.
* `Map` 이므로 시트에 등장한 **순서가 그대로 유지**됩니다.

#### `sortItemColumns(itemNames)`
`localeCompare(…, 'ko', { numeric: true })` 로 오름차순 정렬합니다.
`numeric` 옵션 덕분에 `신장분사A9` 가 `신장분사A20` 보다 앞에 옵니다(문자열 정렬이면 반대가 됩니다).

#### `finalizeRecord(...)`
확정된 컬럼 목록을 기준으로 누락된 항목을 모두 `0` 으로 채웁니다.
따라서 **모든 행이 동일한 키 집합**을 갖게 되어 테이블과 엑셀 출력이 어긋나지 않습니다.

### 4.3 예외 처리 요약

| 상황 | 처리 |
|---|---|
| 빈 셀 | 무시 (`isBlank` / `normalizeText`) |
| 숫자가 없음 | `0` (`DEFAULT_ITEM_COUNT`) |
| 시트의 합계 수식이 깨짐 | 날짜 칸을 직접 합산하고, 값이 다르면 경고 |
| 신장분사 항목이 없는 치료사 | 결과에서 제외 (항목 셀에서만 레코드를 만들기 때문) |
| 시트가 여러 개 | 시트 하나를 한 주차로 보고 모두 추출 (주차 수는 고정하지 않음) |
| 빈 시트 / 항목 없는 시트 | 주차에서 제외 (`시트1`, `시트2` 등) |
| PT번호가 빠진 블록 | 블록 첫 열에서 이름을 찾아 이름으로 묶고 경고 |
| 같은 치료사 중복 등장 | PT 번호 기준으로 한 행에 병합 |
| 같은 항목 중복 등장 | 최초 값 유지 + 경고 |
| PT 번호를 못 찾음 | 해당 항목 제외 + 경고 |
| 셀에 공백 섞임 (`신장분사 C20`) | 공백 제거 후 비교하여 동일 항목으로 인식 |
| `PT팀장` / `PT부팀장` | 숫자가 없으므로 PT 번호로 오인하지 않음 |
| 블록이 좌우로 나란함 | `합계건수` 헤더로 블록 열 범위를 나눠 각각 독립 처리 |
| 시트/데이터 없음 | 명확한 한글 오류 메시지 |

### 4.4 검증 결과

실제 양식(`주간 환자치료 타임 현황관리판`)과 동일하게 **좌우 2열 × 위아래 3줄 = 6개 블록**을
배치한 샘플 시트로 확인했습니다.

```
columns: 신장분사A20 | 신장분사A25 | 신장분사B20 | 신장분사B25 | 신장분사C20 | 신장분사C25

허정훈  PT288  → 0  0  0  0  0  16     (왼쪽 1줄)
강지은  PT287  → 0  0  3  0  6  3      (오른쪽 1줄)
권문옥  PT183  → 0  0  0  0  1  1      (왼쪽 2줄)
김미정  PT300  → 0  0  0  0  0  0      (오른쪽 2줄)
임우선  PT211  → 0  0  0  0  3  6      (왼쪽 3줄, 합계 수식이 깨진 블록)
이삼명  PT303  → 0  0  0  0  6  3      (오른쪽 3줄, 항목 순서가 달라도 정상)

warnings:
  PT211 '신장분사B20': 시트의 합계건수(3)와 날짜 합계(0)가 다릅니다. 날짜 합계를 사용했습니다.
  PT211 '신장분사B25': 시트의 합계건수(6)와 날짜 합계(0)가 다릅니다. 날짜 합계를 사용했습니다.
  PT211 '신장분사C20': 시트의 합계건수(0)와 날짜 합계(3)가 다릅니다. 날짜 합계를 사용했습니다.
  PT211 '신장분사C25': 시트의 합계건수(0)와 날짜 합계(6)가 다릅니다. 날짜 합계를 사용했습니다.
```

임우선 블록은 `B20`/`B25` 의 합계 수식이 두 행 아래 `C20`/`C25` 의 날짜를 더하고 있어
합계 셀을 그대로 쓰면 값이 밀립니다. 날짜 합산 방식이 이를 바로잡습니다.

날짜 개수·합계 열 위치가 블록마다 다르고, 새 항목(`신장분사D20` 등)이 추가되는 경우도
별도 픽스처로 함께 검증했습니다.

---

## 5. 코드 품질

* **TypeScript strict** + `noUncheckedIndexedAccess`, `noUnusedLocals`, `noUnusedParameters`, `noImplicitOverride`
* **ESLint** `typescript-eslint/strictTypeChecked` + react-hooks (`--max-warnings 0`)
* `any` 미사용 (엑셀 셀 값은 `CellValue` 유니온으로 표현)
* 매직 넘버/매직 스트링은 전부 `src/utils/constants.ts` 에 정의
* 모든 함수에 한글 주석으로 의도 명시

### 참고: `xlsx` 버전

npm 레지스트리의 `xlsx` 는 `0.18.5` 가 마지막입니다.
최신 버전이 필요하면 SheetJS 공식 CDN 을 사용하세요.

```bash
npm i https://cdn.sheetjs.com/xlsx-0.20.3/xlsx-0.20.3.tgz
```
