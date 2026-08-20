import type { ParseResult, TherapistRecord } from '../types/excel';

/**
 * 엑셀에서 복사한 인원 목록을 읽어 결과 표의 순서를 맞춘다.
 *
 * 엑셀에서 세로줄(한 열)을 복사해 붙여넣으면 이런 모양으로 들어온다.
 *   "허정훈\r\n강지은\r\n권문옥\r\n"
 *
 * 실제로는 아래가 섞여 들어오므로 모두 걸러 낸다.
 * - 줄바꿈이 \r\n / \n / \r 중 무엇이든 올 수 있다
 * - 마지막에 빈 줄이 하나 더 붙는다
 * - 가로줄(한 행)을 복사하면 탭으로 구분되어 온다
 * - 셀에 쉼표 등이 있으면 엑셀이 따옴표로 감싼다
 * - 셀 안팎의 공백('윤 송', ' 허정훈 ')
 */

/** 줄바꿈과 탭 — 엑셀 붙여넣기에서 값을 나누는 구분자 */
const SEPARATOR_PATTERN = /[\r\n\t]+/;

/** 값을 감싼 따옴표 */
const SURROUNDING_QUOTES = /^["']+|["']+$/g;

/** 비교용으로 공백을 모두 없앤다. ('윤 송' 과 '윤송' 을 같게 본다) */
function compact(text: string): string {
  return text.replace(/\s+/g, '');
}

/**
 * 붙여넣은 글자에서 인원 목록을 뽑는다.
 * 화면에 보여 줄 원본 표기를 그대로 유지하고, 중복은 처음 것만 남긴다.
 */
export function parseNameList(text: string): string[] {
  const seen = new Set<string>();
  const names: string[] = [];

  for (const chunk of text.split(SEPARATOR_PATTERN)) {
    const name = chunk.replace(SURROUNDING_QUOTES, '').trim();
    const key = compact(name);
    if (key === '' || seen.has(key)) {
      continue;
    }
    seen.add(key);
    names.push(name);
  }

  return names;
}

/**
 * 붙여넣은 순서대로 행을 정렬한다.
 *
 * 목록에 없는 사람은 빼지 않고 **뒤에 붙인다.** 시트에는 있는데 목록에서 빠진 사람이
 * 조용히 사라지면 실적이 누락된 줄 모르고 넘어가기 때문이다.
 * 이름이 없으면 PT 번호로도 맞춰 본다.
 */
export function orderRowsByNames(
  rows: readonly TherapistRecord[],
  names: readonly string[],
): TherapistRecord[] {
  if (names.length === 0) {
    return [...rows];
  }

  const rank = new Map<string, number>();
  names.forEach((name, index) => {
    rank.set(compact(name), index);
  });

  const rankOf = (row: TherapistRecord): number =>
    rank.get(compact(row.therapist)) ?? rank.get(compact(row.pt)) ?? Number.MAX_SAFE_INTEGER;

  return [...rows]
    .map((row, index) => ({ row, index, rank: rankOf(row) }))
    // 같은 순위(= 목록에 없는 사람)끼리는 시트에 나온 순서를 유지한다.
    .sort((a, b) => a.rank - b.rank || a.index - b.index)
    .map((entry) => entry.row);
}

/** 결과 전체(모든 주차)에 순서를 적용한다. */
export function applyNameOrder(result: ParseResult, names: readonly string[]): ParseResult {
  if (names.length === 0) {
    return result;
  }

  return {
    ...result,
    weeks: result.weeks.map((week) => ({ ...week, rows: orderRowsByNames(week.rows, names) })),
  };
}

/** 붙여넣은 이름 중 결과에 없는 것 (오타·퇴사자 확인용) */
export function findMissingNames(result: ParseResult, names: readonly string[]): string[] {
  const found = new Set<string>();
  for (const week of result.weeks) {
    for (const row of week.rows) {
      found.add(compact(row.therapist));
      found.add(compact(row.pt));
    }
  }

  return names.filter((name) => !found.has(compact(name)));
}
