import * as XLSX from 'xlsx';
import type { ParseResult } from '../../types/excel';
import {
  COLUMN_LABEL_PT,
  COLUMN_LABEL_THERAPIST,
  DEFAULT_ITEM_COUNT,
  RESULT_FILE_NAME,
  RESULT_SHEET_NAME,
} from '../constants';
import { saveWorkbookData } from './save';

/** 이름/PT번호 컬럼 폭 (문자 수 기준) */
const FIXED_COLUMN_WIDTH = 14;
/** 항목 컬럼 폭 */
const ITEM_COLUMN_WIDTH = 12;

/**
 * 결과를 화면과 동일한 컬럼 순서로 xlsx 바이너리로 만든다.
 * xlsx 는 내부적으로 UTF-8 XML 을 사용하므로 한글이 깨지지 않는다.
 */
export function buildResultWorkbook(result: ParseResult): Uint8Array<ArrayBuffer> {
  const header = [COLUMN_LABEL_THERAPIST, COLUMN_LABEL_PT, ...result.columns];

  const body = result.rows.map((row) => [
    row.therapist,
    row.pt,
    ...result.columns.map((column) => row.items[column] ?? DEFAULT_ITEM_COUNT),
  ]);

  const sheet = XLSX.utils.aoa_to_sheet([header, ...body]);
  sheet['!cols'] = header.map((_label, index) => ({
    wch: index < 2 ? FIXED_COLUMN_WIDTH : ITEM_COLUMN_WIDTH,
  }));

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, sheet, RESULT_SHEET_NAME);

  const output = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' }) as ArrayBuffer;
  return new Uint8Array(output);
}

/**
 * 결과 엑셀을 저장한다.
 *
 * @returns 저장된 경로. 사용자가 취소하면 null.
 */
export async function saveResultWorkbook(result: ParseResult): Promise<string | null> {
  return saveWorkbookData(buildResultWorkbook(result), RESULT_FILE_NAME);
}
