import { contextBridge, ipcRenderer } from 'electron';
import { IPC_SAVE_EXCEL, type SaveExcelRequest, type SaveExcelResponse } from './ipc';

/**
 * 렌더러에 노출되는 최소한의 API.
 * fs / path 등 Node API 는 절대 노출하지 않고, 필요한 동작만 함수 단위로 전달한다.
 */
const electronAPI = {
  /** 결과 엑셀 저장을 메인 프로세스에 요청한다. */
  saveExcel: (request: SaveExcelRequest): Promise<SaveExcelResponse> =>
    ipcRenderer.invoke(IPC_SAVE_EXCEL, request) as Promise<SaveExcelResponse>,
} as const;

export type ElectronAPI = typeof electronAPI;

contextBridge.exposeInMainWorld('electronAPI', electronAPI);
