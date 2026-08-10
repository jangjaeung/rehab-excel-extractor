import { app, BrowserWindow, dialog, ipcMain } from 'electron';
import { writeFile } from 'node:fs/promises';
import * as path from 'node:path';
import { IPC_SAVE_EXCEL, type SaveExcelRequest, type SaveExcelResponse } from './ipc';

/** 개발 서버 주소 (vite.config.ts 의 strictPort 와 동일하게 유지) */
const DEV_SERVER_URL = 'http://localhost:5173';

/** 기본 창 크기 */
const WINDOW_WIDTH = 1100;
const WINDOW_HEIGHT = 760;
const WINDOW_MIN_WIDTH = 760;
const WINDOW_MIN_HEIGHT = 520;

/** dev 스크립트에서 NODE_ENV=development 를 명시적으로 주입한다. */
const isDevelopment = process.env['NODE_ENV'] === 'development';

/**
 * 메인 윈도우를 생성한다.
 * contextIsolation 을 유지한 채 preload 를 통해서만 메인 프로세스 기능을 노출한다.
 */
function createWindow(): void {
  const window = new BrowserWindow({
    width: WINDOW_WIDTH,
    height: WINDOW_HEIGHT,
    minWidth: WINDOW_MIN_WIDTH,
    minHeight: WINDOW_MIN_HEIGHT,
    title: '신장분사 실적 추출기',
    autoHideMenuBar: true,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });

  if (isDevelopment) {
    void window.loadURL(DEV_SERVER_URL);
  } else {
    void window.loadFile(path.join(__dirname, '..', 'dist', 'index.html'));
  }
}

/**
 * 렌더러가 만든 xlsx 바이너리를 저장 다이얼로그를 통해 디스크에 기록한다.
 * 파일 생성은 메인 프로세스에서만 수행하여 렌더러에 fs 접근 권한을 주지 않는다.
 */
async function handleSaveExcel(request: SaveExcelRequest): Promise<SaveExcelResponse> {
  const { canceled, filePath } = await dialog.showSaveDialog({
    title: '결과 엑셀 저장',
    defaultPath: request.defaultFileName,
    filters: [{ name: 'Excel 통합 문서', extensions: ['xlsx'] }],
  });

  if (canceled || filePath === '') {
    return { saved: false };
  }

  try {
    await writeFile(filePath, Buffer.from(request.data));
    return { saved: true, filePath };
  } catch (error) {
    return { saved: false, error: error instanceof Error ? error.message : String(error) };
  }
}

void app.whenReady().then(() => {
  ipcMain.handle(IPC_SAVE_EXCEL, async (_event, request: SaveExcelRequest) => handleSaveExcel(request));

  createWindow();

  // macOS: 독 아이콘 클릭 시 창이 없으면 다시 만든다.
  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});
