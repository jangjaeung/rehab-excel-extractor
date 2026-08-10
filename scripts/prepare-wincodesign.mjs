/**
 * electron-builder 의 코드 서명 도구(winCodeSign) 캐시를 미리 준비한다.
 *
 * 왜 필요한가
 *   electron-builder 는 Windows 설치 파일을 만들기 직전에 winCodeSign 압축을 푸는데,
 *   그 안에 macOS 전용 심볼릭 링크(darwin/.../libcrypto.dylib)가 들어 있다.
 *   Windows 에서 심볼릭 링크를 만들려면 관리자 권한이나 개발자 모드가 필요해서
 *   보통은 아래 오류로 패키징이 멈춘다.
 *
 *     ⨯ cannot execute  cause=exit status 2
 *     ERROR: Cannot create symbolic link : ...\darwin\10.12\lib\libcrypto.dylib
 *
 *   이때 release\win-unpacked 까지만 만들어지고 설치 파일은 나오지 않는다.
 *
 * 무엇을 하는가
 *   Windows 빌드에 darwin 폴더는 필요 없으므로, 그 폴더만 빼고 캐시에 미리 풀어 둔다.
 *   그러면 electron-builder 가 '이미 있다' 고 보고 압축 해제 단계를 건너뛴다.
 *   캐시가 이미 준비돼 있으면 아무 일도 하지 않으므로 매번 실행해도 괜찮다.
 */

import { execFileSync } from 'node:child_process';
import { existsSync, mkdirSync, readdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';

/** electron-builder 25.x 가 쓰는 버전 */
const VERSION = '2.6.0';
const ARCHIVE_URL = `https://github.com/electron-userland/electron-builder-binaries/releases/download/winCodeSign-${VERSION}/winCodeSign-${VERSION}.7z`;

/** 압축을 풀 때 건너뛸 폴더 (macOS 전용) */
const EXCLUDED_DIR = 'darwin';

await main();

async function main() {
  // 서명 도구는 Windows 빌드에서만 쓴다.
  if (process.platform !== 'win32') {
    return;
  }

  const localAppData = process.env.LOCALAPPDATA;
  if (localAppData === undefined) {
    warn('LOCALAPPDATA 를 찾지 못해 건너뜁니다.');
    return;
  }

  const cacheDir = join(localAppData, 'electron-builder', 'Cache', 'winCodeSign');
  const targetDir = join(cacheDir, `winCodeSign-${VERSION}`);

  if (existsSync(targetDir)) {
    console.log(`[winCodeSign] 캐시가 이미 준비되어 있습니다: ${targetDir}`);
    return;
  }

  const sevenZip = findSevenZip();
  if (sevenZip === null) {
    warn('node_modules 에서 7za.exe 를 찾지 못했습니다. pnpm install 후 다시 시도하세요.');
    return;
  }

  try {
    mkdirSync(cacheDir, { recursive: true });

    const archivePath = join(cacheDir, `winCodeSign-${VERSION}.7z`);
    if (!existsSync(archivePath)) {
      console.log('[winCodeSign] 서명 도구를 내려받는 중...');
      await download(ARCHIVE_URL, archivePath);
    }

    console.log(`[winCodeSign] ${EXCLUDED_DIR} 폴더를 빼고 캐시에 푸는 중...`);
    execFileSync(sevenZip, ['x', '-bd', '-y', archivePath, `-o${targetDir}`, `-xr!${EXCLUDED_DIR}`], {
      stdio: 'ignore',
    });

    console.log('[winCodeSign] 준비 완료.');
  } catch (error) {
    // 여기서 실패해도 electron-builder 가 자기 방식으로 다시 시도하므로 빌드를 멈추지는 않는다.
    warn(`준비하지 못했습니다: ${error instanceof Error ? error.message : String(error)}`);
  }
}

/** 파일 하나 내려받기 */
async function download(url, destination) {
  const response = await fetch(url);
  if (!response.ok) {
    throw new Error(`${url} 응답이 ${String(response.status)} 입니다.`);
  }
  writeFileSync(destination, Buffer.from(await response.arrayBuffer()));
}

/**
 * node_modules 에 들어 있는 7za.exe 를 찾는다.
 * npm(평탄한 구조)과 pnpm(.pnpm 구조) 양쪽을 모두 지원한다.
 */
function findSevenZip() {
  const relative = join('7zip-bin', 'win', 'x64', '7za.exe');

  const flat = join('node_modules', relative);
  if (existsSync(flat)) {
    return flat;
  }

  const pnpmDir = join('node_modules', '.pnpm');
  if (!existsSync(pnpmDir)) {
    return null;
  }

  for (const entry of readdirSync(pnpmDir)) {
    if (!entry.startsWith('7zip-bin@')) {
      continue;
    }
    const candidate = join(pnpmDir, entry, 'node_modules', relative);
    if (existsSync(candidate)) {
      return candidate;
    }
  }

  return null;
}

function warn(message) {
  console.warn(`[winCodeSign] ${message}`);
}
