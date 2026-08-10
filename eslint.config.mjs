import js from '@eslint/js';
import globals from 'globals';
import tseslint from 'typescript-eslint';
import reactHooks from 'eslint-plugin-react-hooks';
import reactRefresh from 'eslint-plugin-react-refresh';

export default tseslint.config(
  { ignores: ['dist', 'dist-electron', 'release', 'node_modules'] },

  // --- 렌더러 (React + TypeScript) ---
  {
    files: ['src/**/*.{ts,tsx}'],
    extends: [js.configs.recommended, ...tseslint.configs.strictTypeChecked],
    languageOptions: {
      ecmaVersion: 2022,
      globals: globals.browser,
      parserOptions: {
        project: ['./tsconfig.json'],
        tsconfigRootDir: import.meta.dirname,
      },
    },
    plugins: {
      'react-hooks': reactHooks,
      'react-refresh': reactRefresh,
    },
    rules: {
      ...reactHooks.configs.recommended.rules,
      'react-refresh/only-export-components': ['warn', { allowConstantExport: true }],
      // 매직 넘버 금지 (배열 인덱스와 0/1 같은 자명한 값은 허용)
      'no-magic-numbers': ['warn', { ignore: [0, 1, 2, -1], ignoreArrayIndexes: true, enforceConst: true }],
    },
  },

  // --- 메인 프로세스 (Node) ---
  {
    files: ['electron/**/*.ts'],
    extends: [js.configs.recommended, ...tseslint.configs.strictTypeChecked],
    languageOptions: {
      ecmaVersion: 2022,
      globals: globals.node,
      parserOptions: {
        project: ['./tsconfig.electron.json'],
        tsconfigRootDir: import.meta.dirname,
      },
    },
  },
);
