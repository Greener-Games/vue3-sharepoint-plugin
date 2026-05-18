// eslint.config.js (ESM)
import js from '@eslint/js'
import tseslint from 'typescript-eslint'
import pluginVue from 'eslint-plugin-vue'
import vueParser from 'vue-eslint-parser'         //explicit Vue SFC parser
import globals from 'globals'
import prettier from 'eslint-config-prettier'

export default [
  // Ignore build outputs and config file
  { ignores: ['dist/**', 'node_modules/**', '.vite/**', 'coverage/**', 'eslint.config.js', '.templateScripts/**'] },

  // Base JS rules
  js.configs.recommended,

  // Vue 3 flat preset (must be spread)
  ...pluginVue.configs['flat/recommended'],

  // @typescript-eslint recommended rules that require type info
  ...tseslint.configs.recommendedTypeChecked,

  // Global rule overrides for all files
  {
    rules: {
      'vue/no-v-html': 'off',
      '@typescript-eslint/no-unused-vars': [
        'error',
        {
          argsIgnorePattern: '^_',
          varsIgnorePattern: '^_',
          caughtErrorsIgnorePattern: '^_',
        },
      ],
    },
  },

  // --- Vue SFCs: use vue-eslint-parser and hand <script lang="ts"> to TS ---
  {
    files: ['**/*.vue'],
    languageOptions: {
      parser: vueParser, // ensure SFCs use vue parser
      parserOptions: {
        parser: tseslint.parser, // TS for <script lang="ts">
        projectService: true, // typed-linting (flat config)
        tsconfigRootDir: import.meta.dirname,
        extraFileExtensions: ['.vue'],
        ecmaVersion: 2020,
        sourceType: 'module',
      },
      globals: { ...globals.browser, ...globals.node },
    },
  },

  // --- Pure TS files: use the TS parser directly ---
  {
    files: ['**/*.{ts,tsx}'],
    languageOptions: {
      parser: tseslint.parser,
      parserOptions: {
        projectService: true,
        tsconfigRootDir: import.meta.dirname,
        ecmaVersion: 2020,
        sourceType: 'module',
      },
      globals: { ...globals.browser, ...globals.node },
    },
  },

  // --- Plain JS files: disable type-checked rules ---
  { files: ['**/*.js', '**/*.cjs', '**/*.mjs'], ...tseslint.configs.disableTypeChecked },

  // Prettier last
  prettier,
];
