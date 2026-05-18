export default {
    extends: [
        // Base CSS catching “real” problems
        'stylelint-config-standard',
        // Tailwind-aware adjustments (prevents false positives on @tailwind, @apply, etc.)
        'stylelint-config-tailwindcss',
        // Vue SFC parsing + Vue-specific tuning (MUST be last so its customSyntax is applied)
        'stylelint-config-recommended-vue',
    ],
    plugins: ['stylelint-order'],
    overrides: [
        {
            files: ['**/*.vue', '**/*.html'],
            customSyntax: 'postcss-html', // parse <style> blocks in .vue
        },
    ],
    rules: {
        // add any team preferences here
      "block-no-empty": false,
      "no-empty-source": false
    },
    ignoreFiles: ['dist/**/*', 'node_modules/**/*'],
}