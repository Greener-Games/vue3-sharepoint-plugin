# Project Context: Vue 3 + Vite + Tailwind CSS 4

## Tech Stack
- **Framework**: Vue 3 (Composition API with `<script setup>`)
- **Build Tool**: Vite
- **Styling**: Tailwind CSS 4 (using `@tailwindcss/vite` plugin)
- **State Management**: Pinia
- **Routing**: Vue Router
- **Language**: TypeScript (Strict mode)
- **Linting/Formatting**: ESLint 9 (Flat Config), Prettier, Stylelint
- **Type Checking**: vue-tsc

## Project Structure
- `src/components/`: Reusable UI components.
- `src/components/base/`: Atomic/Base components (e.g., `BaseIcon.vue`, `BaseButton.vue`).
- `src/views/`: Page-level components used by Vue Router.
- `src/composables/`: Shared stateful logic (Composition Functions).
- `src/stores/`: Pinia store definitions.
- `src/router/`: Route definitions.
- `src/types/`: Global TypeScript interfaces and types.
- `src/utils/`: Pure helper functions (formatters, validators, loggers).
- `src/assets/`: Static assets and global CSS (including `main.css`).
- `@/`: Alias for the `src` directory.

## Coding Standards

### Vue Components
- **Block Order**: Always use the following structure:
    1. `<template>`
    2. `<script setup lang="ts">`
    3. `<style>` (only if necessary)
- **Reactivity**: Prefer `ref()` over `reactive()` for all state primitives and objects to maintain clear `.value` usage.
- **Props/Emits**: Use compiler-macros: `defineProps<{ ... }>()` and `defineEmits<{ ... }>()`.
- **Accessibility (a11y)**:
    - Use **Semantic HTML** tags (e.g., `<nav>`, `<main>`, `<article>`, `<button>`).
    - Ensure all interactive elements have visible `focus-visible` states.
    - Provide `aria-label` or `aria-labelledby` for icon-only buttons.
- **Iconography**:
    - **Default**: Use [Lucide Icons](https://lucide.dev/) and [Lucide Animated](https://lucide-animated.com/).
    - **SVG Strategy**: Use neutral/white SVGs as base. Colorize them using Tailwind CSS text-color classes (e.g., `text-primary`).
    - **Custom SVGs**: Ensure custom SVGs use `stroke="currentColor"` or `fill="currentColor"` to allow CSS color injection.
- **Utility First**: Leverage Tailwind utility classes. Avoid `<style scoped>` unless strictly necessary for third-party overrides.

### TypeScript
- Use strict typing. Avoid `any`.
- Prefer `interface` for object shapes and `type` for unions/aliases.

### Clean Code & Documentation
- **Self-Documenting Code**: Prioritize highly descriptive variable and function names (e.g., `isUserAuthenticated` instead of `auth`). The intent should be clear from the name alone.
- **Mandatory Function Comments**: Every function, action, and composable must include a brief JSDoc comment.
    - **Standard**: Comments should be an *addition* to self-documenting code, explaining the "why" or providing a quick summary, not just restating the function name.
- **Concise Logic**: Avoid "death by comments." Keep inline comments extremely brief and only use them when logic is non-obvious.
- **Unit Testing**: Tests (Vitest/Cypress) are only to be generated upon explicit request.

### State Management (Pinia)
- **Setup Store Syntax**: Use the function syntax (`defineStore('id', () => { ... })`).
- **Modular Architecture**: **Avoid "Mega Stores".** Break logic into small, domain-specific stores (e.g., `useAuthStore`, `useUserStore`, `useUIStore`).
- **Cross-Store Communication**: Import and instantiate required stores within the actions/getters of other stores as needed.
- **Destructuring**: Always use `storeToRefs()` when destructuring state or getters to maintain reactivity.

### Styling (Tailwind CSS 4)
- **CSS-First Configuration**: All theme configuration is handled in `src/assets/main.css` via the `@theme` block.
- **Color Variable Hierarchy**:
    1. **Logical/Brand Colors**: Define specific brand identity colors (e.g., `--color-brand-green: #2ecc71;`).
    2. **Semantic/Generic Colors**: Define functional variables (e.g., `--color-primary`, `--color-surface`, `--color-text-main`) that reference the brand variables via `var()`.
- **Semantic-First Approach**: Always use generic/semantic variables for the majority of the application. Brand-specific variables should only be used for fixed identity elements.
- **Responsive Design**: Follow a **Mobile-First** approach using Tailwind responsive prefixes (`sm:`, `md:`, `lg:`).
- **Micro-animations**:
    - Incorporate subtle transitions (hover scales, gentle fades) to make the UI feel fresh and modern.
    - **A11y**: Use `motion-safe:` variants and respect `prefers-reduced-motion`.
- **Organization via Layers**: Wrap global CSS in `@layer components` or `@layer utilities`.

### Error Handling & Logging
- **Resilient Processes**: Wrap asynchronous requests and high-risk logic in `try-catch` blocks.
- **Logging**: Use `console.error` or a dedicated logging utility for developer debugging.
- **User Feedback**: Never allow a silent failure. Always provide UI feedback (Toasts, error banners, or skeleton states) to inform the user of progress or issues.

## Task Management & Documentation
- **Checklist Format**: `[ ]` for pending, `[X]` for completed.
- **Completion Summaries**: Provide a concise summary of the specific actions performed after marking a task as complete.

## Common Commands
- Development: `npm run dev`
- Build: `npm run build`
- Type Check: `npm run type-check`
- Linting: `npm run lint`

## Specific Instructions for Gemini
- **Template Order**: Follow **template -> script -> style**.
- **Naming & JSDoc**: Use intent-revealing names and include a one-line JSDoc header for all functions (e.g., `/** Fetches user profile and updates local state */`).
- **No Automatic Tests**: Do not generate test files unless specifically asked.
- **Robust Error Handling**: Always include `try-catch` blocks with console logging and a UI feedback strategy (e.g., loading states) for async logic.
- **Mobile-First**: Generate responsive layouts by default.
- **Incorporate Motion**: Suggest or implement subtle micro-animations (e.g., `hover:scale-105 transition-all`) for interactive elements.
- **Domain-Specific Stores**: If a request involves new state, suggest a **new specific store** rather than bloating an existing one.
- **Logic Extraction**: If logic in a component exceeds 50 lines, suggest moving it to a `src/composables/` or `src/utils/` file.
- **Route Lazy Loading**: When adding routes to `src/router/index.ts`, always use dynamic imports: `() => import('@/views/MyView.vue')`.
- **Semantic Colors**: Prioritize generic semantic utility classes (e.g., `text-primary`, `bg-background`) over brand-specific ones.
