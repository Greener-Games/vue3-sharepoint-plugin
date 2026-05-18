# Vue 3 + TypeScript + Vite + Plugins

This template supports both standard application development and plugin development.

## Project Structure

- `src/`: Main application source code.
- `plugins/`: Directory for developing plugins.

## Setup Script

This project includes a setup script to help you configure the environment for either Project Development or Plugin Development.

Run the setup script:

```sh
node .templateScripts/setup.cjs
```

### Options

1. **Project Development**: Can remove example plugins and cleanup `main.ts`.
2. **Plugin Development**: Scaffolds a new plugin in the `plugins/` directory and optionally adds it to `src/main.ts`.

## Plugin Development

Plugins reside in the `plugins/` directory. Each plugin should be a folder containing at least:
- `index.ts`: The entry point exporting a Vue `install` function.
- `package.json`: Plugin metadata.

The template is configured to alias `@plugins` to the `plugins/` directory.

### Example Plugin

See `plugins/example-plugin` for a reference implementation.

## Recommended IDE Setup

[VS Code](https://code.visualstudio.com/) + [Volar](https://marketplace.visualstudio.com/items?itemName=Vue.volar) (and disable Vetur).

## Customize configuration

See [Vite Configuration Reference](https://vitejs.dev/config/).

## Project Setup

```sh
npm install
```

### Compile and Hot-Reload for Development

```sh
npm run dev
```

### Type-Check, Compile and Minify for Production

```sh
npm run build
```

### Lint with [ESLint](https://eslint.org/)

```sh
npm run lint
```
