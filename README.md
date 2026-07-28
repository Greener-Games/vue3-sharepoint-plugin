# Vue 3 SharePoint Plugin Repository

[![Vite](https://img.shields.io/badge/vite-%23646CFF.svg?style=flat&logo=vite&logoColor=white)](https://vitejs.dev/)
[![Vue 3](https://img.shields.io/badge/Vue-3-4FC08D.svg?style=flat&logo=vue.js&logoColor=white)](https://vuejs.org/)
[![TypeScript](https://img.shields.io/badge/TypeScript-Ready-3178C6.svg?style=flat&logo=typescript&logoColor=white)](https://www.typescriptlang.org/)
[![npm version](https://img.shields.io/npm/v/@greener-games/vue3-sharepoint-plugin.svg)](https://www.npmjs.com/package/@greener-games/vue3-sharepoint-plugin)
[![npm downloads](https://img.shields.io/npm/dt/@greener-games/vue3-sharepoint-plugin.svg)](https://www.npmjs.com/package/@greener-games/vue3-sharepoint-plugin)
[![GitHub stars](https://img.shields.io/github/stars/Greener-Games/vue3-sharepoint-plugin.svg?style=social&label=Stars)](https://github.com/Greener-Games/vue3-sharepoint-plugin)
[![CI](https://github.com/Greener-Games/vue3-sharepoint-plugin/actions/workflows/ci.yml/badge.svg)](https://github.com/Greener-Games/vue3-sharepoint-plugin/actions/workflows/ci.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

This is the main repository for the **Vue 3 SharePoint Plugin**.

A robust, type-safe wrapper for SharePoint operations in Vue 3 applications. This plugin provides a unified interface for Search, CRUD, File operations, and Batching, supporting both **PnPjs** (Production) and an in-memory **Mock Engine** (Development/Testing).

## 📚 Documentation

For complete installation instructions, configuration details, and API reference, please see the plugin documentation:

👉 **[View Plugin Documentation](./plugins/@greener-games/vue3-sharepoint-plugin/README.md)**

---

## 🛠️ Project Structure

This repository uses a monorepo workspace configuration.

- `src/`: Main application source code. Used as a playground and demo environment to test plugin features.
- `plugins/@greener-games/vue3-sharepoint-plugin/`: The core plugin source code that is compiled and published to npm.

## 🚀 Project Setup (Development)

To run the local playground and develop the plugin alongside it:

```sh
# 1. Install dependencies
npm install

# 2. Start the development server with hot-reload
npm run dev
```

### Build and CI Commands

- **Type-Check & Build**: `npm run build`
- **Lint Scripts**: `npm run lint`
- **Lint Styles**: `npm run lint:styles`
- **Format Code**: `npm run format`

## Recommended IDE Setup

[VS Code](https://code.visualstudio.com/) + [Volar](https://marketplace.visualstudio.com/items?itemName=Vue.volar) (and disable Vetur).

## Customize configuration

See [Vite Configuration Reference](https://vitejs.dev/config/).
