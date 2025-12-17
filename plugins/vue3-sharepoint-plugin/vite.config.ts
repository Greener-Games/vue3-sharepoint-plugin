import { defineConfig } from 'vite'
import dts from 'vite-plugin-dts'
import { resolve } from 'path'

export default defineConfig({
  plugins: [
    dts({
      insertTypesEntry: true,
      rollupTypes: true,
    }),
  ],
  build: {
    lib: {
      // Entry point
      entry: resolve(__dirname, 'src/index.ts'),
      name: 'VueSharePointPlugin',
      fileName: (format) => `vue3-sharepoint-plugin.${format}.js`,
    },
    rollupOptions: {
      // Externalize dependencies you don't want bundled into the library
      external: ['vue'],
      output: {
        globals: {
          vue: 'Vue',
        },
      },
    },
  },
})
