// vite.config.ts
import { defineConfig } from "file:///D:/GG/Web/vue3-sharepoint-plugin/plugins/vue3-sharepoint-plugin/node_modules/vite/dist/node/index.js";
import dts from "file:///D:/GG/Web/vue3-sharepoint-plugin/plugins/vue3-sharepoint-plugin/node_modules/vite-plugin-dts/dist/index.mjs";
import { resolve } from "path";
var __vite_injected_original_dirname = "D:\\GG\\Web\\vue3-sharepoint-plugin\\plugins\\vue3-sharepoint-plugin";
var vite_config_default = defineConfig({
  plugins: [
    dts({
      insertTypesEntry: true,
      rollupTypes: true
    })
  ],
  build: {
    lib: {
      // Entry point
      entry: resolve(__vite_injected_original_dirname, "src/index.ts"),
      name: "VueSharePointPlugin",
      fileName: (format) => `vue3-sharepoint-plugin.${format}.js`
    },
    rollupOptions: {
      // Externalize dependencies you don't want bundled into the library
      external: ["vue"],
      output: {
        globals: {
          vue: "Vue"
        }
      }
    }
  }
});
export {
  vite_config_default as default
};
//# sourceMappingURL=data:application/json;base64,ewogICJ2ZXJzaW9uIjogMywKICAic291cmNlcyI6IFsidml0ZS5jb25maWcudHMiXSwKICAic291cmNlc0NvbnRlbnQiOiBbImNvbnN0IF9fdml0ZV9pbmplY3RlZF9vcmlnaW5hbF9kaXJuYW1lID0gXCJEOlxcXFxHR1xcXFxXZWJcXFxcdnVlMy1zaGFyZXBvaW50LXBsdWdpblxcXFxwbHVnaW5zXFxcXHZ1ZTMtc2hhcmVwb2ludC1wbHVnaW5cIjtjb25zdCBfX3ZpdGVfaW5qZWN0ZWRfb3JpZ2luYWxfZmlsZW5hbWUgPSBcIkQ6XFxcXEdHXFxcXFdlYlxcXFx2dWUzLXNoYXJlcG9pbnQtcGx1Z2luXFxcXHBsdWdpbnNcXFxcdnVlMy1zaGFyZXBvaW50LXBsdWdpblxcXFx2aXRlLmNvbmZpZy50c1wiO2NvbnN0IF9fdml0ZV9pbmplY3RlZF9vcmlnaW5hbF9pbXBvcnRfbWV0YV91cmwgPSBcImZpbGU6Ly8vRDovR0cvV2ViL3Z1ZTMtc2hhcmVwb2ludC1wbHVnaW4vcGx1Z2lucy92dWUzLXNoYXJlcG9pbnQtcGx1Z2luL3ZpdGUuY29uZmlnLnRzXCI7aW1wb3J0IHsgZGVmaW5lQ29uZmlnIH0gZnJvbSAndml0ZSdcclxuaW1wb3J0IGR0cyBmcm9tICd2aXRlLXBsdWdpbi1kdHMnXHJcbmltcG9ydCB7IHJlc29sdmUgfSBmcm9tICdwYXRoJ1xyXG5cclxuZXhwb3J0IGRlZmF1bHQgZGVmaW5lQ29uZmlnKHtcclxuICBwbHVnaW5zOiBbXHJcbiAgICBkdHMoe1xyXG4gICAgICBpbnNlcnRUeXBlc0VudHJ5OiB0cnVlLFxyXG4gICAgICByb2xsdXBUeXBlczogdHJ1ZSxcclxuICAgIH0pLFxyXG4gIF0sXHJcbiAgYnVpbGQ6IHtcclxuICAgIGxpYjoge1xyXG4gICAgICAvLyBFbnRyeSBwb2ludFxyXG4gICAgICBlbnRyeTogcmVzb2x2ZShfX2Rpcm5hbWUsICdzcmMvaW5kZXgudHMnKSxcclxuICAgICAgbmFtZTogJ1Z1ZVNoYXJlUG9pbnRQbHVnaW4nLFxyXG4gICAgICBmaWxlTmFtZTogKGZvcm1hdCkgPT4gYHZ1ZTMtc2hhcmVwb2ludC1wbHVnaW4uJHtmb3JtYXR9LmpzYCxcclxuICAgIH0sXHJcbiAgICByb2xsdXBPcHRpb25zOiB7XHJcbiAgICAgIC8vIEV4dGVybmFsaXplIGRlcGVuZGVuY2llcyB5b3UgZG9uJ3Qgd2FudCBidW5kbGVkIGludG8gdGhlIGxpYnJhcnlcclxuICAgICAgZXh0ZXJuYWw6IFsndnVlJ10sXHJcbiAgICAgIG91dHB1dDoge1xyXG4gICAgICAgIGdsb2JhbHM6IHtcclxuICAgICAgICAgIHZ1ZTogJ1Z1ZScsXHJcbiAgICAgICAgfSxcclxuICAgICAgfSxcclxuICAgIH0sXHJcbiAgfSxcclxufSlcclxuIl0sCiAgIm1hcHBpbmdzIjogIjtBQUEyWCxTQUFTLG9CQUFvQjtBQUN4WixPQUFPLFNBQVM7QUFDaEIsU0FBUyxlQUFlO0FBRnhCLElBQU0sbUNBQW1DO0FBSXpDLElBQU8sc0JBQVEsYUFBYTtBQUFBLEVBQzFCLFNBQVM7QUFBQSxJQUNQLElBQUk7QUFBQSxNQUNGLGtCQUFrQjtBQUFBLE1BQ2xCLGFBQWE7QUFBQSxJQUNmLENBQUM7QUFBQSxFQUNIO0FBQUEsRUFDQSxPQUFPO0FBQUEsSUFDTCxLQUFLO0FBQUE7QUFBQSxNQUVILE9BQU8sUUFBUSxrQ0FBVyxjQUFjO0FBQUEsTUFDeEMsTUFBTTtBQUFBLE1BQ04sVUFBVSxDQUFDLFdBQVcsMEJBQTBCLE1BQU07QUFBQSxJQUN4RDtBQUFBLElBQ0EsZUFBZTtBQUFBO0FBQUEsTUFFYixVQUFVLENBQUMsS0FBSztBQUFBLE1BQ2hCLFFBQVE7QUFBQSxRQUNOLFNBQVM7QUFBQSxVQUNQLEtBQUs7QUFBQSxRQUNQO0FBQUEsTUFDRjtBQUFBLElBQ0Y7QUFBQSxFQUNGO0FBQ0YsQ0FBQzsiLAogICJuYW1lcyI6IFtdCn0K
