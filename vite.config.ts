// vite.config.ts
import { fileURLToPath, URL } from 'node:url'
import { defineConfig } from 'vite'
import vue from '@vitejs/plugin-vue'
import vueJsx from '@vitejs/plugin-vue-jsx'
// Consider removing vueDevTools for production builds
// import vueDevTools from 'vite-plugin-vue-devtools' 

// https://vite.dev/config/
export default defineConfig({
  base: '/EBCG/', // Set to your repository name
  plugins: [
    vue(),
    vueJsx(),
    // vueDevTools(), // Optional: Remove if not needed in production
  ],
  resolve: {
    alias: {
      '@': fileURLToPath(new URL('./src', import.meta.url))
    },
  },
})
