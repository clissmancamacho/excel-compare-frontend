// https://nuxt.com/docs/api/configuration/nuxt-config
export default defineNuxtConfig({
    modules: [
        '@nuxtjs/tailwindcss',
        '@pinia/nuxt',
    ],
    css: [
        'assets/css/main.css'
    ],
    proxy: {
        '/api': {
          target: 'http://127.0.0.1:8000',
          changeOrigin: true,
          pathRewrite: { '^/api': '/' },
        },
      },
    
})
