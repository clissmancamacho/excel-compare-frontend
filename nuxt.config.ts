// https://nuxt.com/docs/api/configuration/nuxt-config
export default defineNuxtConfig({
    modules: [
        '@nuxtjs/tailwindcss',
        '@pinia/nuxt',
    ],
    css: [
        'assets/css/main.css'
    ],
    app: {
        head: {
          charset: 'utf-16',
          title: 'Excel Compare',
          meta: [
            { name: 'description', content: 'Excel compare site' }
          ],
        }
      }
    
})
