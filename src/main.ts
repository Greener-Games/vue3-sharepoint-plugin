import { createApp } from 'vue'
import './style.css'
import App from './App.vue'
import { createMockData, MockSharePointClient, RestSharePointClient, SharePointPlugin } from "@greener-games/vue3-sharepoint-plugin";

export const app = createApp(App)


const isDev = import.meta.env.DEV
let client
if (isDev) {
    const mockData = createMockData(100, '/sites/TestSite')
    console.log('🚧 Running with Mock Data', mockData)
    client = new MockSharePointClient(mockData)
} else {
/*    client = new PnPSharePointClient({
        baseUrl: 'https://atkins.sharepoint.com/sites/TestSite',
        enableCache: true,
    })*/
        client = new RestSharePointClient({
          // Make sure this matches your site URL exactly
          baseUrl: 'https://atkins.sharepoint.com/sites/TestSite',
          enableCache: true
        })
}

app.use(SharePointPlugin, { client })

app.mount('#app')
