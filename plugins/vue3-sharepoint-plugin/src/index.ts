import type { App } from 'vue'
import type { ISharePointClient } from './types'

// Export Types
export * from './types'

// Export Implementations (So the user can instantiate them in main.ts)
export { PnPSharePointClient } from './implementations/PnPSharePointClient'
export { RestSharePointClient } from './implementations/RestSharePointClient'
export { AnonymousSharePointClient } from './implementations/AnonymousSharePointClient'
export { MockSharePointClient } from './implementations/MockSharePointClient'
export type { MockData } from './implementations/MockSharePointClient'

// Export Utils
export {
  createMockData,
  createMockUser,
  createMockGroup,
  createMockDocument,
  MOCK_DIVISIONS,
  MOCK_CAPABILITIES,
  MOCK_MARKETS,
  MOCK_PRACTICES,
  MOCK_REGIONS,
} from './utils/createMockData'


// Export Composables
export { useSharePoint } from './composables/useSharePoint'
export { useSearch } from './composables/useSearch'

// Injection Key
export const SP_CLIENT_SYMBOL = Symbol('SharePointClient')

// Plugin Options
export interface SharePointPluginOptions {
  client: ISharePointClient
}

// Vue Plugin Definition
export const SharePointPlugin = {
  install(app: App, options: SharePointPluginOptions) {
    if (!options.client) {
      throw new Error(
        'SharePointPlugin Error: You must provide a "client" instance in the options.'
      )
    }
    app.provide(SP_CLIENT_SYMBOL, options.client)
  },
}
