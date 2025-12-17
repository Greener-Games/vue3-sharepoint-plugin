
/**
 * Simple Logger class to handle conditional debug output.
 */
export class Logger {
  private enabled: boolean = false
  private prefix: string = '[SharePoint Plugin]'

  constructor(enabled: boolean = false) {
    this.enabled = enabled
  }

  log(message: string, ...data: any[]) {
    if (this.enabled) {
      console.log(`${this.prefix} ${message}`, ...data)
    }
  }

  warn(message: string, ...data: any[]) {
    if (this.enabled) {
      console.warn(`${this.prefix} ${message}`, ...data)
    }
  }

  error(message: string, ...data: any[]) {
    // Errors usually should always be logged, but for consistency we respect the flag or use console.error directly in catch blocks.
    // However, if this is strictly for *debug* info, we might gate it.
    // Usually, real errors should always print. But let's assume this is for tracing.
    console.error(`${this.prefix} ${message}`, ...data)
  }
}
