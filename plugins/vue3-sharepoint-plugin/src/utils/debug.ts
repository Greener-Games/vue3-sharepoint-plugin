/**
 * Simple Logger class to handle conditional debug output.
 */
export class Logger {
  private enabled = false
  private prefix = '[SharePoint Plugin]'

  constructor(enabled = false) {
    this.enabled = enabled
  }

  /** Logs a message to the console if enabled */
  log(message: string, ...data: unknown[]): void {
    if (this.enabled) {
      console.log(`${this.prefix} ${message}`, ...data)
    }
  }

  /** Logs a warning message to the console if enabled */
  warn(message: string, ...data: unknown[]): void {
    if (this.enabled) {
      console.warn(`${this.prefix} ${message}`, ...data)
    }
  }

  /** Logs an error message to the console */
  error(message: string, ...data: unknown[]): void {
    console.error(`${this.prefix} ${message}`, ...data)
  }
}
