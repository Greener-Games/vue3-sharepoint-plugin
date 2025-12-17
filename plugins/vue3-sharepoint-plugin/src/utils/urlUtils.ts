
/**
 * Utility function to normalize paths to be server-relative.
 *
 * Logic:
 * 1. If path starts with 'http' or 'https', it attempts to extract the pathname.
 * 2. If path starts with '/', it assumes it is already server-relative (e.g. '/sites/MySite').
 * 3. If path does not start with '/', it treats it as site-relative and prepends the site's base URL path.
 *
 * @param path - The input path (absolute, server-relative, or site-relative).
 * @param baseUrl - The base URL of the SharePoint site (e.g. 'https://tenant.sharepoint.com/sites/MySite').
 * @returns The normalized server-relative path (e.g. '/sites/MySite/Shared Documents').
 */
export function getServerRelativePath(path: string, baseUrl: string): string {
    if (path.startsWith('http')) {
      try {
        return decodeURIComponent(new URL(path).pathname)
      } catch {
        return path
      }
    }
    if (path.startsWith('/')) {
      // Already server relative
      return path
    }

    // Site relative logic
    let sitePath = ''
    try {
      sitePath = new URL(baseUrl).pathname
    } catch {
      // Fallback if baseUrl is malformed (or relative itself in some edge mock cases)
      if (baseUrl.startsWith('/')) {
          sitePath = baseUrl
      }
    }

    // Ensure no double slash
    const cleanSite = sitePath.replace(/\/$/, '')
    const cleanPath = path.replace(/^\//, '')

    return `${cleanSite}/${cleanPath}`
}
