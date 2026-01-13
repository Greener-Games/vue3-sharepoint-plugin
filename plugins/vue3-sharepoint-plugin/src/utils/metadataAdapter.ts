import type { ISharePointClient, FieldDefinition } from '../types'

export interface FormValue {
  FieldName: string
  FieldValue: string
}

/**
 * Transforms a simple key-value payload into a format suitable for ValidateUpdateListItem.
 * It automatically detects field types (Taxonomy, User, Choice, etc.) by fetching
 * the list definition and formatting the values accordingly.
 *
 * @param client The SharePoint client instance (Rest or PnP)
 * @param listTitle The title of the list/library containing the item
 * @param payload The metadata to update (e.g. { "MyTaxonomy": { Label: "A", TermGuid: "B" }, "AssignedTo": "user@email.com" })
 */
export async function adaptFileMetadata(
  client: ISharePointClient,
  listTitle: string,
  payload: Record<string, any>,
  webUrl?: string
): Promise<FormValue[]> {
  const fields = await client.getListFields(listTitle, webUrl)
  const formValues: FormValue[] = []

  for (const [key, value] of Object.entries(payload)) {
    // 1. Find Field Definition
    // We search by InternalName first, then Title
    const field = fields.find(
      (f) => f.InternalName === key || f.Title === key
    )

    if (!field) {
      // If field is unknown, we pass it as string, assuming InternalName was used correctly
      // but warn the developer.
      console.warn(`[SharePointPlugin] Field '${key}' not found in list '${listTitle}'. Sending value as-is.`)
      formValues.push({
        FieldName: key,
        FieldValue: typeof value === 'string' ? value : JSON.stringify(value)
      })
      continue
    }

    // 2. Adapt Value based on Type
    const adaptedValue = await adaptValue(client, field, value)

    formValues.push({
      FieldName: field.InternalName,
      FieldValue: adaptedValue
    })
  }

  return formValues
}

async function adaptValue(
  client: ISharePointClient,
  field: FieldDefinition,
  value: any
): Promise<string> {
  if (value === null || value === undefined) return ''

  const typeAsString = field.TypeAsString

  switch (typeAsString) {
    case 'TaxonomyFieldType':
    case 'TaxonomyFieldTypeMulti':
      return await formatTaxonomy(client, field, value)

    case 'User':
    case 'UserMulti':
      return await formatUser(client, value)

    case 'MultiChoice':
      // MultiChoice expects "Value1;Value2" for ValidateUpdateListItem
      if (Array.isArray(value)) {
        return value.join(';')
      }
      return String(value)

    case 'Choice':
      return String(value)

    case 'Boolean':
      // ValidateUpdateListItem expects "TRUE" or "FALSE" (case insensitive but standard is caps)
      return value ? 'TRUE' : 'FALSE'

    case 'DateTime':
      // ValidateUpdateListItem works well with ISO strings
      if (value instanceof Date) return value.toISOString()
      return String(value)

    case 'Number':
    case 'Currency':
      return String(value)

    case 'URL':
      // URL field value: "Url, Description"
      if (typeof value === 'object' && value.Url) {
          return `${value.Url}, ${value.Description || ''}`
      }
      return String(value)

    default:
      // Text, Note, etc.
      return String(value)
  }
}

async function formatTaxonomy(client: ISharePointClient, field: FieldDefinition, value: any): Promise<string> {
  // Expected: "Label|Guid;Label|Guid"
  if (Array.isArray(value)) {
    // Use Promise.all to handle multiple async resolutions
    const formatted = await Promise.all(value.map(v => formatSingleTaxonomy(client, field, v)))
    return formatted.join(';')
  }
  return await formatSingleTaxonomy(client, field, value)
}

async function formatSingleTaxonomy(client: ISharePointClient, field: FieldDefinition, val: any): Promise<string> {
  if (typeof val === 'object') {
    // Check for { Label, TermGuid }
    if (val.Label && val.TermGuid) {
      return `${val.Label}|${val.TermGuid}`
    }
    // Fallback
    return JSON.stringify(val)
  }

  // If string, assume it's a Label and we need to find the Guid
  if (typeof val === 'string') {
    // Check if it already looks like "Label|Guid"
    if (val.includes('|') && val.length > 36) { // rough check
       return val
    }

    // Attempt Resolution
    if (field.TermSetId) {
      const term = await client.searchTerm(field.TermSetId, val)
      if (term) {
        return `${term.Label}|${term.TermGuid}`
      }
      console.warn(`[SharePointPlugin] Could not resolve Term '${val}' in TermSet '${field.TermSetId}'.`)
    } else {
      console.warn(`[SharePointPlugin] Cannot resolve Term '${val}' because Field '${field.InternalName}' is missing TermSetId.`)
    }

    // Return as-is if resolution fails (ValidateUpdateListItem might handle new terms if configured)
    return val
  }

  return String(val)
}

async function formatUser(client: ISharePointClient, value: any): Promise<string> {
  // ValidateUpdateListItem requires a specific format for User fields to be robust.
  // Using a JSON stringified array of objects with "Key" (LoginName) is the standard way.
  // Format: '[{"Key":"i:0#.f|membership|user@domain.com"}]'

  if (Array.isArray(value)) {
    const users = await Promise.all(value.map(v => resolveUser(client, v)))
    const payload = users.map(u => ({ Key: u.LoginName }))
    return JSON.stringify(payload)
  }

  const user = await resolveUser(client, value)
  return JSON.stringify([{ Key: user.LoginName }])
}

async function resolveUser(client: ISharePointClient, val: any): Promise<{ LoginName: string }> {
  // If it's an object with LoginName, use it
  if (val && typeof val === 'object' && val.LoginName) {
    return val
  }

  // If string, assume Email or LoginName
  if (typeof val === 'string') {
    // If it looks like a claim "i:0#", return it
    if (val.indexOf('i:0#') > -1) {
      return { LoginName: val }
    }

    // Otherwise try to ensure user (resolve email to claim)
    try {
      const user = await client.ensureUser(val)
      return user
    } catch (e) {
      console.warn(`[SharePointPlugin] Could not resolve user '${val}'. Using as-is.`)
      return { LoginName: val }
    }
  }

  return { LoginName: '' }
}
