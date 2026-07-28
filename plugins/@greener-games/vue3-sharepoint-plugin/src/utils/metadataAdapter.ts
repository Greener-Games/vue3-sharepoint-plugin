import type { ISharePointClient, FieldDefinition } from '../types';

export interface FormValue {
  FieldName: string;
  FieldValue: string;
}

/**
 * Transforms a simple key-value payload into a format suitable for ValidateUpdateListItem.
 * Handles specialized SharePoint fields like Taxonomy, Person, Choice, etc.
 */
export async function adaptFileMetadata(
  client: ISharePointClient,
  listTitle: string,
  payload: Record<string, unknown>,
): Promise<FormValue[]> {
  const fields = await client.getListFields(listTitle);
  const formValues: FormValue[] = [];

  for (const [key, value] of Object.entries(payload)) {
    const field = fields.find((f) => f.InternalName === key || f.Title === key);
    if (!field) {
      // If we don't know the field, pass through as string
      formValues.push({
        FieldName: key,
        // eslint-disable-next-line @typescript-eslint/no-base-to-string
        FieldValue: String(value ?? ''),
      });
      continue;
    }

    const formattedValue = await formatFieldValue(client, field, value);
    formValues.push({
      FieldName: field.InternalName,
      FieldValue: formattedValue,
    });
  }

  return formValues;
}

async function formatFieldValue(
  client: ISharePointClient,
  field: FieldDefinition,
  value: unknown,
): Promise<string> {
  if (value === null || value === undefined) return '';

  switch (field.TypeAsString) {
    case 'TaxonomyFieldType':
    case 'TaxonomyFieldTypeMulti':
      return await formatTaxonomy(client, field, value);

    case 'User':
    case 'UserMulti':
      return await formatUser(client, value);

    case 'MultiChoice':
      // MultiChoice expects "Value1;Value2" for ValidateUpdateListItem
      if (Array.isArray(value)) {
        return value.join(';');
      }
      return typeof value === 'object' && value !== null
        ? JSON.stringify(value)
        : String((value as string | number | boolean | null) ?? '');

    case 'Choice':
      return typeof value === 'object' && value !== null
        ? JSON.stringify(value)
        : String((value as string | number | boolean | null) ?? '');

    case 'Boolean':
      // ValidateUpdateListItem expects "TRUE" or "FALSE" (case insensitive but standard is caps)
      return value ? 'TRUE' : 'FALSE';

    case 'DateTime':
      // ValidateUpdateListItem works well with ISO strings
      if (value instanceof Date) return value.toISOString();
      return typeof value === 'object' && value !== null
        ? JSON.stringify(value)
        : String((value as string | number | boolean | null) ?? '');

    case 'Number':
    case 'Currency':
      return typeof value === 'object' && value !== null
        ? JSON.stringify(value)
        : String((value as string | number | boolean | null) ?? '');

    case 'URL':
      // URL field value: "Url, Description"
      if (typeof value === 'object' && value !== null) {
        const urlObj = value as { Url?: string; Description?: string };
        if (urlObj.Url) {
          return `${urlObj.Url}, ${urlObj.Description || ''}`;
        }
      }
      return typeof value === 'object' && value !== null
        ? JSON.stringify(value)
        : String((value as string | number | boolean | null) ?? '');

    default:
      // Text, Note, etc.
      return typeof value === 'object' && value !== null
        ? JSON.stringify(value)
        : String((value as string | number | boolean | null) ?? '');
  }
}

async function formatTaxonomy(
  client: ISharePointClient,
  field: FieldDefinition,
  value: unknown,
): Promise<string> {
  // Expected: "Label|Guid;Label|Guid"
  if (Array.isArray(value)) {
    // Use Promise.all to handle multiple async resolutions
    const formatted = await Promise.all(value.map((v) => formatSingleTaxonomy(client, field, v)));
    return formatted.join(';');
  }
  return await formatSingleTaxonomy(client, field, value);
}

async function formatSingleTaxonomy(
  client: ISharePointClient,
  field: FieldDefinition,
  val: unknown,
): Promise<string> {
  if (typeof val === 'object' && val !== null) {
    const termObj = val as { Label?: string; TermGuid?: string };
    // Check for { Label, TermGuid }
    if (termObj.Label && termObj.TermGuid) {
      return `${termObj.Label}|${termObj.TermGuid}`;
    }
    // Fallback
    return JSON.stringify(val);
  }

  // If string, assume it's a Label and we need to find the Guid
  if (typeof val === 'string') {
    // Check if it already looks like "Label|Guid"
    if (val.includes('|') && val.length > 36) {
      // rough check
      return val;
    }

    // Attempt Resolution
    if (field.TermSetId) {
      const term = await client.searchTerm(field.TermSetId, val);
      if (term) {
        return `${term.Label}|${term.TermGuid}`;
      }
      console.warn(
        `[SharePointPlugin] Could not resolve Term '${val}' in TermSet '${field.TermSetId}'.`,
      );
    } else {
      console.warn(
        `[SharePointPlugin] Cannot resolve Term '${val}' because Field '${field.InternalName}' is missing TermSetId.`,
      );
    }

    // Return as-is if resolution fails (ValidateUpdateListItem might handle new terms if configured)
    return val;
  }

  if (val === null || val === undefined) return '';
  // eslint-disable-next-line @typescript-eslint/no-base-to-string
  return typeof val === 'object' ? JSON.stringify(val) : String(val);
}

async function formatUser(client: ISharePointClient, value: unknown): Promise<string> {
  // ValidateUpdateListItem requires a specific format for User fields to be robust.
  // Using a JSON stringified array of objects with "Key" (LoginName) is the standard way.
  // Format: '[{"Key":"i:0#.f|membership|user@domain.com"}]'

  if (Array.isArray(value)) {
    const users = await Promise.all(value.map((v) => resolveUser(client, v)));
    const payload = users.map((u) => ({ Key: u.LoginName }));
    return JSON.stringify(payload);
  }

  const user = await resolveUser(client, value);
  return JSON.stringify([{ Key: user.LoginName }]);
}

async function resolveUser(
  client: ISharePointClient,
  val: unknown,
): Promise<{ LoginName: string }> {
  if (typeof val === 'object' && val !== null) {
    const userObj = val as { LoginName?: string; Email?: string };
    if (userObj.LoginName) return { LoginName: userObj.LoginName };
    if (userObj.Email) {
      try {
        const user = await client.ensureUser(userObj.Email);
        return { LoginName: user.LoginName };
      } catch {
        return { LoginName: userObj.Email };
      }
    }
  }

  if (typeof val === 'string') {
    // If it's an email or login name, try to resolve it
    try {
      const user = await client.ensureUser(val);
      return user;
    } catch (_e: unknown) {
      console.warn(`[SharePointPlugin] Could not resolve user '${val}'. Using as-is.`);
      return { LoginName: val };
    }
  }

  return { LoginName: '' };
}
