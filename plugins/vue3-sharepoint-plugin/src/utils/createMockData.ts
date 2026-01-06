import type { FileVersion, UserInfo, SiteGroup } from '../types'
import type { MockData } from '../implementations/MockSharePointClient'

// --- 1. Export Constants ---
export const MOCK_DIVISIONS = [
  'Sales', 'Marketing', 'Human Resources', 'IT', 'Operations', 'Finance',
]
export const MOCK_PRACTICES = [
  'Digital Transformation', 'Cloud Computing', 'Data Analytics', 'Cyber Security',
]
export const MOCK_MARKETS = [
  'North America', 'Europe', 'Asia Pacific',
]
export const MOCK_CAPABILITIES = [
  'Strategic Planning', 'Process Optimization', 'Compliance',
]
export const MOCK_REGIONS = [
  'East Coast', 'West Coast', 'International',
]
export const MOCK_STATUSES = [
  'Published', 'Published', 'Published', 'Pending', 'Draft'
  // Weighted to have more published docs
]

// --- 2. Internal Helpers ---
const ROLES = [
  { title: 'Software Engineer', ext: 'docx', desc: 'Technical design specs.' },
  { title: 'Product Manager', ext: 'pdf', desc: 'Product roadmap Q3.' },
  { title: 'QA Analyst', ext: 'docx', desc: 'Test plan execution.' },
  { title: 'System Architect', ext: 'pptx', desc: 'Cloud infrastructure.' },
  { title: 'HR Coordinator', ext: 'docx', desc: 'Onboarding checklist.' },
]

export const getRandom = <T>(arr: T[]) => arr[Math.floor(Math.random() * arr.length)]

// --- 3. Granular Helper Functions ---

/** Creates a single mock user */
export function createMockUser(id: number, name: string, email: string): UserInfo {
  return {
    Id: id,
    Title: name,
    Email: email,
    LoginName: `i:0#.f|membership|${email}`,
  }
}

/** Creates a mock SharePoint group */
export function createMockGroup(id: number, title: string): SiteGroup {
  return { Id: id, Title: title, LoginName: title }
}

/**
 * Extended definition for our internal logic to track reporting lines
 */
interface ExtendedUser extends UserInfo {
  ReportsToEmail: string | null // Email of the line manager
}

/** Returns a fully formed document list item */
export function createMockDocument(
    id: number,
    baseUrl: string,
    allUsers: ExtendedUser[], // We pass the pool of users to pick an owner
    libraryName = 'Shared Documents'
): {
  item: any
  file: Blob
  versions: FileVersion[]
  serverRelativeUrl: string
} {
  // 1. Pick a Random Owner from our defined user pool
  const ownerUser = getRandom(allUsers)!

  // 2. Determine Line Manager based on the user's structure
  const managerEmail = ownerUser.ReportsToEmail || 'ceo@local'

  const roleObj = getRandom(ROLES)!
  const division = getRandom(MOCK_DIVISIONS)!
  const status = getRandom(MOCK_STATUSES)!

  const fileName = `${roleObj.title}_${id}.${roleObj.ext}`
  const siteUrl = baseUrl.replace(/\/$/, '')
  const serverRelativeUrl = `${siteUrl}/${libraryName}/${division}/${fileName}`

  // 3. Metadata Item
  const item = {
    Id: id,
    UniqueId: `uuid-${id}`,
    Title: `Document - ${division} - ${roleObj.title}`,
    FileRef: serverRelativeUrl,
    Path: serverRelativeUrl,

    // SharePoint Standard Fields
    AuthorId: ownerUser.Id,
    Author: ownerUser.Email, // In search results, Author is often the email or name
    EditorId: ownerUser.Id,
    Created: new Date(Date.now() - Math.floor(Math.random() * 1000000000)).toISOString(),

    // Search Summary
    HitHighlightedSummary: `${roleObj.desc} Owner: ${ownerUser.Title}. Status: ${status}.`,

    // Custom Columns (Mapped Properties)
    // These match the keys used in your Dashboard Component
    DocumentType: roleObj.ext === 'pdf' ? 'PDF' : 'Word',
    DocLanguage: 'English',
    DocOwner: ownerUser.Title,       // Display Name
    DocOwnerEmail: ownerUser.Email,  // For filtering
    LineManager: managerEmail,       // For "My Team" filtering
    DocStatus: status,               // For "Action" button logic

    // Extra Metadata
    Practice: getRandom(MOCK_PRACTICES),
    Markets: getRandom(MOCK_MARKETS),
    RefinableString01: division,     // Simulate mapped RefinableString
  }

  // 4. File Blob
  const file = new Blob([`Fake content for ${fileName}`], { type: 'text/plain' })

  // 5. Versions
  const versions: FileVersion[] = [
    {
      VersionLabel: '1.0',
      Created: item.Created,
      CheckInComment: 'Initial Upload',
      IsCurrentVersion: true,
      Size: 1024 + id,
      Url: serverRelativeUrl,
      CreatedBy: {
        Id: ownerUser.Id,
        Title: ownerUser.Title,
        Email: ownerUser.Email,
        LoginName: ownerUser.LoginName
      },
    }
  ]

  return { item, file, versions, serverRelativeUrl }
}

// --- 4. Main Factory Function ---

export function createMockData(
    docCount = 60,
    baseUrl = '/sites/Intranet'
): MockData {
  const siteUrl = baseUrl.replace(/\/$/, '')

  // A. Setup Rich User Hierarchy
  // 1. The Current Logged in User
  const currentUser: ExtendedUser = {
    ...createMockUser(1, 'Dev User', 'dev@local'),
    ReportsToEmail: 'director@local'
  }

  // 2. People who report TO Dev User (For "My Team" Dashboard)
  const teamMember1: ExtendedUser = {
    ...createMockUser(3, 'Alice TeamMember', 'alice@local'),
    ReportsToEmail: 'dev@local'
  }
  const teamMember2: ExtendedUser = {
    ...createMockUser(4, 'Bob TeamMember', 'bob@local'),
    ReportsToEmail: 'dev@local'
  }

  // 3. Peers (Report to Director, not Dev User)
  const peerUser: ExtendedUser = {
    ...createMockUser(5, 'Carol Peer', 'carol@local'),
    ReportsToEmail: 'director@local'
  }

  // 4. The Boss
  const directorUser: ExtendedUser = {
    ...createMockUser(6, 'Director Dan', 'director@local'),
    ReportsToEmail: null
  }

  const allUsers = [currentUser, teamMember1, teamMember2, peerUser, directorUser]

  // Standard Arrays for MockClient
  const siteUsers = allUsers.map(({ ReportsToEmail, ...u }) => u) // strip extra prop for standard return
  const userGroups: Record<string, SiteGroup[]> = {
    [currentUser.Email]: [createMockGroup(1, 'Site Members')],
  }

  // B. Containers
  const documentsList: any[] = []
  const filesMap: Record<string, Blob | string> = {}
  const versionsMap: Record<string, FileVersion[]> = {}

  // C. Generate Documents
  for (let i = 0; i < docCount; i++) {
    const { item, file, versions, serverRelativeUrl } = createMockDocument(
        1000 + i,
        siteUrl,
        allUsers // Pass our specific user pool
    )

    documentsList.push(item)
    filesMap[serverRelativeUrl] = file
    versionsMap[serverRelativeUrl] = versions
  }

  // D. Return Config
  return {
    currentUser,
    siteUsers,
    userGroups,
    delay: 400, // Simulate network latency
    lists: {
      [`${siteUrl}/Shared Documents`]: documentsList,
      [`${siteUrl}/Lists/News`]: [
        { Id: 1, Title: 'Company Announcement', Created: new Date().toISOString() },
      ],
    },
    files: filesMap,
    fileVersions: versionsMap,
  }
}