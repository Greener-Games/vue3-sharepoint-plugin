import { FileVersion, UserInfo, SiteGroup } from '../types'
import { MockData } from '../implementations/MockSharePointClient'

// --- 1. Export Constants (Useful for UI Filters/Dropdowns in user apps) ---
export const MOCK_DIVISIONS = [
  'Sales',
  'Marketing',
  'Human Resources',
  'IT',
  'Operations',
  'Finance',
]
export const MOCK_PRACTICES = [
  'Digital Transformation',
  'Cloud Computing',
  'Data Analytics',
  'Cyber Security',
  'Product Development',
]
export const MOCK_MARKETS = [
  'North America',
  'Europe',
  'Asia Pacific',
  'Latin America',
  'Middle East',
]
export const MOCK_CAPABILITIES = [
  'Strategic Planning',
  'Process Optimization',
  'Project Management',
  'Compliance',
  'Research & Development',
]
export const MOCK_REGIONS = [
  'East Coast',
  'West Coast',
  'Midwest',
  'International',
  'Headquarters',
]

// --- 2. Internal Helpers for Randomization ---
const ROLES = [
  {
    title: 'Software Engineer',
    ext: 'docx',
    desc: 'Developing scalable web applications and microservices.',
  },
  {
    title: 'Product Manager',
    ext: 'pdf',
    desc: 'Defining product roadmap and strategy for Q3 release.',
  },
  {
    title: 'QA Analyst',
    ext: 'docx',
    desc: 'Executing regression testing and automated test scripts.',
  },
  {
    title: 'Data Scientist',
    ext: 'pdf',
    desc: 'Analyzing customer behavior trends using machine learning models.',
  },
  {
    title: 'System Architect',
    ext: 'pptx',
    desc: 'Designing high-availability cloud infrastructure on Azure.',
  },
  {
    title: 'HR Coordinator',
    ext: 'docx',
    desc: 'Managing employee onboarding and benefits administration.',
  },
  {
    title: 'Financial Analyst',
    ext: 'pdf',
    desc: 'Preparing quarterly financial reports and budget forecasts.',
  },
]

const FIRST_NAMES = [
  'John',
  'Sarah',
  'Tom',
  'Emily',
  'Michael',
  'David',
  'Jessica',
]
const LAST_NAMES = [
  'Smith',
  'Jones',
  'Wilson',
  'Brown',
  'Taylor',
  'Evans',
  'Davies',
]

export const getRandom = <T>(arr: T[]) =>
  arr[Math.floor(Math.random() * arr.length)]

// --- 3. Granular Helper Functions ---

/** Creates a single mock user */
export function createMockUser(
  id: number,
  name: string,
  email: string
): UserInfo {
  return {
    Id: id,
    Title: name,
    Email: email,
    LoginName: `i:0#.f|membership|${email}`,
  }
}

/** Creates a mock SharePoint group */
export function createMockGroup(id: number, title: string): SiteGroup {
  return {
    Id: id,
    Title: title,
    LoginName: title,
  }
}

/** Returns a fully formed document list item, plus its file blob and versions */
export function createMockDocument(
  id: number,
  baseUrl: string,
  libraryName = 'Shared Documents'
): {
  item: any
  file: Blob
  versions: FileVersion[]
  serverRelativeUrl: string
} {
  const roleObj = getRandom(ROLES)
  const division = getRandom(MOCK_DIVISIONS)
  const fileName = `${roleObj.title}_${id}.${roleObj.ext}`
  const siteUrl = baseUrl.replace(/\/$/, '')
  const serverRelativeUrl = `${siteUrl}/${libraryName}/${division}/${fileName}`

  // 1. Metadata Item
  const item = {
    Id: id,
    UniqueId: `uuid-${id}`,
    Title: `Document - ${division} - ${roleObj.title}`,
    FileRef: serverRelativeUrl,
    Path: serverRelativeUrl,
    AuthorId: 1,
    Created: new Date(Date.now() - id * 86400000).toISOString(),
    HitHighlightedSummary: `${roleObj.desc} Focused on ${getRandom(
      MOCK_CAPABILITIES
    )} within the ${getRandom(MOCK_MARKETS)} market.`,
    RefinableString01: getRandom(MOCK_REGIONS),
    RefinableString02: division,
    RefinableString03: 'English',
    RefinableString04: roleObj.ext === 'pdf' ? 'PDF' : 'Word',
    Practice: getRandom(MOCK_PRACTICES),
    Markets: getRandom(MOCK_MARKETS),
    Capabilities: getRandom(MOCK_CAPABILITIES),
    EmployeeDetails: `${getRandom(FIRST_NAMES)} ${getRandom(
      LAST_NAMES
    )}, Senior Lead`,
  }

  // 2. File Blob
  const file = new Blob([`Fake content for ${fileName}`], {
    type: 'text/plain',
  })

  // 3. Versions
  const versions: FileVersion[] = [
    {
      VersionLabel: '2.0',
      Created: new Date().toISOString(),
      CheckInComment: 'Major release',
      IsCurrentVersion: true,
      Size: 1024 + id,
      Url: serverRelativeUrl,
      CreatedBy: {
        Id: 1,
        Title: 'Dev User',
        Email: 'dev@local',
        LoginName: 'i:0#.f|dev',
      },
    },
    {
      VersionLabel: '1.0',
      Created: new Date(Date.now() - 100000000).toISOString(),
      CheckInComment: 'Initial Draft',
      IsCurrentVersion: false,
      Size: 512 + id,
      Url: `${serverRelativeUrl}?v=1.0`,
      CreatedBy: {
        Id: 2,
        Title: 'Admin',
        Email: 'admin@local',
        LoginName: 'i:0#.f|admin',
      },
    },
  ]

  return { item, file, versions, serverRelativeUrl }
}

// --- 4. Main Factory Function ---

/**
 * Generates a complete Mock SharePoint Environment.
 * @param docCount Number of documents to generate (default: 60)
 * @param baseUrl Base URL for the site (default: /sites/Intranet)
 * @returns A full MockData object ready for the Client
 */
export function createMockData(
  docCount = 60,
  baseUrl = '/sites/Intranet'
): MockData {
  const siteUrl = baseUrl.replace(/\/$/, '')

  // A. Setup Users & Groups using helpers
  const currentUser = createMockUser(1, 'Dev User', 'dev@local')
  const adminUser = createMockUser(2, 'Site Admin', 'admin@local')
  const siteUsers = [currentUser, adminUser]

  const membersGroup = createMockGroup(1, 'Site Members')
  const ownersGroup = createMockGroup(2, 'Site Owners')

  const userGroups: Record<string, SiteGroup[]> = {
    [currentUser.Email]: [membersGroup],
    [adminUser.Email]: [ownersGroup],
  }

  // B. Containers
  const documentsList: any[] = []
  const filesMap: Record<string, Blob | string> = {}
  const versionsMap: Record<string, FileVersion[]> = {}

  // C. Generate Documents using helper
  for (let i = 0; i < docCount; i++) {
    const { item, file, versions, serverRelativeUrl } = createMockDocument(
      1000 + i,
      siteUrl
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
    delay: 400,
    lists: {
      [`${siteUrl}/Shared Documents`]: documentsList,
      [`${siteUrl}/Lists/News`]: [
        {
          Id: 1,
          Title: 'Company Announcement',
          Created: new Date().toISOString(),
        },
        { Id: 2, Title: 'Holiday Schedule', Created: new Date().toISOString() },
      ],
    },
    files: filesMap,
    fileVersions: versionsMap,
  }
}
