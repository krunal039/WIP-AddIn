# Code Structure Documentation

## Directory Structure

```
WIP-AddIn/
├── src/                          # Source code
│   ├── index.tsx                 # Application entry point
│   ├── App.tsx                   # Main React component (auth wrapper)
│   ├── auth/                     # Authentication modules
│   │   ├── msalConfig.ts         # MSAL configuration
│   │   ├── msalInstance.ts       # MSAL instance singleton
│   │   └── getToken.ts           # Token helper functions
│   ├── components/               # React components
│   │   ├── WorkbenchLanding.tsx  # Main UI component
│   │   ├── WorkbenchDialogs.tsx  # Dialog components
│   │   ├── WorkbenchHeader.tsx   # Header component
│   │   ├── BUProductsSection.tsx # Business unit/products section
│   │   ├── LandingSection.tsx    # Landing page section
│   │   ├── SpinnerOverlay.tsx    # Loading overlay
│   │   ├── ErrorBoundary.tsx     # Error boundary component
│   │   └── SharedGrid.css        # Shared styles
│   ├── config/                   # Configuration
│   │   ├── environment.ts        # Environment variable access
│   │   └── runtimeConfig.ts      # Runtime configuration loader
│   ├── constants/                # Constants
│   │   └── index.ts              # Application constants
│   ├── hooks/                    # Custom React hooks
│   │   └── useLocalStorage.ts    # Local storage hook
│   ├── service/                  # Business logic services
│   │   ├── ServiceBase.ts        # Base service class
│   │   ├── AuthService.ts        # Authentication service
│   │   ├── WorkbenchService.ts   # Main orchestration service
│   │   ├── PlacementApiService.ts # Placement API client
│   │   ├── GraphEmailService.ts  # Microsoft Graph API client
│   │   ├── EmailConverterService.ts # Email format conversion
│   │   ├── FileValidationService.ts # File validation logic
│   │   ├── OfficeEmailService.ts # Office.js email operations
│   │   ├── OfficeModeService.ts  # Office mode detection
│   │   ├── OfficeIdConverterService.ts # Office ID conversion
│   │   ├── LoggingService.ts     # Logging service
│   │   ├── DebugService.ts       # Debug logging service
│   │   └── ApiClient.ts          # Base API client
│   ├── types/                    # TypeScript type definitions
│   │   ├── index.ts              # Shared types
│   │   └── office.d.ts           # Office.js type extensions
│   └── utils/                    # Utility functions
│       ├── emailHelpers.ts       # Email helper functions
│       ├── emailStamping.ts       # Email stamping utilities
│       ├── duplicateDetection.ts # Duplicate detection logic
│       ├── fileInspector.ts      # File inspection utilities
│       ├── outlookNotification.ts # Outlook notification utilities
│       └── errorHandler.ts       # Error handling utilities
├── docs/                         # Documentation
├── Manifests/                    # Office Add-in manifests
│   ├── Manifest.Local.xml
│   ├── Manifest.Dev.xml
│   ├── Manifest.QA.xml
│   ├── Manifest.UAT.xml
│   └── Manifest.PRD.xml
├── app-settings/                 # Application settings (per environment)
├── secret-settings/              # Secret settings (per environment)
├── public/                       # Public assets
│   ├── index.html                # HTML template
│   └── assets/                   # Icons and images
├── dist/                         # Build output
├── tests/                        # Test files
├── scripts/                      # Build scripts
│   └── log-env.js                # Environment logging script
├── package.json                  # Dependencies and scripts
├── tsconfig.json                 # TypeScript configuration
├── webpack.config.js             # Webpack configuration
└── environments.json             # Runtime environment configuration
```

## File Organization Principles

### 1. Separation of Concerns
- **Components** (`src/components/`) - UI presentation only
- **Services** (`src/service/`) - Business logic and API calls
- **Utils** (`src/utils/`) - Reusable helper functions
- **Types** (`src/types/`) - TypeScript type definitions
- **Config** (`src/config/`) - Configuration management

### 2. Service Layer Pattern
All services follow a consistent pattern:
- Singleton instance pattern
- Static `getInstance()` method
- Private constructor
- Service-specific methods

### 3. Component Structure
Components are organized by feature:
- Main components in root of `components/`
- Shared components (if any) in subdirectories
- Each component is self-contained with its own types

## Key Files Explained

### Entry Point: `src/index.tsx`
- Initializes Office.js context
- Loads runtime configuration
- Creates MSAL instance
- Renders React app with providers

**Key Responsibilities:**
- Office.js initialization
- Runtime config loading
- MSAL setup
- Error boundary setup

### Main App: `src/App.tsx`
- Handles authentication state
- Manages token refresh
- Renders authentication UI or main UI
- Token lifecycle management

**Key State:**
- `isAuthenticated` - Authentication status
- `apiToken` - Placement API token
- `graphToken` - Microsoft Graph token
- `authError` - Authentication errors

### Main UI: `src/components/WorkbenchLanding.tsx`
- Primary user interface
- Form handling (BU, Product selection)
- File validation UI
- Submission orchestration
- Success/error dialogs

**Key Features:**
- 15+ useState hooks (state management)
- File validation integration
- Duplicate detection
- Email metadata extraction
- Submission flow coordination

## Service Architecture

### Service Base Class
All services extend or follow the pattern from `ServiceBase.ts`:
```typescript
class ServiceName {
  private static instance: ServiceName;
  private constructor() {}
  public static getInstance(): ServiceName { ... }
}
```

### Service Responsibilities

#### AuthService (`src/service/AuthService.ts`)
- MSAL authentication
- Token acquisition (API + Graph)
- Token caching and refresh
- Office SSO fallback
- **Lines of code:** ~664

#### WorkbenchService (`src/service/WorkbenchService.ts`)
- Main orchestration service
- Submission flow coordination
- Email processing pipeline
- Error handling
- **Lines of code:** ~425

#### PlacementApiService (`src/service/PlacementApiService.ts`)
- Placement API communication
- Request/response handling
- Error handling
- **Lines of code:** ~186

#### GraphEmailService (`src/service/GraphEmailService.ts`)
- Microsoft Graph API calls
- Email forwarding
- Email search
- **Lines of code:** ~748

#### FileValidationService (`src/service/FileValidationService.ts`)
- File type validation
- Encryption detection
- Password protection detection
- ZIP file detection
- **Lines of code:** ~594

#### EmailConverterService (`src/service/EmailConverterService.ts`)
- Office.js email to EML conversion
- Attachment handling
- Email metadata extraction

#### DebugService (`src/service/DebugService.ts`)
- Centralized logging
- Debug level control
- Environment-aware logging
- **Lines of code:** ~243

#### LoggingService (`src/service/LoggingService.ts`)
- Error logging to external API
- Placement request logging
- Audit trail

## Component Architecture

### Component Hierarchy
```
App
  └── WorkbenchLanding
      ├── WorkbenchHeader
      ├── LandingSection
      ├── BUProductsSection
      ├── WorkbenchDialogs
      │   ├── SuccessMessage
      │   ├── ErrorMessage
      │   ├── ConfirmationDialog
      │   └── RetryButton
      └── SpinnerOverlay
```

### Component Responsibilities

#### WorkbenchLanding
- Main container component
- State management (15+ useState hooks)
- Form handling
- Submission orchestration
- **Lines of code:** ~581

#### WorkbenchDialogs
- Success/error message dialogs
- Confirmation dialogs
- Retry functionality
- **Lines of code:** ~265

#### BUProductsSection
- Business unit selection
- Product selection
- Form validation

#### LandingSection
- Initial landing page
- Welcome message
- Navigation to main form

## Configuration Files

### `environments.json`
Runtime configuration loaded based on URL:
```json
{
  "environments": {
    "dev": { ... },
    "qa": { ... },
    "uat": { ... },
    "prod": { ... }
  },
  "urlPatterns": {
    "dev": ["dev.example.com"],
    "qa": ["qa.example.com"],
    ...
  }
}
```

### `tsconfig.json`
TypeScript compiler configuration:
- Target: ES2020
- Module: ES2020
- JSX: react
- Strict mode enabled

### `webpack.config.js`
Build configuration:
- Entry: `src/index.tsx`
- Output: `dist/`
- Loaders: ts-loader, css-loader, style-loader
- Plugins: HtmlWebpackPlugin, CopyWebpackPlugin

## Constants and Types

### Constants (`src/constants/index.ts`)
- `BU_OPTIONS` - Business unit dropdown options
- `PRODUCT_OPTIONS` - Product dropdown options
- `DEFAULTS` - Default values
- `STORAGE_KEYS` - LocalStorage keys

### Types (`src/types/index.ts`)
- `EmailMetadata` - Email metadata interface
- Shared type definitions

## Utility Functions

### `src/utils/emailHelpers.ts`
- `getSender()` - Extract email sender
- `getSubject()` - Extract email subject
- `getCreatedDate()` - Extract creation date
- `detectSharedMailbox()` - Detect shared mailbox

### `src/utils/emailStamping.ts`
- `stampEmailWithWorkbenchId()` - Add Workbench ID to email

### `src/utils/duplicateDetection.ts`
- `checkDuplicateSubmission()` - Check for duplicate submissions

### `src/utils/fileInspector.ts`
- `detectFileProtectionFromBase64()` - Detect file protection
- `isSupportedFileType()` - Check file type support

## Build and Deployment

### Build Process
1. **Pre-build:** `log-env.js` logs environment variables
2. **Build:** Webpack compiles TypeScript and bundles assets
3. **Post-build:** Copy `staticwebapp.config.json` to dist

### Output Structure
```
dist/
├── bundle.js              # Main application bundle
├── *.bundle.js            # Code-split chunks
├── index.html             # HTML template
├── environments.json      # Runtime config
├── staticwebapp.config.json
└── assets/                # Static assets
```

## Code Metrics

### File Sizes (approximate)
- `WorkbenchLanding.tsx`: 581 lines
- `WorkbenchService.ts`: 425 lines
- `FileValidationService.ts`: 594 lines
- `GraphEmailService.ts`: 748 lines
- `AuthService.ts`: 664 lines
- `WorkbenchDialogs.tsx`: 265 lines

### Complexity Areas
1. **WorkbenchLanding.tsx** - High state complexity (15+ useState)
2. **AuthService.ts** - Complex token management logic
3. **WorkbenchService.ts** - Large orchestration method
4. **GraphEmailService.ts** - Complex email forwarding logic

## Code Quality Considerations

### Areas for Improvement
1. **State Management** - Consider useReducer for WorkbenchLanding
2. **Service Methods** - Split large methods into smaller functions
3. **Error Handling** - More granular error handling
4. **Type Safety** - Additional type definitions where needed
5. **Testing** - Expand test coverage

### Best Practices Followed
1. Singleton pattern for services
2. TypeScript for type safety
3. Separation of concerns
4. Centralized error handling
5. Environment-based configuration
6. Consistent naming conventions

