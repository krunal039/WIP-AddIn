# Architecture Overview

## High-Level Solution Architecture

The **Outlook Add-in: Send to Workbench** is a cross-platform Office Add-in that enables users to submit emails to the Workbench placement system directly from Outlook. The application is built using React, TypeScript, and integrates with Microsoft Graph API and Azure AD for authentication.

## System Components

```
┌─────────────────────────────────────────────────────────────┐
│                    Outlook Add-in UI                         │
│  ┌──────────────┐  ┌──────────────┐  ┌──────────────┐      │
│  │   App.tsx    │  │ Workbench    │  │   Dialogs    │      │
│  │  (Auth UI)   │  │  Landing     │  │  (Success/   │      │
│  │              │  │  (Main UI)   │  │   Error)     │      │
│  └──────────────┘  └──────────────┘  └──────────────┘      │
└─────────────────────────────────────────────────────────────┘
                            │
                            ▼
┌─────────────────────────────────────────────────────────────┐
│                    Service Layer                            │
│  ┌──────────────┐  ┌──────────────┐  ┌──────────────┐    │
│  │ AuthService  │  │ Workbench     │  │   Graph      │    │
│  │  (MSAL)      │  │  Service      │  │   Email      │    │
│  │              │  │               │  │   Service    │    │
│  └──────────────┘  └──────────────┘  └──────────────┘    │
│  ┌──────────────┐  ┌──────────────┐  ┌──────────────┐    │
│  │ Placement    │  │   File       │  │   Email      │    │
│  │   API        │  │ Validation   │  │  Converter  │    │
│  │  Service     │  │  Service     │  │  Service    │    │
│  └──────────────┘  └──────────────┘  └──────────────┘    │
└─────────────────────────────────────────────────────────────┘
                            │
                            ▼
┌─────────────────────────────────────────────────────────────┐
│                  External Services                           │
│  ┌──────────────┐  ┌──────────────┐  ┌──────────────┐    │
│  │  Azure AD    │  │  Microsoft   │  │  Placement   │    │
│  │  (MSAL)      │  │  Graph API   │  │  API         │    │
│  │              │  │              │  │              │    │
│  └──────────────┘  └──────────────┘  └──────────────┘    │
└─────────────────────────────────────────────────────────────┘
```

## Application Flow

### 1. Initialization Flow
```
Office.onReady()
    ↓
RuntimeConfig.initialize() → Load environments.json
    ↓
MSAL Instance Creation
    ↓
App Component Mount
    ↓
AuthService.acquireBothTokens()
    ↓
WorkbenchLanding Component (Main UI)
```

### 2. Submission Flow
```
User selects email → Clicks "Send to Workbench"
    ↓
File Validation (parallel checks)
    ↓
Duplicate Detection
    ↓
Email Conversion (Office → EML)
    ↓
Placement API Submission
    ↓
Email Stamping (add Workbench ID)
    ↓
Email Forwarding (if enabled)
    ↓
Success/Error UI Feedback
```

## Technology Stack

### Frontend
- **React 18.3.1** - UI framework
- **TypeScript 5.8.3** - Type safety
- **Fluent UI 8/9** - Microsoft design system
- **Office.js** - Outlook Add-in API

### Authentication
- **@azure/msal-browser 4.18.0** - Azure AD authentication
- **@azure/msal-react 3.0.15** - React integration
- **Office SSO** - Fallback authentication method

### Build Tools
- **Webpack 5** - Module bundler
- **ts-loader** - TypeScript compilation
- **webpack-dev-server** - Development server

### Utilities
- **jszip 3.10.1** - ZIP file handling
- **axios 1.10.0** - HTTP client
- **form-data 4.0.4** - Form data handling

## Key Design Patterns

### 1. Singleton Pattern
All services use the singleton pattern to ensure single instance:
- `AuthService.getInstance()`
- `WorkbenchService.getInstance()`
- `DebugService.getInstance()`
- `PlacementApiService.getInstance()`
- etc.

### 2. Service Layer Pattern
Business logic is separated into service classes:
- Services handle API calls, data transformation, and business rules
- Components focus on UI rendering and user interaction
- Clear separation of concerns

### 3. Runtime Configuration
Configuration is loaded at runtime based on URL:
- Single build works across multiple environments
- Environment detection via URL patterns
- Configuration stored in `environments.json`

### 4. Error Handling
- Centralized error handling via `ErrorBoundary`
- Service-level error logging via `LoggingService`
- User-friendly error messages

## Data Flow

### Authentication Tokens
```
User Login
    ↓
MSAL Authentication
    ↓
API Token (for Placement API)
Graph Token (for Graph API)
    ↓
Token Caching (in-memory)
    ↓
Auto-refresh (5 minutes before expiry)
```

### Email Processing
```
Office.Item (Office.js)
    ↓
EmailConverterService → EML Format
    ↓
FileValidationService → Validation Results
    ↓
WorkbenchService → Orchestration
    ↓
PlacementApiService → API Submission
    ↓
GraphEmailService → Email Forwarding
```

## Environment Configuration

The application supports multiple environments:
- **localhost** - Local development
- **dev** - Development environment
- **qa** - Quality assurance
- **uat** - User acceptance testing
- **prod** - Production

Each environment has its own configuration in `environments.json`:
- Azure AD client ID and authority
- API endpoints
- Feature flags
- Debug settings

## Security Considerations

1. **Token Management**
   - Tokens stored in-memory only (not persisted)
   - Auto-refresh before expiry
   - Secure token transmission via HTTPS

2. **API Security**
   - All API calls use Bearer token authentication
   - Subscription keys for Placement API
   - CORS configuration for cross-origin requests

3. **File Validation**
   - Prevents submission of encrypted files
   - Blocks password-protected documents
   - Validates file types before processing

## Performance Optimizations

1. **Parallel Operations**
   - File validation checks run in parallel
   - Email stamping and notification banner shown simultaneously
   - Email metadata extraction parallelized

2. **Token Caching**
   - Tokens cached in-memory to reduce API calls
   - Request deduplication prevents multiple simultaneous token requests

3. **Lazy Loading**
   - Components loaded on demand
   - Services initialized only when needed

## Deployment Architecture

### Build Process
```
Source Code
    ↓
Webpack Build (production mode)
    ↓
Bundle Optimization
    ↓
dist/ folder (static files)
    ↓
Azure Static Web Apps / CDN
```

### Manifest Files
Multiple manifest files for different environments:
- `Manifest.Local.xml` - Local development
- `Manifest.Dev.xml` - Development
- `Manifest.QA.xml` - QA
- `Manifest.UAT.xml` - UAT
- `Manifest.PRD.xml` - Production

## Integration Points

1. **Office.js API**
   - Email item access
   - Attachment handling
   - Office SSO authentication

2. **Microsoft Graph API**
   - Email forwarding
   - Email search
   - User profile access

3. **Placement API**
   - Placement request submission
   - Placement ID generation
   - Status tracking

4. **Logging API**
   - Error logging
   - Placement request logging
   - Audit trail

## Future Considerations

- State management refactoring (useReducer)
- Code splitting for better performance
- Bundle size optimization
- Enhanced error recovery
- Offline support capabilities

