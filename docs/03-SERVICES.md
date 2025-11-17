# Services Documentation

This document provides detailed information about all services in the application, their purposes, methods, and usage patterns.

## Service Overview

All services follow the singleton pattern and are accessed via `getInstance()` method. Services are located in `src/service/` directory.

## Core Services

### 1. AuthService
**File:** `src/service/AuthService.ts`  
**Lines:** ~664  
**Purpose:** Manages authentication and token acquisition for both Placement API and Microsoft Graph API.

#### Key Responsibilities
- Azure AD authentication via MSAL
- Token acquisition (API token + Graph token)
- Token caching and refresh
- Office SSO fallback support
- Interaction management (prevents multiple popups)

#### Main Methods

##### `acquireBothTokens()`
Acquires both API and Graph tokens in a single session to prevent multiple authentication popups.

```typescript
const { apiToken, graphToken } = await AuthService.acquireBothTokens();
```

**Returns:**
- `{ apiToken: AuthenticationResult | null, graphToken: AuthenticationResult | null }`

**Behavior:**
- Checks for valid cached tokens first
- Attempts silent token acquisition
- Falls back to popup if silent fails
- Uses combined scopes to get both tokens in one popup

##### `getApiToken()`
Gets API token for Placement API calls.

```typescript
const token = await AuthService.getApiToken();
```

**Returns:** `AuthenticationResult | null`

##### `getGraphToken()`
Gets Graph token for Microsoft Graph API calls.

```typescript
const token = await AuthService.getGraphToken();
```

**Returns:** `AuthenticationResult | null`

##### `ensureAuthenticated()`
Ensures user is authenticated (tries Office SSO first, then MSAL).

```typescript
const isAuth = await AuthService.ensureAuthenticated();
```

**Returns:** `boolean`

##### `isOfficeSSOAvailable()`
Checks if Office SSO is available.

**Returns:** `boolean`

##### `clearTokens()`
Clears cached tokens (useful for logout or error recovery).

##### `getTokenInfo()`
Gets token status information for debugging.

**Returns:**
```typescript
{
  apiTokenValid: boolean;
  graphTokenValid: boolean;
  apiExpiry: number;
  graphExpiry: number;
}
```

#### Token Management
- Tokens cached in-memory
- Auto-refresh 5 minutes before expiry
- Request deduplication prevents multiple simultaneous requests
- Interaction lock prevents concurrent popups

#### Scopes
- **API Scopes:** From `environment.AZURE_API_SCOPES`
- **Graph Scopes:** From `environment.AZURE_GRAPH_SCOPES`

---

### 2. WorkbenchService
**File:** `src/service/WorkbenchService.ts`  
**Lines:** ~425  
**Purpose:** Main orchestration service that coordinates the entire placement submission flow.

#### Key Responsibilities
- Orchestrates email submission workflow
- Coordinates multiple services
- Handles email conversion
- Manages placement submission
- Handles email forwarding
- Error handling and recovery

#### Main Methods

##### `submitPlacement()`
Main entry point for submitting a placement request.

```typescript
const result = await WorkbenchService.submitPlacement(
  apiToken: string,
  graphToken: string,
  item: Office.Item,
  productCode: string,
  sendCopyToCyberAdmin: boolean
);
```

**Returns:** `WorkbenchSubmissionResult`

**Result Structure:**
```typescript
{
  success: boolean;
  placementId?: string;
  error?: string;
  forwardingFailed?: boolean;
  forwardingFailedReason?: string;
  lastPlacementId?: string;
  lastGraphItemId?: string;
  lastSharedMailbox?: string;
}
```

**Workflow:**
1. Validates input parameters
2. Converts email to EML format
3. Extracts email metadata (sender, subject, date)
4. Submits to Placement API
5. Stamps email with Workbench ID
6. Shows notification banner
7. Forwards email (if enabled)

##### `retryForward()`
Retries email forwarding after a failed attempt.

```typescript
const result = await WorkbenchService.retryForward(
  graphToken: string,
  placementId: string,
  graphItemId: string,
  sharedMailbox: string
);
```

**Returns:** `WorkbenchSubmissionResult`

#### Dependencies
- `EmailConverterService` - Email conversion
- `PlacementApiService` - API submission
- `GraphEmailService` - Email forwarding
- `LoggingService` - Logging
- `DebugService` - Debug logging

---

### 3. PlacementApiService
**File:** `src/service/PlacementApiService.ts`  
**Lines:** ~186  
**Purpose:** Handles communication with the Placement API backend.

#### Key Responsibilities
- Placement request submission
- API authentication
- Request/response handling
- Error handling

#### Main Methods

##### `submitPlacementRequest()`
Submits a placement request to the API.

```typescript
const result = await PlacementApiService.submitPlacementRequest(
  apiToken: string,
  data: PlacementRequestData
);
```

**Request Data:**
```typescript
{
  productCode: string;
  emailSender: string;
  emailSubject: string;
  emailReceivedDateTime: string;
  emlContent: string; // Base64 encoded EML
}
```

**Returns:** `PlacementResponse`
```typescript
{
  placementId: string;
  ingestionId: string;
  runId: string;
}
```

#### Configuration
- **API URL:** From `environment.PLACEMENT_API_URL`
- **Subscription Key:** From `environment.PLACEMENT_API_KEY`
- **Authentication:** Bearer token from `apiToken`

#### Error Handling
- Validates API response
- Throws descriptive errors
- Logs errors via DebugService

---

### 4. GraphEmailService
**File:** `src/service/GraphEmailService.ts`  
**Lines:** ~748  
**Purpose:** Handles Microsoft Graph API operations for email management.

#### Key Responsibilities
- Email forwarding via Graph API
- Email search and retrieval
- Shared mailbox detection
- Email ID conversion (Office ID → Graph ID)

#### Main Methods

##### `forwardEmailWithGraphToken()`
Forwards an email to a shared mailbox using Graph API.

```typescript
await GraphEmailService.forwardEmailWithGraphToken(
  graphToken: string,
  data: ForwardEmailData
);
```

**Request Data:**
```typescript
{
  emailId: string;        // Graph API email ID
  sharedMailbox: string;   // Target mailbox email
  uwwbID: string;         // Workbench placement ID
  internetMessageId?: string; // Optional message ID
}
```

**Behavior:**
- Converts Office ID to Graph ID if needed
- Forwards email with Workbench ID in subject
- Handles draft emails differently
- Falls back to search method if direct ID fails

##### `searchEmailByConversationId()`
Searches for email by conversation ID (fallback method).

```typescript
const emailId = await GraphEmailService.searchEmailByConversationId(
  graphToken: string,
  conversationId: string
);
```

**Returns:** `string | null` (Graph email ID)

##### `convertOfficeIdToGraphId()`
Converts Office.js item ID to Microsoft Graph API email ID.

```typescript
const graphId = await GraphEmailService.convertOfficeIdToGraphId(
  graphToken: string,
  officeId: string
);
```

**Returns:** `string | null`

#### Graph API Endpoints Used
- `GET /me/messages/{id}` - Get email by ID
- `POST /me/messages/{id}/forward` - Forward email
- `GET /me/messages` - Search emails
- `GET /users/{mailbox}/messages` - Access shared mailbox

---

### 5. EmailConverterService
**File:** `src/service/EmailConverterService.ts`  
**Lines:** ~408  
**Purpose:** Converts Office.js email items to EML format for API submission.

#### Key Responsibilities
- Office.js email to EML conversion
- Attachment handling
- Email metadata extraction
- Size limit enforcement

#### Main Methods

##### `convertEmailToEml()`
Converts an Office.js email item to EML format.

```typescript
const emlData = await EmailConverterService.convertEmailToEml(item);
```

**Returns:** `EmlEmailData`
```typescript
{
  content: string; // Complete EML file as string
}
```

**Process:**
1. Extracts email metadata (headers, sender, recipient, etc.)
2. Gets email body (HTML or text)
3. Processes attachments (with size limits)
4. Builds RFC 822 compliant EML content
5. Validates EML structure

#### Size Limits
- **Max Attachment Size:** 25MB per attachment
- **Max Base64 Length:** 35MB (to prevent stack overflow)
- **Truncation:** Large attachments are truncated with warning

#### Attachment Handling
- Base64 encoding
- Content-Type detection
- Size validation
- Truncation for oversized files

---

### 6. FileValidationService
**File:** `src/service/FileValidationService.ts`  
**Lines:** ~594  
**Purpose:** Validates email attachments for various restrictions before submission.

#### Key Responsibilities
- File type validation
- Encryption detection
- Password protection detection
- ZIP/compressed file detection
- Unsupported file type detection

#### Main Methods

##### `validateEmailFiles()`
Validates all attachments in an email item.

```typescript
const result = await FileValidationService.validateEmailFiles(item);
```

**Returns:** `FileValidationResult`
```typescript
{
  isValid: boolean;
  errors: FileValidationError[];
}
```

**Error Types:**
- `zip` - Compressed files detected
- `unsupported` - Unsupported file types
- `encrypted` - Encrypted files detected
- `password_protected` - Password-protected files

#### Validation Checks

1. **ZIP/Compressed Files**
   - Detects: `.zip`, `.rar`, `.7z`, `.tar`, `.gz`, etc.
   - Reason: Compressed files may contain restricted content

2. **Unsupported File Types**
   - Validates against whitelist of supported extensions
   - Common supported: `.pdf`, `.doc`, `.docx`, `.xls`, `.xlsx`, `.txt`, etc.

3. **Encrypted Files**
   - Detects: `.gpg`, `.pgp`, `.encrypted`
   - Basic pattern matching

4. **Password-Protected Files**
   - Office documents (Word, Excel, PowerPoint)
   - PDF files
   - Files with "password" in filename

#### Helper Methods

##### `getAllErrorMessages()`
Gets all error messages from validation result.

```typescript
const messages = FileValidationService.getAllErrorMessages(result.errors);
```

**Returns:** `string[]`

##### `isSupportedFileType()`
Checks if a file type is supported.

```typescript
const isSupported = FileValidationService.isSupportedFileType(filename);
```

**Returns:** `boolean`

---

### 7. OfficeModeService
**File:** `src/service/OfficeModeService.ts`  
**Lines:** ~99  
**Purpose:** Detects and manages Office.js mode (compose vs read).

#### Key Responsibilities
- Detect current Office mode
- Distinguish between draft and sent emails
- Provide mode-specific utilities

#### Main Methods

##### `getCurrentMode()`
Gets the current Office.js mode.

```typescript
const mode = OfficeModeService.getCurrentMode();
```

**Returns:** `OfficeMode` enum
- `COMPOSE` - Draft/compose mode
- `READ` - Read/sent email mode
- `UNKNOWN` - Unable to determine

##### `isComposeMode()`
Checks if currently in compose (draft) mode.

```typescript
const isDraft = OfficeModeService.isComposeMode();
```

**Returns:** `boolean`

##### `isReadMode()`
Checks if currently in read mode.

```typescript
const isRead = OfficeModeService.isReadMode();
```

**Returns:** `boolean`

##### `isMessageItem()`
Checks if current item is a message (email).

```typescript
const isMessage = OfficeModeService.isMessageItem();
```

**Returns:** `boolean`

#### Usage
Used throughout the application to handle draft emails differently from sent emails, especially for:
- Email ID retrieval
- Email forwarding
- Email stamping

---

### 8. OfficeIdConverterService
**File:** `src/service/OfficeIdConverterService.ts`  
**Purpose:** Converts Office.js item IDs to various formats needed by different APIs.

#### Key Responsibilities
- Office ID to Graph ID conversion
- Office ID to REST ID conversion
- ID format normalization

#### Main Methods

##### `convertToGraphId()`
Converts Office.js item ID to Microsoft Graph API ID.

```typescript
const graphId = await OfficeIdConverterService.convertToGraphId(officeId);
```

**Returns:** `string | null`

##### `convertToRestId()`
Converts Office.js item ID to REST API ID.

```typescript
const restId = await OfficeIdConverterService.convertToRestId(officeId);
```

**Returns:** `string | null`

---

### 9. OfficeEmailService
**File:** `src/service/OfficeEmailService.ts`  
**Purpose:** Wrapper for Office.js email operations.

#### Key Responsibilities
- Office.js API abstraction
- Email item access
- Attachment access
- Office.js error handling

#### Main Methods
- Email item retrieval
- Attachment access
- Email property access

---

### 10. DebugService
**File:** `src/service/DebugService.ts`  
**Lines:** ~243  
**Purpose:** Centralized debug logging service with environment-aware controls.

#### Key Responsibilities
- Centralized console logging
- Debug level control
- Environment-aware logging
- Structured logging

#### Configuration
- **DEBUG_ENABLED:** From `environment.DEBUG_ENABLED`
- **DEBUG_LEVEL:** From `environment.DEBUG_LEVEL` (error, warn, info, debug, trace)

#### Main Methods

##### Logging Methods
```typescript
DebugService.error(message, ...args);      // Always logged
DebugService.warn(message, ...args);      // If level >= warn
DebugService.info(message, ...args);      // If level >= info
DebugService.debug(message, ...args);    // If level >= debug
DebugService.trace(message, ...args);     // If level >= trace
```

##### Specialized Logging
```typescript
DebugService.auth(event, details);        // Authentication events
DebugService.email(operation, details);   // Email operations
DebugService.placement(operation, details); // Placement operations
DebugService.graph(operation, details);   // Graph API operations
DebugService.api(method, url, data);       // API calls
DebugService.section(title);               // Section header
DebugService.object(label, obj);          // Object logging
DebugService.timing(operation, startTime); // Performance timing
```

#### Debug Levels
- **error (0):** Errors only
- **warn (1):** Warnings and errors
- **info (2):** Info, warnings, and errors (default)
- **debug (3):** Debug, info, warnings, and errors
- **trace (4):** All logs (most verbose)

#### Usage Pattern
```typescript
DebugService.section('Starting Placement Submission');
DebugService.object('Submission parameters', { productCode, sendCopyToCyberAdmin });
DebugService.email('Converting email to EML');
DebugService.api('POST', '/api/placements', data);
```

---

### 11. LoggingService
**File:** `src/service/LoggingService.ts`  
**Lines:** ~119  
**Purpose:** External logging service for audit trail and error tracking.

#### Key Responsibilities
- Log errors to external API
- Log placement requests
- Log user actions
- Audit trail

#### Main Methods

##### `logError()`
Logs an error to the external logging API.

```typescript
await LoggingService.logError(error, context, extra);
```

##### `logPlacementRequest()`
Logs a placement request event.

```typescript
await LoggingService.logPlacementRequest(placementId, userId);
```

##### `logTrace()`
Logs a trace message.

```typescript
await LoggingService.logTrace(message, severityLevel, properties);
```

**Severity Levels:**
- `Information`
- `Verbose`
- `Event`
- `Exception`

##### `logEvent()`
Logs an event.

```typescript
await LoggingService.logEvent(trackingName, properties, metrics);
```

##### `logException()`
Logs an exception.

```typescript
await LoggingService.logException(message, properties, metrics);
```

#### Configuration
- **API URL:** From `environment.LOGGING_API_URL`
- **API Key:** From `environment.LOGGING_API_KEY`
- **Authentication:** Uses API token from AuthService

#### Error Handling
- Swallows errors to avoid breaking app flow
- Logs errors via DebugService if logging fails
- Non-blocking (doesn't throw errors)

---

### 12. ApiClient
**File:** `src/service/ApiClient.ts`  
**Purpose:** Base HTTP client for API requests.

#### Key Responsibilities
- HTTP request handling
- Response parsing
- Error handling
- Request/response logging

#### Main Methods
- `get()`, `post()`, `put()`, `delete()` - HTTP methods
- Error handling and retry logic
- Response type conversion

---

## Service Interaction Patterns

### Typical Submission Flow
```
WorkbenchService.submitPlacement()
    ↓
EmailConverterService.convertEmailToEml()
    ↓
FileValidationService.validateEmailFiles()
    ↓
PlacementApiService.submitPlacementRequest()
    ↓
GraphEmailService.forwardEmailWithGraphToken()
    ↓
LoggingService.logPlacementRequest()
```

### Authentication Flow
```
AuthService.acquireBothTokens()
    ↓
MSAL Authentication (if needed)
    ↓
Token Caching
    ↓
Token Refresh (auto)
```

## Service Dependencies

### Dependency Graph
```
WorkbenchService
  ├── EmailConverterService
  ├── PlacementApiService
  ├── GraphEmailService
  ├── LoggingService
  └── DebugService

PlacementApiService
  ├── ApiClient
  └── DebugService

GraphEmailService
  ├── OfficeIdConverterService
  ├── OfficeModeService
  └── DebugService

AuthService
  ├── MSAL Instance
  └── DebugService

All Services
  └── DebugService (for logging)
```

## Best Practices

1. **Always use getInstance()** - Services are singletons
2. **Handle errors gracefully** - Services throw errors that need catching
3. **Use DebugService for logging** - Don't use console.log directly
4. **Check return values** - Services may return null on failure
5. **Await async methods** - All service methods are async

## Common Patterns

### Service Usage Pattern
```typescript
// Get service instance
const service = ServiceName.getInstance();

// Call service method
try {
  const result = await service.methodName(params);
  // Handle success
} catch (error) {
  // Handle error
  DebugService.error('Service error:', error);
}
```

### Service Initialization
Services are initialized lazily on first use. No explicit initialization needed.

### Error Handling
Services throw errors that should be caught by calling code. Use try-catch blocks.

