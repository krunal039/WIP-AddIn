# Debugging Guide

This guide provides comprehensive information for debugging the Outlook Add-in application.

## Debug Service Overview

The application uses a centralized `DebugService` for all logging. This service respects environment configuration and provides structured logging.

## Enabling Debug Logging

### Runtime Configuration
Debug logging is controlled by `environments.json`:

```json
{
  "environments": {
    "dev": {
      "DEBUG_ENABLED": true,
      "DEBUG_LEVEL": "debug"
    },
    "prod": {
      "DEBUG_ENABLED": false,
      "DEBUG_LEVEL": "error"
    }
  }
}
```

### Debug Levels
- **error (0):** Errors only (always logged)
- **warn (1):** Warnings and errors
- **info (2):** Info, warnings, and errors (default)
- **debug (3):** Debug, info, warnings, and errors
- **trace (4):** All logs (most verbose)

### Force Debug via URL
Add `?env=dev` to URL to force development environment:
```
https://localhost:3001?env=dev
```

## Console Logging

### DebugService Methods

#### Basic Logging
```typescript
DebugService.error('Error message', error);
DebugService.warn('Warning message');
DebugService.info('Info message');
DebugService.debug('Debug message');
DebugService.trace('Trace message');
```

#### Specialized Logging
```typescript
// Authentication events
DebugService.auth('User authenticated', { userId: 'user@example.com' });

// Email operations
DebugService.email('Email converted', { emailId: '123' });

// Placement operations
DebugService.placement('Placement submitted', { placementId: 'abc' });

// Graph API operations
DebugService.graph('Email forwarded', { emailId: '123' });

// API calls
DebugService.api('POST', '/api/placements', { data });

// Structured logging
DebugService.section('Starting Submission');
DebugService.subsection('File Validation');
DebugService.object('Validation Result', result);
DebugService.timing('Submission', startTime);
```

### Console Output Format
```
🔴 [ERROR] Error message
🟡 [WARN] Warning message
🔵 [INFO] Info message
🟢 [DEBUG] Debug message
⚪ [TRACE] Trace message
🔐 [AUTH] Authentication event
📧 [EMAIL] Email operation
📦 [PLACEMENT] Placement operation
📈 [GRAPH] Graph API operation
🌐 [API] POST /api/placements
```

## Common Debugging Scenarios

### 1. Authentication Issues

#### Problem: Authentication popup not appearing
**Debug Steps:**
1. Check console for auth logs:
   ```
   DebugService.auth('Starting MSAL authentication');
   ```
2. Check MSAL configuration:
   ```typescript
   DebugService.debug('[AuthService] MSAL config:', {
     clientId: config.auth?.clientId,
     authority: config.auth?.authority,
     redirectUri: config.auth?.redirectUri
   });
   ```
3. Check for interaction_in_progress errors
4. Verify Office SSO availability:
   ```typescript
   const isAvailable = AuthService.isOfficeSSOAvailable();
   DebugService.auth('Office SSO available:', isAvailable);
   ```

#### Problem: Token acquisition failing
**Debug Steps:**
1. Check token status:
   ```typescript
   const tokenInfo = AuthService.getTokenInfo();
   DebugService.object('Token Info', tokenInfo);
   ```
2. Check token expiry:
   ```typescript
   DebugService.debug('Token expiry:', {
     apiExpiry: tokenInfo.apiExpiry,
     graphExpiry: tokenInfo.graphExpiry,
     currentTime: Math.floor(Date.now() / 1000)
   });
   ```
3. Check for token errors in console

### 2. Email Submission Issues

#### Problem: Submission failing
**Debug Steps:**
1. Check submission flow logs:
   ```
   DebugService.section('Starting Placement Submission');
   ```
2. Check file validation:
   ```typescript
   DebugService.object('File Validation Result', fileValidationResult);
   ```
3. Check email conversion:
   ```typescript
   DebugService.email('Email converted', { emlLength: emlData.content.length });
   ```
4. Check API submission:
   ```typescript
   DebugService.api('POST', '/api/placements', placementData);
   ```

#### Problem: Email forwarding failing
**Debug Steps:**
1. Check Graph token:
   ```typescript
   DebugService.graph('Graph token status', { hasToken: !!graphToken });
   ```
2. Check email ID conversion:
   ```typescript
   DebugService.graph('Office ID to Graph ID', { officeId, graphId });
   ```
3. Check forwarding request:
   ```typescript
   DebugService.graph('Forwarding email', { emailId, sharedMailbox });
   ```

### 3. File Validation Issues

#### Problem: Valid files being rejected
**Debug Steps:**
1. Check validation logs:
   ```typescript
   DebugService.object('File Validation', {
     files: attachments.map(a => a.name),
     result: fileValidationResult
   });
   ```
2. Check file type detection:
   ```typescript
   DebugService.debug('File type check', {
     filename: 'test.pdf',
     isSupported: FileValidationService.isSupportedFileType('test.pdf')
   });
   ```
3. Check encryption detection:
   ```typescript
   DebugService.debug('Encryption check', {
     filename: 'test.doc',
     isEncrypted: await detectFileProtectionFromBase64(base64Content)
   });
   ```

### 4. Configuration Issues

#### Problem: Wrong environment detected
**Debug Steps:**
1. Check environment detection:
   ```typescript
   DebugService.info('Environment detected:', runtimeConfig.getEnvironment());
   ```
2. Check URL patterns:
   ```typescript
   DebugService.debug('Hostname:', window.location.hostname);
   ```
3. Check configuration loading:
   ```typescript
   DebugService.info('Configuration loaded:', {
     environment: runtimeConfig.getEnvironment(),
     keys: Object.keys(runtimeConfig.getAll()).length
   });
   ```

## Browser Developer Tools

### Chrome DevTools

#### Console Tab
- View all DebugService logs
- Filter by log level
- Search for specific messages

#### Network Tab
- Monitor API calls
- Check request/response headers
- Verify authentication tokens
- Check for CORS errors

**Key Endpoints to Monitor:**
- `/api/placements` - Placement API
- `/api/trace`, `/api/event`, `/api/exception` - Logging API
- `https://graph.microsoft.com/*` - Graph API
- `/environments.json` - Configuration

#### Application Tab
- **Local Storage:** Check stored preferences
- **Session Storage:** Check temporary data
- **Cookies:** Check authentication cookies

#### Sources Tab
- Set breakpoints in TypeScript files
- Step through code execution
- Inspect variables
- Watch expressions

### Debugging Tips

1. **Use Source Maps:** Ensure source maps are enabled for readable debugging
2. **Set Breakpoints:** Use breakpoints in service methods
3. **Watch Variables:** Watch token values, state variables
4. **Network Inspection:** Monitor all API calls

## Office.js Debugging

### Office.js Console
Access via: `Office.context.mailbox.item`

#### Check Office Context
```typescript
DebugService.debug('Office context:', {
  itemType: Office.context.mailbox.item.itemType,
  itemId: (Office.context.mailbox.item as any).itemId,
  subject: Office.context.mailbox.item.subject
});
```

#### Check Office Mode
```typescript
const mode = OfficeModeService.getCurrentMode();
DebugService.debug('Office mode:', mode);
```

### Office.js Errors
- Check for Office.js API errors
- Verify Office.js is loaded: `typeof Office !== 'undefined'`
- Check Office.js version compatibility

## API Debugging

### Placement API

#### Check API Request
```typescript
DebugService.api('POST', '/api/placements', {
  productCode,
  emailSender,
  emailSubject,
  emlContentLength: emlContent.length
});
```

#### Check API Response
```typescript
DebugService.object('Placement Response', {
  placementId: response.placementId,
  ingestionId: response.ingestionId,
  runId: response.runId
});
```

### Graph API

#### Check Graph Token
```typescript
DebugService.graph('Graph token', {
  hasToken: !!graphToken,
  tokenLength: graphToken?.length
});
```

#### Check Graph Request
```typescript
DebugService.graph('Forwarding email', {
  emailId,
  sharedMailbox,
  uwwbID: placementId
});
```

## Error Handling Debugging

### Error Boundary
Check ErrorBoundary component for caught errors:
```typescript
componentDidCatch(error: Error, errorInfo: React.ErrorInfo) {
  DebugService.error('Error caught by boundary:', error);
  DebugService.error('Error info:', errorInfo);
}
```

### Service Errors
All services log errors via DebugService:
```typescript
try {
  // Service operation
} catch (error) {
  DebugService.errorWithStack('Service error:', error);
  // Error handling
}
```

## Performance Debugging

### Timing Operations
```typescript
const startTime = Date.now();
// Operation
DebugService.timing('Operation name', startTime);
```

### Performance Metrics
```typescript
DebugService.performance('Submission time', duration, 'ms');
DebugService.performance('File validation time', validationTime, 'ms');
```

## Common Error Messages

### Authentication Errors
- `"interaction_in_progress"` - Another auth popup is open
- `"interaction_required"` - User interaction needed
- `"Failed to acquire API token"` - Token acquisition failed

### API Errors
- `"Failed to load environments.json"` - Configuration not found
- `"Environment not found"` - Invalid environment
- `"API request failed"` - API call failed

### Office.js Errors
- `"Office.js not loaded"` - Office.js not available
- `"Item not available"` - Email item not accessible
- `"Invalid item type"` - Wrong Office.js item type

## Debugging Checklist

### Before Debugging
- [ ] Enable debug logging in environment config
- [ ] Open browser DevTools
- [ ] Check console for errors
- [ ] Verify environment configuration

### During Debugging
- [ ] Check DebugService logs
- [ ] Monitor network requests
- [ ] Verify token status
- [ ] Check Office.js context
- [ ] Inspect component state

### After Debugging
- [ ] Document findings
- [ ] Create bug report if needed
- [ ] Update error handling if needed

## Debugging Tools

### Browser Extensions
- React Developer Tools
- Redux DevTools (if using Redux)
- Office Add-in Debugger

### VS Code Debugging
Configure `.vscode/launch.json`:
```json
{
  "type": "chrome",
  "request": "launch",
  "name": "Debug Add-in",
  "url": "https://localhost:3001",
  "webRoot": "${workspaceFolder}",
  "sourceMaps": true
}
```

## Logging Best Practices

1. **Use Appropriate Log Levels:**
   - Error: Actual errors
   - Warn: Warnings
   - Info: Important events
   - Debug: Detailed debugging
   - Trace: Very detailed tracing

2. **Include Context:**
   ```typescript
   DebugService.error('Submission failed', {
     placementId,
     error: error.message,
     userId: Office.context.mailbox.userProfile.emailAddress
   });
   ```

3. **Use Structured Logging:**
   ```typescript
   DebugService.object('Submission data', {
     productCode,
     emailSubject,
     fileCount: attachments.length
   });
   ```

4. **Don't Log Sensitive Data:**
   - Don't log full tokens
   - Don't log passwords
   - Don't log personal information

## Getting Help

### Information to Collect
1. Browser console logs
2. Network request/response logs
3. Environment configuration
4. Office.js context information
5. Error stack traces
6. Steps to reproduce

### Support Resources
- Check documentation in `docs/` folder
- Review service documentation
- Check component documentation
- Review error messages in console

