# Configuration Guide

This guide explains how to configure the Outlook Add-in for different environments and scenarios.

## Configuration Overview

The application uses **runtime configuration** loaded from `environments.json`. This allows a single build to work across multiple environments (dev, qa, uat, prod).

## Configuration Files

### 1. environments.json
**Location:** Root directory (copied to `dist/` during build)  
**Purpose:** Runtime environment configuration

#### Structure
```json
{
  "environments": {
    "localhost": {
      "REACT_APP_AZURE_CLIENT_ID": "...",
      "REACT_APP_AZURE_AUTHORITY": "...",
      "REACT_APP_AZURE_REDIRECT_URI": "...",
      "REACT_APP_PLACEMENT_API_URL": "...",
      "REACT_APP_PLACEMENT_API_KEY": "...",
      "REACT_APP_LOGGING_API_URL": "...",
      "REACT_APP_LOGGING_API_KEY": "...",
      "DEBUG_ENABLED": true,
      "DEBUG_LEVEL": "debug"
    },
    "dev": { ... },
    "qa": { ... },
    "uat": { ... },
    "prod": { ... }
  },
  "urlPatterns": {
    "localhost": ["localhost", "127.0.0.1"],
    "dev": ["dev.example.com"],
    "qa": ["qa.example.com"],
    "uat": ["uat.example.com"],
    "prod": ["prod.example.com"]
  }
}
```

### 2. app-settings/
**Location:** `app-settings/` directory  
**Purpose:** Application settings per environment (non-secret)

Files:
- `dev-weu.json`
- `qa-weu.json`
- `uat-weu.json`
- `prod-weu.json`

### 3. secret-settings/
**Location:** `secret-settings/` directory  
**Purpose:** Secret settings per environment (API keys, tokens)

Files:
- `dev-weu.json`
- `qa-weu.json`
- `uat-weu.json`
- `prod-weu.json`

## Configuration Variables

### Required Variables

#### Azure AD Configuration
```typescript
REACT_APP_AZURE_CLIENT_ID: string        // Azure AD App Registration Client ID
REACT_APP_AZURE_AUTHORITY: string        // Azure AD Authority URL
REACT_APP_AZURE_REDIRECT_URI: string     // Redirect URI for MSAL
AZURE_API_SCOPES: string                 // Comma-separated API scopes
AZURE_GRAPH_SCOPES: string               // Comma-separated Graph scopes
```

#### API Configuration
```typescript
REACT_APP_PLACEMENT_API_URL: string      // Placement API base URL
REACT_APP_PLACEMENT_API_KEY: string      // Placement API subscription key
REACT_APP_LOGGING_API_URL: string        // Logging API base URL
REACT_APP_LOGGING_API_KEY: string        // Logging API key
```

#### Email Configuration
```typescript
CYBER_MRSNA_MAILBOX: string               // Cyber admin mailbox email
DEFAULT_SHARED_MAILBOX: string           // Default shared mailbox
```

#### Debug Configuration
```typescript
DEBUG_ENABLED: boolean                    // Enable/disable debug logging
DEBUG_LEVEL: string                      // Debug level (error, warn, info, debug, trace)
```

### Optional Variables
```typescript
// Additional configuration as needed
```

## Environment Detection

### Automatic Detection
The application automatically detects the environment based on the URL hostname:

1. Checks `urlPatterns` in `environments.json`
2. Matches hostname against patterns
3. Loads corresponding environment configuration
4. Falls back to `localhost` for local development

### Manual Override
Force environment via URL parameter:
```
https://localhost:3001?env=dev
```

### Detection Logic
```typescript
// From runtimeConfig.ts
private detectEnvironment(urlPatterns: { [key: string]: string[] }): string {
  const hostname = window.location.hostname.toLowerCase();
  
  // Check URL param first
  const urlParams = new URLSearchParams(window.location.search);
  const envParam = urlParams.get('env');
  if (envParam && ['dev', 'qa', 'uat', 'prod', 'localhost'].includes(envParam)) {
    return envParam;
  }
  
  // Check URL patterns
  for (const [env, patterns] of Object.entries(urlPatterns)) {
    for (const pattern of patterns) {
      if (hostname.includes(pattern.toLowerCase())) {
        return env;
      }
    }
  }
  
  // Default
  return hostname === 'localhost' || hostname === '127.0.0.1' ? 'localhost' : 'dev';
}
```

## Configuration Access

### Runtime Configuration
Access configuration via `runtimeConfig`:

```typescript
import runtimeConfig from './config/runtimeConfig';

// Initialize (called in index.tsx)
await runtimeConfig.initialize();

// Get configuration value
const apiUrl = runtimeConfig.getString('REACT_APP_PLACEMENT_API_URL');

// Get all configuration
const allConfig = runtimeConfig.getAll();

// Get current environment
const env = runtimeConfig.getEnvironment();
```

### Environment Helper
Access via `environment` helper:

```typescript
import { environment } from './config/environment';

const apiUrl = environment.PLACEMENT_API_URL;
const debugEnabled = environment.DEBUG_ENABLED;
```

## Environment-Specific Configuration

### Localhost (Development)
```json
{
  "REACT_APP_AZURE_CLIENT_ID": "local-dev-client-id",
  "REACT_APP_AZURE_AUTHORITY": "https://login.microsoftonline.com/tenant-id",
  "REACT_APP_AZURE_REDIRECT_URI": "https://localhost:3001",
  "REACT_APP_PLACEMENT_API_URL": "https://localhost:4001/api/placements",
  "DEBUG_ENABLED": true,
  "DEBUG_LEVEL": "debug"
}
```

### Development
```json
{
  "REACT_APP_AZURE_CLIENT_ID": "dev-client-id",
  "REACT_APP_AZURE_AUTHORITY": "https://login.microsoftonline.com/tenant-id",
  "REACT_APP_AZURE_REDIRECT_URI": "https://dev.example.com",
  "REACT_APP_PLACEMENT_API_URL": "https://dev-api.example.com/api/placements",
  "DEBUG_ENABLED": true,
  "DEBUG_LEVEL": "info"
}
```

### Production
```json
{
  "REACT_APP_AZURE_CLIENT_ID": "prod-client-id",
  "REACT_APP_AZURE_AUTHORITY": "https://login.microsoftonline.com/tenant-id",
  "REACT_APP_AZURE_REDIRECT_URI": "https://prod.example.com",
  "REACT_APP_PLACEMENT_API_URL": "https://api.example.com/api/placements",
  "DEBUG_ENABLED": false,
  "DEBUG_LEVEL": "error"
}
```

## Azure AD Configuration

### App Registration Setup
1. Create Azure AD App Registration
2. Configure redirect URIs:
   - `https://localhost:3001` (local)
   - `https://dev.example.com` (dev)
   - `https://prod.example.com` (prod)

3. Configure API permissions:
   - Microsoft Graph API permissions
   - Custom API permissions

4. Get Client ID and Tenant ID

### MSAL Configuration
```typescript
// From msalConfig.ts
{
  auth: {
    clientId: runtimeConfig.getString('REACT_APP_AZURE_CLIENT_ID'),
    authority: runtimeConfig.getString('REACT_APP_AZURE_AUTHORITY'),
    redirectUri: runtimeConfig.getString('REACT_APP_AZURE_REDIRECT_URI')
  },
  cache: {
    cacheLocation: 'sessionStorage',
    storeAuthStateInCookie: false
  }
}
```

## API Configuration

### Placement API
```typescript
// Base URL
REACT_APP_PLACEMENT_API_URL: "https://api.example.com/api/placements"

// Subscription Key

// Usage
const headers = {
  'Ocp-Apim-Subscription-Key': apiKey,
  'Authorization': `Bearer ${apiToken}`
};
```

### Logging API
```typescript
// Base URL
REACT_APP_LOGGING_API_URL: "https://logging-api.example.com/api"

// API Key
REACT_APP_LOGGING_API_KEY: "your-logging-api-key"

// Endpoints
POST /api/trace
POST /api/event
POST /api/exception
```

## Debug Configuration

### Debug Levels
- **error (0):** Errors only
- **warn (1):** Warnings and errors
- **info (2):** Info, warnings, and errors (default)
- **debug (3):** Debug, info, warnings, and errors
- **trace (4):** All logs (most verbose)

### Configuration
```json
{
  "DEBUG_ENABLED": true,    // Enable/disable debug logging
  "DEBUG_LEVEL": "debug"    // Set debug level
}
```

## Manifest Configuration

### Manifest Files
Located in `Manifests/` directory:
- `Manifest.Local.xml` - Local development
- `Manifest.Dev.xml` - Development
- `Manifest.QA.xml` - QA
- `Manifest.UAT.xml` - UAT
- `Manifest.PRD.xml` - Production

### Manifest Settings
```xml
<OfficeApp>
  <Id>guid-here</Id>
  <Version>1.0.0</Version>
  <ProviderName>Your Company</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Workbench Add-in" />
  <Description DefaultValue="Send emails to Workbench" />
  <IconUrl DefaultValue="https://your-domain.com/assets/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="https://your-domain.com/assets/icon-80.png" />
  <SupportUrl DefaultValue="https://your-domain.com/support" />
  <AppDomains>
    <AppDomain>https://your-domain.com</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>
  <Requirements>
    <Sets>
      <Set Name="Mailbox" MinVersion="1.1" />
    </Sets>
  </Requirements>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://your-domain.com/index.html" />
  </DefaultSettings>
  <Permissions>ReadWriteMailbox</Permissions>
</OfficeApp>
```

## Build Configuration

### Webpack Configuration
**File:** `webpack.config.js`

Key settings:
- Entry: `src/index.tsx`
- Output: `dist/`
- Mode: `production` or `development`
- Source maps: Enabled for debugging

### TypeScript Configuration
**File:** `tsconfig.json`

Key settings:
- Target: `ES2020`
- Module: `ES2020`
- JSX: `react`
- Strict mode: Enabled

## Deployment Configuration

### Static Web App Configuration
**File:** `staticwebapp.config.json`

```json
{
  "navigationFallback": {
    "rewrite": "/index.html",
    "exclude": ["/assets/*", "/*.{js,css,png,gif,jpg,svg}"]
  },
  "routes": [
    {
      "route": "/environments.json",
      "headers": {
        "cache-control": "no-cache"
      }
    }
  ]
}
```

## Configuration Validation

### Required Keys Check
The application validates required configuration keys:

```typescript
const requiredKeys = [
  'REACT_APP_AZURE_CLIENT_ID',
  'REACT_APP_AZURE_AUTHORITY',
  'REACT_APP_AZURE_REDIRECT_URI',
  'REACT_APP_PLACEMENT_API_URL',
  'REACT_APP_LOGGING_API_URL'
];
```

### Configuration Errors
If configuration is invalid:
- Errors are logged via DebugService
- Application may use fallback values
- Check console for warnings

## Security Considerations

### Secrets Management
- **Never commit secrets** to version control
- Use `secret-settings/` for sensitive data
- Use environment variables in CI/CD
- Rotate API keys regularly

### Configuration Security
- Validate configuration on load
- Sanitize user inputs
- Use HTTPS for all API calls
- Validate redirect URIs

## Troubleshooting Configuration

### Problem: Wrong Environment Detected
**Solution:**
1. Check `urlPatterns` in `environments.json`
2. Verify hostname matches pattern
3. Use `?env=dev` URL parameter to override

### Problem: Configuration Not Loading
**Solution:**
1. Check `environments.json` exists in `dist/`
2. Verify file is accessible at `/environments.json`
3. Check browser console for errors
4. Verify CORS settings

### Problem: Missing Configuration Keys
**Solution:**
1. Check required keys are present
2. Verify key names match exactly
3. Check for typos in configuration

### Problem: API Calls Failing
**Solution:**
1. Verify API URLs are correct
2. Check API keys are valid
3. Verify authentication tokens
4. Check CORS configuration

## Configuration Best Practices

1. **Environment Separation:** Keep configurations separate per environment
2. **Version Control:** Don't commit secrets, use `.gitignore`
3. **Validation:** Validate configuration on load
4. **Documentation:** Document all configuration variables
5. **Defaults:** Provide sensible defaults where possible
6. **Security:** Never expose secrets in client-side code

## Configuration Checklist

### Before Deployment
- [ ] All required keys are set
- [ ] API URLs are correct
- [ ] Azure AD configuration is correct
- [ ] Debug settings are appropriate
- [ ] Secrets are not committed
- [ ] Manifest files are updated
- [ ] URL patterns are correct

### After Deployment
- [ ] Verify environment detection
- [ ] Test API connectivity
- [ ] Verify authentication
- [ ] Check debug logging
- [ ] Validate configuration loading

