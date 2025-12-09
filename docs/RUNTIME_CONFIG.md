# Runtime Environment Configuration

## Overview

This application now uses a **runtime environment configuration system** that allows a single build to work across multiple environments (dev, qa, uat, prod). The environment is automatically detected based on the URL where the application is running.

## How It Works

1. **Single Build**: Build the application once with `npm run build`
2. **Runtime Detection**: When the app loads, it detects the environment based on the current URL
3. **Configuration Loading**: The app loads the appropriate configuration from `environments.json` based on the detected environment
4. **Automatic Configuration**: All environment variables are automatically set based on the detected environment

## File Structure

- `environments.json` - Contains all environment configurations in one file
- `src/config/runtimeConfig.ts` - Runtime configuration loader that detects environment and loads config
- `src/config/environment.ts` - Environment variable accessor (uses runtime config)
- `src/index.tsx` - Initializes runtime config before app renders

## Environment Detection

The system detects the environment by matching the current hostname against URL patterns defined in `environments.json`:

- **dev**: `gsi-email-ingestion-request-dev.munichre.com`, `localhost` (any port, e.g., `localhost:3035`), `127.0.0.1`
- **qa**: `gsi-email-ingestion-request-qa.munichre.com`
- **uat**: `gsi-email-ingestion-request-uat.munichre.com`
- **prod**: `gsi-email-ingestion-request.munichre.com`

### Manual Override

You can force a specific environment by adding `?env=<environment>` to the URL:
- `https://your-app.com/?env=dev`
- `https://your-app.com/?env=qa`
- `https://your-app.com/?env=uat`
- `https://your-app.com/?env=prod`

## Configuration File

The `environments.json` file contains:
- `environments`: Object with configuration for each environment
- `urlPatterns`: Object mapping environment names to URL patterns for detection

### Example Structure

```json
{
  "environments": {
    "dev": {
      "REACT_APP_AZURE_CLIENT_ID": "...",
      "REACT_APP_PLACEMENT_API_URL": "...",
      ...
    },
    "qa": { ... },
    "uat": { ... },
    "prod": { ... }
  },
  "urlPatterns": {
    "dev": ["localhost", "127.0.0.1", "..."],
    "qa": ["..."],
    "uat": ["..."],
    "prod": ["..."]
  }
}
```

## Usage in Code

Instead of using `process.env.REACT_APP_*` directly, use the `environment` object:

```typescript
import { environment } from '../config/environment';

// Access environment variables
const apiUrl = environment.PLACEMENT_API_URL;
const clientId = environment.AZURE_CLIENT_ID;
const debugEnabled = environment.DEBUG_ENABLED;
```

## Migration from .env Files

The system maintains backward compatibility:
- If runtime config is not initialized yet, it falls back to `process.env`
- This ensures smooth initialization during app startup
- All existing code using `environment` object will automatically use runtime config once initialized

## Deployment

1. **Build once**: `npm run build`
2. **Deploy the same build** to all environments (dev, qa, uat, prod)
3. **Ensure `environments.json` is included** in the dist folder (handled by webpack)
4. The app will automatically detect and use the correct configuration

## Benefits

✅ **Single Build**: Build once, deploy everywhere  
✅ **No Build-time Secrets**: Configuration is loaded at runtime  
✅ **Easy Environment Management**: All configs in one place  
✅ **URL-based Detection**: Automatic environment detection  
✅ **Backward Compatible**: Falls back to process.env if needed  
✅ **Type Safe**: TypeScript support for all environment variables  

## Troubleshooting

### Config not loading
- Check browser console for errors
- Verify `environments.json` is in the `dist` folder
- Check network tab to see if the file is being fetched

### Wrong environment detected
- Check the URL patterns in `environments.json`
- Use `?env=<environment>` to manually override
- Check browser console for detection logs

### Missing environment variables
- Verify the variable exists in `environments.json` for the detected environment
- Check that the variable name matches exactly (case-sensitive)

