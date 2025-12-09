import { Configuration, LogLevel } from "@azure/msal-browser";
import { environment } from "../config/environment";
import runtimeConfig from "../config/runtimeConfig";
import DebugService from "../service/DebugService";

/**
 * Validates and normalizes a redirect URI
 */
function normalizeRedirectUri(uri: string | undefined): string {
  if (!uri || uri.trim() === '') {
    // Fallback to current origin
    const fallback = window.location.origin;
    DebugService.warn('[MSAL Config] ⚠️ Redirect URI is empty, using fallback:', fallback);
    return fallback;
  }

  const trimmed = uri.trim();
  
  // Remove trailing slash if present (MSAL prefers no trailing slash)
  const normalized = trimmed.endsWith('/') ? trimmed.slice(0, -1) : trimmed;
  
  // Validate URL format
  try {
    const url = new URL(normalized);
    DebugService.debug('[MSAL Config] ✅ Redirect URI validated:', {
      original: uri,
      normalized: normalized,
      protocol: url.protocol,
      hostname: url.hostname,
      port: url.port
    });
    return normalized;
  } catch (error) {
    DebugService.error('[MSAL Config] ❌ Invalid redirect URI format:', {
      uri: uri,
      error: error instanceof Error ? error.message : String(error)
    });
    
    // Try to fix common issues
    let fixed = trimmed;
    if (!fixed.startsWith('http://') && !fixed.startsWith('https://')) {
      // Add protocol if missing
      fixed = window.location.protocol + '//' + fixed;
      DebugService.warn('[MSAL Config] ⚠️ Added missing protocol:', fixed);
    }
    
    try {
      const url = new URL(fixed);
      const normalized = fixed.endsWith('/') ? fixed.slice(0, -1) : fixed;
      DebugService.debug('[MSAL Config] ✅ Fixed redirect URI:', normalized);
      return normalized;
    } catch (fixError) {
      // Last resort: use current origin
      const fallback = window.location.origin;
      DebugService.error('[MSAL Config] ❌ Could not fix redirect URI, using fallback:', fallback);
      return fallback;
    }
  }
}

/**
 * Gets MSAL configuration dynamically from environment
 * This ensures we use the correct values from runtime config
 */
export function getMsalConfig(): Configuration {
  // Check if runtime config is initialized
  const isConfigInitialized = runtimeConfig.isInitialized();
  if (!isConfigInitialized) {
    DebugService.warn('[MSAL Config] ⚠️ Runtime config not initialized yet. Config values may be empty.');
    DebugService.warn('[MSAL Config] Make sure runtimeConfig.initialize() is called before creating MSAL instance.');
  }

  const redirectUriRaw = environment.AZURE_REDIRECT_URI;
  const clientId = environment.AZURE_CLIENT_ID;
  const authority = environment.AZURE_AUTHORITY;

  // Normalize and validate redirect URI
  const redirectUri = normalizeRedirectUri(redirectUriRaw);

  // Log config values for debugging
  DebugService.debug('[MSAL Config] Creating MSAL configuration:', {
    clientId: clientId || '(empty)',
    authority: authority || '(empty)',
    redirectUriRaw: redirectUriRaw || '(empty)',
    redirectUri: redirectUri,
    hasClientId: !!clientId && clientId !== '',
    hasAuthority: !!authority && authority !== '',
    hasRedirectUri: !!redirectUri && redirectUri !== '',
    isConfigInitialized: isConfigInitialized,
    detectedEnvironment: isConfigInitialized ? runtimeConfig.getEnvironment() : 'not loaded'
  });

  // Only log errors if config is initialized and values are still empty (real configuration problem)
  // If config isn't initialized yet, it's just a timing issue and will be resolved
  if (isConfigInitialized) {
    if (!clientId || clientId === '') {
      const detectedEnv = runtimeConfig.getEnvironment();
      const allConfig = runtimeConfig.getAll();
      DebugService.error('[MSAL Config] ❌ ERROR: AZURE_CLIENT_ID is empty!', {
        detectedEnvironment: detectedEnv,
        availableKeys: Object.keys(allConfig),
        hasClientIdKey: 'REACT_APP_AZURE_CLIENT_ID' in allConfig,
        clientIdValue: allConfig.REACT_APP_AZURE_CLIENT_ID
      });
    }
    if (!authority || authority === '') {
      const detectedEnv = runtimeConfig.getEnvironment();
      const allConfig = runtimeConfig.getAll();
      DebugService.error('[MSAL Config] ❌ ERROR: AZURE_AUTHORITY is empty!', {
        detectedEnvironment: detectedEnv,
        availableKeys: Object.keys(allConfig),
        hasAuthorityKey: 'REACT_APP_AZURE_AUTHORITY' in allConfig,
        authorityValue: allConfig.REACT_APP_AZURE_AUTHORITY
      });
    }
  }

  return {
    auth: {
      clientId: clientId || '',
      authority: authority || '',
      redirectUri: redirectUri,
    },
    cache: {
      cacheLocation: "localStorage",
      storeAuthStateInCookie: true, // Enable for better SSO support
    },
    system: {
      allowRedirectInIframe: true, // Allow redirects in Office add-in iframe
      loggerOptions: {
        loggerCallback: (level, message, containsPii) => {
          if (containsPii) return;
          
          // Respect debug environment variables
          const debugEnabled = environment.DEBUG_ENABLED;
          const debugLevel = environment.DEBUG_LEVEL;
          
          // If debug is disabled, only show errors
          if (!debugEnabled && level !== LogLevel.Error) {
            return;
          }
          
          // If debug level is 'error', only show errors
          if (debugLevel === 'error' && level !== LogLevel.Error) {
            return;
          }
          
          // If debug level is 'warn', show warnings and errors
          if (debugLevel === 'warn' && level !== LogLevel.Error && level !== LogLevel.Warning) {
            return;
          }
          
          // Use DebugService for MSAL logging
          switch (level) {
            case LogLevel.Error:
              DebugService.error(`[MSAL] ${message}`);
              break;
            case LogLevel.Info:
              DebugService.info(`[MSAL] ${message}`);
              break;
            case LogLevel.Verbose:
              DebugService.debug(`[MSAL] ${message}`);
              break;
            case LogLevel.Warning:
              DebugService.warn(`[MSAL] ${message}`);
              break;
          }
        },
      },
    },
  };
}

/**
 * Gets login request configuration
 */
export function getLoginRequest() {
  return {
    scopes: environment.AZURE_GRAPH_SCOPES.split(",").map(s => s.trim()),
    prompt: "select_account", // Force account selection for better SSO
  };
}

/**
 * Gets silent token request configuration
 */
export function getSilentRequest() {
  return {
    scopes: environment.AZURE_GRAPH_SCOPES.split(",").map(s => s.trim()),
    forceRefresh: false, // Don't force refresh unless needed
  };
}

// Export for backward compatibility (but will use empty values if called before config loads)
export const loginRequest = getLoginRequest();
export const silentRequest = getSilentRequest();

// Default export for backward compatibility (but should use getMsalConfig() instead)
export default getMsalConfig();