// src/auth/msalInstance.ts
import { PublicClientApplication } from "@azure/msal-browser";
import { getMsalConfig } from "./msalConfig";
import runtimeConfig from "../config/runtimeConfig";
import DebugService from "../service/DebugService";

// Create MSAL instance lazily to ensure config is loaded first
let msalInstance: PublicClientApplication | null = null;

/**
 * Gets or creates the MSAL instance
 * This ensures the instance is created with the correct config from runtime environment
 */
export function getMsalInstance(): PublicClientApplication {
  if (!msalInstance) {
    // Check if runtime config is initialized
    if (!runtimeConfig.isInitialized()) {
      const error = new Error('Runtime config not initialized. Call runtimeConfig.initialize() before creating MSAL instance.');
      DebugService.error('[MSAL Instance] ❌', error.message);
      DebugService.error('[MSAL Instance] Current state:', {
        isInitialized: runtimeConfig.isInitialized(),
        currentUrl: window.location.href,
        hostname: window.location.hostname
      });
      throw error;
    }

    try {
      const config = getMsalConfig();
      
      // Validate config before creating instance
      if (!config.auth.clientId || config.auth.clientId === '') {
        throw new Error('MSAL clientId is empty or invalid. Check that REACT_APP_AZURE_CLIENT_ID is set in environments.json');
      }
      if (!config.auth.authority || config.auth.authority === '') {
        throw new Error('MSAL authority is empty or invalid. Check that REACT_APP_AZURE_AUTHORITY is set in environments.json');
      }
      if (!config.auth.redirectUri || config.auth.redirectUri === '') {
        throw new Error('MSAL redirectUri is empty or invalid. Check that REACT_APP_AZURE_REDIRECT_URI is set in environments.json');
      }
      
      // Validate redirect URI is a valid URL
      try {
        new URL(config.auth.redirectUri);
      } catch (urlError) {
        throw new Error(`MSAL redirectUri is not a valid URL: ${config.auth.redirectUri}. Error: ${urlError instanceof Error ? urlError.message : String(urlError)}`);
      }
      
      DebugService.debug('[MSAL Instance] Creating PublicClientApplication with validated config:', {
        clientId: config.auth.clientId,
        authority: config.auth.authority,
        redirectUri: config.auth.redirectUri,
        redirectUriType: typeof config.auth.redirectUri,
        redirectUriLength: config.auth.redirectUri.length
      });
      
      msalInstance = new PublicClientApplication(config);
      DebugService.info('[MSAL Instance] ✅ PublicClientApplication created successfully');
    } catch (error) {
      DebugService.error('[MSAL Instance] ❌ Failed to create PublicClientApplication:', error);
      DebugService.error('[MSAL Instance] Error details:', {
        message: error instanceof Error ? error.message : String(error),
        stack: error instanceof Error ? error.stack : undefined
      });
      throw error;
    }
  }
  return msalInstance;
}

// Default export for backward compatibility (returns function, doesn't call it)
// This prevents immediate execution during module import
export default getMsalInstance;