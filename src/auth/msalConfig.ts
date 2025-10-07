import { Configuration, LogLevel } from "@azure/msal-browser";

const msalConfig: Configuration = {
  auth: {
    clientId: process.env.REACT_APP_AZURE_CLIENT_ID!,
    authority: process.env.REACT_APP_AZURE_AUTHORITY!,
    redirectUri: process.env.REACT_APP_AZURE_REDIRECT_URI!,
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
        const debugEnabled = process.env.REACT_APP_DEBUG_ENABLED === 'true';
        const debugLevel = process.env.REACT_APP_DEBUG_LEVEL || 'info';
        
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
        
        switch (level) {
          case LogLevel.Error:
            console.error(`[MSAL] ${message}`);
            break;
          case LogLevel.Info:
            console.info(`[MSAL] ${message}`);
            break;
          case LogLevel.Verbose:
            console.debug(`[MSAL] ${message}`);
            break;
          case LogLevel.Warning:
            console.warn(`[MSAL] ${message}`);
            break;
        }
      },
    },
  },
};

export const loginRequest = {
  scopes: (process.env.REACT_APP_AZURE_SCOPES || "User.Read").split(","),
  prompt: "select_account", // Force account selection for better SSO
};

// Silent token request configuration
export const silentRequest = {
  scopes: (process.env.REACT_APP_AZURE_SCOPES || "User.Read").split(","),
  forceRefresh: false, // Don't force refresh unless needed
};

export default msalConfig;