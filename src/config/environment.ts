// Environment configuration for the Outlook add-in
// Uses runtime configuration loaded from environments.json
import runtimeConfig from './runtimeConfig';

// Helper function to get environment value
const getEnvValue = (key: string, defaultValue: string = ''): string => {
  return runtimeConfig.getString(key, defaultValue);
};

const getEnvBoolean = (key: string, defaultValue: boolean = false): boolean => {
  return runtimeConfig.getBoolean(key, defaultValue);
};

export const environment = {
  // Azure AD Configuration
  get AZURE_CLIENT_ID(): string {
    return getEnvValue('REACT_APP_AZURE_CLIENT_ID', '');
  },
  get AZURE_AUTHORITY(): string {
    return getEnvValue('REACT_APP_AZURE_AUTHORITY', '');
  },
  get AZURE_REDIRECT_URI(): string {
    return getEnvValue('REACT_APP_AZURE_REDIRECT_URI', '');
  },
  
  // API Scopes
  get AZURE_API_SCOPES(): string {
    return getEnvValue('REACT_APP_AZURE_API_SCOPES', 'api://d3398715-8435-43df-ac85-d28afd62f0e3/access_as_user');
  },
  get AZURE_GRAPH_SCOPES(): string {
    return getEnvValue('REACT_APP_AZURE_GRAPH_SCOPES', 'https://graph.microsoft.com/Mail.Send');
  },
  get AZURE_GLOBAL_SCOPES(): string {
    return getEnvValue('REACT_APP_AZURE_GLOBAL_SCOPES', 'openid,profile,offline_access');
  },
  
  // API Keys
  get PLACEMENT_API_KEY(): string {
    return getEnvValue('REACT_APP_PLACEMENT_API_KEY', '');
  },
  get LOGGING_API_KEY(): string {
    return getEnvValue('REACT_APP_LOGGING_API_KEY', '');
  },
  
  // Debug Configuration
  get DEBUG_ENABLED(): boolean {
    return getEnvBoolean('REACT_APP_DEBUG_ENABLED', false);
  },
  get DEBUG_LEVEL(): string {
    return getEnvValue('REACT_APP_DEBUG_LEVEL', 'info');
  },
  
  // API Endpoints
  get PLACEMENT_API_URL(): string {
    return getEnvValue('REACT_APP_PLACEMENT_API_URL', '');
  },
  get LOGGING_API_URL(): string {
    return getEnvValue('REACT_APP_LOGGING_API_URL', '');
  },
  
  // Additional environment variables
  get CYBER_MRSNA_MAILBOX(): string {
    return getEnvValue('REACT_APP_CYBER_MRSNA_MAILBOX', '');
  },
  get DEFAULT_SHARED_MAILBOX(): string {
    return getEnvValue('REACT_APP_DEFAULT_SHARED_MAILBOX', '');
  },
  get DRAFT_EMAIL_SUBJECT(): string {
    return getEnvValue('REACT_APP_DRAFT_EMAIL_SUBJECT', '');
  },
};

// Helper function to get scopes as array
export const getScopesArray = (scopesString: string): string[] => {
  return scopesString.split(',').map(scope => scope.trim());
}; 