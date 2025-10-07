// Environment configuration for the Outlook add-in
export const environment = {
  // Azure AD Configuration
  AZURE_CLIENT_ID: process.env.REACT_APP_AZURE_CLIENT_ID || '',
  AZURE_AUTHORITY: process.env.REACT_APP_AZURE_AUTHORITY || '',
  AZURE_REDIRECT_URI: process.env.REACT_APP_AZURE_REDIRECT_URI || '',
  
  // API Scopes
  AZURE_API_SCOPES: process.env.REACT_APP_AZURE_API_SCOPES || 'api://d3398715-8435-43df-ac85-d28afd62f0e3/access_as_user',
  AZURE_GRAPH_SCOPES: process.env.REACT_APP_AZURE_GRAPH_SCOPES || 'https://graph.microsoft.com/Mail.Send',
  AZURE_GLOBAL_SCOPES: process.env.REACT_APP_AZURE_GLOBAL_SCOPES || 'openid,profile,offline_access',
  
  // API Keys
  PLACEMENT_API_KEY: process.env.REACT_APP_PLACEMENT_API_KEY || '',
  LOGGING_API_KEY: process.env.REACT_APP_LOGGING_API_KEY || '',
  
  // Debug Configuration
  DEBUG_ENABLED: process.env.REACT_APP_DEBUG_ENABLED === 'true',
  DEBUG_LEVEL: process.env.REACT_APP_DEBUG_LEVEL || 'info',
  
  // API Endpoints
  PLACEMENT_API_URL: process.env.REACT_APP_PLACEMENT_API_URL || '',
  LOGGING_API_URL: process.env.REACT_APP_LOGGING_API_URL || '',
};

// Helper function to get scopes as array
export const getScopesArray = (scopesString: string): string[] => {
  return scopesString.split(',').map(scope => scope.trim());
}; 