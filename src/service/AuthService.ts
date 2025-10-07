import msalInstance from '../auth/msalInstance';
import { AccountInfo, AuthenticationResult } from '@azure/msal-browser';
import DebugService from './DebugService';
import { environment, getScopesArray } from '../config/environment';

class AuthService {
  private static instance: AuthService;
  private apiToken: AuthenticationResult | null = null;
  private graphToken: AuthenticationResult | null = null;
  private apiTokenExpiry: number = 0;
  private graphTokenExpiry: number = 0;
  private refreshBuffer: number = 300; // 5 minutes before expiry to refresh
  private isInitialized: boolean = false;
  
  // Request deduplication
  private apiTokenPromise: Promise<AuthenticationResult | null> | null = null;
  private graphTokenPromise: Promise<AuthenticationResult | null> | null = null;

  // Scopes
  private readonly apiScopes: string[] = getScopesArray(environment.AZURE_API_SCOPES);
  private readonly graphScopes: string[] = getScopesArray(environment.AZURE_GRAPH_SCOPES);

  private constructor() {
    this.initializeMsal();
  }

  private async initializeMsal(): Promise<void> {
    if (this.isInitialized) {
      return;
    }

    try {
      DebugService.auth('Initializing MSAL');
      await msalInstance.initialize();
      this.isInitialized = true;
      DebugService.auth('MSAL initialized successfully');
    } catch (error) {
      DebugService.errorWithStack('Failed to initialize MSAL', error as Error);
      throw error;
    }
  }

  public static getInstance(): AuthService {
    if (!AuthService.instance) {
      AuthService.instance = new AuthService();
    }
    return AuthService.instance;
  }

  private getCurrentAccount(): AccountInfo | null {
    const accounts = msalInstance.getAllAccounts();
    return accounts.length > 0 ? accounts[0] : null;
  }

  private isTokenValid(token: AuthenticationResult | null, expiry: number): boolean {
    if (!token || !token.accessToken) return false;
    const now = Math.floor(Date.now() / 1000);
    return expiry > now + this.refreshBuffer;
  }

  // Check if Office SSO is available
  public isOfficeSSOAvailable(): boolean {
    return !!(Office && Office.context && Office.context.auth);
  }

  // Get token using Office SSO
  public async getOfficeToken(scopes: string[]): Promise<string | null> {
    if (!this.isOfficeSSOAvailable()) {
      return null;
    }

    try {
      DebugService.auth('Attempting Office SSO token acquisition');
      return new Promise((resolve, reject) => {
        Office.context.auth.getAccessTokenAsync({
          scopes: scopes,
          callback: (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              DebugService.auth('Office SSO token acquired successfully');
              resolve(result.value || null);
            } else {
              DebugService.auth(`Office SSO failed: ${result.error?.message || 'Unknown error'}`);
              reject(new Error(result.error?.message || 'Unknown error'));
            }
          }
        });
      });
    } catch (error) {
      DebugService.errorWithStack('Office SSO token acquisition failed', error as Error);
      return null;
    }
  }

  // Ensure user is authenticated (either via Office SSO or MSAL)
  public async ensureAuthenticated(): Promise<boolean> {
    await this.initializeMsal();

    // Try Office SSO first
    if (this.isOfficeSSOAvailable()) {
      try {
        const officeToken = await this.getOfficeToken(this.apiScopes);
        if (officeToken) {
          DebugService.auth('Authentication successful via Office SSO');
          return true;
        }
      } catch (error) {
        DebugService.auth('Office SSO failed, falling back to MSAL');
      }
    }

    // Fallback to MSAL
    const account = this.getCurrentAccount();
    if (account) {
      return true; // Already authenticated
    }

    try {
      DebugService.auth('Starting MSAL authentication');
      await msalInstance.loginPopup({
        scopes: this.apiScopes,
        prompt: "select_account"
      });
      DebugService.auth('MSAL authentication successful');
      return true;
    } catch (error) {
      DebugService.errorWithStack('MSAL authentication failed', error as Error);
      return false;
    }
  }

  // Acquire both tokens in a single session to prevent multiple popups
  public async acquireBothTokens(): Promise<{ apiToken: AuthenticationResult | null; graphToken: AuthenticationResult | null }> {
    await this.initializeMsal();

    // Check if we have valid cached tokens
    const apiTokenValid = this.isTokenValid(this.apiToken, this.apiTokenExpiry);
    const graphTokenValid = this.isTokenValid(this.graphToken, this.graphTokenExpiry);
    
    if (apiTokenValid && graphTokenValid) {
      DebugService.auth('Using cached API and Graph tokens');
      return { apiToken: this.apiToken, graphToken: this.graphToken };
    }

    // Ensure authentication first
    const isAuthenticated = await this.ensureAuthenticated();
    if (!isAuthenticated) {
      console.error('Authentication failed, cannot acquire tokens');
      return { apiToken: null, graphToken: null };
    }

    const account = this.getCurrentAccount();
    if (!account) {
      DebugService.warn('No account found after authentication');
      return { apiToken: null, graphToken: null };
    }

    // Try to acquire both tokens silently first
    let apiToken = null;
    let graphToken = null;

    try {
      // Try API token first
      if (!apiTokenValid) {
        DebugService.auth('Attempting silent API token acquisition');
        apiToken = await msalInstance.acquireTokenSilent({
          account: account,
          scopes: this.apiScopes,
          forceRefresh: false
        });
        this.apiToken = apiToken;
        this.apiTokenExpiry = apiToken.expiresOn ? Math.floor(apiToken.expiresOn.getTime() / 1000) : Math.floor(Date.now() / 1000) + 3600;
        DebugService.auth('API token acquired silently');
      } else {
        apiToken = this.apiToken;
      }

      // Try Graph token
      if (!graphTokenValid) {
        DebugService.auth('Attempting silent Graph token acquisition');
        graphToken = await msalInstance.acquireTokenSilent({
          account: account,
          scopes: this.graphScopes,
          forceRefresh: false
        });
        this.graphToken = graphToken;
        this.graphTokenExpiry = graphToken.expiresOn ? Math.floor(graphToken.expiresOn.getTime() / 1000) : Math.floor(Date.now() / 1000) + 3600;
        DebugService.auth('Graph token acquired silently');
      } else {
        graphToken = this.graphToken;
      }

      return { apiToken, graphToken };
    } catch (error) {
      DebugService.errorWithStack('Silent token acquisition failed', error as Error);
      
      // If silent acquisition fails, we need to show a popup
      // Use the combined scopes to get both tokens in one popup
      if (error instanceof Error && error.message.includes('interaction_required')) {
        try {
          DebugService.auth('Showing popup for combined token acquisition');
          const combinedScopes = [...this.apiScopes, ...this.graphScopes];
          const result = await msalInstance.loginPopup({
            scopes: combinedScopes,
            prompt: "select_account"
          });
          
          // Store the result for both tokens (they'll have the same access token but different scopes)
          this.apiToken = result;
          this.graphToken = result;
          this.apiTokenExpiry = result.expiresOn ? Math.floor(result.expiresOn.getTime() / 1000) : Math.floor(Date.now() / 1000) + 3600;
          this.graphTokenExpiry = this.apiTokenExpiry;
          
          DebugService.auth('Both tokens acquired via single popup');
          return { apiToken: result, graphToken: result };
        } catch (popupError) {
          DebugService.errorWithStack('Popup login failed for combined token acquisition', popupError as Error);
          return { apiToken: null, graphToken: null };
        }
      }
      
      return { apiToken: null, graphToken: null };
    }
  }

  public async getApiToken(): Promise<AuthenticationResult | null> {
    await this.initializeMsal();

    // Check if we have a valid cached token
    if (this.isTokenValid(this.apiToken, this.apiTokenExpiry)) {
      DebugService.auth('Using cached API token');
      return this.apiToken;
    }

    // Check if there's already a request in progress
    if (this.apiTokenPromise) {
      DebugService.auth('API token request already in progress, waiting...');
      return this.apiTokenPromise;
    }

    // Ensure authentication first
    const isAuthenticated = await this.ensureAuthenticated();
    if (!isAuthenticated) {
      console.error('Authentication failed, cannot get API token');
      return null;
    }

    // Start new request
    DebugService.auth('Starting new API token request');
    this.apiTokenPromise = this.acquireApiToken();
    
    try {
      const result = await this.apiTokenPromise;
      return result;
    } finally {
      this.apiTokenPromise = null;
    }
  }

  public async getGraphToken(): Promise<AuthenticationResult | null> {
    await this.initializeMsal();

    // Check if we have a valid cached token
    if (this.isTokenValid(this.graphToken, this.graphTokenExpiry)) {
      DebugService.auth('Using cached Graph token');
      return this.graphToken;
    }

    // Check if there's already a request in progress
    if (this.graphTokenPromise) {
      DebugService.auth('Graph token request already in progress, waiting...');
      return this.graphTokenPromise;
    }

    // Ensure authentication first
    const isAuthenticated = await this.ensureAuthenticated();
    if (!isAuthenticated) {
      console.error('Authentication failed, cannot get Graph token');
      return null;
    }

    // Start new request
    DebugService.auth('Starting new Graph token request');
    this.graphTokenPromise = this.acquireGraphToken();
    
    try {
      const result = await this.graphTokenPromise;
      return result;
    } finally {
      this.graphTokenPromise = null;
    }
  }

  private async acquireApiToken(): Promise<AuthenticationResult | null> {
    try {
      const account = this.getCurrentAccount();
      if (!account) {
        DebugService.warn('No account found after authentication');
        return null;
      }

      // Try silent acquisition first
      DebugService.auth('Attempting silent API token acquisition');
      const result = await msalInstance.acquireTokenSilent({
        account: account,
        scopes: this.apiScopes,
        forceRefresh: false
      });

      this.apiToken = result;
      this.apiTokenExpiry = result.expiresOn ? Math.floor(result.expiresOn.getTime() / 1000) : Math.floor(Date.now() / 1000) + 3600;
      DebugService.auth('API token acquired successfully');
      return result;
    } catch (error) {
      DebugService.errorWithStack('Silent token acquisition failed for API token', error as Error);
      
      // Only show popup if it's an interaction required error
      if (error instanceof Error && error.message.includes('interaction_required')) {
        try {
          DebugService.auth('Showing popup for API token');
          const result = await msalInstance.loginPopup({
            scopes: this.apiScopes,
            prompt: "select_account"
          });
          
          this.apiToken = result;
          this.apiTokenExpiry = result.expiresOn ? Math.floor(result.expiresOn.getTime() / 1000) : Math.floor(Date.now() / 1000) + 3600;
          DebugService.auth('API token acquired via popup');
          return result;
        } catch (popupError) {
          DebugService.errorWithStack('Popup login failed for API token', popupError as Error);
          return null;
        }
      }
      
      return null;
    }
  }

  private async acquireGraphToken(): Promise<AuthenticationResult | null> {
    try {
      const account = this.getCurrentAccount();
      if (!account) {
        DebugService.warn('No account found after authentication');
        return null;
      }

      // Try silent acquisition first
      DebugService.auth('Attempting silent Graph token acquisition');
      const result = await msalInstance.acquireTokenSilent({
        account: account,
        scopes: this.graphScopes,
        forceRefresh: false
      });

      this.graphToken = result;
      this.graphTokenExpiry = result.expiresOn ? Math.floor(result.expiresOn.getTime() / 1000) : Math.floor(Date.now() / 1000) + 3600;
      DebugService.auth('Graph token acquired silently');
      return result;
    } catch (error) {
      DebugService.errorWithStack('Silent token acquisition failed for Graph token', error as Error);
      
      // Only show popup if it's an interaction required error
      if (error instanceof Error && error.message.includes('interaction_required')) {
        try {
          DebugService.auth('Showing popup for Graph token');
          const result = await msalInstance.loginPopup({
            scopes: this.graphScopes,
            prompt: "select_account"
          });
          
          this.graphToken = result;
          this.graphTokenExpiry = result.expiresOn ? Math.floor(result.expiresOn.getTime() / 1000) : Math.floor(Date.now() / 1000) + 3600;
          DebugService.auth('Graph token acquired via popup');
          return result;
        } catch (popupError) {
          DebugService.errorWithStack('Popup login failed for Graph token', popupError as Error);
          return null;
        }
      }
      
      return null;
    }
  }

  // Method to clear cached tokens
  public clearTokens(): void {
    this.apiToken = null;
    this.graphToken = null;
    this.apiTokenExpiry = 0;
    this.graphTokenExpiry = 0;
    this.apiTokenPromise = null;
    this.graphTokenPromise = null;
  }

  // Get authentication status
  public isAuthenticated(): boolean {
    const account = this.getCurrentAccount();
    return !!account;
  }

  // Get token info for debugging
  public getTokenInfo(): { apiTokenValid: boolean; graphTokenValid: boolean; apiExpiry: number; graphExpiry: number } {
    return {
      apiTokenValid: this.isTokenValid(this.apiToken, this.apiTokenExpiry),
      graphTokenValid: this.isTokenValid(this.graphToken, this.graphTokenExpiry),
      apiExpiry: this.apiTokenExpiry,
      graphExpiry: this.graphTokenExpiry
    };
  }
}

export default AuthService.getInstance(); 