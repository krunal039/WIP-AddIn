import { getMsalInstance } from '../auth/msalInstance';
import { AccountInfo, AuthenticationResult } from '@azure/msal-browser';
import DebugService from './DebugService';
import { environment, getScopesArray } from '../config/environment';
import runtimeConfig from '../config/runtimeConfig';

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
  
  // Interaction lock to prevent concurrent popups
  private interactionInProgress: boolean = false;
  private interactionPromise: Promise<boolean> | null = null;

  // Scopes - get dynamically to ensure runtime config is loaded
  private get apiScopes(): string[] {
    return getScopesArray(environment.AZURE_API_SCOPES);
  }
  
  private get graphScopes(): string[] {
    return getScopesArray(environment.AZURE_GRAPH_SCOPES);
  }

  private constructor() {
    // Don't initialize MSAL immediately - wait for runtime config
    // MSAL will be initialized lazily when first needed
  }

  private async initializeMsal(): Promise<void> {
    if (this.isInitialized) {
      return;
    }

    try {
      // Wait for runtime config to be initialized before creating MSAL instance
      if (!runtimeConfig.isInitialized()) {
        DebugService.auth('Waiting for runtime config to initialize...');
        // Wait a bit and check again (config should be loading)
        let attempts = 0;
        const maxAttempts = 50; // 5 seconds max wait
        while (!runtimeConfig.isInitialized() && attempts < maxAttempts) {
          await new Promise(resolve => setTimeout(resolve, 100));
          attempts++;
        }
        
        if (!runtimeConfig.isInitialized()) {
          throw new Error('Runtime config not initialized after waiting. Make sure runtimeConfig.initialize() is called before using AuthService.');
        }
        DebugService.auth('Runtime config initialized, proceeding with MSAL initialization');
      }

      DebugService.auth('Initializing MSAL');
      const msalInstance = getMsalInstance();
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
    const msalInstance = getMsalInstance();
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

    // Check if there's already an interaction in progress
    if (this.interactionInProgress && this.interactionPromise) {
      DebugService.auth('Waiting for existing interaction to complete...');
      try {
        const result = await this.interactionPromise;
        return result;
      } catch (error) {
        DebugService.warn('Existing interaction failed, starting new one');
      }
    }

    // Start new interaction
    this.interactionInProgress = true;
    this.interactionPromise = this.performMsalLogin();

    try {
      const result = await this.interactionPromise;
      return result;
    } finally {
      this.interactionInProgress = false;
      this.interactionPromise = null;
    }
  }

  /**
   * Performs MSAL login popup with proper error handling
   */
  private async performMsalLogin(): Promise<boolean> {
    try {
      DebugService.auth('Starting MSAL authentication');
      const msalInstance = getMsalInstance();
      
      // Log the current config to debug redirect URI issues
      const config = msalInstance.getConfiguration();
      DebugService.debug('[AuthService] MSAL config before loginPopup:', {
        clientId: config.auth?.clientId,
        authority: config.auth?.authority,
        redirectUri: config.auth?.redirectUri,
        redirectUriType: typeof config.auth?.redirectUri
      });
      
      await msalInstance.loginPopup({
        scopes: this.apiScopes,
        prompt: "select_account"
      });
      DebugService.auth('MSAL authentication successful');
      return true;
    } catch (error) {
      const msalInstance = getMsalInstance();
      const config = msalInstance.getConfiguration();
      
      // Handle interaction_in_progress error
      if (error instanceof Error && error.message.includes('interaction_in_progress')) {
        DebugService.warn('MSAL interaction already in progress, waiting...');
        
        // Wait a bit and check if we have an account now
        await new Promise(resolve => setTimeout(resolve, 1000));
        const account = this.getCurrentAccount();
        if (account) {
          DebugService.auth('Account found after waiting for interaction');
          return true;
        }
        
        // If still no account, try to handle the error
        DebugService.warn('No account after interaction_in_progress, may need user to complete existing popup');
        return false;
      }
      
      DebugService.errorWithStack('MSAL authentication failed', error as Error);
      DebugService.error('[AuthService] MSAL config at error time:', {
        clientId: config.auth?.clientId,
        authority: config.auth?.authority,
        redirectUri: config.auth?.redirectUri,
        redirectUriType: typeof config.auth?.redirectUri,
        error: error instanceof Error ? error.message : String(error)
      });
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
        DebugService.error('Authentication failed, cannot acquire tokens');
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

    const msalInstance = getMsalInstance();
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
      if (error instanceof Error && (error.message.includes('interaction_required') || error.message.includes('interaction_in_progress'))) {
        // Check if interaction is already in progress
        if (this.interactionInProgress && this.interactionPromise) {
          DebugService.auth('Waiting for existing interaction to complete before token acquisition');
          try {
            await this.interactionPromise;
            // Retry silent acquisition after interaction completes
            const account = this.getCurrentAccount();
            if (account) {
              try {
                apiToken = await msalInstance.acquireTokenSilent({
                  account: account,
                  scopes: this.apiScopes,
                  forceRefresh: false
                });
                graphToken = await msalInstance.acquireTokenSilent({
                  account: account,
                  scopes: this.graphScopes,
                  forceRefresh: false
                });
                this.apiToken = apiToken;
                this.graphToken = graphToken;
                this.apiTokenExpiry = apiToken.expiresOn ? Math.floor(apiToken.expiresOn.getTime() / 1000) : Math.floor(Date.now() / 1000) + 3600;
                this.graphTokenExpiry = graphToken.expiresOn ? Math.floor(graphToken.expiresOn.getTime() / 1000) : Math.floor(Date.now() / 1000) + 3600;
                return { apiToken, graphToken };
              } catch (retryError) {
                DebugService.warn('Silent acquisition failed after interaction, will show popup');
              }
            }
          } catch (waitError) {
            DebugService.warn('Error waiting for interaction:', waitError);
          }
        }
        
        try {
          DebugService.auth('Showing popup for combined token acquisition');
          const combinedScopes = [...this.apiScopes, ...this.graphScopes];
          const msalInstance = getMsalInstance();
          
          // Set interaction flag
          this.interactionInProgress = true;
          const result = await msalInstance.loginPopup({
            scopes: combinedScopes,
            prompt: "select_account"
          });
          this.interactionInProgress = false;
          
          // Store the result for both tokens (they'll have the same access token but different scopes)
          this.apiToken = result;
          this.graphToken = result;
          this.apiTokenExpiry = result.expiresOn ? Math.floor(result.expiresOn.getTime() / 1000) : Math.floor(Date.now() / 1000) + 3600;
          this.graphTokenExpiry = this.apiTokenExpiry;
          
          DebugService.auth('Both tokens acquired via single popup');
          return { apiToken: result, graphToken: result };
        } catch (popupError) {
          this.interactionInProgress = false;
          if (popupError instanceof Error && popupError.message.includes('interaction_in_progress')) {
            DebugService.warn('Popup blocked due to interaction_in_progress, user may need to complete existing popup');
            // Wait and check for account
            await new Promise(resolve => setTimeout(resolve, 1000));
            const account = this.getCurrentAccount();
            if (account) {
              try {
                const apiToken = await msalInstance.acquireTokenSilent({
                  account: account,
                  scopes: this.apiScopes,
                  forceRefresh: false
                });
                const graphToken = await msalInstance.acquireTokenSilent({
                  account: account,
                  scopes: this.graphScopes,
                  forceRefresh: false
                });
                return { apiToken, graphToken };
              } catch (silentError) {
                DebugService.errorWithStack('Silent acquisition failed after interaction_in_progress', silentError as Error);
              }
            }
          }
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
      DebugService.error('Authentication failed, cannot get API token');
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
      DebugService.error('Authentication failed, cannot get Graph token');
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
      const msalInstance = getMsalInstance();
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
      if (error instanceof Error && (error.message.includes('interaction_required') || error.message.includes('interaction_in_progress'))) {
        // Check if interaction is already in progress
        if (this.interactionInProgress && this.interactionPromise) {
          DebugService.auth('Waiting for existing interaction before API token popup');
          try {
            await this.interactionPromise;
            const account = this.getCurrentAccount();
            if (account) {
              try {
                const msalInstance = getMsalInstance();
                const result = await msalInstance.acquireTokenSilent({
                  account: account,
                  scopes: this.apiScopes,
                  forceRefresh: false
                });
                this.apiToken = result;
                this.apiTokenExpiry = result.expiresOn ? Math.floor(result.expiresOn.getTime() / 1000) : Math.floor(Date.now() / 1000) + 3600;
                return result;
              } catch (retryError) {
                DebugService.warn('Silent acquisition failed after interaction');
              }
            }
          } catch (waitError) {
            DebugService.warn('Error waiting for interaction');
          }
        }
        
        try {
          DebugService.auth('Showing popup for API token');
          const msalInstance = getMsalInstance();
          this.interactionInProgress = true;
          const result = await msalInstance.loginPopup({
            scopes: this.apiScopes,
            prompt: "select_account"
          });
          this.interactionInProgress = false;
          
          this.apiToken = result;
          this.apiTokenExpiry = result.expiresOn ? Math.floor(result.expiresOn.getTime() / 1000) : Math.floor(Date.now() / 1000) + 3600;
          DebugService.auth('API token acquired via popup');
          return result;
        } catch (popupError) {
          this.interactionInProgress = false;
          if (popupError instanceof Error && popupError.message.includes('interaction_in_progress')) {
            DebugService.warn('API token popup blocked due to interaction_in_progress');
            await new Promise(resolve => setTimeout(resolve, 1000));
            const account = this.getCurrentAccount();
            if (account) {
              try {
                const msalInstance = getMsalInstance();
                const result = await msalInstance.acquireTokenSilent({
                  account: account,
                  scopes: this.apiScopes,
                  forceRefresh: false
                });
                return result;
              } catch (silentError) {
                DebugService.errorWithStack('Silent acquisition failed after interaction_in_progress', silentError as Error);
              }
            }
          }
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
      const msalInstance = getMsalInstance();
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
      if (error instanceof Error && (error.message.includes('interaction_required') || error.message.includes('interaction_in_progress'))) {
        // Check if interaction is already in progress
        if (this.interactionInProgress && this.interactionPromise) {
          DebugService.auth('Waiting for existing interaction before Graph token popup');
          try {
            await this.interactionPromise;
            const account = this.getCurrentAccount();
            if (account) {
              try {
                const msalInstance = getMsalInstance();
                const result = await msalInstance.acquireTokenSilent({
                  account: account,
                  scopes: this.graphScopes,
                  forceRefresh: false
                });
                this.graphToken = result;
                this.graphTokenExpiry = result.expiresOn ? Math.floor(result.expiresOn.getTime() / 1000) : Math.floor(Date.now() / 1000) + 3600;
                return result;
              } catch (retryError) {
                DebugService.warn('Silent acquisition failed after interaction');
              }
            }
          } catch (waitError) {
            DebugService.warn('Error waiting for interaction');
          }
        }
        
        try {
          DebugService.auth('Showing popup for Graph token');
          const msalInstance = getMsalInstance();
          this.interactionInProgress = true;
          const result = await msalInstance.loginPopup({
            scopes: this.graphScopes,
            prompt: "select_account"
          });
          this.interactionInProgress = false;
          
          this.graphToken = result;
          this.graphTokenExpiry = result.expiresOn ? Math.floor(result.expiresOn.getTime() / 1000) : Math.floor(Date.now() / 1000) + 3600;
          DebugService.auth('Graph token acquired via popup');
          return result;
        } catch (popupError) {
          this.interactionInProgress = false;
          if (popupError instanceof Error && popupError.message.includes('interaction_in_progress')) {
            DebugService.warn('Graph token popup blocked due to interaction_in_progress');
            await new Promise(resolve => setTimeout(resolve, 1000));
            const account = this.getCurrentAccount();
            if (account) {
              try {
                const msalInstance = getMsalInstance();
                const result = await msalInstance.acquireTokenSilent({
                  account: account,
                  scopes: this.graphScopes,
                  forceRefresh: false
                });
                return result;
              } catch (silentError) {
                DebugService.errorWithStack('Silent acquisition failed after interaction_in_progress', silentError as Error);
              }
            }
          }
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

  // Method to clear interaction state (useful if popup is blocked or closed)
  public clearInteractionState(): void {
    this.interactionInProgress = false;
    this.interactionPromise = null;
    DebugService.auth('Interaction state cleared');
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