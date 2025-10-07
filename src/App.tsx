import * as React from 'react';
import { useState, useEffect, useCallback } from 'react';
import WorkbenchLanding from './components/WorkbenchLanding';
import { initializeIcons } from '@fluentui/react/lib/Icons';
import { Spinner, SpinnerSize, MessageBar, MessageBarType, PrimaryButton } from '@fluentui/react';
import AuthService from './service/AuthService';
import './components/SharedGrid.css';

initializeIcons();

export const App: React.FC = () => {
  const [isAuthenticated, setIsAuthenticated] = useState<boolean | null>(null);
  const [apiToken, setApiToken] = useState<string | null>(null);
  const [graphToken, setGraphToken] = useState<string | null>(null);
  const [authError, setAuthError] = useState<string | null>(null);
  const [isRetrying, setIsRetrying] = useState(false);
  const [isInitializing, setIsInitializing] = useState(true);

  const TOKEN_REFRESH_INTERVAL = 5 * 60 * 1000; // 5 minutes

  const refreshTokens = useCallback(async () => {
    try {
      console.log('Refreshing tokens...');
      const { apiToken: newApiToken, graphToken: newGraphToken } = await AuthService.acquireBothTokens();
      
      if (newApiToken?.accessToken && newGraphToken?.accessToken) {
        setApiToken(newApiToken.accessToken);
        setGraphToken(newGraphToken.accessToken);
        console.log('Tokens refreshed successfully');
      } else {
        throw new Error('Failed to refresh tokens');
      }
      
    } catch (error) {
      console.error('Token refresh failed:', error);
      // If refresh fails, try to re-authenticate
      await authenticate();
    }
  }, []);

  const authenticate = async () => {
    try {
      setIsRetrying(false);
      setAuthError(null);
      setIsAuthenticated(null); // Set to loading state
      
      console.log('Starting authentication...');
      
      // Acquire both tokens in a single session to prevent multiple popups
      const { apiToken: apiTokenResult, graphToken: graphTokenResult } = await AuthService.acquireBothTokens();
      
      if (!apiTokenResult || !apiTokenResult.accessToken) {
        throw new Error('Failed to acquire API token');
      }
      
      if (!graphTokenResult || !graphTokenResult.accessToken) {
        throw new Error('Failed to acquire Graph token');
      }
      
      setApiToken(apiTokenResult.accessToken);
      setGraphToken(graphTokenResult.accessToken);
      setIsAuthenticated(true);
      console.log('Authentication successful - both API and Graph tokens acquired');
      
    } catch (error) {
      console.error('Authentication error:', error);
      setAuthError(error instanceof Error ? error.message : 'Authentication failed. Please try again.');
      setIsAuthenticated(false);
    }
  };

  const handleRetry = async () => {
    setIsRetrying(true);
    setIsAuthenticated(null);
    setApiToken(null);
    setGraphToken(null);
    setAuthError(null);
    await authenticate();
  };

  // Set up token auto-refresh
  useEffect(() => {
    if (isAuthenticated && apiToken && graphToken) {
      const interval = setInterval(refreshTokens, TOKEN_REFRESH_INTERVAL);
      console.log('Token auto-refresh interval set up');
      
      return () => {
        clearInterval(interval);
        console.log('Token auto-refresh interval cleared');
      };
    }
  }, [isAuthenticated, apiToken, graphToken, refreshTokens]);

  // Initial setup - check if user is already authenticated
  useEffect(() => {
    const initializeAuth = async () => {
      try {
        setIsInitializing(true);
        
        // Check if user is already authenticated (has valid tokens)
        const { apiToken: cachedApiToken, graphToken: cachedGraphToken } = await AuthService.acquireBothTokens();
        
        if (cachedApiToken?.accessToken && cachedGraphToken?.accessToken) {
          // User already has valid tokens, no need for authentication
          setApiToken(cachedApiToken.accessToken);
          setGraphToken(cachedGraphToken.accessToken);
          setIsAuthenticated(true);
          console.log('User already authenticated with valid tokens');
        } else {
          // User needs to authenticate
          setIsAuthenticated(false);
          console.log('User needs to authenticate');
        }
      } catch (error) {
        console.error('Initial auth check failed:', error);
        // If initial check fails, user needs to authenticate
        setIsAuthenticated(false);
      } finally {
        setIsInitializing(false);
      }
    };

    initializeAuth();
  }, []);

  // Show loading spinner while initializing
  if (isInitializing) {
    return (
      <div className="App" style={{ 
        display: 'flex', 
        justifyContent: 'center', 
        alignItems: 'center', 
        height: '100vh',
        flexDirection: 'column',
        gap: '16px'
      }}>
        <Spinner size={SpinnerSize.large} label="Initializing..." />
        <div style={{ fontSize: '14px', color: '#666', textAlign: 'center' }}>
          <div>Setting up secure connection...</div>
        </div>
      </div>
    );
  }

  // Show loading spinner while authenticating
  if (isAuthenticated === null) {
    return (
      <div className="App" style={{ 
        display: 'flex', 
        justifyContent: 'center', 
        alignItems: 'center', 
        height: '100vh',
        flexDirection: 'column',
        gap: '16px'
      }}>
        <Spinner size={SpinnerSize.large} label="Authenticating..." />
        <div style={{ fontSize: '14px', color: '#666', textAlign: 'center' }}>
          <div>Setting up secure connection...</div>
          <div>Please complete the sign-in process if prompted</div>
        </div>
      </div>
    );
  }

  // Show authentication prompt if user needs to authenticate
  if (isAuthenticated === false) {
    return (
      <div className="App" style={{ padding: '16px' }}>
        {authError ? (
          // Show error message only after user has tried to authenticate
          <>
            <MessageBar messageBarType={MessageBarType.error} isMultiline={true}>
              {authError}
            </MessageBar>
            <div style={{ marginTop: '16px', textAlign: 'center' }}>
              <PrimaryButton 
                onClick={handleRetry} 
                disabled={isRetrying}
                text={isRetrying ? "Retrying..." : "Retry Authentication"}
              />
            </div>
          </>
        ) : (
          // Show initial authentication prompt
          <div style={{ 
            display: 'flex', 
            flexDirection: 'column', 
            alignItems: 'center', 
            gap: '16px',
            padding: '24px',
            textAlign: 'center'
          }}>
            <div style={{ fontSize: '16px', fontWeight: 'bold', color: '#333' }}>
              Welcome to the Workbench Add-in
            </div>
            <div style={{ fontSize: '14px', color: '#666', maxWidth: '300px' }}>
              To use this add-in, you need to sign in with your MunichRe account. 
              This will allow you to submit placement requests and forward emails securely.
            </div>
            <PrimaryButton 
              onClick={authenticate} 
              disabled={isRetrying}
              text={isRetrying ? "Signing In..." : "Sign In"}
              style={{ marginTop: '8px' }}
            />
          </div>
        )}
      </div>
    );
  }

  // Show main UI only after successful authentication with both tokens
  return (
    <div className="App">
      <WorkbenchLanding 
        apiToken={apiToken} 
        graphToken={graphToken}
      />
    </div>
  );
};