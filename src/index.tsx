import * as React from 'react';
import { createRoot } from 'react-dom/client';
import { MsalProvider } from '@azure/msal-react';
import { getMsalInstance } from './auth/msalInstance';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import { App } from './App';
import ErrorBoundary from './components/ErrorBoundary';
import LoggingService from './service/LoggingService';
import runtimeConfig from './config/runtimeConfig';
import DebugService from './service/DebugService';

// Initialize runtime configuration before rendering the app
Office.onReady().then(async () => {
  // Use DebugService for all logging - it will respect DEBUG_ENABLED setting
  DebugService.info('Office.onReady triggered, initializing runtime configuration...');
  DebugService.debug('Current URL:', window.location.href);
  DebugService.debug('Hostname:', window.location.hostname);
  
  try {
    // Load environment configuration based on URL
    await runtimeConfig.initialize();
    
    const detectedEnv = runtimeConfig.getEnvironment();
    const config = runtimeConfig.getAll();
    
    // Now that config is loaded, DebugService will respect DEBUG_ENABLED
    DebugService.info(`✅ Configuration loaded - Environment: ${detectedEnv}, Keys: ${Object.keys(config).length}`);

    // Create MSAL instance AFTER config is loaded to ensure correct redirect URI
    DebugService.info('Creating MSAL instance with runtime config...');
    const msalInstance = getMsalInstance();
    DebugService.info('✅ MSAL instance created');
  } catch (error) {
    // Errors should always be logged, even if DEBUG_ENABLED is false
    DebugService.error('❌ Failed to initialize runtime configuration:', error);
    DebugService.error('Error details:', {
      message: error instanceof Error ? error.message : String(error),
      stack: error instanceof Error ? error.stack : undefined,
      currentUrl: window.location.href,
      hostname: window.location.hostname
    });
    // Continue anyway - the app will use fallback values
  }

  // Get MSAL instance (will be created with correct config now)
  const msalInstance = getMsalInstance();
  
  const container = document.getElementById('root');
  if (container) {
    const root = createRoot(container);
    root.render(
      <React.StrictMode>
        <ErrorBoundary>
          <MsalProvider instance={msalInstance}>
            <FluentProvider theme={webLightTheme}>
              <App />
            </FluentProvider>
          </MsalProvider>
        </ErrorBoundary>
      </React.StrictMode>,
    );
  }
});

// Add global error listeners
window.addEventListener('unhandledrejection', (event) => {
  LoggingService.logError(
    event.reason instanceof Error ? event.reason : new Error(String(event.reason)),
    'unhandledrejection'
  );
});

window.addEventListener('error', (event) => {
  LoggingService.logError(
    event.error instanceof Error ? event.error : new Error(String(event.error)),
    'window.onerror'
  );
});
