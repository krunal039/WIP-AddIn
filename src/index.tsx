import * as React from 'react';
import { createRoot } from 'react-dom/client';
import { MsalProvider } from '@azure/msal-react';
import msalInstance from './auth/msalInstance';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import { App } from './App';
import ErrorBoundary from './components/ErrorBoundary';
import LoggingService from './service/LoggingService';

Office.onReady().then(() => {
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
