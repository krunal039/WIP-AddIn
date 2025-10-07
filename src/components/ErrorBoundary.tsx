import React from 'react';
import { MessageBar, MessageBarType, PrimaryButton } from '@fluentui/react';
import LoggingService from '../service/LoggingService';

interface ErrorBoundaryProps {
  children: React.ReactNode;
}

interface ErrorBoundaryState {
  hasError: boolean;
  error?: Error;
}

class ErrorBoundary extends React.Component<ErrorBoundaryProps, ErrorBoundaryState> {
  constructor(props: ErrorBoundaryProps) {
    super(props);
    this.state = { hasError: false };
  }

  static getDerivedStateFromError(error: Error): ErrorBoundaryState {
    return { hasError: true, error };
  }

  async componentDidCatch(error: Error, info: React.ErrorInfo) {
    await LoggingService.logError(error, 'ErrorBoundary', { info: JSON.stringify(info) });
  }

  private handleRetry = () => {
    this.setState({ hasError: false, error: undefined });
  };

  render() {
    if (this.state.hasError) {
      return (
        <div style={{ padding: '20px' }}>
          <MessageBar
            messageBarType={MessageBarType.error}
            isMultiline={true}
            styles={{
              root: {
                backgroundColor: "#FBE9E7",
                border: "1px solid #F4F4F4",
                color: "#0F5132",
                fontWeight: "normal",
                marginBottom: 20,
              },
              icon: {
                color: "#D7422F",
                fontSize: 28,
                height: 28,
                width: 28,
                alignSelf: "center",
              },
            }}
          >
            <div>
              <strong>Something went wrong</strong>
              <div>
                An unexpected error occurred. Please try refreshing the page or contact support if the problem persists.
              </div>
              {this.state.error && (
                <div style={{ marginTop: 10, fontSize: '12px', color: '#666' }}>
                  Error: {this.state.error.message}
                </div>
              )}
            </div>
          </MessageBar>
          <PrimaryButton
            text="Try Again"
            onClick={this.handleRetry}
            styles={{
              root: {
                backgroundColor: "#0F1E32",
                borderRadius: 4,
              },
              rootHovered: { backgroundColor: "#0F1E32" },
              rootPressed: { backgroundColor: "#0F1E32" },
            }}
          />
        </div>
      );
    }

    return this.props.children;
  }
}

export default ErrorBoundary; 