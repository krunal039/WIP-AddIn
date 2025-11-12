/**
 * Centralized error handling utility
 * Provides consistent error handling patterns across the application
 */
import DebugService from '../service/DebugService';
import LoggingService from '../service/LoggingService';

export interface ErrorContext {
  [key: string]: unknown;
}

/**
 * Standard error response format
 */
export interface StandardError {
  message: string;
  code?: string;
  context?: ErrorContext;
  originalError?: Error;
}

/**
 * Handle and log errors consistently
 */
export class ErrorHandler {
  /**
   * Handle an error with logging and optional user notification
   */
  static async handleError(
    error: unknown,
    context: string,
    userContext?: ErrorContext
  ): Promise<StandardError> {
    const errorObj = error instanceof Error ? error : new Error(String(error));
    
    // Log error
    DebugService.error(`[${context}] ${errorObj.message}`, errorObj);
    
    // Log to backend if available
    try {
      // Convert ErrorContext to Record<string, string> for LoggingService
      const logContext = userContext ? Object.fromEntries(
        Object.entries(userContext).map(([k, v]) => [k, String(v)])
      ) : {};
      await LoggingService.logError(errorObj, context, logContext);
    } catch (logError) {
      DebugService.warn('Failed to log error to backend:', logError);
    }

    return {
      message: errorObj.message,
      context: userContext,
      originalError: errorObj,
    };
  }

  /**
   * Create a standardized error from a message
   */
  static createError(message: string, code?: string): StandardError {
    return {
      message,
      code,
    };
  }

  /**
   * Check if error is a network error
   */
  static isNetworkError(error: unknown): boolean {
    if (error instanceof Error) {
      return (
        error.message.includes('fetch') ||
        error.message.includes('network') ||
        error.message.includes('NetworkError') ||
        error.message.includes('Failed to fetch')
      );
    }
    return false;
  }

  /**
   * Check if error is an authentication error
   */
  static isAuthError(error: unknown): boolean {
    if (error instanceof Error) {
      return (
        error.message.includes('authentication') ||
        error.message.includes('unauthorized') ||
        error.message.includes('token') ||
        error.message.includes('401') ||
        error.message.includes('403')
      );
    }
    return false;
  }

  /**
   * Get user-friendly error message
   */
  static getUserMessage(error: StandardError): string {
    if (ErrorHandler.isNetworkError(error.originalError || error)) {
      return 'Network error. Please check your connection and try again.';
    }
    if (ErrorHandler.isAuthError(error.originalError || error)) {
      return 'Authentication error. Please sign in again.';
    }
    return error.message || 'An unexpected error occurred.';
  }
}

