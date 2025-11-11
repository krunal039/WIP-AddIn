import AuthService from './AuthService';
import { environment } from '../config/environment';
import DebugService from './DebugService';

export type SeverityLevel = 'Information' | 'Verbose' | 'Event' | 'Exception';

class LoggingService {
  private static instance: LoggingService;

  private constructor() {
    // Don't initialize values here - read them dynamically when needed
  }

  /**
   * Gets the logging API URL dynamically from environment config
   */
  private getBaseUrl(): string {
    const url = environment.LOGGING_API_URL || '';
    if (!url) {
      DebugService.warn('[LoggingService] LOGGING_API_URL is empty, logging will be disabled');
    }
    return url;
  }

  /**
   * Gets the logging API key dynamically from environment config
   */
  private getApiKey(): string {
    const key = environment.LOGGING_API_KEY || '';
    if (!key) {
      DebugService.warn('[LoggingService] LOGGING_API_KEY is empty');
    }
    return key;
  }

  public static getInstance(): LoggingService {
    if (!LoggingService.instance) {
      LoggingService.instance = new LoggingService();
    }
    return LoggingService.instance;
  }

  private async getAuthHeaders(): Promise<Record<string, string>> {
    const headers: Record<string, string> = {
      'Content-Type': 'application/json',
      'Ocp-Apim-Subscription-Key': this.getApiKey(),
    };

    try {
      // Get API token for authorization
      const apiTokenResult = await AuthService.getApiToken();
      if (apiTokenResult?.accessToken) {
        headers['Authorization'] = `Bearer ${apiTokenResult.accessToken}`;
      }
    } catch (error) {
      // If token acquisition fails, log without authorization
      // This ensures logging doesn't break the application
      DebugService.warn('Failed to get API token for logging:', error);
    }

    return headers;
  }

  private async postLog(logTable: 'trace' | 'event' | 'exception', body: any) {
    const baseUrl = this.getBaseUrl();
    if (!baseUrl) return;
    
    try {
      const headers = await this.getAuthHeaders();
      
      await fetch(`${baseUrl}/${logTable}`, {
        method: 'POST',
        headers,
        body: JSON.stringify(body),
      });
    } catch (err) {
      // Swallow errors to avoid breaking app flow
      // Log errors through DebugService
      DebugService.error('LoggingService error:', err);
    }
  }

  public async logTrace(message: string, severityLevel: SeverityLevel = 'Information', properties: Record<string, string> = {}) {
    await this.postLog('trace', { severityLevel, message, properties });
  }

  public async logEvent(trackingName: string, properties: Record<string, string> = {}, metrics: Record<string, number> = {}) {
    await this.postLog('event', { trackingName, properties, metrics });
  }

  public async logException(message: string, properties: Record<string, string> = {}, metrics: Record<string, number> = {}) {
    await this.postLog('exception', { message, properties, metrics });
  }

  // Example wrappers for common app actions
  public async logUserToken(userId: string, tokenType: string) {
    await this.logEvent('UserTokenAcquired', { userId, tokenType });
  }

  public async logApiToken(token: string) {
    await this.logTrace('API token acquired', 'Information', { token });
  }

  public async logPlacementRequest(placementId: string, userId: string) {
    await this.logEvent('PlacementRequest', { placementId, userId });
  }

  public async logEmailStamped(placementId: string, emailId: string) {
    await this.logTrace('Email stamped with placementId', 'Information', { placementId, emailId });
  }

  public async logEmailForwarded(placementId: string, emailId: string, mailbox: string) {
    await this.logEvent('EmailForwarded', { placementId, emailId, mailbox });
  }

  public async logError(error: Error, context: string, extra: Record<string, string> = {}) {
    await this.logException(error.message, { context, ...extra });
  }
}

export default LoggingService.getInstance(); 