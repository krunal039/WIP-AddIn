/**
 * Base API client for making HTTP requests
 * Provides common functionality for API services
 */
import DebugService from './DebugService';
import { ErrorHandler } from '../utils/errorHandler';

export interface RequestOptions {
  method?: 'GET' | 'POST' | 'PUT' | 'DELETE' | 'PATCH';
  headers?: Record<string, string>;
  body?: any;
  timeout?: number;
}

export interface ApiResponse<T = unknown> {
  data: T;
  status: number;
  statusText: string;
  headers: Headers;
}

export class ApiClient {
  /**
   * Make an HTTP request
   */
  protected async request<T = unknown>(
    url: string,
    options: RequestOptions = {}
  ): Promise<ApiResponse<T>> {
    const {
      method = 'GET',
      headers = {},
      body,
      timeout = 30000, // 30 seconds default
    } = options;

    const requestHeaders: Record<string, string> = {
      'Content-Type': 'application/json',
      ...headers,
    };

    const requestOptions: RequestInit = {
      method,
      headers: requestHeaders,
    };

    if (body && method !== 'GET') {
      requestOptions.body = typeof body === 'string' ? body : JSON.stringify(body);
    }

    try {
      DebugService.debug(`[API] ${method} ${url}`);
      
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), timeout);
      
      const response = await fetch(url, {
        ...requestOptions,
        signal: controller.signal,
      });

      clearTimeout(timeoutId);

      if (!response.ok) {
        const errorText = await response.text().catch(() => 'Unknown error');
        throw new Error(`API request failed: ${response.status} ${response.statusText} - ${errorText}`);
      }

      const data = await response.json().catch(() => ({}));

      DebugService.debug(`[API] ${method} ${url} - Success (${response.status})`);

      return {
        data,
        status: response.status,
        statusText: response.statusText,
        headers: response.headers,
      };
    } catch (error) {
      DebugService.error(`[API] ${method} ${url} - Failed`, error);
      await ErrorHandler.handleError(error, 'ApiClient.request', { url, method });
      throw error;
    }
  }

  /**
   * GET request
   */
  protected async get<T = unknown>(url: string, headers?: Record<string, string>): Promise<ApiResponse<T>> {
    return this.request<T>(url, { method: 'GET', headers });
  }

  /**
   * POST request
   */
  protected async post<T = unknown>(
    url: string,
    body?: any,
    headers?: Record<string, string>
  ): Promise<ApiResponse<T>> {
    return this.request<T>(url, { method: 'POST', body, headers });
  }

  /**
   * PUT request
   */
  protected async put<T = unknown>(
    url: string,
    body?: any,
    headers?: Record<string, string>
  ): Promise<ApiResponse<T>> {
    return this.request<T>(url, { method: 'PUT', body, headers });
  }

  /**
   * DELETE request
   */
  protected async delete<T = unknown>(
    url: string,
    headers?: Record<string, string>
  ): Promise<ApiResponse<T>> {
    return this.request<T>(url, { method: 'DELETE', headers });
  }
}

