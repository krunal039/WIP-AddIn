/**
 * Runtime Configuration Loader
 * 
 * This module loads environment configuration at runtime based on the current URL.
 * It allows a single build to work across multiple environments (dev, qa, uat, prod).
 */

export interface EnvironmentConfig {
  [key: string]: string | number | boolean;
}

export interface EnvironmentsData {
  environments: {
    [env: string]: EnvironmentConfig;
  };
  urlPatterns: {
    [env: string]: string[];
  };
}

import DebugService from '../service/DebugService';

class RuntimeConfig {
  private static instance: RuntimeConfig;
  private config: EnvironmentConfig | null = null;
  private environment: string = 'dev';
  private initialized: boolean = false;
  private initPromise: Promise<void> | null = null;

  private constructor() {}

  /**
   * Logs a message using DebugService
   * DebugService handles DEBUG_ENABLED checks internally
   */
  private log(level: 'info' | 'debug' | 'warn' | 'error', message: string, ...args: any[]): void {
    try {
      switch (level) {
        case 'info':
          DebugService.info(message, ...args);
          break;
        case 'debug':
          DebugService.debug(message, ...args);
          break;
        case 'warn':
          DebugService.warn(message, ...args);
          break;
        case 'error':
          DebugService.error(message, ...args);
          break;
      }
    } catch (error) {
      // Fallback to console only for errors if DebugService fails
      if (level === 'error') {
        console.error(`[RuntimeConfig] ${message}`, ...args);
      }
    }
  }

  public static getInstance(): RuntimeConfig {
    if (!RuntimeConfig.instance) {
      RuntimeConfig.instance = new RuntimeConfig();
    }
    return RuntimeConfig.instance;
  }

  /**
   * Detects the environment based on the current URL
   * @param urlPatterns URL patterns from environments.json (required)
   */
  private detectEnvironment(urlPatterns: { [key: string]: string[] }): string {
    const hostname = window.location.hostname.toLowerCase();
    
    // Check for explicit environment in URL search params (for testing)
    const urlParams = new URLSearchParams(window.location.search);
    const envParam = urlParams.get('env');
    if (envParam && ['dev', 'qa', 'uat', 'prod', 'localhost'].includes(envParam)) {
      this.log('info', `Environment forced via URL param: ${envParam}`);
      return envParam;
    }

    // Check each environment's URL patterns
    for (const [env, envPatterns] of Object.entries(urlPatterns)) {
      for (const pattern of envPatterns) {
        if (hostname.includes(pattern.toLowerCase())) {
          this.log('info', `✅ Detected environment: ${env}`);
          return env;
        }
      }
    }

    // Default to localhost for local development, then dev
    const defaultEnv = hostname === 'localhost' || hostname === '127.0.0.1' ? 'localhost' : 'dev';
    this.log('warn', `⚠️ Could not detect environment from hostname "${hostname}", defaulting to ${defaultEnv}`);
    return defaultEnv;
  }

  /**
   * Loads the environments.json file and extracts the config for the detected environment
   */
  private async loadConfig(): Promise<void> {
    this.log('info', '🔄 Loading runtime configuration...');
    
    const response = await fetch('/environments.json');
    if (!response.ok) {
      throw new Error(`Failed to load environments.json: ${response.status} ${response.statusText}`);
    }

    const data: EnvironmentsData = await response.json();
    
    // Validate urlPatterns exists
    if (!data.urlPatterns || Object.keys(data.urlPatterns).length === 0) {
      throw new Error('environments.json is missing urlPatterns');
    }
    
    // Detect environment using URL patterns from the loaded file
    this.environment = this.detectEnvironment(data.urlPatterns);

    // Get the config for the detected environment
    if (!data.environments[this.environment]) {
      throw new Error(`Environment "${this.environment}" not found in environments.json. Available: ${Object.keys(data.environments).join(', ')}`);
    }

    this.config = data.environments[this.environment];
    this.initialized = true;

    // Validate required keys
    const requiredKeys = [
      'REACT_APP_AZURE_CLIENT_ID',
      'REACT_APP_AZURE_AUTHORITY',
      'REACT_APP_AZURE_REDIRECT_URI',
      'REACT_APP_PLACEMENT_API_URL',
      'REACT_APP_LOGGING_API_URL'
    ];
    
    const missingKeys = requiredKeys.filter(key => !(key in this.config!));
    if (missingKeys.length > 0) {
      this.log('warn', '⚠️ Missing required configuration keys:', missingKeys);
    }

    this.log('info', `✅ Configuration loaded - Environment: ${this.environment}, Keys: ${Object.keys(this.config).length}`);
  }

  /**
   * Initializes the configuration (loads environments.json)
   * This should be called ONCE before the app starts
   * Uses singleton pattern to ensure only one load happens
   */
  public async initialize(): Promise<void> {
    // If already initialized, return immediately
    if (this.initialized) {
      return Promise.resolve();
    }

    // If initialization is in progress, wait for it
    if (this.initPromise) {
      return this.initPromise;
    }

    // Start the one-time load
    this.initPromise = this.loadConfig();
    return this.initPromise;
  }

  /**
   * Gets a configuration value
   */
  public get(key: string): string | number | boolean | undefined {
    if (!this.initialized || !this.config) {
      return undefined;
    }
    return this.config[key];
  }

  /**
   * Gets a configuration value as a string
   */
  public getString(key: string, defaultValue: string = ''): string {
    const value = this.get(key);
    if (value === undefined || value === null) {
      return defaultValue;
    }
    return String(value);
  }

  /**
   * Gets a configuration value as a boolean
   */
  public getBoolean(key: string, defaultValue: boolean = false): boolean {
    const value = this.get(key);
    if (value === undefined || value === null) {
      return defaultValue;
    }
    if (typeof value === 'boolean') {
      return value;
    }
    if (typeof value === 'string') {
      return value.toLowerCase() === 'true' || value === '1';
    }
    return Boolean(value);
  }

  /**
   * Gets a configuration value as a number
   */
  public getNumber(key: string, defaultValue: number = 0): number {
    const value = this.get(key);
    if (value === undefined || value === null) {
      return defaultValue;
    }
    if (typeof value === 'number') {
      return value;
    }
    const parsed = Number(value);
    return isNaN(parsed) ? defaultValue : parsed;
  }

  /**
   * Gets the current environment name
   */
  public getEnvironment(): string {
    return this.environment;
  }

  /**
   * Gets all configuration as an object (for compatibility with process.env)
   */
  public getAll(): EnvironmentConfig {
    return this.config || {};
  }

  /**
   * Checks if the configuration is initialized
   */
  public isInitialized(): boolean {
    return this.initialized;
  }
}

export default RuntimeConfig.getInstance();

