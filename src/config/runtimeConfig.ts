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

class RuntimeConfig {
  private static instance: RuntimeConfig;
  private config: EnvironmentConfig | null = null;
  private environment: string = 'dev';
  private initialized: boolean = false;
  private initPromise: Promise<void> | null = null;
  private debugService: any = null; // Lazy-loaded to avoid circular dependency

  private constructor() {}

  /**
   * Gets DebugService instance (lazy-loaded to avoid circular dependency)
   */
  private getDebugService(): any {
    if (!this.debugService && typeof window !== 'undefined') {
      try {
        // Dynamic import to avoid circular dependency
        const DebugServiceModule = require('../service/DebugService');
        this.debugService = DebugServiceModule.default;
      } catch (error) {
        // DebugService not available yet, will use console fallback
      }
    }
    return this.debugService;
  }

  /**
   * Logs a message using DebugService if available, otherwise uses console
   * Respects DEBUG_ENABLED setting from runtime config
   */
  private log(level: 'info' | 'debug' | 'warn' | 'error', message: string, ...args: any[]): void {
    const debugService = this.getDebugService();
    if (debugService) {
      try {
        // Check if debug is enabled before logging (except for errors)
        const isDebugEnabled = this.config?.REACT_APP_DEBUG_ENABLED !== false;
        if (level !== 'error' && !isDebugEnabled) {
          // Don't log info/debug/warn if DEBUG_ENABLED is false
          return;
        }
        
        switch (level) {
          case 'info':
            debugService.info(message, ...args);
            break;
          case 'debug':
            debugService.debug(message, ...args);
            break;
          case 'warn':
            debugService.warn(message, ...args);
            break;
          case 'error':
            // Errors are always logged
            debugService.error(message, ...args);
            break;
        }
      } catch (error) {
        // Fallback to console only for errors if DebugService fails
        if (level === 'error') {
          console.error(`[RuntimeConfig] ${message}`, ...args);
        }
      }
    } else {
      // Fallback to console during initial load - but only for errors
      // Other logs will be suppressed if DEBUG_ENABLED is false
      if (level === 'error') {
        console.error(`🔴 [RuntimeConfig] ${message}`, ...args);
      }
      // For other levels, check if we should log
      // During initial load, we don't have config yet, so we'll be conservative
      // and only log errors
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
   * @param urlPatterns Optional URL patterns from environments.json
   */
  private detectEnvironment(urlPatterns?: { [key: string]: string[] }): string {
    const hostname = window.location.hostname.toLowerCase();
    const fullUrl = window.location.href;
    
    this.log('debug', 'Starting environment detection', {
      hostname,
      fullUrl,
      hasUrlPatterns: !!urlPatterns
    });
    
    // Check for explicit environment in URL search params (for testing)
    const urlParams = new URLSearchParams(window.location.search);
    const envParam = urlParams.get('env');
    if (envParam && ['dev', 'qa', 'uat', 'prod', 'localhost'].includes(envParam)) {
      this.log('info', `Environment forced via URL param: ${envParam}`);
      return envParam;
    }

    // Use provided URL patterns, or fallback to hardcoded patterns
    const patterns = urlPatterns || {
      localhost: [
        'localhost',
        '127.0.0.1'
      ],
      dev: [
        'gsi-email-ingestion-request-dev.munichre.com'
      ],
      qa: [
        'gsi-email-ingestion-request-qa.munichre.com'
      ],
      uat: [
        'gsi-email-ingestion-request-uat.munichre.com'
      ],
      prod: [
        'gsi-email-ingestion-request.munichre.com'
      ]
    };

    this.log('debug', 'Checking URL patterns', {
      hostname,
      availableEnvironments: Object.keys(patterns),
      patterns: Object.entries(patterns).map(([env, pats]) => ({
        environment: env,
        patterns: pats
      }))
    });

    // Check each environment's URL patterns
    for (const [env, envPatterns] of Object.entries(patterns)) {
      for (const pattern of envPatterns) {
        const patternLower = pattern.toLowerCase();
        const hostnameLower = hostname.toLowerCase();
        const matches = hostnameLower.includes(patternLower);
        
        this.log('debug', `Checking pattern match`, {
          environment: env,
          pattern: patternLower,
          hostname: hostnameLower,
          matches
        });
        
        if (matches) {
          this.log('info', `✅ Detected environment: ${env}`, {
            matchedPattern: pattern,
            hostname,
            fullUrl
          });
          return env;
        }
      }
    }

    // Default to localhost for local development, then dev
    const defaultEnv = hostname === 'localhost' || hostname === '127.0.0.1' ? 'localhost' : 'dev';
    this.log('warn', `⚠️ Could not detect environment from hostname "${hostname}", defaulting to ${defaultEnv}`, {
      hostname,
      fullUrl,
      checkedPatterns: Object.values(patterns).flat()
    });
    return defaultEnv;
  }

  /**
   * Loads the environments.json file and extracts the config for the detected environment
   */
  private async loadConfig(): Promise<void> {
    this.log('info', '🔄 Starting configuration load...');
    const startTime = Date.now();
    
    try {
      this.log('debug', 'Fetching environments.json from /environments.json');
      const response = await fetch('/environments.json');
      
      if (!response.ok) {
        throw new Error(`Failed to load environments.json: ${response.status} ${response.statusText}`);
      }

      this.log('debug', '✅ Successfully fetched environments.json', {
        status: response.status,
        statusText: response.statusText,
        contentType: response.headers.get('content-type')
      });

      const data: EnvironmentsData = await response.json();
      
      this.log('debug', 'Parsed environments.json', {
        availableEnvironments: Object.keys(data.environments),
        urlPatterns: Object.keys(data.urlPatterns || {})
      });
      
      // Detect environment using URL patterns from the loaded file
      this.log('info', '🔍 Detecting environment...');
      this.environment = this.detectEnvironment(data.urlPatterns);

      this.log('info', `📍 Detected environment: ${this.environment}`, {
        detectedEnvironment: this.environment,
        availableEnvironments: Object.keys(data.environments)
      });

      // Get the config for the detected environment
      if (!data.environments[this.environment]) {
        const errorMsg = `Environment "${this.environment}" not found in environments.json. Available: ${Object.keys(data.environments).join(', ')}`;
        this.log('error', errorMsg, {
          requestedEnvironment: this.environment,
          availableEnvironments: Object.keys(data.environments)
        });
        throw new Error(errorMsg);
      }

      this.config = data.environments[this.environment];
      this.initialized = true;

      const loadTime = Date.now() - startTime;
      
      // Validate all required configuration keys are present
      // At this point, config should not be null, but TypeScript needs assurance
      if (!this.config) {
        throw new Error('Configuration object is null after assignment');
      }
      
      const requiredKeys = [
        'REACT_APP_AZURE_CLIENT_ID',
        'REACT_APP_AZURE_AUTHORITY',
        'REACT_APP_AZURE_REDIRECT_URI',
        'REACT_APP_PLACEMENT_API_URL',
        'REACT_APP_LOGGING_API_URL'
      ];
      
      const missingKeys = requiredKeys.filter(key => !(key in this.config!));
      if (missingKeys.length > 0) {
        this.log('warn', '⚠️ Some required configuration keys are missing', {
          missingKeys,
          availableKeys: Object.keys(this.config)
        });
      }

      this.log('info', `✅ Configuration loaded successfully (ONE-TIME LOAD)`, {
        environment: this.environment,
        loadTimeMs: loadTime,
        configKeyCount: Object.keys(this.config).length,
        allConfigKeys: Object.keys(this.config).sort()
      });

      // Log ALL configuration values for verification (masking sensitive data)
      const allConfigValues: any = {};
      for (const key of Object.keys(this.config)) {
        const value = this.config[key];
        // Mask sensitive values
        if (key.includes('KEY') || key.includes('SECRET') || key.includes('TOKEN')) {
          allConfigValues[key] = value ? '***MASKED***' : '(empty)';
        } else {
          allConfigValues[key] = value;
        }
      }
      
      this.log('info', '📋 All loaded configuration values:', allConfigValues);

      // Log key configuration values for quick verification
      const keyConfigValues = {
        AZURE_CLIENT_ID: this.config.REACT_APP_AZURE_CLIENT_ID,
        AZURE_AUTHORITY: this.config.REACT_APP_AZURE_AUTHORITY,
        AZURE_REDIRECT_URI: this.config.REACT_APP_AZURE_REDIRECT_URI,
        PLACEMENT_API_URL: this.config.REACT_APP_PLACEMENT_API_URL,
        LOGGING_API_URL: this.config.REACT_APP_LOGGING_API_URL,
        DEBUG_ENABLED: this.config.REACT_APP_DEBUG_ENABLED,
        DEBUG_LEVEL: this.config.REACT_APP_DEBUG_LEVEL
      };
      
      this.log('info', '🔑 Key configuration values:', keyConfigValues);

      // Log full config if debug is enabled (after config is loaded)
      if (this.config.REACT_APP_DEBUG_ENABLED) {
        this.log('debug', 'Full configuration object (debug mode):', this.config);
      }

      // After config is loaded, refresh DebugService to use new config
      this.debugService = null; // Reset to reload with new config
      
      // Verify DebugService will respect the new config
      this.log('info', '✅ Configuration ready - DebugService will now respect DEBUG_ENABLED and DEBUG_LEVEL settings');
    } catch (error) {
      const loadTime = Date.now() - startTime;
      this.log('error', `❌ Failed to load configuration`, {
        error: error instanceof Error ? error.message : String(error),
        stack: error instanceof Error ? error.stack : undefined,
        loadTimeMs: loadTime,
        currentEnvironment: this.environment
      });
      
      // Fallback to empty config to prevent app crash
      this.config = {};
      this.initialized = true;
      this.log('warn', '⚠️ Using empty fallback configuration - app may not work correctly');
    }
  }

  /**
   * Initializes the configuration (loads environments.json)
   * This should be called ONCE before the app starts
   * Uses singleton pattern to ensure only one load happens
   */
  public async initialize(): Promise<void> {
    // If already initialized, return immediately
    if (this.initialized) {
      this.log('debug', 'Configuration already initialized, skipping reload', {
        environment: this.environment,
        configKeyCount: Object.keys(this.config || {}).length
      });
      return Promise.resolve();
    }

    // If initialization is in progress, wait for it
    if (this.initPromise) {
      this.log('debug', 'Initialization already in progress, waiting for existing promise');
      return this.initPromise;
    }

    this.log('info', '🚀 Initializing runtime configuration (one-time load)...', {
      currentUrl: window.location.href,
      hostname: window.location.hostname,
      timestamp: new Date().toISOString()
    });

    // Start the one-time load
    this.initPromise = this.loadConfig();
    return this.initPromise;
  }

  /**
   * Gets a configuration value
   */
  public get(key: string): string | number | boolean | undefined {
    if (!this.initialized) {
      this.log('warn', `Config not initialized yet, returning undefined for key: ${key}`);
      return undefined;
    }
    const value = this.config?.[key];
    if (value === undefined) {
      this.log('debug', `Config key not found: ${key}`, {
        key,
        availableKeys: Object.keys(this.config || {}),
        environment: this.environment
      });
    }
    return value;
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

