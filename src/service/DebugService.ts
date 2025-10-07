/**
 * Centralized debugging service for the Outlook Add-in
 * Controls console logging based on environment variables
 */
export class DebugService {
  private static instance: DebugService;
  
  // Environment variable to control debug logging
  private readonly DEBUG_ENABLED: boolean = process.env.REACT_APP_DEBUG_ENABLED === 'true';
  private readonly DEBUG_LEVEL: string = process.env.REACT_APP_DEBUG_LEVEL || 'info';
  
  // Debug levels
  private readonly DEBUG_LEVELS = {
    error: 0,
    warn: 1,
    info: 2,
    debug: 3,
    trace: 4
  };

  private constructor() {
    // Log debug service initialization
    if (this.DEBUG_ENABLED) {
      console.log('🔧 DebugService initialized:', {
        enabled: this.DEBUG_ENABLED,
        level: this.DEBUG_LEVEL,
        currentLevel: this.DEBUG_LEVELS[this.DEBUG_LEVEL as keyof typeof this.DEBUG_LEVELS] || 2
      });
    }
  }

  public static getInstance(): DebugService {
    if (!DebugService.instance) {
      DebugService.instance = new DebugService();
    }
    return DebugService.instance;
  }

  /**
   * Check if debug logging is enabled
   */
  public isEnabled(): boolean {
    return this.DEBUG_ENABLED;
  }

  /**
   * Check if a specific debug level should be logged
   */
  private shouldLog(level: string): boolean {
    if (!this.DEBUG_ENABLED) return false;
    
    const currentLevel = this.DEBUG_LEVELS[this.DEBUG_LEVEL as keyof typeof this.DEBUG_LEVELS] || 2;
    const messageLevel = this.DEBUG_LEVELS[level as keyof typeof this.DEBUG_LEVELS] || 2;
    
    return messageLevel <= currentLevel;
  }

  /**
   * Log error messages
   */
  public error(message: string, ...args: any[]): void {
    if (this.shouldLog('error')) {
      console.error(`🔴 [ERROR] ${message}`, ...args);
    }
  }

  /**
   * Log warning messages
   */
  public warn(message: string, ...args: any[]): void {
    if (this.shouldLog('warn')) {
      console.warn(`🟡 [WARN] ${message}`, ...args);
    }
  }

  /**
   * Log info messages
   */
  public info(message: string, ...args: any[]): void {
    if (this.shouldLog('info')) {
      console.log(`🔵 [INFO] ${message}`, ...args);
    }
  }

  /**
   * Log debug messages
   */
  public debug(message: string, ...args: any[]): void {
    if (this.shouldLog('debug')) {
      console.log(`🟢 [DEBUG] ${message}`, ...args);
    }
  }

  /**
   * Log trace messages (most verbose)
   */
  public trace(message: string, ...args: any[]): void {
    if (this.shouldLog('trace')) {
      console.log(`⚪ [TRACE] ${message}`, ...args);
    }
  }

  /**
   * Log with custom prefix
   */
  public log(prefix: string, message: string, ...args: any[]): void {
    if (this.shouldLog('info')) {
      console.log(`📝 [${prefix}] ${message}`, ...args);
    }
  }

  /**
   * Log section headers for better organization
   */
  public section(title: string): void {
    if (this.shouldLog('info')) {
      console.log(`\n${'='.repeat(50)}`);
      console.log(`📋 ${title}`);
      console.log(`${'='.repeat(50)}`);
    }
  }

  /**
   * Log subsection headers
   */
  public subsection(title: string): void {
    if (this.shouldLog('info')) {
      console.log(`\n${'-'.repeat(30)}`);
      console.log(`📌 ${title}`);
      console.log(`${'-'.repeat(30)}`);
    }
  }

  /**
   * Log object data in a formatted way
   */
  public object(label: string, obj: any): void {
    if (this.shouldLog('debug')) {
      console.log(`📊 [OBJECT] ${label}:`, obj);
    }
  }

  /**
   * Log API calls
   */
  public api(method: string, url: string, data?: any): void {
    if (this.shouldLog('debug')) {
      console.log(`🌐 [API] ${method} ${url}`, data ? { data } : '');
    }
  }

  /**
   * Log authentication events
   */
  public auth(event: string, details?: any): void {
    if (this.shouldLog('info')) {
      console.log(`🔐 [AUTH] ${event}`, details || '');
    }
  }

  /**
   * Log email operations
   */
  public email(operation: string, details?: any): void {
    if (this.shouldLog('info')) {
      console.log(`📧 [EMAIL] ${operation}`, details || '');
    }
  }

  /**
   * Log placement operations
   */
  public placement(operation: string, details?: any): void {
    if (this.shouldLog('info')) {
      console.log(`📦 [PLACEMENT] ${operation}`, details || '');
    }
  }

  /**
   * Log Graph API operations
   */
  public graph(operation: string, details?: any): void {
    if (this.shouldLog('debug')) {
      console.log(`📈 [GRAPH] ${operation}`, details || '');
    }
  }

  /**
   * Log service operations
   */
  public service(serviceName: string, operation: string, details?: any): void {
    if (this.shouldLog('debug')) {
      console.log(`⚙️ [${serviceName.toUpperCase()}] ${operation}`, details || '');
    }
  }

  /**
   * Log timing information
   */
  public timing(operation: string, startTime: number): void {
    if (this.shouldLog('debug')) {
      const duration = Date.now() - startTime;
      console.log(`⏱️ [TIMING] ${operation} took ${duration}ms`);
    }
  }

  /**
   * Log error with stack trace
   */
  public errorWithStack(message: string, error: Error): void {
    if (this.shouldLog('error')) {
      console.error(`🔴 [ERROR] ${message}:`, error);
      if (error.stack) {
        console.error(`🔴 [STACK] ${error.stack}`);
      }
    }
  }

  /**
   * Log performance metrics
   */
  public performance(metric: string, value: number, unit: string = 'ms'): void {
    if (this.shouldLog('debug')) {
      console.log(`📊 [PERF] ${metric}: ${value}${unit}`);
    }
  }
}

// Export singleton instance
export default DebugService.getInstance(); 