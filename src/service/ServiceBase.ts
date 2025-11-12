/**
 * Base class for singleton services
 * Provides common singleton pattern implementation
 */
export abstract class ServiceBase {
  private static instances: Map<typeof ServiceBase, ServiceBase> = new Map();

  protected constructor() {
    // Prevent direct instantiation
  }

  /**
   * Get or create singleton instance
   * @param ServiceClass The service class to get instance for
   * @returns Singleton instance
   */
  protected static getInstance<T extends ServiceBase>(
    ServiceClass: new () => T
  ): T {
    if (!ServiceBase.instances.has(ServiceClass as any)) {
      ServiceBase.instances.set(ServiceClass as any, new ServiceClass());
    }
    return ServiceBase.instances.get(ServiceClass as any) as T;
  }
}

