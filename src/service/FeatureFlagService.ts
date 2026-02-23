/**
 * Feature Flag Service
 * 
 * Provides feature flag functionality based on environment configuration.
 * Feature flags are defined in environments.json per environment.
 */
import runtimeConfig from '../config/runtimeConfig';

export type FeatureFlag =
  | 'newPlacement'
  | 'updatePlacement'
  | 'productCyber'
  | 'productNaLpl'
  | 'productNaMpl';

const FEATURE_FLAG_MAP: Record<FeatureFlag, string> = {
  newPlacement: 'FEATURE_NEW_PLACEMENT',
  updatePlacement: 'FEATURE_UPDATE_PLACEMENT',
  productCyber: 'FEATURE_PRODUCT_CYBER',
  productNaLpl: 'FEATURE_PRODUCT_NA_LPL',
  productNaMpl: 'FEATURE_PRODUCT_NA_MPL',
};

class FeatureFlagService {
  /**
   * Check if a feature flag is enabled
   * @param flag The feature flag to check
   * @returns true if enabled, false if disabled (defaults to true if not configured)
   */
  isEnabled(flag: FeatureFlag): boolean {
    const configKey = FEATURE_FLAG_MAP[flag];
    return runtimeConfig.getBoolean(configKey, true);
  }

  /**
   * Get list of enabled product keys based on feature flags
   * @returns Array of product keys that are enabled
   */
  getEnabledProducts(): string[] {
    const products: string[] = [];
    if (this.isEnabled('productCyber')) products.push('20001');
    if (this.isEnabled('productNaLpl')) products.push('10013');
    if (this.isEnabled('productNaMpl')) products.push('10012');
    return products;
  }

  /**
   * Get all feature flags and their current status
   * @returns Object with all feature flags and their boolean values
   */
  getAllFlags(): Record<FeatureFlag, boolean> {
    const flags = {} as Record<FeatureFlag, boolean>;
    for (const flag of Object.keys(FEATURE_FLAG_MAP) as FeatureFlag[]) {
      flags[flag] = this.isEnabled(flag);
    }
    return flags;
  }
}

export default new FeatureFlagService();
