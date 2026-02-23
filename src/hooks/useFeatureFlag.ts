import { useMemo } from 'react';
import featureFlagService, { FeatureFlag } from '../service/FeatureFlagService';

/**
 * Hook to check if a single feature flag is enabled
 * @param flag The feature flag to check
 * @returns boolean indicating if the feature is enabled
 */
export function useFeatureFlag(flag: FeatureFlag): boolean {
  return useMemo(() => featureFlagService.isEnabled(flag), [flag]);
}

/**
 * Hook to get all feature flags and their status
 * @returns Object with all feature flags and their boolean values
 */
export function useFeatureFlags(): Record<FeatureFlag, boolean> {
  return useMemo(() => featureFlagService.getAllFlags(), []);
}

/**
 * Hook to get list of enabled product keys
 * @returns Array of product keys that are enabled by feature flags
 */
export function useEnabledProducts(): string[] {
  return useMemo(() => featureFlagService.getEnabledProducts(), []);
}
