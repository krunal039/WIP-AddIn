/**
 * Application constants
 * Centralized location for all magic numbers, strings, and configuration values
 */

// Token refresh settings
export const TOKEN_REFRESH_INTERVAL_MS = 5 * 60 * 1000; // 5 minutes
export const TOKEN_REFRESH_BUFFER_SECONDS = 300; // 5 minutes before expiry

// File size limits
export const MAX_ATTACHMENT_SIZE_BYTES = 25 * 1024 * 1024; // 25MB
export const MAX_BASE64_LENGTH = 35 * 1024 * 1024; // 35MB in base64 characters

// LocalStorage keys
export const STORAGE_KEYS = {
  SELECTED_PRODUCT: 'wb.selectedProduct',
  SELECTED_BU: 'wb.selectedBU',
  SEND_COPY_TO_CYBER_ADMIN: 'wb.sendCopyToCyberAdmin',
} as const;

import { IDropdownOption } from '@fluentui/react';

// BU Options
export const BU_OPTIONS: IDropdownOption[] = [
  { key: 'MRSNA', text: 'MRSNA' },
  { key: 'MRSGM', text: 'MRSGM' },
];

// Product Options
export const PRODUCT_OPTIONS: IDropdownOption[] = [
  { key: '20001', text: 'Cyber' },
  { key: '10013', text: 'NA LPL' },
  { key: '10012', text: 'NA MPL' },
];

// Default values
export const DEFAULTS = {
  PRODUCT: '20001',
  BU: 'MRSGM',
  SEND_COPY_TO_CYBER_ADMIN: false,
} as const;

// Error messages
export const ERROR_MESSAGES = {
  OFFICE_NOT_READY: 'Office.js failed to initialize',
  INVALID_ITEM_TYPE: 'Invalid item type. Expected Message.',
  EMPTY_SUBJECT: 'Email subject is required',
  FILE_VALIDATION_FAILED: 'File validation failed',
  DUPLICATE_SUBMISSION: 'This email has already been submitted',
  TOKEN_REFRESH_FAILED: 'Token refresh failed',
  MSAL_INIT_FAILED: 'Failed to initialize MSAL',
  CONFIG_LOAD_FAILED: 'Failed to load configuration',
} as const;

// API endpoints
export const API_ENDPOINTS = {
  GRAPH_ME: 'https://graph.microsoft.com/v1.0/me',
  GRAPH_USERS: (email: string) => `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(email)}`,
} as const;

// Duplicate detection markers
export const DUPLICATE_MARKERS = {
  CUSTOM_PROPERTY_UWWBID: 'UWWBID',
  CUSTOM_PROPERTY_X_UWWBID: 'X-UWWBID',
  INTERNET_HEADER_X_UWWBID: 'X-UWWBID',
  SUBJECT_WBID_PATTERN: /WBID:\s*([A-Z0-9-]+)/i,
  SUBJECT_UWWBID_PATTERN: /UWWBID:\s*([A-Z0-9-]+)/i,
} as const;

