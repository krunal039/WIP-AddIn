// Utility for duplicate detection in Outlook add-in
// Checks CustomProperties (UWWBID/X-UWWBID), internetHeaders (X-UWWBID), and getAllInternetHeadersAsync fallback
// Handles drafts, forwarded emails, and already filed emails

import DebugService from '../service/DebugService';

/**
 * Check CustomProperties for duplicate markers
 */
async function checkCustomProperties(item: any): Promise<boolean> {
  return new Promise((resolve) => {
    if (typeof item.loadCustomPropertiesAsync !== 'function') {
      resolve(false);
      return;
    }

    item.loadCustomPropertiesAsync((result: Office.AsyncResult<Office.CustomProperties>) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const cp = result.value;
        const cpValue = cp?.get('UWWBID') || cp?.get('X-UWWBID');
        if (cpValue) {
          DebugService.debug('Found duplicate in CustomProperties');
          resolve(true);
          return;
        }
      }
      resolve(false);
    });
  });
}

/**
 * Check subject line for WBID/UWWBID patterns
 */
async function checkSubjectLine(item: any): Promise<boolean> {
  return new Promise((resolve) => {
    if (!item.subject || typeof item.subject.getAsync !== 'function') {
      resolve(false);
      return;
    }

    item.subject.getAsync((result: Office.AsyncResult<string>) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const subject = result.value || '';
        const wbidMatch = subject.match(/WBID:\s*([A-Z0-9-]+)/i);
        const uwwbidMatch = subject.match(/UWWBID:\s*([A-Z0-9-]+)/i);
        
        if (wbidMatch?.[1] || uwwbidMatch?.[1]) {
          DebugService.debug('Found duplicate in Subject line');
          resolve(true);
          return;
        }
      }
      resolve(false);
    });
  });
}

/**
 * Check internet headers for X-UWWBID
 */
async function checkInternetHeaders(item: any): Promise<boolean> {
  return new Promise((resolve) => {
    if (item.internetHeaders && typeof item.internetHeaders.getAsync === 'function') {
      item.internetHeaders.getAsync(['X-UWWBID'], (result: Office.AsyncResult<any>) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const value = result.value && result.value['X-UWWBID'];
          if (value) {
            DebugService.debug('Found duplicate in Internet Headers');
            resolve(true);
            return;
          }
        }
        resolve(false);
      });
    } else if (typeof item.getAllInternetHeadersAsync === 'function') {
      // Fallback: get all headers and parse
      item.getAllInternetHeadersAsync((result: Office.AsyncResult<string>) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const headers = result.value || '';
          const match = headers.match(/^X-UWWBID:\s*(.+)$/im);
          const xUwwbId = match ? match[1].trim() : null;
          if (xUwwbId) {
            DebugService.debug('Found duplicate in All Internet Headers');
            resolve(true);
            return;
          }
        }
        resolve(false);
      });
    } else {
      resolve(false);
    }
  });
}

/**
 * Check if email has already been submitted to Workbench
 * @param item Office.js mailbox item
 * @returns Promise<boolean> true if duplicate found
 */
export async function checkDuplicateSubmission(item: any | null): Promise<boolean> {
  if (!item) {
    return false;
  }

  const isDraft = !(item as any).itemId;
  DebugService.debug(`Checking for duplicates (${isDraft ? 'draft' : 'non-draft'})`);

  if (isDraft) {
    // Drafts: check CustomProperties and subject line only
    const [hasCustomProps, hasSubject] = await Promise.all([
      checkCustomProperties(item),
      checkSubjectLine(item)
    ]);
    
    const isDuplicate = hasCustomProps || hasSubject;
    DebugService.debug(`Duplicate check result (draft): ${isDuplicate}`);
    return isDuplicate;
  }

  // Non-drafts: check all methods
  const [hasCustomProps, hasSubject, hasHeaders] = await Promise.all([
    checkCustomProperties(item),
    checkSubjectLine(item),
    checkInternetHeaders(item)
  ]);

  const isDuplicate = hasCustomProps || hasSubject || hasHeaders;
  DebugService.debug(`Duplicate check result (non-draft): ${isDuplicate}`);
  return isDuplicate;
}
