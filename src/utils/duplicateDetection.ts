// Utility for duplicate detection in Outlook add-in
// Checks CustomProperties (UWWBID/X-UWWBID), internetHeaders (X-UWWBID), and getAllInternetHeadersAsync fallback
// Handles drafts, forwarded emails, and already filed emails

export async function checkDuplicateSubmission(item: any, DebugService: any): Promise<boolean> {
  DebugService.debug('=== DUPLICATE DETECTION START ===');
  DebugService.debug('Item type:', item ? 'exists' : 'null');
  DebugService.debug('Item has itemId:', !!(item as any)?.itemId);
  DebugService.debug('Item has subject:', !!(item as any)?.subject);
  DebugService.debug('Item has internetHeaders:', !!(item as any)?.internetHeaders);
  DebugService.debug('Item has loadCustomPropertiesAsync:', typeof item?.loadCustomPropertiesAsync === 'function');
  DebugService.debug('Item has saveAsync:', typeof item?.saveAsync === 'function');
  
  // Drafts: check CustomProperties and subject line
  if (!item || !(item as any).itemId) {
    DebugService.debug('🔍 DETECTED AS DRAFT EMAIL (no itemId)');
    DebugService.debug('Draft email: checking CustomProperties and subject for duplicate detection');
    return new Promise((resolve) => {
      let checked = 0;
      let found = false;
      
      const finish = () => {
        checked++;
        DebugService.debug(`Draft Check ${checked}/2 completed. Found so far: ${found}`);
        if (checked === 2) {
          DebugService.debug('=== DUPLICATE DETECTION RESULT (DRAFT):', found, '===');
          resolve(found);
        }
      };

      // Check 1: CustomProperties (UWWBID or X-UWWBID)
      DebugService.debug('Draft Check 1: CustomProperties');
      if (typeof item.loadCustomPropertiesAsync === 'function') {
        item.loadCustomPropertiesAsync((result: Office.AsyncResult<Office.CustomProperties>) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            const cp = result.value;
            const cpValue = cp?.get('UWWBID') || cp?.get('X-UWWBID');
            DebugService.debug('Draft CustomProperties UWWBID/X-UWWBID value:', cpValue);
            if (cpValue) {
              found = true;
              DebugService.debug('✅ Found duplicate in Draft CustomProperties');
            }
          } else {
            DebugService.warn('Failed to load CustomProperties for draft:', result.error);
          }
          finish();
        });
      } else {
        DebugService.warn('loadCustomPropertiesAsync not available on draft item');
        finish();
      }

      // Check 2: Subject line for WBID pattern (for forwarded emails in shared mailboxes)
      DebugService.debug('Draft Check 2: Subject line');
      if (item.subject && typeof item.subject.getAsync === 'function') {
        item.subject.getAsync((result: Office.AsyncResult<string>) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            const subject = result.value || '';
            DebugService.debug('Draft Subject line:', subject);
            DebugService.debug('Draft Subject line length:', subject.length);
            DebugService.debug('Draft Subject line contains "WBID":', subject.includes('WBID'));
            DebugService.debug('Draft Subject line contains "UWWBID":', subject.includes('UWWBID'));
            
            const wbidMatch = subject.match(/WBID:\s*([A-Z0-9-]+)/i);
            const uwwbidMatch = subject.match(/UWWBID:\s*([A-Z0-9-]+)/i);
            DebugService.debug('Draft Subject WBID match:', wbidMatch);
            DebugService.debug('Draft Subject UWWBID match:', uwwbidMatch);
            DebugService.debug('Draft Subject WBID match[1]:', wbidMatch?.[1]);
            DebugService.debug('Draft Subject UWWBID match[1]:', uwwbidMatch?.[1]);
            
            if (wbidMatch?.[1] || uwwbidMatch?.[1]) {
              found = true;
              DebugService.debug('✅ Found duplicate in Draft Subject line');
            } else {
              DebugService.debug('❌ No duplicate found in Draft Subject line');
            }
          } else {
            DebugService.warn('Failed to read draft subject:', result.error);
          }
          finish();
        });
      } else {
        DebugService.warn('Subject not available on draft item');
        finish();
      }
    });
  }

  // For non-drafts (forwarded/filed emails), check all methods
  return new Promise((resolve) => {
    DebugService.debug('🔍 DETECTED AS NON-DRAFT EMAIL (has itemId)');
    let checked = 0;
    let found = false;
    
    const finish = () => {
      checked++;
      DebugService.debug(`Check ${checked}/4 completed. Found so far: ${found}`);
      if (checked === 4) {
        DebugService.debug('=== DUPLICATE DETECTION RESULT (NON-DRAFT):', found, '===');
        resolve(found);
      }
    };

    // Check 1: CustomProperties (UWWBID or X-UWWBID)
    DebugService.debug('Check 1: CustomProperties');
    if (typeof item.loadCustomPropertiesAsync === 'function') {
      item.loadCustomPropertiesAsync((result: Office.AsyncResult<Office.CustomProperties>) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const cp = result.value;
          const cpValue = cp?.get('UWWBID') || cp?.get('X-UWWBID');
          DebugService.debug('CustomProperties UWWBID/X-UWWBID value:', cpValue);
          if (cpValue) {
            found = true;
            DebugService.debug('✅ Found duplicate in CustomProperties');
          }
        } else {
          DebugService.warn('Failed to load CustomProperties:', result.error);
        }
        finish();
      });
    } else {
      DebugService.warn('loadCustomPropertiesAsync not available on this item');
      finish();
    }

    // Check 2: Subject line for WBID pattern
    DebugService.debug('Check 2: Subject line');
    if (item.subject && typeof item.subject.getAsync === 'function') {
      item.subject.getAsync((result: Office.AsyncResult<string>) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const subject = result.value || '';
          DebugService.debug('Subject line:', subject);
          DebugService.debug('Subject line length:', subject.length);
          DebugService.debug('Subject line contains "WBID":', subject.includes('WBID'));
          DebugService.debug('Subject line contains "UWWBID":', subject.includes('UWWBID'));
          
          const wbidMatch = subject.match(/WBID:\s*([A-Z0-9-]+)/i);
          const uwwbidMatch = subject.match(/UWWBID:\s*([A-Z0-9-]+)/i);
          DebugService.debug('Subject WBID match:', wbidMatch);
          DebugService.debug('Subject UWWBID match:', uwwbidMatch);
          DebugService.debug('Subject WBID match[1]:', wbidMatch?.[1]);
          DebugService.debug('Subject UWWBID match[1]:', uwwbidMatch?.[1]);
          
          if (wbidMatch?.[1] || uwwbidMatch?.[1]) {
            found = true;
            DebugService.debug('✅ Found duplicate in Subject line');
          } else {
            DebugService.debug('❌ No duplicate found in Subject line');
          }
        } else {
          DebugService.warn('Failed to read subject:', result.error);
        }
        finish();
      });
    } else {
      DebugService.warn('Subject not available on this item');
      finish();
    }

    // Check 3: internetHeaders (X-UWWBID)
    DebugService.debug('Check 3: Internet Headers');
    if (item.internetHeaders && typeof item.internetHeaders.getAsync === 'function') {
      item.internetHeaders.getAsync(['X-UWWBID'], (result: Office.AsyncResult<any>) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const value = result.value && result.value['X-UWWBID'];
          DebugService.debug('internetHeaders X-UWWBID value:', value);
          if (value) {
            found = true;
            DebugService.debug('✅ Found duplicate in Internet Headers');
          }
        } else {
          DebugService.warn('Failed to read X-UWWBID header:', result.error);
        }
        finish();
      });
    } else if (typeof item.getAllInternetHeadersAsync === 'function') {
      // Check 4: Fallback - get all headers and parse
      DebugService.debug('Check 4: All Internet Headers (fallback)');
      item.getAllInternetHeadersAsync((result: Office.AsyncResult<string>) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const headers = result.value || '';
          DebugService.debug('All headers length:', headers.length);
          const match = headers.match(/^X-UWWBID:\s*(.+)$/im);
          const xUwwbId = match ? match[1].trim() : null;
          DebugService.debug('getAllInternetHeadersAsync X-UWWBID:', xUwwbId);
          if (xUwwbId) {
            found = true;
            DebugService.debug('✅ Found duplicate in All Internet Headers');
          }
        } else {
          DebugService.warn('Failed to get all internet headers:', result.error);
        }
        finish();
      });
    } else {
      DebugService.warn('Neither internetHeaders.getAsync nor getAllInternetHeadersAsync available on this item');
      finish();
    }

    // Fourth check is always called to keep checked==4 logic
    setTimeout(finish, 0); // No-op, just to ensure checked==4
  });
} 