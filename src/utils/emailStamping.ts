// Utility for stamping emails in Outlook add-in
// Stamps CustomProperties (UWWBID) for drafts and all emails, and internetHeaders (X-UWWBID) for non-drafts if available

export async function stampEmailWithWorkbenchId(item: any, workbenchId: string, DebugService: any): Promise<void> {
  DebugService.debug('Starting email stamping with workbenchId:', workbenchId);
  DebugService.debug('Item type:', item ? 'exists' : 'null');
  DebugService.debug('Item has itemId:', !!(item as any)?.itemId);
  
  // For drafts, only stamp CustomProperties
  if (!item || !(item as any).itemId) {
    DebugService.debug('Stamping draft email: CustomProperties only');
    return new Promise((resolve) => {
      if (typeof item.loadCustomPropertiesAsync === 'function') {
        item.loadCustomPropertiesAsync((result: Office.AsyncResult<Office.CustomProperties>) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            const cp = result.value;
            cp?.set('UWWBID', workbenchId);
            cp?.saveAsync((saveResult) => {
              if (saveResult.status === Office.AsyncResultStatus.Succeeded) {
                DebugService.debug('Stamped CustomProperties UWWBID on draft:', workbenchId);
              } else {
                DebugService.warn('Failed to save CustomProperties UWWBID on draft:', saveResult.error);
              }
              resolve();
            });
          } else {
            DebugService.warn('Failed to load CustomProperties for stamping draft:', result.error);
            resolve();
          }
        });
      } else {
        DebugService.warn('loadCustomPropertiesAsync not available for stamping draft');
        resolve();
      }
    });
  }
  // For non-drafts, stamp internetHeaders if possible, and always stamp CustomProperties
  return new Promise((resolve, reject) => {
    if (item.internetHeaders && typeof item.internetHeaders.setAsync === 'function') {
      item.internetHeaders.setAsync({ 'X-UWWBID': workbenchId }, (result: Office.AsyncResult<void>) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          DebugService.debug('Stamped X-UWWBID header:', workbenchId);
        } else {
          DebugService.warn('Failed to stamp X-UWWBID header:', result.error);
        }
        // Always stamp CustomProperties for redundancy
        if (typeof item.loadCustomPropertiesAsync === 'function') {
          item.loadCustomPropertiesAsync((cpResult: Office.AsyncResult<Office.CustomProperties>) => {
            if (cpResult.status === Office.AsyncResultStatus.Succeeded) {
              const cp = cpResult.value;
              cp?.set('UWWBID', workbenchId);
              cp?.saveAsync((saveResult) => {
                if (saveResult.status === Office.AsyncResultStatus.Succeeded) {
                  DebugService.debug('Stamped CustomProperties UWWBID:', workbenchId);
                } else {
                  DebugService.warn('Failed to save CustomProperties UWWBID:', saveResult.error);
                }
                resolve();
              });
            } else {
              DebugService.warn('Failed to load CustomProperties for stamping:', cpResult.error);
              resolve();
            }
          });
        } else {
          DebugService.warn('loadCustomPropertiesAsync not available for stamping');
          resolve();
        }
      });
    } else if (typeof item.loadCustomPropertiesAsync === 'function') {
      // If headers not available, stamp custom properties only
      item.loadCustomPropertiesAsync((cpResult: Office.AsyncResult<Office.CustomProperties>) => {
        if (cpResult.status === Office.AsyncResultStatus.Succeeded) {
          const cp = cpResult.value;
          cp?.set('UWWBID', workbenchId);
          cp?.saveAsync((saveResult) => {
            if (saveResult.status === Office.AsyncResultStatus.Succeeded) {
              DebugService.debug('Stamped CustomProperties UWWBID:', workbenchId);
            } else {
              DebugService.warn('Failed to save CustomProperties UWWBID:', saveResult.error);
            }
            resolve();
          });
        } else {
          DebugService.warn('Failed to load CustomProperties for stamping:', cpResult.error);
          resolve();
        }
      });
    } else {
      DebugService.warn('Neither internetHeaders nor loadCustomPropertiesAsync available for stamping');
      resolve();
    }
  });
} 