import DebugService from '../service/DebugService';

export const getSubject = (item: Office.MessageRead | Office.MessageCompose): Promise<string> => {
  return new Promise((resolve, reject) => {
    try {
      if (typeof (item as Office.MessageCompose).subject?.getAsync === 'function') {
        (item as Office.MessageCompose).subject.getAsync((res) => {
          if (res.status === Office.AsyncResultStatus.Succeeded) {
            const subject = res.value || '';
            DebugService.debug('getSubject (compose mode):', subject, 'type:', typeof subject);
            // Ensure we always return a string
            resolve(String(subject));
          } else {
            DebugService.error('getSubject (compose mode) failed:', res.error);
            // Return empty string instead of rejecting to prevent crashes
            resolve('');
          }
        });
      } else {
        const subject = (item as Office.MessageRead).subject || '';
        DebugService.debug('getSubject (read mode):', subject, 'type:', typeof subject);
        // Ensure we always return a string
        resolve(String(subject));
      }
    } catch (error) {
      DebugService.error('getSubject error:', error);
      // Return empty string instead of rejecting to prevent crashes
      resolve('');
    }
  });
};

export const getSender = (item: Office.MessageRead | Office.MessageCompose): Promise<string> => {
  return new Promise((resolve, reject) => {
    try {
      if (typeof (item as Office.MessageCompose).from?.getAsync === 'function') {
        (item as Office.MessageCompose).from.getAsync((res) => {
          if (res.status === Office.AsyncResultStatus.Succeeded) {
            const sender = res.value?.emailAddress || Office.context.mailbox.userProfile.emailAddress;
            DebugService.debug('getSender (compose mode):', sender, 'type:', typeof sender);
            // Ensure we always return a string
            resolve(String(sender));
          } else {
            DebugService.error('getSender (compose mode) failed:', res.error);
            // Return user email instead of rejecting to prevent crashes
            resolve(String(Office.context.mailbox.userProfile.emailAddress));
          }
        });
      } else {
        const sender = (item as Office.MessageRead).from?.emailAddress || Office.context.mailbox.userProfile.emailAddress;
        DebugService.debug('getSender (read mode):', sender, 'type:', typeof sender);
        // Ensure we always return a string
        resolve(String(sender));
      }
    } catch (error) {
      DebugService.error('getSender error:', error);
      // Return user email instead of rejecting to prevent crashes
      resolve(String(Office.context.mailbox.userProfile.emailAddress));
    }
  });
};

export const getCreatedDate = (item: Office.MessageRead | Office.MessageCompose): Promise<string> => {
  return new Promise((resolve) => {
    const dt = (item as any).dateTimeCreated;
    resolve(dt ? new Date(dt).toISOString() : new Date().toISOString());
  });
};

export const getInternetMessageId = (item: Office.MessageRead | Office.MessageCompose): Promise<string | null> => {
  return new Promise((resolve) => {
    // First try to get it directly from the item
    const directMessageId = (item as any).internetMessageId;
    if (directMessageId) {
      resolve(directMessageId);
      return;
    }

    // If not available directly, try to get it from internet headers
    const messageItem = item as any;
    if (messageItem.internetHeaders && typeof messageItem.internetHeaders.getAsync === 'function') {
      messageItem.internetHeaders.getAsync(['Message-ID'], (result: Office.AsyncResult<any>) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const messageId = result.value && result.value['Message-ID'];
          resolve(messageId || null);
        } else {
          resolve(null);
        }
      });
    } else if (typeof (item as any).getAllInternetHeadersAsync === 'function') {
      // Fallback: get all headers and parse Message-ID
      (item as any).getAllInternetHeadersAsync((result: Office.AsyncResult<string>) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const headers = result.value || '';
          const match = headers.match(/^Message-ID:\s*(.+)$/im);
          const messageId = match ? match[1].trim() : null;
          resolve(messageId);
        } else {
          resolve(null);
        }
      });
    } else {
      resolve(null);
    }
  });
};

/**
 * Detects if the current email is from a shared mailbox
 * @param item The Office.js mailbox item
 * @returns Promise<{isShared: boolean, mailboxEmail?: string}> - Whether it's shared and the mailbox email
 */
export const detectSharedMailbox = (item: Office.MessageRead | Office.MessageCompose): Promise<{isShared: boolean, mailboxEmail?: string}> => {
  return new Promise((resolve) => {
    try {
      // Get the current user's email address
      const currentUserEmail = Office.context.mailbox.userProfile.emailAddress;
      
      // Use Office.js API to detect shared mailbox
      if (typeof item.getSharedPropertiesAsync === 'function') {
        item.getSharedPropertiesAsync((result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded && result.value) {
            // This is a shared mailbox - get the owner's email
            const sharedProperties = result.value;
            const mailboxEmail = sharedProperties.owner || sharedProperties.targetMailbox || currentUserEmail;
            
            DebugService.debug('Shared mailbox detected:', {
              isShared: true,
              mailboxEmail: mailboxEmail,
              owner: sharedProperties.owner,
              targetMailbox: sharedProperties.targetMailbox
            });
            
            resolve({
              isShared: true,
              mailboxEmail: mailboxEmail
            });
          } else {
            // Not a shared mailbox or API not supported
            DebugService.debug('Personal mailbox detected or getSharedPropertiesAsync not supported');
            resolve({
              isShared: false,
              mailboxEmail: currentUserEmail
            });
          }
        });
      } else {
        // Fallback: Check if we're in a different context
        // This is a conservative approach for when the API isn't available
        DebugService.debug('getSharedPropertiesAsync not available, assuming personal mailbox');
        resolve({
          isShared: false,
          mailboxEmail: currentUserEmail
        });
      }
    } catch (error) {
      DebugService.error('Error detecting shared mailbox:', error);
      resolve({
        isShared: false,
        mailboxEmail: Office.context.mailbox.userProfile.emailAddress
      });
    }
  });
}; 