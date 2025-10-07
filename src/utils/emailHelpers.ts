export const getSubject = (item: Office.MessageRead | Office.MessageCompose): Promise<string> => {
  return new Promise((resolve, reject) => {
    if (typeof (item as Office.MessageCompose).subject?.getAsync === 'function') {
      (item as Office.MessageCompose).subject.getAsync((res) => {
        res.status === Office.AsyncResultStatus.Succeeded ? resolve(res.value || '') : reject(res.error?.message || 'Unknown error');
      });
    } else {
      resolve((item as Office.MessageRead).subject || '');
    }
  });
};

export const getSender = (item: Office.MessageRead | Office.MessageCompose): Promise<string> => {
  return new Promise((resolve, reject) => {
    if (typeof (item as Office.MessageCompose).from?.getAsync === 'function') {
      (item as Office.MessageCompose).from.getAsync((res) => {
        res.status === Office.AsyncResultStatus.Succeeded
          ? resolve(res.value?.emailAddress || Office.context.mailbox.userProfile.emailAddress)
          : reject(res.error?.message || 'Unknown error');
      });
    } else {
      resolve((item as Office.MessageRead).from?.emailAddress || Office.context.mailbox.userProfile.emailAddress);
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
    if (item.internetHeaders && typeof item.internetHeaders.getAsync === 'function') {
      item.internetHeaders.getAsync(['Message-ID'], (result: Office.AsyncResult<any>) => {
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