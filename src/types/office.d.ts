declare namespace Office {
  namespace MailboxEnums {
    enum ItemType {
      Message = "message"
    }
    enum ItemMode {
      Read = "read",
      Compose = "compose"
    }
    enum RestVersion {
      v2_0 = "v2.0"
    }
    enum ItemNotificationMessageType {
      InformationalMessage = "informationalMessage"
    }
  }

  interface AsyncResult<T> {
    status: Office.AsyncResultStatus;
    value?: T;
    error?: any;
  }

  enum AsyncResultStatus {
    Succeeded = "succeeded",
    Failed = "failed"
  }

  interface Context {
    mailbox: {
      item: Item;
      userProfile: UserProfile;
      convertToRestId(
        ids: string[],
        restVersion: MailboxEnums.RestVersion,
        callback: (result: AsyncResult<string[]>) => void
      ): void;
    };
    auth: {
      getAccessTokenAsync(options: {
        scopes: string[];
        callback: (result: AsyncResult<string>) => void;
      }): void;
    };
  }

  interface Item {
    itemType: MailboxEnums.ItemType;
    subject?: {
      getAsync(callback: (result: AsyncResult<string>) => void): void;
    };
    internetHeaders?: {
      getAsync(headerNames: string[], callback: (result: AsyncResult<any>) => void): void;
    };
    getAllInternetHeadersAsync?(callback: (result: AsyncResult<string>) => void): void;
    loadCustomPropertiesAsync?(callback: (result: AsyncResult<CustomProperties>) => void): void;
    saveAsync?(callback: (result: AsyncResult<string>) => void): void;
    itemId?: string;
  }

  interface CustomProperties {
    get(key: string): string | undefined;
    set(key: string, value: string): void;
    saveAsync(callback: (result: AsyncResult<void>) => void): void;
  }

  interface UserProfile {
    displayName: string;
    emailAddress: string;
  }
} 