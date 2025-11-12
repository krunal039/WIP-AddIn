// Extended Office.js type definitions for better type safety

/**
 * Extended Office Item type that includes common properties
 */
export interface ExtendedOfficeItem extends Office.Item {
  itemId?: string;
  itemType?: Office.MailboxEnums.ItemType;
  subject?: Office.Subject;
  internetHeaders?: Office.InternetHeaders;
  loadCustomPropertiesAsync?: (callback: (result: Office.AsyncResult<Office.CustomProperties>) => void) => void;
  saveAsync?: (callback?: (result: Office.AsyncResult<void>) => void) => void;
  getAllInternetHeadersAsync?: (callback: (result: Office.AsyncResult<string>) => void) => void;
}

/**
 * Office Message Read type
 */
export type OfficeMessageRead = Office.MessageRead & ExtendedOfficeItem;

/**
 * Office Message Compose type
 */
export type OfficeMessageCompose = Office.MessageCompose & ExtendedOfficeItem;

/**
 * Union type for Office items
 */
export type OfficeItem = Office.Item | ExtendedOfficeItem;
