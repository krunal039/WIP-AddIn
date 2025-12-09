import DebugService from './DebugService';

export enum OfficeMode {
  COMPOSE = 'compose',
  READ = 'read',
  UNKNOWN = 'unknown'
}

class OfficeModeService {
  private static instance: OfficeModeService;

  private constructor() {}

  public static getInstance(): OfficeModeService {
    if (!OfficeModeService.instance) {
      OfficeModeService.instance = new OfficeModeService();
    }
    return OfficeModeService.instance;
  }

  /**
   * Get the current Office.js mode
   * @returns OfficeMode - COMPOSE, READ, or UNKNOWN
   */
  public getCurrentMode(): OfficeMode {
    try {
      const item = Office.context.mailbox.item;
      
      // Check if we're in compose mode by looking for saveAsync method
      if (item.itemType === Office.MailboxEnums.ItemType.Message && (item as any).saveAsync) {
        DebugService.debug('Office mode detected: COMPOSE');
        return OfficeMode.COMPOSE;
      }
      
      // Check if we're in read mode
      if (item.itemType === Office.MailboxEnums.ItemType.Message && !(item as any).saveAsync) {
        DebugService.debug('Office mode detected: READ');
        return OfficeMode.READ;
      }
      
      DebugService.debug('Office mode detected: UNKNOWN');
      return OfficeMode.UNKNOWN;
    } catch (error) {
      DebugService.error('Failed to detect Office mode:', error);
      return OfficeMode.UNKNOWN;
    }
  }

  /**
   * Check if we're currently in compose (draft) mode
   * @returns boolean - true if in compose mode
   */
  public isComposeMode(): boolean {
    return this.getCurrentMode() === OfficeMode.COMPOSE;
  }

  /**
   * Check if we're currently in read mode
   * @returns boolean - true if in read mode
   */
  public isReadMode(): boolean {
    return this.getCurrentMode() === OfficeMode.READ;
  }

  /**
   * Check if we're dealing with a message (email) item
   * @returns boolean - true if it's a message item
   */
  public isMessageItem(): boolean {
    try {
      const item = Office.context.mailbox.item;
      return item.itemType === Office.MailboxEnums.ItemType.Message;
    } catch (error) {
      DebugService.error('Failed to check if item is message:', error);
      return false;
    }
  }

  /**
   * Get a human-readable description of the current mode
   * @returns string - Description of current mode
   */
  public getModeDescription(): string {
    const mode = this.getCurrentMode();
    switch (mode) {
      case OfficeMode.COMPOSE:
        return 'Compose Mode (Draft Email)';
      case OfficeMode.READ:
        return 'Read Mode (Sent Email)';
      case OfficeMode.UNKNOWN:
        return 'Unknown Mode';
      default:
        return 'Unknown Mode';
    }
  }
}

export default OfficeModeService.getInstance(); 