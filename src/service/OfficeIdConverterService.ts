import DebugService from './DebugService';

/**
 * Service for converting Exchange IDs to REST IDs for Graph API calls
 * Handles the conversion of Office.js Exchange IDs to Microsoft Graph REST IDs
 */
class OfficeIdConverterService {
  private static instance: OfficeIdConverterService;

  private constructor() {}

  public static getInstance(): OfficeIdConverterService {
    if (!OfficeIdConverterService.instance) {
      OfficeIdConverterService.instance = new OfficeIdConverterService();
    }
    return OfficeIdConverterService.instance;
  }

  /**
   * Convert Exchange ID to REST ID for Graph API calls
   * @param exchangeId The Exchange ID from Office.js
   * @returns Promise<string> The REST ID for Graph API
   */
  public async convertToRestId(exchangeId: string): Promise<string> {
    return new Promise((resolve, reject) => {
      try {
        (Office.context.mailbox as any).convertToRestId(
          [exchangeId],
          Office.MailboxEnums.RestVersion.v2_0,
          (result: Office.AsyncResult<string[]>) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              const restId = result.value?.[0];
              DebugService.debug('Converted Exchange ID to REST ID:', { exchangeId, restId });
              resolve(restId || exchangeId);
            } else {
              DebugService.warn('Failed to convert Exchange ID to REST ID:', result.error);
              // Fallback: return the original ID if conversion fails
              resolve(exchangeId);
            }
          }
        );
      } catch (error) {
        DebugService.warn('Error converting Exchange ID to REST ID:', error);
        // Fallback: return the original ID if conversion fails
        resolve(exchangeId);
      }
    });
  }

  /**
   * Convert multiple Exchange IDs to REST IDs
   * @param exchangeIds Array of Exchange IDs
   * @returns Promise<string[]> Array of REST IDs
   */
  public async convertMultipleToRestIds(exchangeIds: string[]): Promise<string[]> {
    return new Promise((resolve, reject) => {
      try {
        (Office.context.mailbox as any).convertToRestId(
          exchangeIds,
          Office.MailboxEnums.RestVersion.v2_0,
          (result: Office.AsyncResult<string[]>) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              const restIds = result.value || exchangeIds;
              DebugService.debug('Converted multiple Exchange IDs to REST IDs:', { 
                exchangeIds, 
                restIds 
              });
              resolve(restIds);
            } else {
              DebugService.warn('Failed to convert multiple Exchange IDs to REST IDs:', result.error);
              // Fallback: return the original IDs if conversion fails
              resolve(exchangeIds);
            }
          }
        );
      } catch (error) {
        DebugService.warn('Error converting multiple Exchange IDs to REST IDs:', error);
        // Fallback: return the original IDs if conversion fails
        resolve(exchangeIds);
      }
    });
  }
}

export default OfficeIdConverterService.getInstance(); 