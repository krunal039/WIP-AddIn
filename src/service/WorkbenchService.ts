import DebugService from './DebugService';
import LoggingService from './LoggingService';
import PlacementApiService, { PlacementRequestData } from './PlacementApiService';
import { EmailConverterService } from './EmailConverterService';
import GraphEmailService from './GraphEmailService';
import OfficeModeService from './OfficeModeService';
import OfficeIdConverterService from './OfficeIdConverterService';
import { getSender, getSubject, getCreatedDate, detectSharedMailbox } from '../utils/emailHelpers';
import { stampEmailWithWorkbenchId } from '../utils/emailStamping';
import { showWorkbenchNotificationBanner } from '../utils/outlookNotification';
import { environment } from '../config/environment';

export interface WorkbenchSubmissionResult {
  success: boolean;
  placementId?: string;
  error?: string;
  forwardingFailed?: boolean;
  forwardingFailedReason?: string;
  lastPlacementId?: string;
  lastGraphItemId?: string;
  lastSharedMailbox?: string;
}

export class WorkbenchService {
  private static instance: WorkbenchService;
  
  // Environment variables
  private get sharedMailbox(): string {
    return environment.CYBER_MRSNA_MAILBOX || environment.DEFAULT_SHARED_MAILBOX || '';
  }

  private constructor() {}

  public static getInstance(): WorkbenchService {
    if (!WorkbenchService.instance) {
      WorkbenchService.instance = new WorkbenchService();
    }
    return WorkbenchService.instance;
  }

  public async submitPlacement(
    apiToken: string,
    graphToken: string,
    item: Office.Item,
    productCode: string,
    sendCopyToCyberAdmin: boolean
  ): Promise<WorkbenchSubmissionResult> {
    try {
      DebugService.section('Starting Placement Submission');
      DebugService.object('Submission parameters', { productCode, sendCopyToCyberAdmin });

      // Continue with normal submission flow
      return await this.processSubmission(apiToken, graphToken, item, productCode, sendCopyToCyberAdmin);
    } catch (error) {
      const errObj = error instanceof Error ? error : new Error(String(error));
      await LoggingService.logError(errObj, 'submitPlacement', { 
        user: Office.context.mailbox.userProfile.emailAddress 
      });
      return { success: false, error: errObj.message };
    }
  }

  public async retryForward(
    graphToken: string,
    placementId: string,
    graphItemId: string,
    sharedMailbox: string
  ): Promise<WorkbenchSubmissionResult> {
    try {
      DebugService.section('Retrying Email Forward');
      DebugService.object('Retry parameters', { placementId, graphItemId, sharedMailbox });

      // If we have a valid graphItemId, use the normal method
      if (graphItemId && graphItemId !== 'SEARCH_METHOD') {
        await GraphEmailService.forwardEmailWithGraphToken(graphToken, {
          emailId: graphItemId,
          uwwbID: placementId,
          sharedMailbox: sharedMailbox,
        });
      } else {
        // For draft emails, try to get fresh itemId first, then conversationId, then search
        const item = Office.context.mailbox.item;
        let currentItemId = (item as any).itemId;
        
        // If no itemId and we're in draft mode, try to save and get it
        if (!currentItemId && OfficeModeService.isComposeMode()) {
          try {
            DebugService.email('Getting fresh itemId for retry forwarding (draft mode)');
            currentItemId = await this.saveAndGetItemId(item);
            DebugService.email('Got fresh itemId for retry:', currentItemId);
          } catch (saveError) {
            DebugService.warn('Failed to get fresh itemId, trying conversationId:', saveError);
          }
        }
        
        if (currentItemId) {
          DebugService.email('Using fresh itemId for retry forwarding');
          await GraphEmailService.forwardEmailWithGraphToken(graphToken, {
            emailId: currentItemId,
            uwwbID: placementId,
            sharedMailbox: sharedMailbox,
          });
        } else {
          // Try conversationId as fallback
          const conversationId = (item as any).conversationId;
          if (conversationId) {
            DebugService.email('Using conversationId for retry forwarding');
            await GraphEmailService.forwardEmailByConversationId(
              graphToken,
              conversationId,
              placementId,
              sharedMailbox
            );
          } else {
            // Final fallback: search by subject and createdDate
            DebugService.debug('Using search fallback for retry with subject and createdDate');
            try {
              const emailSubject = await getSubject(item as any);
              const emailSender = await getSender(item as any);
              const emailCreatedDateTime = await getCreatedDate(item as any);
              
              DebugService.debug('Search fallback parameters:', { emailSubject, emailSender, emailCreatedDateTime });
              
              await GraphEmailService.forwardEmailBySearchFallback(
                graphToken,
                emailSubject,
                emailSender,
                emailCreatedDateTime,
                placementId,
                sharedMailbox
              );
            } catch (error) {
              DebugService.error('Search fallback failed:', error);
              const errorMessage = error instanceof Error ? error.message : String(error);
              throw new Error(`Search fallback failed: ${errorMessage}`);
            }
          }
        }
      }
      
      DebugService.email('Forwarding retry successful');
      return { success: true };
    } catch (error) {
      const errObj = error instanceof Error ? error : new Error(String(error));
      await LoggingService.logError(errObj, 'retryForward', { 
        user: Office.context.mailbox.userProfile.emailAddress 
      });
      return { 
        success: false, 
        error: errObj.message,
        forwardingFailed: true,
        forwardingFailedReason: errObj.message
      };
    }
  }

  private async saveAndGetItemId(item: Office.Item | undefined, retry = 0): Promise<string> {
    if (!item) {
      throw new Error('Item is undefined');
    }
    return new Promise((resolve, reject) => {
      (item as any).saveAsync((result: any) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const exchangeId = result.value;
          if (exchangeId) {
            DebugService.debug('Got Exchange ID from saveAsync:', exchangeId);
            
            // Convert Exchange ID to REST ID for Graph API
            (Office.context.mailbox as any).convertToRestId(
              [exchangeId],
              Office.MailboxEnums.RestVersion.v2_0,
              (convertResult: Office.AsyncResult<string[]>) => {
                if (convertResult.status === Office.AsyncResultStatus.Succeeded) {
                  const restId = convertResult.value?.[0];
                  DebugService.debug('Converted Exchange ID to REST ID:', { exchangeId, restId });
                  
                  // Add a small delay to ensure itemId is properly set on the item
                  setTimeout(() => {
                    resolve(restId || exchangeId);
                  }, 100);
                } else {
                  DebugService.warn('Failed to convert Exchange ID to REST ID, using original:', convertResult.error);
                  // Fallback: use original ID if conversion fails
                  setTimeout(() => {
                    resolve(exchangeId);
                  }, 100);
                }
              }
            );
          } else if (retry < 5) {
            // Enhanced retry with exponential backoff
            const delay = Math.min(1000 * Math.pow(2, retry), 10000); // Max 10s
            DebugService.warn(`Item ID was null, retrying in ${delay}ms... (attempt ${retry + 1}/5)`);
            setTimeout(() => {
              this.saveAndGetItemId(item, retry + 1).then(resolve).catch(reject);
            }, delay);
          } else {
            DebugService.error('Item ID is still null after 5 retries');
            reject(new Error('Failed to get itemId after 5 retries'));
          }
        } else {
          DebugService.error('Failed to save draft:', result.error);
          reject(result.error);
        }
      });
    });
  }

  // Method to trigger early save when user interacts with UI
  public async attemptEarlySave(item: Office.Item): Promise<void> {
    try {
      // Only attempt save if we're in compose mode (draft) and don't have itemId
      if (OfficeModeService.isComposeMode() && !(item as any).itemId) {
        DebugService.debug('Attempting early save on user interaction (draft mode)');
        await this.saveAndGetItemId(item);
        DebugService.debug('Early draft save successful');
      }
    } catch (error) {
      DebugService.warn('Early save failed, will retry later:', error);
      // Don't throw - let user continue with their workflow
    }
  }

  private async processSubmission(
    apiToken: string,
    graphToken: string,
    item: Office.Item,
    productCode: string,
    sendCopyToCyberAdmin: boolean
  ): Promise<WorkbenchSubmissionResult> {
    // Step 1: Save draft if not already saved (compose mode only) and get itemId
    let itemId = (item as any).itemId;
    if (OfficeModeService.isComposeMode() && !itemId) {
      try {
        itemId = await this.saveAndGetItemId(item);
        DebugService.debug('Initial draft saved successfully with itemId:', itemId);
      } catch (error) {
        LoggingService.logError(error as Error, 'Saving draft email');
        throw error;
      }
    } else if (itemId) {
      // Convert existing itemId to REST ID if it hasn't been converted yet
      try {
        const restId = await OfficeIdConverterService.convertToRestId(itemId);
        if (restId !== itemId) {
          DebugService.debug('Converted existing itemId to REST ID:', { itemId, restId });
          itemId = restId;
        }
      } catch (error) {
        DebugService.warn('Failed to convert existing itemId, using original:', error);
      }
    }

    DebugService.subsection('Item ID Debug');
    DebugService.debug('itemId:', itemId);
    DebugService.debug('item type:', (item as any).itemType);
    DebugService.debug('item has itemId:', !!(item as any).itemId);
    DebugService.debug('item has saveAsync:', !!(item as any).saveAsync);

    // Step 2: Log API token for debugging
    DebugService.auth('Using provided API token for placement submission');
    await LoggingService.logApiToken(apiToken);

    // Step 3: Convert email and submit placement
    DebugService.placement('Converting email and submitting placement');
    const converter = new EmailConverterService();
    const emlData = await converter.convertEmailToEml(item);

    // Use util helpers for subject, sender, and created date (parallel execution)
    const [emailSender, emailSubject, emailReceivedDateTime] = await Promise.all([
      getSender(item as any),
      getSubject(item as any),
      getCreatedDate(item as any)
    ]);

    // Ensure all values are strings to prevent .replace() errors
    const validatedEmailSender = String(emailSender || '');
    const validatedEmailSubject = String(emailSubject || '');
    const validatedEmailReceivedDateTime = String(emailReceivedDateTime || new Date().toISOString());

    DebugService.debug('Validated email data:', {
      emailSender: validatedEmailSender,
      emailSubject: validatedEmailSubject,
      emailReceivedDateTime: validatedEmailReceivedDateTime
    });

    const data: PlacementRequestData = {
      productCode,
      emailSender: validatedEmailSender,
      emailSubject: validatedEmailSubject,
      emailReceivedDateTime: validatedEmailReceivedDateTime,
      emlContent: emlData.content,
    };

    // Use the service with provided API token
    const placementData = await PlacementApiService.submitPlacementRequest(apiToken, data);
    await LoggingService.logPlacementRequest(placementData.placementId, emailSender);

    // Step 4: Stamp the email with workbench ID BEFORE forwarding
    DebugService.debug('Stamping email with workbench ID before forwarding');
    if (item) {
      await stampEmailWithWorkbenchId(item, placementData.placementId, DebugService);
    }

    // Step 4.5: Show Outlook notification banner with WBID
    DebugService.debug('Showing Outlook notification banner with WBID');
    if (item) {
      await showWorkbenchNotificationBanner(item, placementData.placementId, DebugService);
    }

    // Step 5: Handle email forwarding only if needed
    if (productCode === "20001" && sendCopyToCyberAdmin) {
      DebugService.email('Cyber product with forwarding enabled - using provided Graph token');
      
      // Try to get itemId for forwarding
      let currentItemId = itemId;

     if (!currentItemId && OfficeModeService.isComposeMode()) {
        DebugService.email('Item ID not available - attempting to save draft first');
        
        try {
          // Use the improved saveAndGetItemId function (draft mode only)
          currentItemId = await this.saveAndGetItemId(item);
          DebugService.email('Draft saved successfully with itemId:', currentItemId);
        } catch (saveError) {
          DebugService.errorWithStack('Failed to save draft for forwarding', saveError as Error);
          return {
            success: true,
            placementId: placementData.placementId,
            forwardingFailed: true,
            forwardingFailedReason: saveError instanceof Error ? saveError.message : 'Failed to save draft',
            lastPlacementId: placementData.placementId,
            lastSharedMailbox: this.sharedMailbox
          };
        }
      }

      // Forward the email using the provided Graph token
      try {
        DebugService.email('Forwarding email using Graph token');
        
        // Detect if the current email is from a shared mailbox
        const { isShared, mailboxEmail } = await detectSharedMailbox(item as any);
        DebugService.debug(`Detected mailbox type - isShared: ${isShared}, mailboxEmail: ${mailboxEmail}`);
        
        // We should now have a valid itemId for all cases
        if (currentItemId) {
          // Get internetMessageId for better fallback handling
          const { getInternetMessageId } = await import('../utils/emailHelpers');
          const internetMessageId = await getInternetMessageId(item as any);
          
          await GraphEmailService.forwardEmailWithGraphToken(graphToken, {
            emailId: currentItemId,
            uwwbID: placementData.placementId,
            sharedMailbox: this.sharedMailbox,
            internetMessageId: internetMessageId || undefined,
          }, mailboxEmail, isShared);
        } else {
          // Fallback: if somehow we still don't have itemId, try conversationId
          const conversationId = (item as any).conversationId;
          if (conversationId) {
            DebugService.debug('Using conversationId as fallback:', conversationId);
            // Get internetMessageId for better fallback handling
            const { getInternetMessageId } = await import('../utils/emailHelpers');
            const internetMessageId = await getInternetMessageId(item as any);
            
            await GraphEmailService.forwardEmailByConversationId(
              graphToken,
              conversationId,
              placementData.placementId,
              this.sharedMailbox
            );
          } else {
            // Final fallback: search by subject and createdDate
            DebugService.debug('Using search fallback with subject and createdDate');
            try {
              const emailSubject = await getSubject(item as any);
              const emailSender = await getSender(item as any);
              const emailCreatedDateTime = await getCreatedDate(item as any);
              
              DebugService.debug('Search fallback parameters:', { emailSubject, emailSender, emailCreatedDateTime });
              
              await GraphEmailService.forwardEmailBySearchFallback(
                graphToken,
                emailSubject,
                emailSender,
                emailCreatedDateTime,
                placementData.placementId,
                this.sharedMailbox
              );
            } catch (error) {
              DebugService.error('Search fallback failed:', error);
              const errorMessage = error instanceof Error ? error.message : String(error);
              throw new Error(`Search fallback failed: ${errorMessage}`);
            }
          }
        }
        
        DebugService.email('Email forwarding successful');
        return {
          success: true,
          placementId: placementData.placementId,
          forwardingFailed: false
        };
      } catch (forwardError) {
        DebugService.errorWithStack('Email forwarding failed', forwardError as Error);
        return {
          success: true,
          placementId: placementData.placementId,
          forwardingFailed: true,
          forwardingFailedReason: forwardError instanceof Error ? forwardError.message : 'Forwarding failed',
          lastPlacementId: placementData.placementId,
          lastGraphItemId: currentItemId,
          lastSharedMailbox: this.sharedMailbox
        };
      }
    }

    // Success without forwarding
    return {
      success: true,
      placementId: placementData.placementId,
      forwardingFailed: false
    };
  }
}

export default WorkbenchService.getInstance(); 