import DebugService from './DebugService';
import OfficeIdConverterService from './OfficeIdConverterService';
import OfficeModeService from './OfficeModeService';

export interface ForwardEmailData {
  emailId: string;
  sharedMailbox: string;
  uwwbID: string;
  internetMessageId?: string;
}

class GraphEmailService {
  private static instance: GraphEmailService;

  private constructor() {}

  public static getInstance(): GraphEmailService {
    if (!GraphEmailService.instance) {
      GraphEmailService.instance = new GraphEmailService();
    }
    return GraphEmailService.instance;
  }

  private getAuthHeaders(graphToken: string): Record<string, string> {
    return {
      'Authorization': `Bearer ${graphToken}`,
      'Content-Type': 'application/json'
    };
  }

  public async forwardEmailWithGraphToken(graphToken: string, data: ForwardEmailData): Promise<void> {
    try {
      DebugService.service('GraphEmailService', 'forwardEmailWithGraphToken started');
      DebugService.object('Forward parameters', data);
      DebugService.debug(`Starting email forwarding for emailId: ${data.emailId}`);
      
      const headers = this.getAuthHeaders(graphToken);
      
      // Step 1: Get the original email with attachments (with retry for timing)
      DebugService.debug('Step 1: Getting original email...');
      const originalEmail = await this.getEmailWithRetry(graphToken, data.emailId);
      
      // Only add delay for draft emails (not for sent emails)
      const isDraft = OfficeModeService.isComposeMode();
      if (isDraft) {
        DebugService.debug('Draft email detected - waiting 5 seconds before creating draft...');
        await new Promise(resolve => setTimeout(resolve, 5000));
      }
      
      // Step 2: Create draft with attachments
      DebugService.debug('Step 2: Creating draft...');
      const draftBody: { 
        subject: string; 
        body: any; 
        toRecipients: any[]; 
        attachments?: any[]; 
      } = {
        subject: `Ingestion Requested(${data.uwwbID}): ${originalEmail.subject || 'No Subject'}`,
        body: originalEmail.body || {
          contentType: "HTML",
          content: "<p>Email forwarded via Workbench</p>"
        },
        toRecipients: [
          {
            emailAddress: {
              address: data.sharedMailbox
            }
          }
        ]
      };

      // Add attachments if they exist
      if (originalEmail.hasAttachments && originalEmail.attachments && originalEmail.attachments.length > 0) {
        const attachments = [];
        for (const attachment of originalEmail.attachments) {
          if (attachment['@odata.type'] === '#microsoft.graph.fileAttachment') {
            attachments.push({
              '@odata.type': '#microsoft.graph.fileAttachment',
              name: attachment.name,
              contentType: attachment.contentType,
              contentBytes: attachment.contentBytes,
              size: attachment.size
            });
          }
        }
        if (attachments.length > 0) {
          draftBody.attachments = attachments;
        }
      }

      // Step 3: Create draft
      const draftUrl = "https://graph.microsoft.com/v1.0/me/messages";
      DebugService.api('POST', draftUrl);
      DebugService.object('Draft body', draftBody);

      const draftRes = await fetch(draftUrl, {
        method: "POST",
        headers: {
          ...headers,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify(draftBody)
      });

      if (!draftRes.ok) {
        const err = await draftRes.json();
        throw new Error(`Failed to create draft: ${err.error?.message}`);
      }

      const draft = await draftRes.json();
      DebugService.debug(`Draft created successfully: ${draft.id}`);

      // Only add delay for draft emails before sending
      if (isDraft) {
        DebugService.debug('Draft email detected - waiting 5 seconds before sending...');
        await new Promise(resolve => setTimeout(resolve, 5000));
      }

      // Step 4: Send the draft
      DebugService.debug('Step 3: Sending draft...');
      const sendUrl = `https://graph.microsoft.com/v1.0/me/messages/${draft.id}/send`;
      DebugService.api('POST', sendUrl);

      const sendRes = await fetch(sendUrl, {
        method: "POST",
        headers
      });

      if (!sendRes.ok) {
        const err = await sendRes.json();
        throw new Error(`Failed to send draft: ${err.error?.message}`);
      }

      DebugService.debug('Email forwarded successfully with attachments');
      DebugService.email('Email forwarded successfully with attachments');
    } catch (error) {
      DebugService.errorWithStack('Email forwarding failed', error as Error);
      throw error;
    }
  }

  // Helper method to get email with retry for timing issues
  private async getEmailWithRetry(graphToken: string, emailId: string, retry = 0): Promise<any> {
    const headers = this.getAuthHeaders(graphToken);
    
    // Convert Exchange ID to REST ID for Graph API
    const restId = await OfficeIdConverterService.convertToRestId(emailId);
    
    // Debug logging to understand ID conversion issues
    DebugService.debug(`ID Conversion - Original: "${emailId}", Converted: "${restId}"`);
    DebugService.debug(`ID contains special chars - Original: ${/[\/\\]/.test(emailId)}, Converted: ${/[\/\\]/.test(restId)}`);
    
    // Check if conversion failed or returned invalid ID
    if (!restId || restId === emailId) {
      DebugService.warn('ID conversion may have failed - using original emailId');
      // If conversion failed, the service should have logged a warning
      // We'll proceed with the original ID and let the API call fail gracefully
    }
    const emailUrl = `https://graph.microsoft.com/v1.0/me/messages/${restId}?$expand=attachments`;
    
    try {
      DebugService.api('GET', emailUrl);
      DebugService.debug(`Attempting to get email ${restId} (original: ${emailId}) (attempt ${retry + 1}/10)`);
      
      const emailRes = await fetch(emailUrl, {
        method: "GET",
        headers,
      });
      
      if (!emailRes.ok) {
        const err = await emailRes.json();
        
        if (err.error?.code === 'RequestBroker--ParseUri' && (emailRes.status === 500 || emailRes.status === 400)) {
          DebugService.error(`Graph API returned 500 error for emailId: ${emailId}, restId: ${restId}`);
          DebugService.error('500 error details:', err);
          
          // If this is the first attempt and we have a 500 error, try internetMessageId fallback first
          if (retry === 0) {
            DebugService.warn('500 error on first attempt - trying internetMessageId fallback method');
            try {
              // Try to get internetMessageId from the current context
              const { getInternetMessageId } = await import('../utils/emailHelpers');
              const internetMessageId = await getInternetMessageId(Office.context.mailbox.item as any);
              
              if (internetMessageId) {
                DebugService.debug('Found internetMessageId, attempting search:', internetMessageId);
                return await this.getEmailByInternetMessageId(graphToken, internetMessageId);
              } else {
                DebugService.warn('No internetMessageId available, trying folder search fallback');
                // Fall back to folder-specific search for drafts only
                if (OfficeModeService.isComposeMode()) {
                  return await this.getEmailBySearchFallback(graphToken, emailId);
                }
              }
            } catch (fallbackError) {
              DebugService.warn('InternetMessageId fallback also failed:', fallbackError);
              // Continue with normal retry logic
            }
          }
        }

        // Handle specific error cases
        if (emailRes.status === 404 && retry < 10) {
          const delay = 5000; // 5 seconds between each retry
          DebugService.warn(`Email not found in Graph API, retrying in ${delay}ms... (attempt ${retry + 1}/10)`);
          await new Promise(resolve => setTimeout(resolve, delay));
          return this.getEmailWithRetry(graphToken, emailId, retry + 1);
        }
        
        // Handle 400 errors with malformed ID
        if (emailRes.status === 400 && err.error?.code === 'ErrorInvalidIdMalformed') {
          DebugService.error(`Malformed ID error - Original: "${emailId}", Converted: "${restId}"`);
          DebugService.error('400 Malformed ID error details:', err);
          
          // Try search fallback for malformed ID errors only if we're in draft mode
          if (retry === 0 && OfficeModeService.isComposeMode()) {
            DebugService.warn('Malformed ID error in draft mode - trying search fallback method');
            try {
              return await this.getEmailBySearchFallback(graphToken, emailId);
            } catch (fallbackError) {
              DebugService.warn('Search fallback also failed:', fallbackError);
              // Continue with normal retry logic
            }
          }
        }
        
        // Handle 500 errors that might be related to special characters in email ID
        
        
        DebugService.error(`Failed to get email after ${retry + 1} attempts:`, err.error?.message);
        throw new Error(`Failed to get email: ${err.error?.message} (Status: ${emailRes.status})`);
      }
      
      const originalEmail = await emailRes.json();
      DebugService.debug(`Email retrieved successfully on attempt ${retry + 1}`);
      DebugService.debug('Email retrieved successfully');
      return originalEmail;
      
    } catch (error) {
      if (retry < 10) {
        const delay = 5000; // 5 seconds between each retry
        DebugService.warn(`Failed to get email, retrying in ${delay}ms... (attempt ${retry + 1}/10)`, error);
        await new Promise(resolve => setTimeout(resolve, delay));
        return this.getEmailWithRetry(graphToken, emailId, retry + 1);
      }
      
      DebugService.error(`Failed to get email after ${retry + 1} attempts:`, error);
      throw error;
    }
  }

  /**
   * Alternative method to get email using internetMessageId
   * This method searches for emails using the internetMessageId which is more stable
   */
  private async getEmailByInternetMessageId(
    graphToken: string, 
    internetMessageId: string, 
    retry = 0
  ): Promise<any> {
    const headers = this.getAuthHeaders(graphToken);
    
    try {
      DebugService.debug(`Attempting internetMessageId search for: ${internetMessageId} (attempt ${retry + 1}/3)`);
      
      // Search for the email using internetMessageId
      const searchUrl = `https://graph.microsoft.com/v1.0/me/messages?$filter=internetMessageId eq '${internetMessageId}'&$expand=attachments&$top=1`;
      DebugService.api('GET', searchUrl);
      
      const searchRes = await fetch(searchUrl, {
        method: "GET",
        headers,
      });
      
      if (searchRes.ok) {
        const searchResult = await searchRes.json();
        if (searchResult.value && searchResult.value.length > 0) {
          DebugService.debug('Email found via internetMessageId search');
          return searchResult.value[0];
        }
      }
      
      throw new Error('Email not found via internetMessageId search');
      
    } catch (error) {
      if (retry < 3) {
        const delay = 2000; // 2 seconds between retries for search
        DebugService.warn(`InternetMessageId search failed, retrying in ${delay}ms... (attempt ${retry + 1}/3)`, error);
        await new Promise(resolve => setTimeout(resolve, delay));
        return this.getEmailByInternetMessageId(graphToken, internetMessageId, retry + 1);
      }
      
      DebugService.error(`InternetMessageId search failed after ${retry + 1} attempts:`, error);
      throw error;
    }
  }

  /**
   * Alternative method to get email when direct ID access fails
   * This method tries to find the email using search criteria
   */
  private async getEmailBySearchFallback(
    graphToken: string, 
    emailId: string, 
    retry = 0
  ): Promise<any> {
    const headers = this.getAuthHeaders(graphToken);
    
    try {
      DebugService.debug(`Attempting search fallback for emailId: ${emailId} (attempt ${retry + 1}/3)`);
      
      // Try to search for the email in drafts folder first
      const searchUrl = `https://graph.microsoft.com/v1.0/me/mailFolders('drafts')/messages?$filter=id eq '${emailId}'&$top=1`;
      DebugService.api('GET', searchUrl);
      
      const searchRes = await fetch(searchUrl, {
        method: "GET",
        headers,
      });
      
      if (searchRes.ok) {
        const searchResult = await searchRes.json();
        if (searchResult.value && searchResult.value.length > 0) {
          DebugService.debug('Email found via search fallback in drafts folder');
          return searchResult.value[0];
        }
      }
      
      // If not found in drafts, try inbox
      const inboxSearchUrl = `https://graph.microsoft.com/v1.0/me/mailFolders('inbox')/messages?$filter=id eq '${emailId}'&$top=1`;
      DebugService.api('GET', inboxSearchUrl);
      
      const inboxSearchRes = await fetch(inboxSearchUrl, {
        method: "GET",
        headers,
      });
      
      if (inboxSearchRes.ok) {
        const inboxSearchResult = await inboxSearchRes.json();
        if (inboxSearchResult.value && inboxSearchResult.value.length > 0) {
          DebugService.debug('Email found via search fallback in inbox');
          return inboxSearchResult.value[0];
        }
      }
      
      throw new Error('Email not found via search fallback');
      
    } catch (error) {
      if (retry < 3) {
        const delay = 2000; // 2 seconds between retries for search
        DebugService.warn(`Search fallback failed, retrying in ${delay}ms... (attempt ${retry + 1}/3)`, error);
        await new Promise(resolve => setTimeout(resolve, delay));
        return this.getEmailBySearchFallback(graphToken, emailId, retry + 1);
      }
      
      DebugService.error(`Search fallback failed after ${retry + 1} attempts:`, error);
      throw error;
    }
  }

  public async forwardEmailBySearchFallback(
    graphToken: string,
    subject: string,
    sender: string,
    createdDateTime: string,
    uwwbID: string,
    sharedMailbox: string
  ): Promise<void> {
    try {
      DebugService.service('GraphEmailService', 'forwardEmailBySearchFallback started');
      DebugService.object('Search fallback parameters', { subject, sender, createdDateTime, uwwbID, sharedMailbox });
      
      const headers = this.getAuthHeaders(graphToken);
      
      // Search in drafts folder with subject and createdDate filtering
      const escapedSubject = subject.replace(/'/g, "''");
      const escapedSender = sender.replace(/'/g, "''");
      
      // Use $filter with subject, sender, and createdDateTime for exact match
      const filterQuery = `subject eq '${escapedSubject}' and from/emailAddress/address eq '${escapedSender}' and createdDateTime ge ${createdDateTime}`;
      const searchUrl = `https://graph.microsoft.com/v1.0/me/mailFolders('drafts')/messages?$filter=${encodeURIComponent(filterQuery)}&$orderby=createdDateTime desc&$top=1`;
      
      DebugService.api('GET', searchUrl);
      DebugService.debug('Search fallback query:', filterQuery);
      
      const searchResult = await this.searchWithRetry(graphToken, searchUrl);
      
      if (!searchResult.value || searchResult.value.length === 0) {
        throw new Error('Email not found in drafts folder with subject and date criteria');
      }
      
      const emailId = searchResult.value[0].id;
      DebugService.debug('Found email ID by search fallback:', emailId);
      
      // Forward the found email
      await this.forwardEmailWithGraphToken(graphToken, {
        emailId,
        sharedMailbox,
        uwwbID
      });
      
    } catch (error) {
      DebugService.errorWithStack('Email forwarding by search fallback failed', error as Error);
      throw error;
    }
  }

  public async forwardEmailByConversationId(
    graphToken: string,
    conversationId: string,
    uwwbID: string,
    sharedMailbox: string
  ): Promise<void> {
    try {
      DebugService.service('GraphEmailService', 'forwardEmailByConversationId started');
      DebugService.object('ConversationId parameters', { conversationId, uwwbID, sharedMailbox });
      
      const headers = this.getAuthHeaders(graphToken);
      
      // Search for email in drafts folder with retry for timing
      const searchUrl = `https://graph.microsoft.com/v1.0/me/mailFolders('drafts')/messages?$filter=conversationId eq '${conversationId}'&$top=1`;
      DebugService.api('GET', searchUrl);
      
      const searchResult = await this.searchWithRetry(graphToken, searchUrl);
      
      if (!searchResult.value || searchResult.value.length === 0) {
        throw new Error('Email not found in drafts folder with conversationId');
      }
      
      const emailId = searchResult.value[0].id;
      DebugService.debug('Found email ID by conversationId:', emailId);
      
      // Forward the found email
      await this.forwardEmailWithGraphToken(graphToken, { emailId, sharedMailbox, uwwbID });
      
    } catch (error) {
      DebugService.errorWithStack('Email forwarding by conversationId failed', error as Error);
      throw error;
    }
  }

  // Helper method to search with retry for timing issues
  private async searchWithRetry(graphToken: string, searchUrl: string, retry = 0): Promise<any> {
    const headers = this.getAuthHeaders(graphToken);
    
    try {
      DebugService.debug(`Attempting to search (attempt ${retry + 1}/10)`);
      
      const searchRes = await fetch(searchUrl, {
        method: "GET",
        headers,
      });
      
      if (!searchRes.ok) {
        const err = await searchRes.json();
        
        // If it's a 404 or empty results and we're in retry mode, wait and retry
        if ((searchRes.status === 404 || (searchRes.status === 200 && err.value?.length === 0)) && retry < 10) {
          const delay = 5000; // 5 seconds between each retry
          DebugService.warn(`Search returned no results, retrying in ${delay}ms... (attempt ${retry + 1}/10)`);
          await new Promise(resolve => setTimeout(resolve, delay));
          return this.searchWithRetry(graphToken, searchUrl, retry + 1);
        }
        
        DebugService.error(`Failed to search after ${retry + 1} attempts:`, err.error?.message);
        throw new Error(`Failed to search for email: ${err.error?.message}`);
      }
      
      const searchResult = await searchRes.json();
      DebugService.debug(`Search completed successfully on attempt ${retry + 1}`);
      DebugService.debug('Search completed successfully');
      return searchResult;
      
    } catch (error) {
      if (retry < 10) {
        const delay = 5000; // 5 seconds between each retry
        DebugService.warn(`Search failed, retrying in ${delay}ms... (attempt ${retry + 1}/10)`, error);
        await new Promise(resolve => setTimeout(resolve, delay));
        return this.searchWithRetry(graphToken, searchUrl, retry + 1);
      }
      DebugService.error(`Failed to search after ${retry + 1} attempts:`, error);
      throw error;
    }
  }
}

export default GraphEmailService.getInstance(); 