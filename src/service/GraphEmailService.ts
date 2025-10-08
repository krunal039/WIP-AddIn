import DebugService from './DebugService';
import OfficeIdConverterService from './OfficeIdConverterService';
import OfficeModeService from './OfficeModeService';
import { detectSharedMailbox } from '../utils/emailHelpers';

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

  /**
   * Debug the Graph token and user context
   */
  private async debugTokenAndUser(graphToken: string): Promise<void> {
    try {
      // Decode the token to see user information (basic JWT decode)
      const tokenParts = graphToken.split('.');
      if (tokenParts.length === 3) {
        const payload = JSON.parse(atob(tokenParts[1]));
        DebugService.debug('Token payload:', {
          sub: payload.sub,
          oid: payload.oid,
          upn: payload.upn,
          email: payload.email,
          aud: payload.aud,
          iss: payload.iss
        });
      }
      
      // Try to get user info from Graph API
      const userInfoUrl = 'https://graph.microsoft.com/v1.0/me';
      const headers = this.getAuthHeaders(graphToken);
      
      const userRes = await fetch(userInfoUrl, {
        method: 'GET',
        headers
      });
      
      if (userRes.ok) {
        const userInfo = await userRes.json();
        DebugService.debug('Graph API user info:', {
          id: userInfo.id,
          userPrincipalName: userInfo.userPrincipalName,
          mail: userInfo.mail,
          displayName: userInfo.displayName
        });
      } else {
        DebugService.warn('Failed to get user info from Graph API:', userRes.status);
      }
    } catch (error) {
      DebugService.warn('Error debugging token and user:', error);
    }
  }

  /**
   * Get the correct Graph API endpoint based on mailbox type
   * @param mailboxEmail The email address of the mailbox (personal or shared)
   * @param isShared Whether the mailbox is shared
   * @returns The Graph API endpoint
   */
  private getMailboxEndpoint(mailboxEmail: string, isShared: boolean): string {
    if (isShared) {
      return `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(mailboxEmail)}`;
    } else {
      return 'https://graph.microsoft.com/v1.0/me';
    }
  }

  public async forwardEmailWithGraphToken(graphToken: string, data: ForwardEmailData, sourceMailboxEmail?: string, isSourceShared?: boolean): Promise<void> {
    try {
      DebugService.service('GraphEmailService', 'forwardEmailWithGraphToken started');
      DebugService.object('Forward parameters', data);
      DebugService.debug(`Starting email forwarding for emailId: ${data.emailId}`);
      
      // Debug token and user context
      await this.debugTokenAndUser(graphToken);
      
      const headers = this.getAuthHeaders(graphToken);
      
      // Step 1: Get the original email with attachments (with retry for timing)
      DebugService.debug('Step 1: Getting original email...');
      const originalEmail = await this.getEmailWithRetry(graphToken, data.emailId, 0, sourceMailboxEmail, isSourceShared);
      
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
      // Use the correct endpoint for the target mailbox (where we're sending to)
      const targetMailboxEndpoint = this.getMailboxEndpoint(data.sharedMailbox, true); // Always shared for target
      const draftUrl = `${targetMailboxEndpoint}/messages`;
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
      const sendUrl = `${targetMailboxEndpoint}/messages/${draft.id}/send`;
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
  private async getEmailWithRetry(graphToken: string, emailId: string, retry = 0, mailboxEmail?: string, isShared?: boolean): Promise<any> {
    const headers = this.getAuthHeaders(graphToken);
    
    // Convert Exchange ID to REST ID for Graph API
    let restId;
    
    // Check if the ID is already a REST ID (contains no special characters)
    if (!emailId.includes('+') && !emailId.includes('/') && !emailId.includes('\\')) {
      DebugService.debug('ID appears to be already a REST ID, skipping conversion');
      restId = emailId;
    } else {
      try {
        restId = await OfficeIdConverterService.convertToRestId(emailId);
        DebugService.debug(`ID Conversion successful - Original: "${emailId}", Converted: "${restId}"`);
      } catch (conversionError) {
        DebugService.error('ID conversion failed:', conversionError);
        restId = emailId; // Fallback to original ID
      }
    }
    
    // Debug logging to understand ID conversion issues
    DebugService.debug(`ID contains special chars - Original: ${/[\/\\]/.test(emailId)}, Converted: ${/[\/\\]/.test(restId)}`);
    DebugService.debug(`Mailbox info - mailboxEmail: "${mailboxEmail}", isShared: ${isShared}`);
    
    // Check if conversion failed or returned invalid ID
    if (!restId || restId === emailId) {
      DebugService.warn('ID conversion may have failed - using original emailId');
      // If conversion failed, the service should have logged a warning
      // We'll proceed with the original ID and let the API call fail gracefully
    }
    
    // Use the correct endpoint based on mailbox type
    const baseEndpoint = mailboxEmail && isShared !== undefined 
      ? this.getMailboxEndpoint(mailboxEmail, isShared)
      : 'https://graph.microsoft.com/v1.0/me';
    
    const emailUrl = `${baseEndpoint}/messages/${restId}?$expand=attachments`;
    
    DebugService.debug(`Using endpoint: ${baseEndpoint} for email ID: ${restId}`);
    
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
                return await this.getEmailByInternetMessageId(graphToken, internetMessageId, 0, mailboxEmail, isShared);
              } else {
                DebugService.warn('No internetMessageId available, trying folder search fallback');
                // Fall back to folder-specific search for drafts only
                if (OfficeModeService.isComposeMode()) {
                  return await this.getEmailBySearchFallback(graphToken, emailId, 0, mailboxEmail, isShared);
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
          DebugService.warn(`404 Error details: ${JSON.stringify(err)}`);
          
          // Try different approaches for 404 errors
          if (retry === 0) {
            DebugService.warn('First 404 attempt - trying alternative ID approaches');
            
            // Try with original email ID (no conversion)
            try {
              DebugService.warn('Trying with original email ID (no conversion)');
              const originalEmailUrl = `https://graph.microsoft.com/v1.0/me/messages/${emailId}?$expand=attachments`;
              
              DebugService.api('GET', originalEmailUrl);
              const originalEmailRes = await fetch(originalEmailUrl, {
                method: "GET",
                headers,
              });
              
              if (originalEmailRes.ok) {
                const originalEmail = await originalEmailRes.json();
                DebugService.debug('Email retrieved successfully using original ID');
                return originalEmail;
              } else {
                DebugService.warn(`Original ID also failed with status: ${originalEmailRes.status}`);
              }
            } catch (originalError) {
              DebugService.warn('Original ID approach failed:', originalError);
            }
            
            // Try with a longer delay for timing issues
            DebugService.warn('Trying with extended delay for timing issues');
            await new Promise(resolve => setTimeout(resolve, 10000)); // 10 seconds
          }
          
          await new Promise(resolve => setTimeout(resolve, delay));
          return this.getEmailWithRetry(graphToken, emailId, retry + 1, mailboxEmail, isShared);
        }
        
        // Handle 400 errors with malformed ID
        if (emailRes.status === 400 && err.error?.code === 'ErrorInvalidIdMalformed') {
          DebugService.error(`Malformed ID error - Original: "${emailId}", Converted: "${restId}"`);
          DebugService.error('400 Malformed ID error details:', err);
          
          // Try search fallback for malformed ID errors only if we're in draft mode
          if (retry === 0 && OfficeModeService.isComposeMode()) {
            DebugService.warn('Malformed ID error in draft mode - trying search fallback method');
            try {
              return await this.getEmailBySearchFallback(graphToken, emailId, 0, mailboxEmail, isShared);
            } catch (fallbackError) {
              DebugService.warn('Search fallback also failed:', fallbackError);
              // Continue with normal retry logic
            }
          }
        }

        // Handle 400 errors with invalid mailbox item ID
        if (emailRes.status === 400 && err.error?.code === 'ErrorInvalidMailboxItemId') {
          DebugService.error(`Invalid mailbox item ID error - Item doesn't belong to targeted mailbox`);
          DebugService.error('400 Invalid Mailbox Item ID error details:', err);
          
          // This could mean:
          // 1. We're trying to access a shared mailbox email with /me/ endpoint
          // 2. The email ID is corrupted or invalid
          // 3. There's a timing issue with email ID conversion
          
          if (retry === 0) {
            DebugService.warn('Invalid mailbox item ID - trying alternative approaches');
            
            // First, try with the personal mailbox endpoint but with different ID handling
            try {
              DebugService.warn('Trying with original email ID (no conversion)');
              const originalEmailUrl = `https://graph.microsoft.com/v1.0/me/messages/${emailId}?$expand=attachments`;
              
              DebugService.api('GET', originalEmailUrl);
              const originalEmailRes = await fetch(originalEmailUrl, {
                method: "GET",
                headers,
              });
              
              if (originalEmailRes.ok) {
                const originalEmail = await originalEmailRes.json();
                DebugService.debug('Email retrieved successfully using original ID');
                return originalEmail;
              }
            } catch (originalError) {
              DebugService.warn('Original ID approach failed:', originalError);
            }
            
            // If that fails and we think it might be a shared mailbox, try shared mailbox endpoint
            if (mailboxEmail && mailboxEmail !== Office.context.mailbox.userProfile.emailAddress) {
              try {
                DebugService.warn('Trying with shared mailbox endpoint');
                const sharedEndpoint = this.getMailboxEndpoint(mailboxEmail, true);
                const sharedEmailUrl = `${sharedEndpoint}/messages/${restId}?$expand=attachments`;
                
                DebugService.api('GET', sharedEmailUrl);
                const sharedEmailRes = await fetch(sharedEmailUrl, {
                  method: "GET",
                  headers,
                });
                
                if (sharedEmailRes.ok) {
                  const sharedEmail = await sharedEmailRes.json();
                  DebugService.debug('Email retrieved successfully using shared mailbox endpoint');
                  return sharedEmail;
                }
              } catch (fallbackError) {
                DebugService.warn('Shared mailbox endpoint also failed:', fallbackError);
              }
            }
          }
        }
        
        // Handle 500 errors that might be related to special characters in email ID
        
        
        // Before throwing the final error, try one more approach
        if (retry === 9) { // Last attempt
          DebugService.warn('Final attempt - trying search-based approach');
          try {
            // Try to get the email using search instead of direct ID access
            const searchUrl = `https://graph.microsoft.com/v1.0/me/messages?$filter=id eq '${restId}'&$expand=attachments&$top=1`;
            DebugService.api('GET', searchUrl);
            
            const searchRes = await fetch(searchUrl, {
              method: "GET",
              headers,
            });
            
            if (searchRes.ok) {
              const searchResult = await searchRes.json();
              if (searchResult.value && searchResult.value.length > 0) {
                DebugService.debug('Email found via search approach');
                return searchResult.value[0];
              }
            }
            
            // If search also fails, try to get user info and compare with Office context
            DebugService.warn('Search failed, checking user context mismatch');
            try {
              const userInfoUrl = 'https://graph.microsoft.com/v1.0/me';
              const userRes = await fetch(userInfoUrl, { method: 'GET', headers });
              if (userRes.ok) {
                const userInfo = await userRes.json();
                const officeUserEmail = Office.context.mailbox.userProfile.emailAddress;
                DebugService.warn('User context comparison:', {
                  graphUserId: userInfo.id,
                  graphUserEmail: userInfo.userPrincipalName,
                  officeUserEmail: officeUserEmail,
                  emailId: emailId,
                  restId: restId
                });
                
                if (userInfo.userPrincipalName !== officeUserEmail) {
                  DebugService.error('User context mismatch detected! Graph token is for different user than Office context');
                }
              }
            } catch (userError) {
              DebugService.warn('Failed to get user info for comparison:', userError);
            }
          } catch (searchError) {
            DebugService.warn('Search approach also failed:', searchError);
          }
        }
        
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
        return this.getEmailWithRetry(graphToken, emailId, retry + 1, mailboxEmail, isShared);
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
    retry = 0,
    mailboxEmail?: string,
    isShared?: boolean
  ): Promise<any> {
    const headers = this.getAuthHeaders(graphToken);
    
    try {
      DebugService.debug(`Attempting internetMessageId search for: ${internetMessageId} (attempt ${retry + 1}/3)`);
      
      // Use the correct endpoint based on mailbox type
      const baseEndpoint = mailboxEmail && isShared !== undefined 
        ? this.getMailboxEndpoint(mailboxEmail, isShared)
        : 'https://graph.microsoft.com/v1.0/me';
      
      // Search for the email using internetMessageId
      const searchUrl = `${baseEndpoint}/messages?$filter=internetMessageId eq '${internetMessageId}'&$expand=attachments&$top=1`;
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
        return this.getEmailByInternetMessageId(graphToken, internetMessageId, retry + 1, mailboxEmail, isShared);
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
    retry = 0,
    mailboxEmail?: string,
    isShared?: boolean
  ): Promise<any> {
    const headers = this.getAuthHeaders(graphToken);
    
    try {
      DebugService.debug(`Attempting search fallback for emailId: ${emailId} (attempt ${retry + 1}/3)`);
      
      // Use the correct endpoint based on mailbox type
      const baseEndpoint = mailboxEmail && isShared !== undefined 
        ? this.getMailboxEndpoint(mailboxEmail, isShared)
        : 'https://graph.microsoft.com/v1.0/me';
      
      // Try to search for the email in drafts folder first
      const searchUrl = `${baseEndpoint}/mailFolders('drafts')/messages?$filter=id eq '${emailId}'&$top=1`;
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
      const inboxSearchUrl = `${baseEndpoint}/mailFolders('inbox')/messages?$filter=id eq '${emailId}'&$top=1`;
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
        return this.getEmailBySearchFallback(graphToken, emailId, retry + 1, mailboxEmail, isShared);
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
      
      // Validate parameters before using them
      DebugService.debug('Validating parameters:', { 
        subject: subject, 
        subjectType: typeof subject, 
        sender: sender, 
        senderType: typeof sender 
      });
      
      if (!subject || typeof subject !== 'string') {
        DebugService.error('Invalid subject parameter:', { subject, type: typeof subject });
        throw new Error(`Subject is required and must be a string. Got: ${typeof subject} - ${subject}`);
      }
      if (!sender || typeof sender !== 'string') {
        DebugService.error('Invalid sender parameter:', { sender, type: typeof sender });
        throw new Error(`Sender is required and must be a string. Got: ${typeof sender} - ${sender}`);
      }
      
      const headers = this.getAuthHeaders(graphToken);
      
      // Search in drafts folder with subject and createdDate filtering
      DebugService.debug('About to call replace on subject and sender');
      const escapedSubject = subject.replace(/'/g, "''");
      const escapedSender = sender.replace(/'/g, "''");
      DebugService.debug('Successfully escaped subject and sender:', { escapedSubject, escapedSender });
      
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