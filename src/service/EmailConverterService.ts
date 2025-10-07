import { EmailMetadata } from '../types';
import { getSubject, getSender, getCreatedDate } from '../utils/emailHelpers';
import DebugService from './DebugService';

export interface EmlEmailData {
  content: string;
}

export interface AttachmentData {
  name: string;
  contentType: string;
  content: string; // Base64 encoded
  size: number;
  isTruncated?: boolean;
  originalSize?: number;
}

export class EmailConverterService {
  // Maximum attachment size in bytes (25MB)
  private readonly MAX_ATTACHMENT_SIZE = 25 * 1024 * 1024;
  
  // Maximum base64 content length to prevent stack overflow (increased for larger files)
  // 25MB file ≈ 33.5MB in base64 (25 * 1.34), so we set limit to 35MB in base64 chars
  private readonly MAX_BASE64_LENGTH = 35 * 1024 * 1024; // 35MB in base64 characters

  /**
   * Convert Office.js email to .eml format with all attachments
   */
  public async convertEmailToEml(item: Office.Item): Promise<EmlEmailData> {
    try {
      DebugService.email('Starting EML conversion');
      
      // Get email metadata
      const metadata = await this.getEmailMetadata(item);
      DebugService.object('Email metadata', metadata);
      
      // Get email body
      const body = await this.getEmailBody(item);
      DebugService.debug('Email body length:', body.length);
      DebugService.debug('Email body preview:', body.substring(0, 200) + '...');
      
      // Get attachments with size limits
      const attachments = await this.getEmailAttachments(item);
      DebugService.debug('Number of attachments:', attachments.length);
      
      // Generate .eml content
      const emlContent = this.buildEmlContent(metadata, body, attachments);
      
      // Validate EML content
      DebugService.debug('EML content length:', emlContent.length);
      DebugService.debug('EML content preview:', emlContent.substring(0, 500) + '...');
      
      // Check if EML content is valid
      if (!emlContent || emlContent.length === 0) {
        throw new Error('Generated EML content is empty');
      }
      
      if (!emlContent.includes('From:') || !emlContent.includes('Subject:')) {
        throw new Error('Generated EML content is missing required headers');
      }
      
      DebugService.email('EML conversion completed successfully');
      
      return {
        content: emlContent
      };
    } catch (error) {
      DebugService.errorWithStack('Failed to convert email to EML', error as Error);
      throw new Error(`Failed to convert email to EML: ${error}`);
    }
  }

  /**
   * Get email metadata from Office.js item
   */
  private async getEmailMetadata(item: Office.Item): Promise<EmailMetadata> {
    // Use util helpers for subject, sender, and created date
    const subject = await getSubject(item as any);
    const sender = await getSender(item as any);
    const receivedTime = await getCreatedDate(item as any);
    // Compose mode: recipient and attachments may differ
    let recipient = 'Unknown Recipient';
    let messageId = (item as any).itemId || `msg_${Date.now()}`;
    let hasAttachments = false;
    let attachmentCount = 0;
    let attachmentNames: string[] = [];
    if ((item as any).to) {
      // Read mode
      recipient = (item as any).to?.[0]?.emailAddress || 'Unknown Recipient';
      if ((item as any).attachments) {
        hasAttachments = (item as any).attachments.length > 0;
        attachmentCount = (item as any).attachments.length;
        attachmentNames = (item as any).attachments.map((att: any) => att.name);
      }
    } else if ((item as any).toRecipients) {
      // Compose mode
      recipient = (item as any).toRecipients?.[0]?.emailAddress || 'Unknown Recipient';
      if ((item as any).attachments) {
        hasAttachments = (item as any).attachments.length > 0;
        attachmentCount = (item as any).attachments.length;
        attachmentNames = (item as any).attachments.map((att: any) => att.name);
      }
    }
    return {
      messageId,
      subject,
      sender,
      senderEmail: sender,
      receivedDate: receivedTime,
      receivedTime,
      hasAttachments,
      attachmentCount,
      attachmentNames,
      recipient,
      workbenchId: `UWWB_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`
    };
  }

  /**
   * Get email body as text
   */
  private async getEmailBody(item: Office.Item): Promise<string> {
    return new Promise((resolve, reject) => {
      (item as any).body.getAsync(Office.CoercionType.Text, (result: Office.AsyncResult<string>) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          reject(new Error(`Failed to get email body: ${result.error?.message || 'Unknown error'}`));
          return;
        }
        resolve(result.value || '');
      });
    });
  }

  /**
   * Get email attachments (fetches actual content as base64)
   */
  private async getEmailAttachments(item: Office.Item): Promise<AttachmentData[]> {
    // Compose and read both have attachments array, but content fetching may differ
    return new Promise((resolve, reject) => {
      const attachmentsArr = (item as any).attachments || [];
      if (!attachmentsArr.length) {
        resolve([]);
        return;
      }
      const attachments: AttachmentData[] = [];
      let processed = 0;
      const total = attachmentsArr.length;
      const finish = () => { if (processed === total) resolve(attachments); };
      
      attachmentsArr.forEach((attachment: any, idx: number) => {
        const attachmentSize = attachment.size || 0;
        const attachmentName = attachment.name || 'Unknown';
        
        // Check if attachment is too large
        if (attachmentSize > this.MAX_ATTACHMENT_SIZE) {
          DebugService.warn(`Attachment '${attachmentName}' (${this.formatFileSize(attachmentSize)}) exceeds maximum size limit. Skipping content.`);
          attachments.push({
            name: attachmentName,
            contentType: attachment.contentType || 'application/octet-stream',
            content: '',
            size: attachmentSize,
            isTruncated: true,
            originalSize: attachmentSize
          });
          processed++;
          finish();
          return;
        }
        
        // Only file attachments (not item or reference)
        if (attachment.attachmentType !== 'file') {
          attachments.push({
            name: attachmentName,
            contentType: attachment.contentType || 'application/octet-stream',
            content: '',
            size: attachmentSize
          });
          processed++;
          finish();
          return;
        }
        
        // Compose mode: getAttachmentContentAsync may not be available, fallback to base64 if present
        if (typeof (item as any).getAttachmentContentAsync === 'function') {
          (item as any).getAttachmentContentAsync(
            attachment.id,
            { asyncContext: { idx } },
            (result: Office.AsyncResult<Office.AttachmentContent>) => {
              if (result.status === Office.AsyncResultStatus.Succeeded && result.value) {
                let base64Content = '';
                if (result.value.format === Office.MailboxEnums.AttachmentContentFormat.Base64) {
                  base64Content = this.truncateBase64Content(result.value.content, attachmentName);
                } else if (result.value.format === Office.MailboxEnums.AttachmentContentFormat.Url) {
                  base64Content = '';
                }
                attachments.push({
                  name: attachmentName,
                  contentType: attachment.contentType || 'application/octet-stream',
                  content: base64Content,
                  size: attachmentSize,
                  isTruncated: base64Content.length < result.value.content.length,
                  originalSize: result.value.content.length
                });
              } else {
                attachments.push({
                  name: attachmentName,
                  contentType: attachment.contentType || 'application/octet-stream',
                  content: '',
                  size: attachmentSize
                });
              }
              processed++;
              finish();
            }
          );
        } else if (attachment.contentBytes) {
          // Compose mode: attachment content may be available as base64
          const truncatedContent = this.truncateBase64Content(attachment.contentBytes, attachmentName);
          attachments.push({
            name: attachmentName,
            contentType: attachment.contentType || 'application/octet-stream',
            content: truncatedContent,
            size: attachmentSize,
            isTruncated: truncatedContent.length < attachment.contentBytes.length,
            originalSize: attachment.contentBytes.length
          });
          processed++;
          finish();
        } else {
          attachments.push({
            name: attachmentName,
            contentType: attachment.contentType || 'application/octet-stream',
            content: '',
            size: attachmentSize
          });
          processed++;
          finish();
        }
      });
    });
  }

  /**
   * Build .eml content with headers, body, and attachments
   */
  private buildEmlContent(metadata: EmailMetadata, body: string, attachments: AttachmentData[]): string {
    DebugService.email('Building EML content');
    DebugService.debug('Metadata:', { subject: metadata.subject, sender: metadata.sender, recipient: metadata.recipient });
    DebugService.debug('Body length:', body.length);
    DebugService.debug('Attachments count:', attachments.length);
    
    const lines: string[] = [];
    
    // Add email headers
    lines.push(`From: ${metadata.sender}`);
    lines.push(`To: ${metadata.recipient || 'Unknown'}`);
    lines.push(`Subject: ${metadata.subject}`);
    lines.push(`Date: ${metadata.receivedTime || metadata.receivedDate || new Date().toISOString()}`);
    lines.push(`Message-ID: <${metadata.messageId}@workbench.local>`);
    lines.push(`MIME-Version: 1.0`);
    
    // Add workbench ID as custom header
    if (metadata.workbenchId) {
      lines.push(`X-Workbench-ID: ${metadata.workbenchId}`);
    }
    
    if (attachments.length > 0) {
      // Defensive: log number of attachments
      DebugService.email(`Processing ${attachments.length} attachments`);
      // Limit number of attachments processed
      const maxAttachments = 50;
      const safeAttachments = attachments.slice(0, maxAttachments);
      if (attachments.length > maxAttachments) {
        DebugService.warn(`Too many attachments (${attachments.length}), only processing first ${maxAttachments}`);
      }
      // Multipart message with attachments
      const boundary = `boundary_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
      lines.push(`Content-Type: multipart/mixed; boundary="${boundary}"`);
      lines.push('');
      lines.push(`--${boundary}`);
      lines.push('Content-Type: text/plain; charset=utf-8');
      lines.push('Content-Transfer-Encoding: 7bit');
      lines.push('');
      lines.push(body);
      
      // Deep defensive logging
      try {
        DebugService.debug('Attachments isArray:', Array.isArray(safeAttachments), 'length:', safeAttachments.length);
        safeAttachments.forEach((att, idx) => {
          const ctor = att && att.constructor ? att.constructor.name : typeof att;
          const keys = att && typeof att === 'object' ? Object.keys(att) : [];
          const contentType = att && typeof att === 'object' ? typeof att.content : 'undefined';
          DebugService.debug(`Attachment[${idx}]: ctor=${ctor}, keys=${keys.join(',')}, contentType=${contentType}`);
        });
      } catch (logErr) {
        DebugService.warn('Error during attachment structure logging:', logErr);
      }
      
      // Add attachment placeholders
      safeAttachments.forEach((attachment, idx) => {
        lines.push('');
        lines.push(`--${boundary}`);
        lines.push(`Content-Type: ${attachment.contentType}; name="${attachment.name}"`);
        lines.push('Content-Transfer-Encoding: base64');
        lines.push(`Content-Disposition: attachment; filename="${attachment.name}"`);
        lines.push('');
        // Log type and length
        const contentType = typeof attachment.content;
        const contentLen = contentType === 'string' ? attachment.content.length : 0;
        DebugService.debug(`Attachment #${idx+1}: ${attachment.name}, type=${contentType}, length=${contentLen}`);
        // Defensive: skip very large attachments
        const maxContentLen = 100 * 1024 * 1024; // 100MB
        if (contentType === 'string' && contentLen > maxContentLen) {
          DebugService.warn(`Skipping attachment '${attachment.name}' due to excessive size (${contentLen} bytes)`);
          lines.push(`[Attachment: ${attachment.name} - ${attachment.size} bytes - Skipped due to excessive size]`);
        } else if (attachment.content && contentType === 'string') {
          // Check if attachment was truncated
          if (attachment.isTruncated) {
            DebugService.warn(`Attachment '${attachment.name}' was truncated. Original size: ${attachment.originalSize || 0}, Current size: ${attachment.content.length}`);
            lines.push(`[Attachment: ${attachment.name} - TRUNCATED - Original size: ${this.formatFileSize(attachment.originalSize || 0)}]`);
            lines.push('');
          }
          
          // Split base64 content into lines of 76 characters (safe, no regex)
          let base64Lines = this.splitBase64Lines(attachment.content, 76);
          const maxLines = 1000000; // Increased to handle larger files (was 100000)
          // For 25MB file: ~35MB base64 / 76 chars per line = ~460,000 lines
          // Setting to 1M lines to be safe
          if (base64Lines.length > maxLines) {
            DebugService.warn(`Attachment '${attachment.name}' base64 lines truncated from ${base64Lines.length} to ${maxLines} to prevent stack overflow.`);
            base64Lines = base64Lines.slice(0, maxLines);
          }
          
          // Process base64 lines in chunks to avoid stack overflow
          const chunkSize = 10000; // Process 10K lines at a time
          for (let i = 0; i < base64Lines.length; i += chunkSize) {
            const chunk = base64Lines.slice(i, i + chunkSize);
            chunk.forEach(line => lines.push(line));
          }
        } else if (attachment.content) {
          // Log and skip if content is not a string (to prevent stack overflow)
          DebugService.warn(`Attachment content for '${attachment.name}' is not a string. Type:`, contentType, attachment.content);
          lines.push(`[Attachment: ${attachment.name} - ${attachment.size} bytes - Content not included due to invalid type]`);
        } else {
          // Placeholder for attachment content
          lines.push(`[Attachment: ${attachment.name} - ${attachment.size} bytes - Content not included in this export]`);
        }
      });
      
      lines.push('');
      lines.push(`--${boundary}--`);
    } else {
      // Simple text message
      lines.push('Content-Type: text/plain; charset=utf-8');
      lines.push('Content-Transfer-Encoding: 7bit');
      lines.push('');
      lines.push(body);
    }
    
    const finalContent = lines.join('\r\n');
    
    // Final validation
    DebugService.debug('Final EML content length:', finalContent.length);
    DebugService.debug('Final EML content ends with:', finalContent.substring(Math.max(0, finalContent.length - 100)));
    
    // Validate final content
    if (!finalContent || finalContent.length === 0) {
      throw new Error('Generated EML content is empty after joining lines');
    }
    
    if (!finalContent.includes('From:') || !finalContent.includes('Subject:')) {
      throw new Error('Final EML content is missing required headers');
    }
    
    DebugService.email('EML content built successfully');
    
    return finalContent;
  }

  // Helper to split base64 into 76-char lines (safe, no regex)
  private splitBase64Lines(str: string, lineLength = 76): string[] {
    const lines = [];
    for (let i = 0; i < str.length; i += lineLength) {
      lines.push(str.slice(i, i + lineLength));
    }
    return lines;
  }

  // Helper to truncate base64 content to prevent stack overflow
  private truncateBase64Content(content: string, attachmentName: string): string {
    const maxLength = this.MAX_BASE64_LENGTH;
    if (content.length <= maxLength) {
      return content;
    }
    DebugService.warn(`Truncating base64 content for attachment '${attachmentName}' to prevent stack overflow. Original length: ${content.length}, Max length: ${maxLength}`);
    return content.substring(0, maxLength);
  }

  // Helper to format file size
  private formatFileSize(bytes: number): string {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB', 'TB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  }
} 
