export interface EmailMetadata {
  messageId: string;
  subject: string;
  sender: string;
  senderEmail?: string;
  receivedDate?: string;
  receivedTime?: string;
  hasAttachments?: boolean;
  attachmentCount?: number;
  attachmentNames?: string[];
  bodyPreview?: string;
  workbenchId?: string;
  recipient?: string;
  attachments?: Array<{ name: string; size: number; contentType: string }>;
  hasExistingData?: boolean;
}
