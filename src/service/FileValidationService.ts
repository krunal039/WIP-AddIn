import DebugService from './DebugService';
import { detectFileProtectionFromBase64, isSupportedFileType } from '../utils/fileInspector';

/**
 * FileValidationService - Validates email attachments for various restrictions
 * 
 * CAPABILITIES:
 * Zip/Compressed files: Detects .zip, .rar, .7z, .tar, .gz, etc.
 * Unsupported file types: Validates against whitelist of supported extensions
 * Basic encrypted files: Detects .gpg, .pgp, .encrypted files
 * Basic password-protected files: Detects files with "password" in filename
 * Office document encryption: Detects password-protected Word, Excel, PowerPoint files
 * PDF password protection: Detects password-protected PDF files
 *  M365 encryption: Limited detection based on file patterns and content types
 * 
 */

export interface FileValidationError {
  type: 'zip' | 'unsupported' | 'encrypted' | 'password_protected';
  message: string;
  files: string[];
}

export interface FileValidationResult {
  isValid: boolean;
  errors: FileValidationError[];
}

class FileValidationService {
  private static instance: FileValidationService;

  // Supported file extensions
  private readonly supportedExtensions = [
    '.bat', '.bashc', '.c', '.cc', '.chtml', '.cmake', '.cmd', '.cpp',
    '.cs', '.css', '.csv', '.cxx', '.cy', '.dockerfile', '.doc', '.docx',
    '.eml', '.gdoc', '.go', '.gsheet', '.gslides', '.h++', '.hpp', '.html',
    '.java', '.jpeg', '.jpg', '.js', '.json', '.mht', '.mhtml', '.mkfile',
    '.msg', '.pdf', '.perl', '.php', '.plsql', '.png', '.pptx', '.py',
    '.pxi', '.pyx', '.r', '.rd', '.rs', '.rtf', '.ruby', '.tif', '.tiff',
    '.ts', '.txt', '.xls', '.xlsx', '.xml', '.yaml', '.yml', '.zsh'
  ];

  // Compressed file extensions
  private readonly compressedExtensions = [
    '.zip', '.rar', '.7z', '.tar', '.gz', '.bz2', '.xz', '.tar.gz', '.tar.bz2'
  ];

  private constructor() {}

  public static getInstance(): FileValidationService {
    if (!FileValidationService.instance) {
      FileValidationService.instance = new FileValidationService();
    }
    return FileValidationService.instance;
  }

  /**
   * Validate email attachments for various file restrictions
   * @param item Office.js mailbox item
   * @returns Promise<FileValidationResult>
   */
  public async validateEmailAttachments(item: any): Promise<FileValidationResult> {
    try {
      DebugService.debug('Starting file validation for email attachments');
      
      const errors: FileValidationError[] = [];
      
      // Get attachments from the email
      const attachments = await this.getEmailAttachments(item);
      DebugService.debug(`Found ${attachments.length} attachments to validate:`, attachments.map(a => ({ name: a.name, size: a.size, contentType: a.contentType })));
      
      if (attachments.length === 0) {
        return { isValid: true, errors: [] };
      }

      // Parallelize independent validation checks for better performance
      const [zipFiles, unsupportedFiles, encryptedFiles, passwordProtectedFiles] = await Promise.all([
        Promise.resolve(attachments.filter(att => this.isCompressedFile(att.name))),
        Promise.resolve(attachments.filter(att => !this.isSupportedFile(att.name))),
        this.detectEncryptedFiles(attachments),
        this.detectPasswordProtectedFiles(attachments)
      ]);

      DebugService.debug(`Found ${zipFiles.length} compressed files:`, zipFiles.map(f => f.name));
      
      if (zipFiles.length > 0) {
        errors.push({
          type: 'zip',
          message: 'One or multiple .zip files attached. Please remove all .zip file(s) to submit.',
          files: zipFiles.map(f => f.name)
        });
        DebugService.debug('Added zip validation error');
      }

      if (unsupportedFiles.length > 0) {
        errors.push({
          type: 'unsupported',
          message: `One or more unsupported file types attached. Please remove unsupported file(s) to submit.`,
          files: unsupportedFiles.map(f => f.name)
        });
      }

      if (encryptedFiles.length > 0) {
        errors.push({
          type: 'encrypted',
          message: 'One or more files is encrypted. Please remove encrypted file(s) to submit.',
          files: encryptedFiles
        });
      }

      if (passwordProtectedFiles.length > 0) {
        errors.push({
          type: 'password_protected',
          message: 'One or more files is password protected. Please remove password protected file(s) to submit.',
          files: passwordProtectedFiles
        });
      }

      const isValid = errors.length === 0;
      DebugService.debug(`File validation completed. Valid: ${isValid}, Errors: ${errors.length}`, errors);
      
      return { isValid, errors };

    } catch (error) {
      DebugService.error('File validation failed:', error);
      // Return a generic error if validation fails
      return {
        isValid: false,
        errors: [{
          type: 'unsupported',
          message: 'Unable to validate attachments. Please check file types and try again.',
          files: []
        }]
      };
    }
  }

  /**
   * Get email attachments from Office.js item with file content for analysis
   */
  private async getEmailAttachments(item: any): Promise<Array<{name: string, size: number, contentType: string, content?: string}>> {
    return new Promise((resolve) => {
      try {
        DebugService.debug('Getting email attachments with content, item type:', item.itemType);
        
        // For compose mode, try to get attachments
        if (item.attachments && typeof item.attachments.getAsync === 'function') {
          DebugService.debug('Using compose mode attachment access');
          item.attachments.getAsync(async (result: any) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              const attachments = result.value || [];
              DebugService.debug(`Compose mode: Found ${attachments.length} attachments`);
              
              // Get attachment content for password protection detection
              const attachmentsWithContent = await this.getAttachmentContent(item, attachments);
              resolve(attachmentsWithContent);
            } else {
              DebugService.warn('Failed to get attachments in compose mode:', result.error);
              resolve([]);
            }
          });
        } else if (item.itemType === Office.MailboxEnums.ItemType.Message) {
          // For read mode, try to access attachments directly
          DebugService.debug('Using read mode attachment access');
          if (item.attachments && Array.isArray(item.attachments)) {
            DebugService.debug(`Read mode: Found ${item.attachments.length} attachments directly`);
            
            // Get attachment content for password protection detection
            this.getAttachmentContent(item, item.attachments).then(resolve);
          } else {
            DebugService.warn('No attachments found in read mode');
            resolve([]);
          }
        } else {
          DebugService.warn('Unknown item type or no attachment access method');
          resolve([]);
        }
      } catch (error) {
        DebugService.error('Error getting email attachments:', error);
        resolve([]);
      }
    });
  }

  /**
   * Get attachment content for password protection detection
   */
  private async getAttachmentContent(item: any, attachments: any[]): Promise<Array<{name: string, size: number, contentType: string, content?: string}>> {
    const attachmentsWithContent: Array<{name: string, size: number, contentType: string, content?: string}> = [];
    
    for (const attachment of attachments) {
      const attachmentData = {
        name: attachment.name || 'Unknown',
        size: attachment.size || 0,
        contentType: attachment.contentType || 'application/octet-stream',
        content: undefined as string | undefined
      };
      
      // Try to get attachment content for password protection detection
      DebugService.debug(`Attempting to get content for attachment: ${attachment.name}, ID: ${attachment.id}`);
      
      if (typeof item.getAttachmentContentAsync === 'function' && attachment.id) {
        try {
          DebugService.debug(`Calling getAttachmentContentAsync for: ${attachment.name}`);
          const contentResult = await new Promise<Office.AsyncResult<Office.AttachmentContent>>((resolve) => {
            item.getAttachmentContentAsync(
              attachment.id,
              { asyncContext: attachment.id },
              (result: Office.AsyncResult<Office.AttachmentContent>) => {
                DebugService.debug(`getAttachmentContentAsync callback for ${attachment.name}:`, result.status);
                resolve(result);
              }
            );
          });
          
          if (contentResult.status === Office.AsyncResultStatus.Succeeded && contentResult.value) {
            if (contentResult.value.format === Office.MailboxEnums.AttachmentContentFormat.Base64) {
              attachmentData.content = contentResult.value.content;
              DebugService.debug(`Retrieved content for attachment: ${attachment.name} (${contentResult.value.content.length} chars)`);
            } else {
              DebugService.debug(`Attachment ${attachment.name} content format not supported: ${contentResult.value.format}`);
            }
          } else {
            DebugService.warn(`Failed to get content for attachment ${attachment.name}:`, contentResult.error);
          }
        } catch (error) {
          DebugService.warn(`Error getting content for attachment ${attachment.name}:`, error);
        }
      } else {
        DebugService.debug(`Cannot get content for attachment ${attachment.name} - getAttachmentContentAsync: ${typeof item.getAttachmentContentAsync === 'function'}, attachment.id: ${attachment.id}`);
      }
      
      attachmentsWithContent.push(attachmentData);
    }
    
    return attachmentsWithContent;
  }

  /**
   * Check if file is compressed
   */
  private isCompressedFile(filename: string): boolean {
    const extension = this.getFileExtension(filename).toLowerCase();
    return this.compressedExtensions.includes(extension);
  }

  /**
   * Check if file type is supported
   */
  private isSupportedFile(filename: string): boolean {
    const extension = this.getFileExtension(filename).toLowerCase();
    return this.supportedExtensions.includes(extension);
  }

  /**
   * Get file extension from filename
   */
  private getFileExtension(filename: string): string {
    const lastDot = filename.lastIndexOf('.');
    return lastDot !== -1 ? filename.substring(lastDot) : '';
  }

  /**
   * Detect encrypted files including M365 encryption and common encryption patterns
   */
  private async detectEncryptedFiles(attachments: Array<{name: string, size: number, contentType: string, content?: string}>): Promise<string[]> {
    const encryptedFiles: string[] = [];
    
    for (const attachment of attachments) {
      const filename = attachment.name.toLowerCase();
      const contentType = attachment.contentType.toLowerCase();
      
      // Check for explicit encrypted file extensions
      if (filename.includes('.encrypted') || 
          filename.includes('.enc') ||
          filename.includes('.gpg') ||
          filename.includes('.pgp') ||
          filename.includes('.asc') ||
          filename.includes('.key')) {
        encryptedFiles.push(attachment.name);
        continue;
      }
      
      // Check for encrypted content types
      if (contentType === 'application/pgp-encrypted' ||
          contentType === 'application/pkcs7-mime' ||
          contentType === 'application/x-pkcs7-mime' ||
          contentType === 'application/octet-stream' && filename.includes('encrypted')) {
        encryptedFiles.push(attachment.name);
        continue;
      }
      
      // Check for M365 encryption indicators
      // M365 encrypted files often have specific content types or size patterns
      if (this.isM365EncryptedFile(attachment)) {
        encryptedFiles.push(attachment.name);
        continue;
      }
      
      // Check for Office document encryption (password-protected files that are also encrypted)
      if (this.isOfficeEncryptedFile(attachment)) {
        encryptedFiles.push(attachment.name);
      }
    }
    
    return encryptedFiles;
  }

  /**
   * Check if file appears to be M365 encrypted
   * M365 encryption is complex and may not always be detectable via simple patterns
   */
  private isM365EncryptedFile(attachment: {name: string, size: number, contentType: string, content?: string}): boolean {
    const filename = attachment.name.toLowerCase();
    const contentType = attachment.contentType.toLowerCase();
    
    // Only check for explicit M365 encryption indicators to avoid false positives
    // Check for M365-specific encrypted content types
    if (contentType === 'application/x-microsoft-encrypted' ||
        contentType === 'application/x-microsoft-office-encrypted') {
      return true;
    }
    
    // Check for files with explicit encryption indicators in filename
    if (filename.includes('encrypted') && (
        filename.endsWith('.encrypted') ||
        filename.includes('_encrypted') ||
        filename.includes('_enc_')
    )) {
      return true;
    }
    
    return false;
  }

  /**
   * Check if Office document appears to be encrypted (password-protected)
   */
  private isOfficeEncryptedFile(attachment: {name: string, size: number, contentType: string, content?: string}): boolean {
    const filename = attachment.name.toLowerCase();
    const contentType = attachment.contentType.toLowerCase();
    
    // Check for Office document types that might be encrypted
    const officeExtensions = ['.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx', '.pdf'];
    const isOfficeFile = officeExtensions.some(ext => filename.endsWith(ext));
    
    if (!isOfficeFile) return false;
    
    // Check for encrypted Office content types
    if (contentType === 'application/x-microsoft-office-encrypted' ||
        contentType === 'application/x-password-protected') {
      return true;
    }
    
    return false;
  }

  /**
   * Detect password protected files including Office documents (PDF, Excel, Word, PowerPoint)
   * Now uses file content analysis for more accurate detection
   */
  private async detectPasswordProtectedFiles(attachments: Array<{name: string, size: number, contentType: string, content?: string}>): Promise<string[]> {
    const passwordProtectedFiles: string[] = [];
    
    for (const attachment of attachments) {
      const filename = attachment.name.toLowerCase();
      const contentType = attachment.contentType.toLowerCase();
      
      // First, try content-based detection if content is available
      DebugService.debug(`Processing attachment: ${attachment.name}, has content: ${!!attachment.content}, supported: ${isSupportedFileType(attachment.name)}`);
      
      if (attachment.content && isSupportedFileType(attachment.name)) {
        try {
          DebugService.debug(`Starting comprehensive content analysis for: ${attachment.name}`);
          DebugService.debug(`Content length: ${attachment.content.length} characters`);
          DebugService.debug(`Content preview: ${attachment.content.substring(0, 50)}...`);
          
          const result = await detectFileProtectionFromBase64(attachment.content);
          
          DebugService.debug(`Analysis result for ${attachment.name}:`, result);
          
          if (result.encrypted) {
            passwordProtectedFiles.push(attachment.name);
            DebugService.debug(`Comprehensive analysis detected password protection: ${attachment.name} - ${result.details.join(', ')}`);
            continue;
          } else {
            DebugService.debug(`Comprehensive analysis found no password protection: ${attachment.name}`);
          }
        } catch (error) {
          DebugService.warn(`Error in comprehensive analysis for password protection: ${attachment.name}`, error);
        }
      } else {
        DebugService.debug(`Skipping comprehensive analysis for ${attachment.name} - content: ${!!attachment.content}, supported: ${isSupportedFileType(attachment.name)}`);
      }
      
      // Fallback to filename and content type based detection
      // Check for explicit password protection indicators in filename
      if (filename.includes('password') ||
          filename.includes('protected') ||
          filename.includes('locked') ||
          filename.includes('secure') ||
          filename.includes('encrypted') ||
          filename.includes('_pw_') ||
          filename.includes('_pass_')) {
        passwordProtectedFiles.push(attachment.name);
        DebugService.debug(`Filename analysis detected password protection: ${attachment.name}`);
        continue;
      }
      
      // Check for password-protected content types
      if (contentType === 'application/x-password-protected' ||
          contentType === 'application/x-microsoft-office-encrypted' ||
          contentType.includes('password') ||
          contentType.includes('protected')) {
        passwordProtectedFiles.push(attachment.name);
        DebugService.debug(`Content type analysis detected password protection: ${attachment.name}`);
        continue;
      }
      
      // Check for Office documents that might be password-protected (fallback)
      if (this.isOfficePasswordProtectedFile(attachment)) {
        passwordProtectedFiles.push(attachment.name);
        DebugService.debug(`Office file analysis detected password protection: ${attachment.name}`);
        continue;
      }
      
      // Check for PDF files that might be password-protected (fallback)
      if (this.isPasswordProtectedPDF(attachment)) {
        passwordProtectedFiles.push(attachment.name);
        DebugService.debug(`PDF analysis detected password protection: ${attachment.name}`);
      }
    }
    
    DebugService.debug(`Password protection detection completed. Found ${passwordProtectedFiles.length} protected files:`, passwordProtectedFiles);
    return passwordProtectedFiles;
  }

  /**
   * Check if Office document (Word, Excel, PowerPoint) appears to be password-protected
   */
  private isOfficePasswordProtectedFile(attachment: {name: string, size: number, contentType: string, content?: string}): boolean {
    const filename = attachment.name.toLowerCase();
    const contentType = attachment.contentType.toLowerCase();
    
    // Office document extensions
    const officeExtensions = ['.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx'];
    const isOfficeFile = officeExtensions.some(ext => filename.endsWith(ext));
    
    if (!isOfficeFile) return false;
    
    // Check for encrypted Office content types
    if (contentType === 'application/x-microsoft-office-encrypted' ||
        contentType === 'application/x-password-protected') {
      return true;
    }
    
    return false;
  }

  /**
   * Check if PDF file appears to be password-protected
   */
  private isPasswordProtectedPDF(attachment: {name: string, size: number, contentType: string, content?: string}): boolean {
    const filename = attachment.name.toLowerCase();
    const contentType = attachment.contentType.toLowerCase();
    
    if (!filename.endsWith('.pdf')) return false;
    
    // Check for encrypted PDF content types
    if (contentType === 'application/x-password-protected' ||
        contentType === 'application/pdf-encrypted') {
      return true;
    }
    
    return false;
  }

  /**
   * Get the primary error message to display to user
   */
  public getPrimaryErrorMessage(errors: FileValidationError[]): string | null {
    if (errors.length === 0) return null;
    
    // Priority order: zip > unsupported > encrypted > password protected
    const priorityOrder = ['zip', 'unsupported', 'encrypted', 'password_protected'];
    
    for (const type of priorityOrder) {
      const error = errors.find(e => e.type === type);
      if (error) {
        return error.message;
      }
    }
    
    return errors[0].message;
  }

  /**
   * Get all error messages combined for display to user
   * This shows all validation errors at once instead of one at a time
   */
  public getAllErrorMessages(errors: FileValidationError[]): string | null {
    if (errors.length === 0) return null;
    
    // Always use the structured format, even for single errors
    const errorMessages: string[] = [];
    
    // Add main error message
    errorMessages.push('There\'s one or more unsupported file types attached.');
    
    // Track files that have already been shown to avoid duplicates
    const shownFiles = new Set<string>();
    
    // Priority order: zip > unsupported > encrypted > password protected
    const priorityOrder = ['zip', 'unsupported', 'encrypted', 'password_protected'];
    
    // Add specific file details for each error type in priority order
    priorityOrder.forEach(errorType => {
      const error = errors.find(e => e.type === errorType);
      if (error && error.files.length > 0) {
        // Filter out files that have already been shown
        const newFiles = error.files.filter(file => !shownFiles.has(file));
        if (newFiles.length > 0) {
          const errorTypeLabel = this.getErrorTypeLabel(error.type);
          const fileList = newFiles.join(', ');
          errorMessages.push(`• ${errorTypeLabel}: ${fileList}`);
          
          // Mark these files as shown
          newFiles.forEach(file => shownFiles.add(file));
        }
      }
    });
    
    return errorMessages.join('\n');
  }

  /**
   * Get user-friendly label for error type
   */
  private getErrorTypeLabel(errorType: string): string {
    switch (errorType) {
      case 'zip':
        return 'Zip files';
      case 'unsupported':
        return 'Unsupported file';
      case 'encrypted':
        return 'Encrypted file';
      case 'password_protected':
        return 'Password protected file';
      default:
        return 'Invalid file';
    }
  }

  /**
   * Get detailed error information for debugging
   * This helps understand what types of files were detected as problematic
   */
  public getDetailedErrorInfo(errors: FileValidationError[]): string {
    if (errors.length === 0) return 'No validation errors';
    
    const details = errors.map(error => {
      const fileList = error.files.length > 0 ? ` (${error.files.join(', ')})` : '';
      return `${error.type}: ${error.message}${fileList}`;
    });
    
    return details.join('; ');
  }

  /**
   * Get validation statistics for monitoring
   */
  public getValidationStats(result: FileValidationResult): {
    totalFiles: number;
    validFiles: number;
    errorCounts: Record<string, number>;
  } {
    const totalFiles = result.errors.reduce((sum, error) => sum + error.files.length, 0);
    const validFiles = totalFiles - result.errors.reduce((sum, error) => sum + error.files.length, 0);
    
    const errorCounts: Record<string, number> = {};
    result.errors.forEach(error => {
      errorCounts[error.type] = (errorCounts[error.type] || 0) + error.files.length;
    });
    
    return {
      totalFiles,
      validFiles,
      errorCounts
    };
  }
}

export default FileValidationService.getInstance();
