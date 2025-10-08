import DebugService from './DebugService';

/**
 * FileValidationService - Validates email attachments for various restrictions
 * 
 * CAPABILITIES:
 * ✅ Zip/Compressed files: Detects .zip, .rar, .7z, .tar, .gz, etc.
 * ✅ Unsupported file types: Validates against whitelist of supported extensions
 * ✅ Basic encrypted files: Detects .gpg, .pgp, .encrypted files
 * ✅ Basic password-protected files: Detects files with "password" in filename
 * ✅ Office document encryption: Detects password-protected Word, Excel, PowerPoint files
 * ✅ PDF password protection: Detects password-protected PDF files
 * ⚠️  M365 encryption: Limited detection based on file patterns and content types
 * 
 * LIMITATIONS:
 * ❌ M365 encryption: Microsoft 365 encryption is complex and may not always be detectable
 *    through simple file analysis. M365 encrypted files might appear as normal files
 *    with standard content types.
 * ❌ Advanced password protection: Some password-protected files may not be detected
 *    if they don't have obvious indicators in filename or content type.
 * ❌ File content analysis: This service only analyzes metadata (filename, size, content type)
 *    and cannot read actual file content to determine encryption status.
 * 
 * RECOMMENDATIONS:
 * - For production use, consider implementing server-side file analysis
 * - Add user education about file restrictions
 * - Consider allowing users to override certain warnings for legitimate files
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

      // Check for compressed files
      const zipFiles = attachments.filter(att => this.isCompressedFile(att.name));
      DebugService.debug(`Found ${zipFiles.length} compressed files:`, zipFiles.map(f => f.name));
      if (zipFiles.length > 0) {
        errors.push({
          type: 'zip',
          message: 'One or multiple .zip files attached. Please remove all .zip file(s) to submit.',
          files: zipFiles.map(f => f.name)
        });
        DebugService.debug('Added zip validation error');
      }

      // Check for unsupported file types
      const unsupportedFiles = attachments.filter(att => !this.isSupportedFile(att.name));
      if (unsupportedFiles.length > 0) {
        errors.push({
          type: 'unsupported',
          message: `One or more unsupported file types attached. Please remove unsupported file(s) to submit.`,
          files: unsupportedFiles.map(f => f.name)
        });
      }

      // Check for encrypted files (basic detection)
      const encryptedFiles = await this.detectEncryptedFiles(attachments);
      if (encryptedFiles.length > 0) {
        errors.push({
          type: 'encrypted',
          message: 'One or more files is encrypted. Please remove encrypted file(s) to submit.',
          files: encryptedFiles
        });
      }

      // Check for password protected files (basic detection)
      const passwordProtectedFiles = await this.detectPasswordProtectedFiles(attachments);
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
   * Get email attachments from Office.js item
   */
  private async getEmailAttachments(item: any): Promise<Array<{name: string, size: number, contentType: string}>> {
    return new Promise((resolve) => {
      try {
        DebugService.debug('Getting email attachments, item type:', item.itemType);
        
        // For compose mode, try to get attachments
        if (item.attachments && typeof item.attachments.getAsync === 'function') {
          DebugService.debug('Using compose mode attachment access');
          item.attachments.getAsync((result: any) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              const attachments = result.value || [];
              DebugService.debug(`Compose mode: Found ${attachments.length} attachments`);
              resolve(attachments.map((att: any) => ({
                name: att.name || 'Unknown',
                size: att.size || 0,
                contentType: att.contentType || 'application/octet-stream'
              })));
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
            resolve(item.attachments.map((att: any) => ({
              name: att.name || 'Unknown',
              size: att.size || 0,
              contentType: att.contentType || 'application/octet-stream'
            })));
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
  private async detectEncryptedFiles(attachments: Array<{name: string, size: number, contentType: string}>): Promise<string[]> {
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
  private isM365EncryptedFile(attachment: {name: string, size: number, contentType: string}): boolean {
    const filename = attachment.name.toLowerCase();
    const contentType = attachment.contentType.toLowerCase();
    
    // M365 encrypted files might have these characteristics:
    // 1. Very small file size (encrypted metadata)
    // 2. Specific content types that indicate encryption
    // 3. Filename patterns that suggest encryption
    
    // Check for very small files that might be encrypted metadata
    if (attachment.size < 1000 && (
        contentType === 'application/octet-stream' ||
        contentType === 'text/plain' ||
        filename.endsWith('.encrypted')
    )) {
      return true;
    }
    
    // Check for M365-specific encrypted content types
    if (contentType.includes('encrypted') || 
        contentType.includes('cipher') ||
        contentType === 'application/x-microsoft-encrypted') {
      return true;
    }
    
    // Check for files that might be M365 encrypted based on naming patterns
    if (filename.includes('encrypted') || 
        filename.includes('cipher') ||
        filename.match(/^[a-f0-9]{32,}\./) || // Hex-like filenames
        filename.includes('_enc_')) {
      return true;
    }
    
    return false;
  }

  /**
   * Check if Office document appears to be encrypted (password-protected)
   */
  private isOfficeEncryptedFile(attachment: {name: string, size: number, contentType: string}): boolean {
    const filename = attachment.name.toLowerCase();
    const contentType = attachment.contentType.toLowerCase();
    
    // Check for Office document types that might be encrypted
    const officeExtensions = ['.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx', '.pdf'];
    const isOfficeFile = officeExtensions.some(ext => filename.endsWith(ext));
    
    if (!isOfficeFile) return false;
    
    // Password-protected Office files often have these characteristics:
    // 1. Smaller than expected size for the content
    // 2. Specific content types that indicate encryption
    // 3. Filename patterns
    
    // Check for encrypted Office content types
    if (contentType === 'application/x-microsoft-office-encrypted' ||
        contentType === 'application/x-password-protected' ||
        contentType.includes('encrypted')) {
      return true;
    }
    
    // Check for suspiciously small Office files (might be encrypted/compressed)
    if (attachment.size < 5000 && isOfficeFile) {
      return true;
    }
    
    return false;
  }

  /**
   * Detect password protected files including Office documents (PDF, Excel, Word, PowerPoint)
   */
  private async detectPasswordProtectedFiles(attachments: Array<{name: string, size: number, contentType: string}>): Promise<string[]> {
    const passwordProtectedFiles: string[] = [];
    
    for (const attachment of attachments) {
      const filename = attachment.name.toLowerCase();
      const contentType = attachment.contentType.toLowerCase();
      
      // Check for explicit password protection indicators in filename
      if (filename.includes('password') ||
          filename.includes('protected') ||
          filename.includes('locked') ||
          filename.includes('secure') ||
          filename.includes('encrypted') ||
          filename.includes('_pw_') ||
          filename.includes('_pass_')) {
        passwordProtectedFiles.push(attachment.name);
        continue;
      }
      
      // Check for password-protected content types
      if (contentType === 'application/x-password-protected' ||
          contentType === 'application/x-microsoft-office-encrypted' ||
          contentType.includes('password') ||
          contentType.includes('protected')) {
        passwordProtectedFiles.push(attachment.name);
        continue;
      }
      
      // Check for Office documents that might be password-protected
      if (this.isOfficePasswordProtectedFile(attachment)) {
        passwordProtectedFiles.push(attachment.name);
        continue;
      }
      
      // Check for PDF files that might be password-protected
      if (this.isPasswordProtectedPDF(attachment)) {
        passwordProtectedFiles.push(attachment.name);
      }
    }
    
    return passwordProtectedFiles;
  }

  /**
   * Check if Office document (Word, Excel, PowerPoint) appears to be password-protected
   */
  private isOfficePasswordProtectedFile(attachment: {name: string, size: number, contentType: string}): boolean {
    const filename = attachment.name.toLowerCase();
    const contentType = attachment.contentType.toLowerCase();
    
    // Office document extensions
    const officeExtensions = ['.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx'];
    const isOfficeFile = officeExtensions.some(ext => filename.endsWith(ext));
    
    if (!isOfficeFile) return false;
    
    // Password-protected Office files often have these characteristics:
    // 1. Smaller file size than expected (due to encryption)
    // 2. Specific content types
    // 3. Unusual file structure patterns
    
    // Check for encrypted Office content types
    if (contentType === 'application/x-microsoft-office-encrypted' ||
        contentType === 'application/x-password-protected' ||
        contentType.includes('encrypted') ||
        contentType.includes('password')) {
      return true;
    }
    
    // Check for suspiciously small Office files (might be password-protected)
    // Normal Office files are usually larger than 5KB
    if (attachment.size < 5000 && isOfficeFile) {
      return true;
    }
    
    // Check for Office files with unusual content types
    if (isOfficeFile && (
        contentType === 'application/octet-stream' ||
        contentType === 'text/plain' ||
        contentType.includes('encrypted')
    )) {
      return true;
    }
    
    return false;
  }

  /**
   * Check if PDF file appears to be password-protected
   */
  private isPasswordProtectedPDF(attachment: {name: string, size: number, contentType: string}): boolean {
    const filename = attachment.name.toLowerCase();
    const contentType = attachment.contentType.toLowerCase();
    
    if (!filename.endsWith('.pdf')) return false;
    
    // Password-protected PDFs often have these characteristics:
    // 1. Smaller file size than expected
    // 2. Specific content types
    // 3. Unusual file structure
    
    // Check for encrypted PDF content types
    if (contentType === 'application/x-password-protected' ||
        contentType === 'application/pdf-encrypted' ||
        contentType.includes('encrypted') ||
        contentType.includes('password')) {
      return true;
    }
    
    // Check for suspiciously small PDF files (might be password-protected)
    // Normal PDFs are usually larger than 2KB
    if (attachment.size < 2000) {
      return true;
    }
    
    // Check for PDFs with unusual content types
    if (contentType === 'application/octet-stream' ||
        contentType === 'text/plain' ||
        contentType.includes('encrypted')) {
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
    
    if (errors.length === 1) {
      return errors[0].message;
    }
    
    // Combine all error messages with file lists
    const errorMessages = errors.map(error => {
      const fileList = error.files.length > 0 ? ` (${error.files.join(', ')})` : '';
      return `${error.message}${fileList}`;
    });
    
    // Join with line breaks for better readability
    return errorMessages.join('\n');
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
