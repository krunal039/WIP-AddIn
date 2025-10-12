/**
 * File Inspector Utility
 * 
 * Simple password-protected file detection for Office and PDF files.
 * Supports DOCX, XLSX, PPTX, PDF, DOC, XLS, PPT file formats.
 */

// Convert Base64 to ArrayBuffer
function base64ToArrayBuffer(base64: string): ArrayBuffer {
  const binaryString = atob(base64);
  const len = binaryString.length;
  const bytes = new Uint8Array(len);
  for (let i = 0; i < len; i++) bytes[i] = binaryString.charCodeAt(i);
  return bytes.buffer;
}

// Simple PDF scan for /Encrypt token
function isPdfEncrypted(ab: ArrayBuffer): boolean {
  if (!ab || ab.byteLength < 4) return false;
  
  const headStr = new TextDecoder().decode(ab.slice(0, 16));
  if (!headStr.startsWith('%PDF-')) return false;

  const scanLen = Math.min(ab.byteLength, 256 * 1024); // 256 KB scan
  const content = new TextDecoder().decode(ab.slice(0, scanLen));
  
  return content.includes('/Encrypt');
}

// Simple Office file scan
async function isOfficeEncrypted(ab: ArrayBuffer): Promise<boolean> {
  const header = new Uint8Array(ab.slice(0, 8));
  const zipSig = [0x50, 0x4B, 0x03, 0x04];
  const oleSig = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
  const matchSig = (sig: number[]) => sig.every((v, i) => header[i] === v);

  // OOXML (ZIP) files
  if (matchSig(zipSig)) {
    try {
      const JSZip = (await import('jszip')).default;
      const zip = await JSZip.loadAsync(ab);
      
      // Check for encryption streams
      const fileNames = Object.keys(zip.files);
      return fileNames.some(name => 
        name.toLowerCase().includes('encryptedpackage') || 
        name.toLowerCase().includes('encryptioninfo')
      );
    } catch (error) {
      console.warn('[FileInspector] Error parsing ZIP file:', error);
      return false;
    }
  }

  // Legacy OLE files
  if (matchSig(oleSig)) {
    try {
      const CFB = (await import('cfb')).default;
      const cfb = CFB.read(new Uint8Array(ab), { type: "array" });
      const names = cfb.FullPaths.map(n => n.toLowerCase());
      
      // Check for encryption streams - this is the most reliable indicator
      return names.some(name => 
        name.includes('/encryptedpackage') || 
        name.includes('/encryptioninfo')
      );
    } catch (error) {
      console.warn('[FileInspector] Error parsing OLE file:', error);
      return false;
    }
  }

  return false;
}

/**
 * Get file extension from filename
 */
export function getFileExtension(filename: string): string {
  const lastDotIndex = filename.lastIndexOf('.');
  if (lastDotIndex === -1) {
    return '';
  }
  return filename.substring(lastDotIndex + 1).toLowerCase();
}

/**
 * Check if file type is supported for password protection detection
 */
export function isSupportedFileType(filename: string): boolean {
  const extension = getFileExtension(filename);
  const supportedExtensions = ['docx', 'xlsx', 'pptx', 'pdf', 'doc', 'xls', 'ppt'];
  return supportedExtensions.includes(extension);
}

/**
 * Simple password protection detection
 */
export async function detectFileProtectionFromBase64(base64: string): Promise<{ fileName: string; type: string; encrypted: boolean; details: string[] }> {
  const ab = base64ToArrayBuffer(base64);

  // Check if it's a PDF
  if (isPdfEncrypted(ab)) {
    return { 
      fileName: "unknown.pdf", 
      type: "PDF", 
      encrypted: true, 
      details: ["PDF is encrypted"] 
    };
  }

  // Check if it's an Office file
  const isEncrypted = await isOfficeEncrypted(ab);
  if (isEncrypted) {
    return {
      fileName: "unknown.office",
      type: "Office file",
      encrypted: true,
      details: ["Office file is encrypted"]
    };
  }

  return {
    fileName: "unknown",
    type: "Unknown",
    encrypted: false,
    details: ["No encryption detected"]
  };
}