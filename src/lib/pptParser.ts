/**
 * PPT Parser - Extract data from legacy PowerPoint files (.ppt)
 * PPT files use a binary format (OLE Compound Document)
 * Uses CFB (Compound File Binary) for proper OLE parsing
 */

import type { ExtractedPresentation, SlideContent, PresentationMetadata } from '../types';
import CFB from 'cfb';

// eslint-disable-next-line @typescript-eslint/no-explicit-any
let codepage: any = null;

// Dynamic import of codepage for text decoding
async function loadCodepage() {
  if (!codepage) {
    try {
      codepage = await import('codepage');
    } catch (err) {
      console.warn('codepage not available:', err);
    }
  }
  return codepage;
}

/**
 * Parse a legacy PPT file using CFB
 */
export async function parsePPT(file: File): Promise<ExtractedPresentation> {
  const buffer = await file.arrayBuffer();
  const data = new Uint8Array(buffer);
  
  try {
    // Try to parse with CFB
    const cfb = CFB.read(data, { type: 'array' });
    const result = await parseCFBDocument(cfb, file);
    return result;
  } catch (err) {
    console.warn('CFB parsing failed, falling back to basic extraction:', err);
    // Fallback to basic extraction
    return parseWithBasicExtraction(data, file);
  }
}

/**
 * Parse PPT using CFB library
 */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
async function parseCFBDocument(cfb: any, file: File): Promise<ExtractedPresentation> {
  const cp = await loadCodepage();
  const texts: string[] = [];
  
  // Try to find the PowerPoint Document stream
  const pptDoc = cfb.find('PowerPoint Document');
  
  if (pptDoc && pptDoc.content) {
    // Parse the PowerPoint Document stream
    const extractedTexts = extractTextFromPPTStream(pptDoc.content, cp);
    texts.push(...extractedTexts);
  }
  
  // Also try to extract from other streams
  const currentUser = cfb.find('Current User');
  if (currentUser) {
    // Current User stream contains some metadata
  }
  
  // Create slides from extracted text
  const slides = createSlidesFromText(texts);
  const metadata = extractMetadataFromCFB(cfb, file);
  
  return {
    id: crypto.randomUUID(),
    fileName: file.name,
    fileSize: file.size,
    fileType: 'ppt',
    extractedAt: new Date().toISOString(),
    metadata: {
      ...metadata,
      totalSlides: slides.length,
    },
    slides,
    media: [],
    themes: [],
    masterSlides: [],
    customProperties: {
      parsedWith: 'cfb',
      note: 'Legacy PPT format - using CFB parser',
    },
  };
}

/**
 * Extract text from PowerPoint Document stream
 * Based on MS-PPT specification
 */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
function extractTextFromPPTStream(content: Uint8Array, cp: any): string[] {
  const texts: string[] = [];
  
  // PPT uses record-based format
  // Record header: 2 bytes version/instance, 2 bytes type, 4 bytes length
  
  // Text record types we're interested in:
  // 0x0FA0 = TextCharsAtom (Unicode)
  // 0x0FA8 = TextBytesAtom (ASCII)
  
  const view = new DataView(content.buffer, content.byteOffset, content.byteLength);
  let offset = 0;
  
  while (offset < content.length - 8) {
    try {
      // Skip version/instance field
      offset += 0; // recVerInst at offset, we read recType next
      const recType = view.getUint16(offset + 2, true);
      const recLen = view.getUint32(offset + 4, true);
      
      offset += 8;
      
      if (offset + recLen > content.length) break;
      
      // TextCharsAtom - Unicode text (UTF-16LE)
      if (recType === 0x0FA0 && recLen > 0) {
        const textData = content.slice(offset, offset + recLen);
        let text = '';
        
        if (cp && cp.utils) {
          text = cp.utils.decode(1200, textData); // 1200 = UTF-16LE
        } else {
          // Fallback: manual UTF-16LE decoding
          for (let i = 0; i < textData.length - 1; i += 2) {
            const charCode = textData[i] | (textData[i + 1] << 8);
            if (charCode > 0 && charCode < 0xFFFF) {
              text += String.fromCharCode(charCode);
            }
          }
        }
        
        text = cleanText(text);
        if (text && text.length > 2 && !isGarbage(text)) {
          texts.push(text);
        }
      }
      
      // TextBytesAtom - ASCII/ANSI text
      if (recType === 0x0FA8 && recLen > 0) {
        const textData = content.slice(offset, offset + recLen);
        let text = '';
        
        for (let i = 0; i < textData.length; i++) {
          const charCode = textData[i];
          if (charCode >= 32 && charCode < 127) {
            text += String.fromCharCode(charCode);
          } else if (charCode === 13 || charCode === 10) {
            text += '\n';
          }
        }
        
        text = cleanText(text);
        if (text && text.length > 2 && !isGarbage(text)) {
          texts.push(text);
        }
      }
      
      offset += recLen;
    } catch {
      offset++;
    }
  }
  
  // Also do a broader scan for Unicode strings
  const unicodeTexts = extractUnicodeStrings(content);
  texts.push(...unicodeTexts);
  
  // Remove duplicates
  return [...new Set(texts)];
}

/**
 * Extract Unicode strings from binary data
 */
function extractUnicodeStrings(data: Uint8Array): string[] {
  const texts: string[] = [];
  const minLength = 4;
  
  let currentUnicode = '';
  for (let i = 0; i < data.length - 1; i += 2) {
    const charCode = data[i] | (data[i + 1] << 8);
    
    if ((charCode >= 32 && charCode <= 126) || 
        (charCode >= 0x00A0 && charCode < 0x0600) || 
        charCode === 9 || charCode === 10 || charCode === 13) {
      currentUnicode += String.fromCharCode(charCode);
    } else {
      if (currentUnicode.length >= minLength) {
        const cleaned = cleanText(currentUnicode);
        if (cleaned && !isGarbage(cleaned)) {
          texts.push(cleaned);
        }
      }
      currentUnicode = '';
    }
  }
  
  if (currentUnicode.length >= minLength) {
    const cleaned = cleanText(currentUnicode);
    if (cleaned && !isGarbage(cleaned)) {
      texts.push(cleaned);
    }
  }
  
  return texts;
}

/**
 * Extract metadata from CFB document
 */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
function extractMetadataFromCFB(cfb: any, file: File): PresentationMetadata {
  let title = file.name.replace(/\.ppt$/i, '');
  let creator = '';
  let application = 'Microsoft PowerPoint (Legacy)';
  
  // Try to get Summary Information stream
  const summary = cfb.find('\x05SummaryInformation');
  if (summary && summary.content) {
    // Parse OLE property set (simplified)
    const content = summary.content;
    const text = new TextDecoder('utf-8', { fatal: false }).decode(content);
    
    const titleMatch = text.match(/[\x00-\x20]([A-Za-z][A-Za-z0-9 ]{3,50})[\x00]/);
    if (titleMatch) {
      title = cleanText(titleMatch[1]) || title;
    }
  }
  
  return {
    title,
    subject: '',
    creator,
    lastModifiedBy: '',
    created: '',
    modified: new Date(file.lastModified).toISOString(),
    revision: '',
    category: '',
    keywords: '',
    description: 'Legacy PowerPoint format (.ppt)',
    application,
    appVersion: '',
    company: '',
    manager: '',
    totalSlides: 0,
    totalWords: 0,
    totalParagraphs: 0,
    presentationFormat: 'PPT (Legacy)',
    template: '',
  };
}

/**
 * Fallback: Basic text extraction from binary
 */
function parseWithBasicExtraction(data: Uint8Array, file: File): ExtractedPresentation {
  const textContent = extractTextFromBinary(data);
  const metadata = extractBasicMetadata(data, file);
  const slides = createSlidesFromText(textContent);
  
  return {
    id: crypto.randomUUID(),
    fileName: file.name,
    fileSize: file.size,
    fileType: 'ppt',
    extractedAt: new Date().toISOString(),
    metadata,
    slides,
    media: [],
    themes: [],
    masterSlides: [],
    customProperties: {
      parsedWith: 'basic-extraction',
      note: 'Legacy PPT format - basic text extraction',
    },
  };
}

/**
 * Extract readable text from binary data (fallback)
 */
function extractTextFromBinary(data: Uint8Array): string[] {
  const texts: string[] = [];
  const minLength = 4;
  
  // Extract ASCII strings
  let currentAscii = '';
  for (let i = 0; i < data.length; i++) {
    const byte = data[i];
    if ((byte >= 32 && byte <= 126) || byte === 9 || byte === 10 || byte === 13) {
      currentAscii += String.fromCharCode(byte);
    } else {
      if (currentAscii.length >= minLength) {
        const cleaned = cleanText(currentAscii);
        if (cleaned && !isGarbage(cleaned)) {
          texts.push(cleaned);
        }
      }
      currentAscii = '';
    }
  }
  
  // Extract Unicode (UTF-16LE) strings
  let currentUnicode = '';
  for (let i = 0; i < data.length - 1; i += 2) {
    const charCode = data[i] | (data[i + 1] << 8);
    if ((charCode >= 32 && charCode <= 126) || charCode === 9 || charCode === 10 || charCode === 13) {
      currentUnicode += String.fromCharCode(charCode);
    } else if (charCode > 126 && charCode < 0xFFFF) {
      try {
        currentUnicode += String.fromCharCode(charCode);
      } catch {
        if (currentUnicode.length >= minLength) {
          const cleaned = cleanText(currentUnicode);
          if (cleaned && !isGarbage(cleaned)) {
            texts.push(cleaned);
          }
        }
        currentUnicode = '';
      }
    } else {
      if (currentUnicode.length >= minLength) {
        const cleaned = cleanText(currentUnicode);
        if (cleaned && !isGarbage(cleaned)) {
          texts.push(cleaned);
        }
      }
      currentUnicode = '';
    }
  }
  
  return [...new Set(texts)].filter(t => t.length >= minLength);
}

/**
 * Clean extracted text
 */
function cleanText(text: string): string {
  return text
    .replace(/[\x00-\x1F\x7F]/g, ' ') // Remove control characters
    .replace(/\s+/g, ' ') // Normalize whitespace
    .trim();
}

/**
 * Check if text is likely garbage/metadata rather than content
 */
function isGarbage(text: string): boolean {
  // Common patterns that indicate non-content
  const garbagePatterns = [
    /^[A-F0-9]{8,}$/i, // Hex strings
    /^[\x00-\x1F]+$/, // Control characters only
    /^[_\-\.]+$/, // Only special chars
    /^Root Entry$/i,
    /^PowerPoint Document$/i,
    /^Current User$/i,
    /^SummaryInformation$/i,
    /^DocumentSummaryInformation$/i,
    /^\d+$/i, // Only numbers
    /^[a-z]$/i, // Single letter
    /^\s*$/,  // Only whitespace
  ];
  
  return garbagePatterns.some(pattern => pattern.test(text));
}

/**
 * Extract basic metadata from PPT binary
 */
function extractBasicMetadata(data: Uint8Array, file: File): PresentationMetadata {
  // Try to find common metadata strings
  const dataString = new TextDecoder('utf-8', { fatal: false }).decode(data);
  
  let title = '';
  let creator = '';
  let application = 'Microsoft PowerPoint (Legacy)';
  
  // Look for title in property stream (simplified)
  const titleMatch = dataString.match(/Title[:\s]*([^\x00\n]+)/i);
  if (titleMatch) {
    title = cleanText(titleMatch[1]);
  }
  
  // Look for author
  const authorMatch = dataString.match(/Author[:\s]*([^\x00\n]+)/i);
  if (authorMatch) {
    creator = cleanText(authorMatch[1]);
  }
  
  return {
    title: title || file.name.replace(/\.ppt$/i, ''),
    subject: '',
    creator,
    lastModifiedBy: '',
    created: '',
    modified: new Date(file.lastModified).toISOString(),
    revision: '',
    category: '',
    keywords: '',
    description: 'Legacy PowerPoint format (.ppt)',
    application,
    appVersion: '',
    company: '',
    manager: '',
    totalSlides: 0,
    totalWords: 0,
    totalParagraphs: 0,
    presentationFormat: 'PPT (Legacy)',
    template: '',
  };
}

/**
 * Create slide objects from extracted text
 */
function createSlidesFromText(texts: string[]): SlideContent[] {
  // Group texts into potential slides
  // This is a heuristic approach since PPT structure is complex
  
  const slides: SlideContent[] = [];
  const contentTexts = texts.filter(t => t.length > 10 && !isMetadataString(t));
  
  if (contentTexts.length === 0) {
    return [{
      slideNumber: 1,
      title: 'Extracted Content',
      textContent: texts.filter(t => t.length >= 4),
      notes: '',
      shapes: [],
      images: [],
      tables: [],
    }];
  }
  
  // Create slides from content (estimate based on text length patterns)
  let currentSlide: SlideContent = {
    slideNumber: 1,
    title: '',
    textContent: [],
    notes: '',
    shapes: [],
    images: [],
    tables: [],
  };
  
  for (const text of contentTexts) {
    // Heuristic: Short text at start might be a title
    if (!currentSlide.title && text.length < 100 && text.length > 3) {
      currentSlide.title = text;
    } else {
      currentSlide.textContent.push(text);
    }
    
    // Create new slide after accumulating content
    if (currentSlide.textContent.length >= 5) {
      slides.push(currentSlide);
      currentSlide = {
        slideNumber: slides.length + 1,
        title: '',
        textContent: [],
        notes: '',
        shapes: [],
        images: [],
        tables: [],
      };
    }
  }
  
  // Add remaining content
  if (currentSlide.title || currentSlide.textContent.length > 0) {
    slides.push(currentSlide);
  }
  
  // Update metadata
  if (slides.length > 0) {
    slides.forEach((slide, index) => {
      slide.slideNumber = index + 1;
      if (!slide.title) {
        slide.title = `Slide ${index + 1}`;
      }
    });
  }
  
  return slides.length > 0 ? slides : [{
    slideNumber: 1,
    title: 'No Content Extracted',
    textContent: ['Unable to extract structured content from this legacy PPT file.'],
    notes: '',
    shapes: [],
    images: [],
    tables: [],
  }];
}

/**
 * Check if text is likely metadata rather than slide content
 */
function isMetadataString(text: string): boolean {
  const metaPatterns = [
    /microsoft/i,
    /powerpoint/i,
    /arial/i,
    /times new roman/i,
    /calibri/i,
    /^font$/i,
    /^\d{4}$/,
    /copyright/i,
  ];
  
  return metaPatterns.some(pattern => pattern.test(text));
}
