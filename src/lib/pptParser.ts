/**
 * PPT Parser - Extract data from legacy PowerPoint files (.ppt)
 * PPT files use a binary format (OLE Compound Document)
 * This is a basic parser that extracts what it can from the binary format
 */

import type { ExtractedPresentation, SlideContent, PresentationMetadata } from '../types';

/**
 * Parse a legacy PPT file
 * Note: Full PPT parsing requires complex binary parsing.
 * This implementation extracts basic text and metadata where possible.
 */
export async function parsePPT(file: File): Promise<ExtractedPresentation> {
  const buffer = await file.arrayBuffer();
  const data = new Uint8Array(buffer);
  
  // Extract text content using string scanning
  const textContent = extractTextFromBinary(data);
  const metadata = extractBasicMetadata(data, file);
  
  // Create slides from extracted text (best effort)
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
      note: 'Legacy PPT format - limited extraction capabilities',
    },
  };
}

/**
 * Extract readable text from binary data
 * Looks for ASCII and Unicode text strings
 */
function extractTextFromBinary(data: Uint8Array): string[] {
  const texts: string[] = [];
  const minLength = 4; // Minimum string length to consider
  
  // Extract ASCII strings
  let currentAscii = '';
  for (let i = 0; i < data.length; i++) {
    const byte = data[i];
    // Printable ASCII range (space to ~) plus common whitespace
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
  
  // Extract Unicode (UTF-16LE) strings - common in Office documents
  let currentUnicode = '';
  for (let i = 0; i < data.length - 1; i += 2) {
    const charCode = data[i] | (data[i + 1] << 8);
    if ((charCode >= 32 && charCode <= 126) || charCode === 9 || charCode === 10 || charCode === 13) {
      currentUnicode += String.fromCharCode(charCode);
    } else if (charCode > 126 && charCode < 0xFFFF) {
      // Extended characters
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
  
  // Remove duplicates and filter
  const unique = [...new Set(texts)];
  return unique.filter(t => t.length >= minLength);
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
