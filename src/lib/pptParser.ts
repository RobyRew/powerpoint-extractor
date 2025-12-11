/**
 * PPT Parser - Complete implementation for legacy PowerPoint files (.ppt)
 * Based on [MS-PPT]: PowerPoint (.ppt) Binary File Format
 * 
 * PPT files are OLE Compound Documents (CFB) containing multiple streams.
 * The main content is in the "PowerPoint Document" stream which contains
 * records in a hierarchical structure.
 */

import type { ExtractedPresentation, SlideContent, PresentationMetadata, MediaInfo } from '../types';
import CFB from 'cfb';

// ============================================================================
// RECORD TYPES from [MS-PPT] specification
// ============================================================================

const RecordType = {
  // Text records - these contain the actual text content
  RT_TextCharsAtom: 0x0FA0,      // Unicode text (UTF-16LE)
  RT_TextBytesAtom: 0x0FA8,      // ASCII/ANSI text (Windows-1252)
  RT_CString: 0x0FBA,            // Unicode string
  RT_TextHeaderAtom: 0x0F9F,     // Text type indicator
  
  // Container records
  RT_Document: 0x03E8,
  RT_Slide: 0x03EE,
  RT_SlideListWithText: 0x0FF0,
  RT_MainMaster: 0x03F8,
  RT_Notes: 0x03F0,
  RT_Drawing: 0x040C,
  RT_List: 0x07D0,
  
  // Office Art containers
  OfficeArtSpContainer: 0xF004,
  OfficeArtClientTextbox: 0xF00D,
  OfficeArtDgContainer: 0xF002,
  OfficeArtSpgrContainer: 0xF003,
  
  // Blip (image) records
  OfficeArtBlipJPEG: 0xF01D,
  OfficeArtBlipJPEG2: 0xF02A,
  OfficeArtBlipPNG: 0xF01E,
  OfficeArtBlipEMF: 0xF01A,
  OfficeArtBlipWMF: 0xF01B,
  OfficeArtBlipDIB: 0xF01F,
};

// Container record types that we should recurse into
const CONTAINER_TYPES = new Set([
  RecordType.RT_Document,
  RecordType.RT_Slide,
  RecordType.RT_SlideListWithText,
  RecordType.RT_MainMaster,
  RecordType.RT_Notes,
  RecordType.RT_Drawing,
  RecordType.RT_List,
  RecordType.OfficeArtSpContainer,
  RecordType.OfficeArtClientTextbox,
  RecordType.OfficeArtDgContainer,
  RecordType.OfficeArtSpgrContainer,
  0x03F2, // RT_Environment
  0x040B, // RT_DrawingGroup
  0x07D5, // RT_FontCollection
  0x0FD9, // RT_HeadersFooters
  0x1388, // RT_ProgTags
  0xF000, // OfficeArtDggContainer
  0xF001, // OfficeArtBStoreContainer
]);

// ============================================================================
// BINARY READER
// ============================================================================

class BinaryReader {
  private data: Uint8Array;
  private view: DataView;
  public pos: number = 0;

  constructor(data: Uint8Array | ArrayBuffer) {
    this.data = data instanceof Uint8Array ? data : new Uint8Array(data);
    this.view = new DataView(this.data.buffer, this.data.byteOffset, this.data.byteLength);
  }

  get length(): number {
    return this.data.length;
  }

  get remaining(): number {
    return this.data.length - this.pos;
  }

  seek(pos: number): void {
    this.pos = Math.max(0, Math.min(pos, this.data.length));
  }

  skip(n: number): void {
    this.pos += n;
  }

  readUInt8(): number {
    if (this.pos >= this.data.length) return 0;
    return this.data[this.pos++];
  }

  readUInt16LE(): number {
    if (this.pos + 2 > this.data.length) return 0;
    const val = this.view.getUint16(this.pos, true);
    this.pos += 2;
    return val;
  }

  readUInt32LE(): number {
    if (this.pos + 4 > this.data.length) return 0;
    const val = this.view.getUint32(this.pos, true);
    this.pos += 4;
    return val;
  }

  readInt32LE(): number {
    if (this.pos + 4 > this.data.length) return 0;
    const val = this.view.getInt32(this.pos, true);
    this.pos += 4;
    return val;
  }

  readBytes(n: number): Uint8Array {
    const end = Math.min(this.pos + n, this.data.length);
    const bytes = this.data.slice(this.pos, end);
    this.pos = end;
    return bytes;
  }

  /**
   * Read UTF-16LE string (2 bytes per character)
   */
  readUTF16LE(byteLength: number): string {
    const bytes = this.readBytes(byteLength);
    let result = '';
    for (let i = 0; i + 1 < bytes.length; i += 2) {
      const code = bytes[i] | (bytes[i + 1] << 8);
      if (code === 0) break;
      result += String.fromCharCode(code);
    }
    return result;
  }

  /**
   * Read Windows-1252 encoded string (1 byte per character)
   * This is the encoding used in TextBytesAtom
   */
  readWindows1252(byteLength: number): string {
    const bytes = this.readBytes(byteLength);
    // Windows-1252 to Unicode mapping for bytes 0x80-0x9F
    const win1252Map: Record<number, number> = {
      0x80: 0x20AC, 0x82: 0x201A, 0x83: 0x0192, 0x84: 0x201E, 0x85: 0x2026,
      0x86: 0x2020, 0x87: 0x2021, 0x88: 0x02C6, 0x89: 0x2030, 0x8A: 0x0160,
      0x8B: 0x2039, 0x8C: 0x0152, 0x8E: 0x017D, 0x91: 0x2018, 0x92: 0x2019,
      0x93: 0x201C, 0x94: 0x201D, 0x95: 0x2022, 0x96: 0x2013, 0x97: 0x2014,
      0x98: 0x02DC, 0x99: 0x2122, 0x9A: 0x0161, 0x9B: 0x203A, 0x9C: 0x0153,
      0x9E: 0x017E, 0x9F: 0x0178,
    };
    
    let result = '';
    for (let i = 0; i < bytes.length; i++) {
      const byte = bytes[i];
      if (byte === 0) break;
      if (byte >= 0x80 && byte <= 0x9F && win1252Map[byte]) {
        result += String.fromCharCode(win1252Map[byte]);
      } else {
        result += String.fromCharCode(byte);
      }
    }
    return result;
  }

  slice(start: number, end: number): BinaryReader {
    return new BinaryReader(this.data.slice(start, end));
  }
}

// ============================================================================
// PPT PARSER
// ============================================================================

interface ParseResult {
  texts: string[];
  images: MediaInfo[];
  slideTexts: Map<number, string[]>;
  metadata: Partial<PresentationMetadata>;
}

/**
 * Parse the PowerPoint Document stream
 */
function parsePPTStream(data: Uint8Array): ParseResult {
  const reader = new BinaryReader(data);
  const result: ParseResult = {
    texts: [],
    images: [],
    slideTexts: new Map(),
    metadata: {},
  };
  
  let currentSlide = 0;
  
  parseRecords(reader, data.length, result, 0, () => currentSlide, (n) => { currentSlide = n; });
  
  return result;
}

/**
 * Parse records recursively
 */
function parseRecords(
  reader: BinaryReader,
  endPos: number,
  result: ParseResult,
  depth: number,
  getCurrentSlide: () => number,
  setCurrentSlide: (n: number) => void
): void {
  const maxDepth = 50;
  if (depth > maxDepth) return;
  
  let iterations = 0;
  const maxIterations = 100000;
  
  while (reader.pos + 8 <= endPos && iterations < maxIterations) {
    iterations++;
    const startPos = reader.pos;
    
    // Read record header
    const recVerInstance = reader.readUInt16LE();
    const recType = reader.readUInt16LE();
    const recLen = reader.readUInt32LE();
    
    const recVer = recVerInstance & 0x0F;
    
    // Validate record
    if (recLen > 100000000 || reader.pos + recLen > endPos + 8) {
      reader.seek(startPos + 1);
      continue;
    }
    
    const recordEnd = reader.pos + recLen;
    
    // Process based on record type
    switch (recType) {
      case RecordType.RT_TextCharsAtom: {
        // Unicode text (UTF-16LE)
        const text = reader.readUTF16LE(recLen);
        const cleaned = cleanText(text);
        if (cleaned && isValidText(cleaned)) {
          result.texts.push(cleaned);
          addToSlide(result.slideTexts, getCurrentSlide(), cleaned);
        }
        break;
      }
      
      case RecordType.RT_TextBytesAtom: {
        // ANSI text (Windows-1252)
        const text = reader.readWindows1252(recLen);
        const cleaned = cleanText(text);
        if (cleaned && isValidText(cleaned)) {
          result.texts.push(cleaned);
          addToSlide(result.slideTexts, getCurrentSlide(), cleaned);
        }
        break;
      }
      
      case RecordType.RT_CString: {
        // Unicode string
        const text = reader.readUTF16LE(recLen);
        const cleaned = cleanText(text);
        if (cleaned && isValidText(cleaned) && !isSystemString(cleaned)) {
          result.texts.push(cleaned);
          addToSlide(result.slideTexts, getCurrentSlide(), cleaned);
        }
        break;
      }
      
      case RecordType.RT_Slide: {
        // New slide
        setCurrentSlide(getCurrentSlide() + 1);
        // Recurse into slide container
        const subReader = reader.slice(reader.pos, recordEnd);
        parseRecords(subReader, recLen, result, depth + 1, getCurrentSlide, setCurrentSlide);
        break;
      }
      
      case RecordType.OfficeArtBlipJPEG:
      case RecordType.OfficeArtBlipJPEG2: {
        // JPEG image
        if (recLen > 17) {
          reader.skip(17); // Skip UID and tag
          const imgData = reader.readBytes(recLen - 17);
          if (imgData.length > 100) {
            result.images.push({
              name: `image_${result.images.length + 1}.jpg`,
              type: 'image/jpeg',
              size: imgData.length,
              extension: 'jpg',
              data: arrayToBase64(imgData),
            });
          }
        }
        break;
      }
      
      case RecordType.OfficeArtBlipPNG: {
        // PNG image
        if (recLen > 17) {
          reader.skip(17);
          const imgData = reader.readBytes(recLen - 17);
          if (imgData.length > 100) {
            result.images.push({
              name: `image_${result.images.length + 1}.png`,
              type: 'image/png',
              size: imgData.length,
              extension: 'png',
              data: arrayToBase64(imgData),
            });
          }
        }
        break;
      }
      
      default: {
        // Check if this is a container we should recurse into
        // Container records have recVer = 0xF
        if (recVer === 0xF || CONTAINER_TYPES.has(recType)) {
          const subReader = reader.slice(reader.pos, recordEnd);
          parseRecords(subReader, recLen, result, depth + 1, getCurrentSlide, setCurrentSlide);
        }
        break;
      }
    }
    
    reader.seek(recordEnd);
  }
}

function addToSlide(slideTexts: Map<number, string[]>, slideNum: number, text: string): void {
  if (!slideTexts.has(slideNum)) {
    slideTexts.set(slideNum, []);
  }
  slideTexts.get(slideNum)!.push(text);
}

/**
 * Clean text content
 */
function cleanText(text: string): string {
  return text
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '')
    .replace(/\x0D/g, '\n')
    .replace(/\t/g, ' ')
    .trim();
}

/**
 * Check if text is valid content (not garbage)
 */
function isValidText(text: string): boolean {
  if (!text || text.length < 2) return false;
  
  // Must have at least some readable characters
  const readableChars = text.replace(/[\s\x00-\x1F]/g, '');
  if (readableChars.length === 0) return false;
  
  // Count character types
  let latin = 0;       // A-Za-z and Latin extended
  let digits = 0;      // 0-9
  let punctuation = 0; // Common punctuation
  let spaces = 0;      // Whitespace
  let highUnicode = 0; // Unusual Unicode ranges
  let control = 0;     // Control characters
  
  for (let i = 0; i < text.length; i++) {
    const code = text.charCodeAt(i);
    
    if ((code >= 0x41 && code <= 0x5A) ||    // A-Z
        (code >= 0x61 && code <= 0x7A) ||    // a-z
        (code >= 0xC0 && code <= 0xFF) ||    // Latin-1 Supplement (À-ÿ)
        (code >= 0x100 && code <= 0x17F) ||  // Latin Extended-A (ă, ș, ț, etc.)
        (code >= 0x180 && code <= 0x24F)) {  // Latin Extended-B
      latin++;
    } else if (code >= 0x30 && code <= 0x39) {
      digits++;
    } else if (code === 0x20 || code === 0x0A || code === 0x0D || code === 0x09) {
      spaces++;
    } else if ((code >= 0x21 && code <= 0x2F) ||  // !"#$%&'()*+,-./
               (code >= 0x3A && code <= 0x40) ||  // :;<=>?@
               (code >= 0x5B && code <= 0x60) ||  // [\]^_`
               (code >= 0x7B && code <= 0x7E) ||  // {|}~
               code === 0x2018 || code === 0x2019 || // ' '
               code === 0x201C || code === 0x201D || // " "
               code === 0x2013 || code === 0x2014 || // – —
               code === 0x2026) {                    // …
      punctuation++;
    } else if (code < 0x20 || (code >= 0x7F && code <= 0x9F)) {
      control++;
    } else if (code >= 0x0400 && code <= 0x04FF) {
      // Cyrillic - count as latin equivalent
      latin++;
    } else if (code >= 0x0370 && code <= 0x03FF) {
      // Greek - count as latin equivalent
      latin++;
    } else if (code > 0x2000) {
      // High Unicode - often garbage in PPT files
      highUnicode++;
    }
  }
  
  const total = text.length;
  const textChars = latin + digits + punctuation;
  
  // Must have at least 50% text characters (letters, digits, punctuation)
  if (textChars / total < 0.5) return false;
  
  // Must have at least some letters (not just numbers and symbols)
  if (latin < 1) return false;
  
  // Too much high unicode is garbage
  if (highUnicode > 0 && highUnicode / total > 0.2) return false;
  
  // Any control characters (except in spaces) is garbage
  if (control > 0) return false;
  
  // Check for known garbage patterns
  const garbagePatterns = [
    /^[\x00-\x1F\x7F-\x9F]+$/,     // Control characters only
    /^[ༀ-࿿]+$/,                    // Tibetan block
    /^[Ḁ-ỿ]{1,3}$/,                // Very short Latin Extended Additional
    /^[℀-⅏]+$/,                    // Letterlike symbols
    /^[⿰-⿻]+$/,                    // CJK Radicals
    /^[가-힣]+$/,                   // Korean
    /^[一-龥]+$/,                   // CJK
    /^[ก-๛]+$/,                    // Thai
    /^[؀-ۿ]+$/,                    // Arabic
    /^[א-ת]+$/,                    // Hebrew
    /^[Ā-ſ]{1,4}$/,                // Very short Latin Extended
    /^[ᄀ-ᇿ]+$/,                   // Hangul Jamo
    /^[㄰-㆏]+$/,                   // CJK compatibility
    /^[＀-￯]+$/,                   // Halfwidth/Fullwidth forms
    /^[Ⰰ-Ⱞ]+$/,                    // Glagolitic
    /^[\uE000-\uF8FF]+$/,          // Private Use Area
    /^PK/,                          // ZIP signature
    /^\[Content_Types\]/,
    /^_rels\//,
    /^drs\//,
    /^ppt\//,
    /\.xml$/i,
    /\.rels$/i,
  ];
  
  if (garbagePatterns.some(p => p.test(text))) return false;
  
  return true;
}

/**
 * Check if text is a system/internal string
 */
function isSystemString(text: string): boolean {
  const patterns = [
    /^Root Entry$/i,
    /^PowerPoint Document$/i,
    /^Current User$/i,
    /^SummaryInformation$/i,
    /^DocumentSummaryInformation$/i,
    /^Pictures$/i,
    /^_VBA_PROJECT/i,
    /^___PPT\d*$/,
    /^\[Content_Types\]/,
    /^_rels[\/\\]/,
    /^drs[\/\\]/,
    /^ppt[\/\\]/,
    /\.xml$/i,
    /\.rels$/i,
    /^PK/,
    /^theme[\/\\]/i,
    /^slideLayout/i,
    /^slideMaster/i,
    /^Rectangle \d+$/i,
    /^Text Box \d+$/i,
    /^WordArt \d+$/i,
    /^Click to edit/i,
    /^Master (title|text|subtitle)/i,
    /^Second level$/i,
    /^Third level$/i,
    /^Fourth level$/i,
    /^Fifth level$/i,
    /^tableStyles/i,
    /^Microsoft/i,
    /^Arial$/i,
    /^Times New Roman$/i,
    /^Calibri$/i,
    /^Tahoma$/i,
    /^Verdana$/i,
    /^No Slide Title$/i,
    /^Amin$/i,
    /^Refren:?$/i,
    /^Presentacion/i,
    /^Presentation/i,
    /^METANOIA$/i,
    /^Personalizad/i,
    /^Fuentes usadas/i,
    /^Tema$/i,
  ];
  
  return patterns.some(p => p.test(text.trim()));
}

/**
 * Convert Uint8Array to base64
 */
function arrayToBase64(arr: Uint8Array): string {
  let binary = '';
  const chunkSize = 8192;
  for (let i = 0; i < arr.length; i += chunkSize) {
    const chunk = arr.slice(i, Math.min(i + chunkSize, arr.length));
    for (let j = 0; j < chunk.length; j++) {
      binary += String.fromCharCode(chunk[j]);
    }
  }
  return btoa(binary);
}

/**
 * Parse OLE property stream for metadata
 */
function parsePropertyStream(data: Uint8Array): Partial<PresentationMetadata> {
  const metadata: Partial<PresentationMetadata> = {};
  
  try {
    const reader = new BinaryReader(data);
    
    // OLE Property Set header
    // Byte order (2), Format version (2), OS Version (4), CLSID (16)
    reader.skip(28);
    
    const numSections = reader.readUInt32LE();
    if (numSections === 0 || numSections > 100) return metadata;
    
    // Read section info (FMTID + Offset)
    reader.skip(16); // FMTID
    const sectionOffset = reader.readUInt32LE();
    
    reader.seek(sectionOffset);
    
    // Section header
    reader.readUInt32LE(); // size
    const numProps = reader.readUInt32LE();
    
    if (numProps > 1000) return metadata;
    
    // Property entries
    const props: { id: number; offset: number }[] = [];
    for (let i = 0; i < numProps; i++) {
      props.push({
        id: reader.readUInt32LE(),
        offset: reader.readUInt32LE(),
      });
    }
    
    // Read property values
    for (const prop of props) {
      reader.seek(sectionOffset + prop.offset);
      const type = reader.readUInt32LE();
      
      if (type === 0x1E || type === 0x1F) { // VT_LPSTR or VT_LPWSTR
        const len = reader.readUInt32LE();
        if (len > 0 && len < 10000) {
          const str = type === 0x1F 
            ? reader.readUTF16LE(len * 2)
            : reader.readWindows1252(len);
          
          switch (prop.id) {
            case 2: metadata.title = str.trim(); break;
            case 3: metadata.subject = str.trim(); break;
            case 4: metadata.creator = str.trim(); break;
            case 5: metadata.keywords = str.trim(); break;
            case 6: metadata.description = str.trim(); break;
            case 8: metadata.lastModifiedBy = str.trim(); break;
            case 9: metadata.revision = str.trim(); break;
            case 18: metadata.application = str.trim(); break;
          }
        }
      }
    }
  } catch {
    // Ignore parsing errors
  }
  
  return metadata;
}

/**
 * Create slides from extracted texts
 */
function createSlides(texts: string[], slideTexts: Map<number, string[]>): SlideContent[] {
  // If we have slide-organized texts, use them
  if (slideTexts.size > 0) {
    const slides: SlideContent[] = [];
    const sortedKeys = Array.from(slideTexts.keys()).sort((a, b) => a - b);
    
    for (const slideNum of sortedKeys) {
      const slideTextList = slideTexts.get(slideNum) || [];
      if (slideTextList.length === 0) continue;
      
      // First text is usually the title
      const title = slideTextList[0] || `Slide ${slides.length + 1}`;
      const content = slideTextList.slice(1);
      
      slides.push({
        slideNumber: slides.length + 1,
        title: title,
        textContent: content,
        notes: '',
        shapes: [],
        images: [],
        tables: [],
      });
    }
    
    if (slides.length > 0) return slides;
  }
  
  // Fallback: organize texts into slides
  if (texts.length === 0) {
    return [{
      slideNumber: 1,
      title: 'No Content Found',
      textContent: ['Could not extract text from this presentation.'],
      notes: '',
      shapes: [],
      images: [],
      tables: [],
    }];
  }
  
  // Deduplicate while preserving order
  const seen = new Set<string>();
  const uniqueTexts = texts.filter(t => {
    const normalized = t.toLowerCase().trim();
    if (seen.has(normalized)) return false;
    seen.add(normalized);
    return true;
  });
  
  // Group into slides (heuristic: ~5-8 items per slide)
  const slides: SlideContent[] = [];
  let currentSlide: SlideContent = {
    slideNumber: 1,
    title: '',
    textContent: [],
    notes: '',
    shapes: [],
    images: [],
    tables: [],
  };
  
  for (const text of uniqueTexts) {
    // If no title yet and text is short, use as title
    if (!currentSlide.title && text.length < 100) {
      currentSlide.title = text;
    } else {
      currentSlide.textContent.push(text);
    }
    
    // Start new slide when we have enough content
    if (currentSlide.textContent.length >= 6) {
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
  
  // Add last slide
  if (currentSlide.title || currentSlide.textContent.length > 0) {
    slides.push(currentSlide);
  }
  
  // Ensure all slides have titles
  slides.forEach((slide, idx) => {
    slide.slideNumber = idx + 1;
    if (!slide.title) {
      slide.title = `Slide ${idx + 1}`;
    }
  });
  
  return slides;
}

// ============================================================================
// MAIN EXPORT FUNCTION
// ============================================================================

export async function parsePPT(file: File): Promise<ExtractedPresentation> {
  const buffer = await file.arrayBuffer();
  const data = new Uint8Array(buffer);
  
  try {
    // Parse as OLE Compound Document
    const cfb = CFB.read(data, { type: 'array' });
    
    let metadata: PresentationMetadata = {
      title: file.name.replace(/\.ppt$/i, ''),
      subject: '',
      creator: '',
      lastModifiedBy: '',
      created: '',
      modified: new Date(file.lastModified).toISOString(),
      revision: '',
      category: '',
      keywords: '',
      description: '',
      application: 'Microsoft PowerPoint',
      appVersion: '',
      company: '',
      manager: '',
      totalSlides: 0,
      totalWords: 0,
      totalParagraphs: 0,
      presentationFormat: '',
      template: '',
    };
    
    // Parse Summary Information for metadata
    try {
      const summaryEntry = cfb.find('\x05SummaryInformation');
      if (summaryEntry?.content) {
        const content = summaryEntry.content;
        const summaryData = content instanceof Uint8Array ? content : new Uint8Array(content);
        const summaryMeta = parsePropertyStream(summaryData);
        metadata = { ...metadata, ...summaryMeta };
      }
    } catch { /* ignore */ }
    
    // Find the PowerPoint Document stream
    const pptEntry = cfb.find('PowerPoint Document');
    if (!pptEntry?.content) {
      throw new Error('PowerPoint Document stream not found');
    }
    
    // Parse the PowerPoint stream
    const pptContent = pptEntry.content;
    const pptData = pptContent instanceof Uint8Array ? pptContent : new Uint8Array(pptContent);
    const parseResult = parsePPTStream(pptData);
    
    // Create slides from parsed data
    const slides = createSlides(parseResult.texts, parseResult.slideTexts);
    
    // Update metadata
    metadata.totalSlides = slides.length;
    
    // Count words
    let wordCount = 0;
    for (const slide of slides) {
      wordCount += countWords(slide.title);
      for (const text of slide.textContent) {
        wordCount += countWords(text);
      }
    }
    metadata.totalWords = wordCount;
    
    return {
      id: crypto.randomUUID(),
      fileName: file.name,
      fileSize: file.size,
      fileType: 'ppt',
      extractedAt: new Date().toISOString(),
      metadata,
      slides,
      media: parseResult.images,
      themes: [],
      masterSlides: [],
      customProperties: {
        parsedWith: 'ppt-parser-v2',
      },
    };
    
  } catch (error) {
    console.error('PPT parsing error:', error);
    
    // Return error result
    return {
      id: crypto.randomUUID(),
      fileName: file.name,
      fileSize: file.size,
      fileType: 'ppt',
      extractedAt: new Date().toISOString(),
      metadata: {
        title: file.name.replace(/\.ppt$/i, ''),
        subject: '',
        creator: '',
        lastModifiedBy: '',
        created: '',
        modified: new Date(file.lastModified).toISOString(),
        revision: '',
        category: '',
        keywords: '',
        description: '',
        application: 'Microsoft PowerPoint (Legacy)',
        appVersion: '',
        company: '',
        manager: '',
        totalSlides: 1,
        totalWords: 0,
        totalParagraphs: 0,
        presentationFormat: '',
        template: '',
      },
      slides: [{
        slideNumber: 1,
        title: 'Error',
        textContent: [
          'Failed to parse this PowerPoint file.',
          error instanceof Error ? error.message : 'Unknown error',
        ],
        notes: '',
        shapes: [],
        images: [],
        tables: [],
      }],
      media: [],
      themes: [],
      masterSlides: [],
      customProperties: {
        parsedWith: 'ppt-parser-v2',
        error: error instanceof Error ? error.message : 'Unknown error',
      },
    };
  }
}

function countWords(text: string): number {
  return text.split(/\s+/).filter(w => w.length > 0).length;
}
