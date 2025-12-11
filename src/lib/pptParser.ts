/**
 * PPT Parser - Full implementation for legacy PowerPoint files (.ppt)
 * Based on [MS-PPT]: PowerPoint (.ppt) Binary File Format
 * https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-ppt/
 * 
 * PPT files are OLE Compound Documents containing:
 * - "PowerPoint Document" stream - main presentation data
 * - "Current User" stream - current user info and edit offset
 * - "\x05SummaryInformation" - Dublin Core metadata
 * - "\x05DocumentSummaryInformation" - Document metadata
 * - "Pictures" stream - embedded images (optional)
 */

import type { ExtractedPresentation, SlideContent, PresentationMetadata, ShapeInfo, TableInfo, MediaInfo } from '../types';
import CFB from 'cfb';

// ============================================================================
// TYPE DEFINITIONS based on MS-PPT specification
// ============================================================================

interface RecordHeader {
  recVer: number;      // 4 bits - record version
  recInstance: number; // 12 bits - record instance
  recType: number;     // 16 bits - record type
  recLen: number;      // 32 bits - record length
}

interface ParsedSlide {
  slideId: number;
  texts: string[];
  title: string;
  notes: string;
  shapes: ShapeInfo[];
  tables: TableInfo[];
  images: MediaInfo[];
}

interface ParseContext {
  slides: Map<number, ParsedSlide>;
  currentSlideId: number;
  documentTexts: string[];
  masterTexts: string[];
  notesTexts: Map<number, string[]>;
  pictures: MediaInfo[];
  metadata: Partial<PresentationMetadata>;
}

// ============================================================================
// RECORD TYPE CONSTANTS from [MS-PPT] 2.13.24
// ============================================================================

const RecordType = {
  // Document records
  RT_Document: 0x03E8,
  RT_DocumentAtom: 0x03E9,
  RT_EndDocumentAtom: 0x03EA,
  RT_Slide: 0x03EE,
  RT_SlideAtom: 0x03EF,
  RT_Notes: 0x03F0,
  RT_NotesAtom: 0x03F1,
  RT_Environment: 0x03F2,
  RT_SlidePersistAtom: 0x03F3,
  RT_MainMaster: 0x03F8,
  RT_SlideShowSlideInfoAtom: 0x03F9,
  RT_SlideViewInfo: 0x03FA,
  RT_GuideAtom: 0x03FB,
  RT_ViewInfoAtom: 0x03FD,
  RT_SlideViewInfoAtom: 0x03FE,
  RT_VbaInfo: 0x03FF,
  RT_VbaInfoAtom: 0x0400,
  RT_SlideShowDocInfoAtom: 0x0401,
  RT_Summary: 0x0402,
  RT_DocRoutingSlipAtom: 0x0406,
  RT_OutlineViewInfo: 0x0407,
  RT_SorterViewInfo: 0x0408,
  RT_ExternalObjectList: 0x0409,
  RT_ExternalObjectListAtom: 0x040A,
  RT_DrawingGroup: 0x040B,
  RT_Drawing: 0x040C,
  RT_GridSpacing10Atom: 0x040D,
  RT_RoundTripTheme12Atom: 0x040E,
  RT_RoundTripColorMapping12Atom: 0x040F,
  RT_NamedShows: 0x0410,
  RT_NamedShow: 0x0411,
  RT_NamedShowSlidesAtom: 0x0412,
  RT_NotesTextViewInfo9: 0x0413,
  RT_NormalViewSetInfo9: 0x0414,
  
  // List records
  RT_List: 0x07D0,
  RT_FontCollection: 0x07D5,
  RT_BookmarkCollection: 0x07E3,
  RT_SoundCollection: 0x07E4,
  RT_SoundCollectionAtom: 0x07E5,
  RT_Sound: 0x07E6,
  RT_SoundDataBlob: 0x07E7,
  RT_BookmarkSeedAtom: 0x07E9,
  RT_ColorSchemeAtom: 0x07F0,
  RT_BlipCollection9: 0x07F8,
  RT_BlipEntity9Atom: 0x07F9,
  
  // Shape records
  RT_ExternalObjectRefAtom: 0x0BC1,
  RT_PlaceholderAtom: 0x0BC3,
  RT_ShapeAtom: 0x0BDB,
  RT_ShapeFlags10Atom: 0x0BDC,
  RT_RoundTripNewPlaceholderId12Atom: 0x0BDD,
  
  // Text records - IMPORTANT for text extraction
  RT_OutlineTextRefAtom: 0x0F9E,
  RT_TextHeaderAtom: 0x0F9F,
  RT_TextCharsAtom: 0x0FA0,      // Unicode text (UTF-16LE)
  RT_StyleTextPropAtom: 0x0FA1,
  RT_MasterTextPropAtom: 0x0FA2,
  RT_TextMasterStyleAtom: 0x0FA3,
  RT_TextCharFormatExceptionAtom: 0x0FA4,
  RT_TextParagraphFormatExceptionAtom: 0x0FA5,
  RT_TextRulerAtom: 0x0FA6,
  RT_TextBookmarkAtom: 0x0FA7,
  RT_TextBytesAtom: 0x0FA8,      // ASCII/ANSI text
  RT_TextSpecialInfoDefaultAtom: 0x0FA9,
  RT_TextSpecialInfoAtom: 0x0FAA,
  RT_DefaultRulerAtom: 0x0FAB,
  RT_StyleTextProp9Atom: 0x0FAC,
  RT_TextMasterStyle9Atom: 0x0FAD,
  RT_OutlineTextProps9: 0x0FAE,
  RT_OutlineTextPropsHeader9Atom: 0x0FAF,
  RT_TextDefaults9Atom: 0x0FB0,
  RT_StyleTextProp10Atom: 0x0FB1,
  RT_TextMasterStyle10Atom: 0x0FB2,
  RT_OutlineTextProps10: 0x0FB3,
  RT_TextDefaults10Atom: 0x0FB4,
  RT_OutlineTextProps11: 0x0FB5,
  RT_StyleTextProp11Atom: 0x0FB6,
  RT_FontEntityAtom: 0x0FB7,
  RT_FontEmbedDataBlob: 0x0FB8,
  RT_CString: 0x0FBA,            // Unicode string (null-terminated)
  RT_MetaFile: 0x0FC1,
  RT_ExternalOleObjectAtom: 0x0FC3,
  RT_Kinsoku: 0x0FC8,
  RT_Handout: 0x0FC9,
  RT_ExternalOleEmbed: 0x0FCC,
  RT_ExternalOleEmbedAtom: 0x0FCD,
  RT_ExternalOleLink: 0x0FCE,
  RT_BookmarkEntityAtom: 0x0FD0,
  RT_ExternalOleLinkAtom: 0x0FD1,
  RT_KinsokuAtom: 0x0FD2,
  RT_ExternalHyperlinkAtom: 0x0FD3,
  RT_ExternalHyperlink: 0x0FD7,
  RT_SlideNumberMCAtom: 0x0FD8,
  RT_HeadersFooters: 0x0FD9,
  RT_HeadersFootersAtom: 0x0FDA,
  
  // Interactive records
  RT_TextInteractiveInfoAtom: 0x0FDF,
  RT_ExternalHyperlink9: 0x0FE4,
  RT_RecolorInfoAtom: 0x0FE7,
  RT_ExternalOleControl: 0x0FEE,
  RT_SlideListWithText: 0x0FF0,
  RT_AnimationInfoAtom: 0x0FF1,
  RT_InteractiveInfo: 0x0FF2,
  RT_InteractiveInfoAtom: 0x0FF3,
  RT_UserEditAtom: 0x0FF5,
  RT_CurrentUserAtom: 0x0FF6,
  RT_DateTimeMCAtom: 0x0FF7,
  RT_GenericDateMCAtom: 0x0FF8,
  RT_HeaderMCAtom: 0x0FF9,
  RT_FooterMCAtom: 0x0FFA,
  RT_ExternalOleControlAtom: 0x0FFB,
  
  // Media records
  RT_ExternalMediaAtom: 0x1004,
  RT_ExternalVideo: 0x1005,
  RT_ExternalAviMovie: 0x1006,
  RT_ExternalMciMovie: 0x1007,
  RT_ExternalMidiAudio: 0x100D,
  RT_ExternalCdAudio: 0x100E,
  RT_ExternalWavAudioEmbedded: 0x100F,
  RT_ExternalWavAudioLink: 0x1010,
  RT_ExternalOleObjectStg: 0x1011,
  RT_ExternalCdAudioAtom: 0x1012,
  RT_ExternalWavAudioEmbeddedAtom: 0x1013,
  RT_AnimationInfo: 0x1014,
  RT_RtfDateTimeMCAtom: 0x1015,
  RT_ExternalHyperlinkFlagsAtom: 0x1018,
  
  // Program tags
  RT_ProgTags: 0x1388,
  RT_ProgStringTag: 0x1389,
  RT_ProgBinaryTag: 0x138A,
  RT_BinaryTagDataBlob: 0x138B,
  
  // Persist records
  RT_PrintOptionsAtom: 0x1770,
  RT_PersistDirectoryAtom: 0x1772,
  RT_PresentationAdvisorFlags9Atom: 0x177A,
  RT_HtmlDocInfo9Atom: 0x177B,
  RT_HtmlPublishInfoAtom: 0x177C,
  RT_HtmlPublishInfo9: 0x177D,
  RT_BroadcastDocInfo9: 0x177E,
  RT_BroadcastDocInfo9Atom: 0x177F,
  RT_EnvelopeFlags9Atom: 0x1784,
  RT_EnvelopeData9Atom: 0x1785,
  
  // Animation records
  RT_VisualShapeAtom: 0x2AFB,
  RT_HashCodeAtom: 0x2B00,
  RT_VisualPageAtom: 0x2B01,
  RT_BuildList: 0x2B02,
  RT_BuildAtom: 0x2B03,
  RT_ChartBuild: 0x2B04,
  RT_ChartBuildAtom: 0x2B05,
  RT_DiagramBuild: 0x2B06,
  RT_DiagramBuildAtom: 0x2B07,
  RT_ParaBuild: 0x2B08,
  RT_ParaBuildAtom: 0x2B09,
  RT_LevelInfoAtom: 0x2B0A,
  RT_RoundTripAnimationAtom12Atom: 0x2B0B,
  RT_RoundTripAnimationHashAtom12Atom: 0x2B0D,
  
  // Comment records
  RT_Comment10: 0x2EE0,
  RT_Comment10Atom: 0x2EE1,
  RT_CommentIndex10: 0x2EE4,
  RT_CommentIndex10Atom: 0x2EE5,
  RT_LinkedShape10Atom: 0x2EE6,
  RT_LinkedSlide10Atom: 0x2EE7,
  RT_SlideFlags10Atom: 0x2EEA,
  RT_SlideTime10Atom: 0x2EEB,
  RT_DiffTree10: 0x2EEC,
  RT_Diff10: 0x2EED,
  RT_Diff10Atom: 0x2EEE,
  RT_SlideListTableSize10Atom: 0x2EEF,
  RT_SlideListEntry10Atom: 0x2EF0,
  RT_SlideListTable10: 0x2EF1,
  
  // Crypto
  RT_CryptSession10Container: 0x2F14,
  RT_FontEmbedFlags10Atom: 0x32C8,
  RT_FilterPrivacyFlags10Atom: 0x36B0,
  RT_DocToolbarStates10Atom: 0x36B1,
  RT_PhotoAlbumInfo10Atom: 0x36B2,
  RT_SmartTagStore11Container: 0x36B3,
  
  // RoundTrip records
  RT_RoundTripSlideSyncInfo12: 0x3714,
  RT_RoundTripSlideSyncInfoAtom12: 0x3715,
  
  // Office Art records [MS-ODRAW]
  OfficeArtDggContainer: 0xF000,
  OfficeArtBStoreContainer: 0xF001,
  OfficeArtDgContainer: 0xF002,
  OfficeArtSpgrContainer: 0xF003,
  OfficeArtSpContainer: 0xF004,
  OfficeArtSolverContainer: 0xF005,
  OfficeArtFDGGBlock: 0xF006,
  OfficeArtFBSE: 0xF007,
  OfficeArtFDG: 0xF008,
  OfficeArtFSPGR: 0xF009,
  OfficeArtFSP: 0xF00A,
  OfficeArtFOPT: 0xF00B,
  OfficeArtClientTextbox: 0xF00D,
  OfficeArtChildAnchor: 0xF00F,
  OfficeArtClientAnchor: 0xF010,
  OfficeArtClientData: 0xF011,
  OfficeArtFConnectorRule: 0xF012,
  OfficeArtFArcRule: 0xF014,
  OfficeArtFCalloutRule: 0xF017,
  
  // Blip records (images)
  OfficeArtBlipEMF: 0xF01A,
  OfficeArtBlipWMF: 0xF01B,
  OfficeArtBlipPICT: 0xF01C,
  OfficeArtBlipJPEG: 0xF01D,
  OfficeArtBlipPNG: 0xF01E,
  OfficeArtBlipDIB: 0xF01F,
  OfficeArtBlipTIFF: 0xF029,
  OfficeArtBlipJPEG2: 0xF02A,
  
  // More Office Art
  OfficeArtFRITContainer: 0xF118,
  OfficeArtFDGSL: 0xF119,
  OfficeArtColorMRUContainer: 0xF11A,
  OfficeArtFPSPL: 0xF11D,
  OfficeArtSplitMenuColorContainer: 0xF11E,
  OfficeArtSecondaryFOPT: 0xF121,
  OfficeArtTertiaryFOPT: 0xF122,
};

// ============================================================================
// BINARY READER CLASS
// ============================================================================

class BinaryReader {
  private data: Uint8Array;
  private view: DataView;
  public position: number;
  public length: number;

  constructor(data: Uint8Array) {
    this.data = data;
    this.view = new DataView(data.buffer, data.byteOffset, data.byteLength);
    this.position = 0;
    this.length = data.length;
  }

  get remaining(): number {
    return this.length - this.position;
  }

  seek(position: number): void {
    this.position = Math.min(Math.max(0, position), this.length);
  }

  skip(bytes: number): void {
    this.position += bytes;
  }

  readUInt8(): number {
    if (this.position >= this.length) return 0;
    return this.data[this.position++];
  }

  readUInt16LE(): number {
    if (this.position + 2 > this.length) return 0;
    const value = this.view.getUint16(this.position, true);
    this.position += 2;
    return value;
  }

  readUInt32LE(): number {
    if (this.position + 4 > this.length) return 0;
    const value = this.view.getUint32(this.position, true);
    this.position += 4;
    return value;
  }

  readInt32LE(): number {
    if (this.position + 4 > this.length) return 0;
    const value = this.view.getInt32(this.position, true);
    this.position += 4;
    return value;
  }

  readBytes(length: number): Uint8Array {
    const end = Math.min(this.position + length, this.length);
    const bytes = this.data.slice(this.position, end);
    this.position = end;
    return bytes;
  }

  // Read UTF-16LE string (Unicode)
  readUnicodeString(byteLength: number): string {
    const bytes = this.readBytes(byteLength);
    let result = '';
    for (let i = 0; i < bytes.length - 1; i += 2) {
      const charCode = bytes[i] | (bytes[i + 1] << 8);
      if (charCode === 0) break; // Null terminator
      result += String.fromCharCode(charCode);
    }
    return result;
  }

  // Read ASCII/ANSI string
  readAsciiString(byteLength: number): string {
    const bytes = this.readBytes(byteLength);
    let result = '';
    for (let i = 0; i < bytes.length; i++) {
      if (bytes[i] === 0) break; // Null terminator
      result += String.fromCharCode(bytes[i]);
    }
    return result;
  }

  // Read record header (8 bytes)
  readRecordHeader(): RecordHeader | null {
    if (this.remaining < 8) return null;
    
    const recVerInstance = this.readUInt16LE();
    const recType = this.readUInt16LE();
    const recLen = this.readUInt32LE();
    
    return {
      recVer: recVerInstance & 0x0F,
      recInstance: (recVerInstance >> 4) & 0x0FFF,
      recType,
      recLen,
    };
  }

  // Create a sub-reader for a specific range
  subReader(length: number): BinaryReader {
    const subData = this.data.slice(this.position, this.position + length);
    return new BinaryReader(subData);
  }
}

// ============================================================================
// MAIN PARSER CLASS
// ============================================================================

class PPTParser {
  private reader: BinaryReader;
  private context: ParseContext;

  constructor(data: Uint8Array) {
    this.reader = new BinaryReader(data);
    this.context = {
      slides: new Map(),
      currentSlideId: 0,
      documentTexts: [],
      masterTexts: [],
      notesTexts: new Map(),
      pictures: [],
      metadata: {},
    };
  }

  parse(): ParseContext {
    this.parseRecords(this.reader, this.reader.length);
    return this.context;
  }

  private parseRecords(reader: BinaryReader, endPosition: number): void {
    const maxIterations = 1000000; // Safety limit
    let iterations = 0;

    while (reader.position < endPosition && reader.remaining >= 8 && iterations < maxIterations) {
      iterations++;
      const recordStartPos = reader.position;
      const header = reader.readRecordHeader();
      
      if (!header) break;
      
      // Sanity check on record length
      if (header.recLen > reader.remaining || header.recLen > 100000000) {
        // Invalid record length, try to skip 1 byte and continue
        reader.seek(recordStartPos + 1);
        continue;
      }

      const recordEnd = reader.position + header.recLen;
      
      try {
        this.processRecord(reader, header);
      } catch (e) {
        // Continue on error
        console.warn(`Error processing record type 0x${header.recType.toString(16)}:`, e);
      }

      // Ensure we move past this record
      reader.seek(recordEnd);
    }
  }

  private processRecord(reader: BinaryReader, header: RecordHeader): void {
    const recordData = reader.subReader(header.recLen);

    switch (header.recType) {
      // Container records - recurse into them
      case RecordType.RT_Document:
      case RecordType.RT_Slide:
      case RecordType.RT_Notes:
      case RecordType.RT_MainMaster:
      case RecordType.RT_SlideListWithText:
      case RecordType.RT_List:
      case RecordType.RT_Environment:
      case RecordType.RT_DrawingGroup:
      case RecordType.RT_Drawing:
      case RecordType.RT_HeadersFooters:
      case RecordType.RT_ProgTags:
      case RecordType.RT_ExternalObjectList:
      case RecordType.RT_FontCollection:
      case RecordType.RT_SoundCollection:
      case RecordType.RT_Handout:
      case RecordType.RT_VbaInfo:
      case RecordType.OfficeArtDggContainer:
      case RecordType.OfficeArtBStoreContainer:
      case RecordType.OfficeArtDgContainer:
      case RecordType.OfficeArtSpgrContainer:
      case RecordType.OfficeArtSpContainer:
      case RecordType.OfficeArtSolverContainer:
        this.parseRecords(recordData, header.recLen);
        break;

      // Text records
      case RecordType.RT_TextCharsAtom:
        this.parseTextCharsAtom(recordData, header);
        break;

      case RecordType.RT_TextBytesAtom:
        this.parseTextBytesAtom(recordData, header);
        break;

      case RecordType.RT_CString:
        this.parseCString(recordData, header);
        break;

      case RecordType.RT_TextHeaderAtom:
        this.parseTextHeaderAtom(recordData, header);
        break;

      // Slide records
      case RecordType.RT_SlidePersistAtom:
        this.parseSlidePersistAtom(recordData, header);
        break;

      case RecordType.RT_SlideAtom:
        this.parseSlideAtom(recordData, header);
        break;

      // Office Art text
      case RecordType.OfficeArtClientTextbox:
        this.parseOfficeArtClientTextbox(recordData, header);
        break;

      // Image records
      case RecordType.OfficeArtBlipJPEG:
      case RecordType.OfficeArtBlipJPEG2:
        this.parseBlipJPEG(recordData, header);
        break;

      case RecordType.OfficeArtBlipPNG:
        this.parseBlipPNG(recordData, header);
        break;

      case RecordType.OfficeArtBlipEMF:
      case RecordType.OfficeArtBlipWMF:
      case RecordType.OfficeArtBlipPICT:
      case RecordType.OfficeArtBlipDIB:
      case RecordType.OfficeArtBlipTIFF:
        this.parseBlipGeneric(recordData, header);
        break;

      // Metadata records
      case RecordType.RT_DocumentAtom:
        this.parseDocumentAtom(recordData, header);
        break;

      // Default: skip unknown records
      default:
        break;
    }
  }

  // ============================================================================
  // TEXT PARSING
  // ============================================================================

  private parseTextCharsAtom(reader: BinaryReader, header: RecordHeader): void {
    // TextCharsAtom contains Unicode (UTF-16LE) text
    const text = reader.readUnicodeString(header.recLen);
    const cleanedText = this.cleanText(text);
    
    if (cleanedText && cleanedText.length > 0) {
      this.context.documentTexts.push(cleanedText);
    }
  }

  private parseTextBytesAtom(reader: BinaryReader, header: RecordHeader): void {
    // TextBytesAtom contains ASCII/ANSI text
    const text = reader.readAsciiString(header.recLen);
    const cleanedText = this.cleanText(text);
    
    if (cleanedText && cleanedText.length > 0) {
      this.context.documentTexts.push(cleanedText);
    }
  }

  private parseCString(reader: BinaryReader, header: RecordHeader): void {
    // CString is a Unicode null-terminated string
    const text = reader.readUnicodeString(header.recLen);
    const cleanedText = this.cleanText(text);
    
    if (cleanedText && cleanedText.length > 0 && !this.isSystemString(cleanedText)) {
      this.context.documentTexts.push(cleanedText);
    }
  }

  private parseTextHeaderAtom(reader: BinaryReader, _header: RecordHeader): void {
    // TextHeaderAtom specifies the type of text that follows
    // Just read and skip the text type
    reader.readUInt32LE();
  }

  private parseOfficeArtClientTextbox(reader: BinaryReader, header: RecordHeader): void {
    // OfficeArtClientTextbox contains text in shapes
    // It's a container with text records inside
    this.parseRecords(reader, header.recLen);
  }

  // ============================================================================
  // SLIDE PARSING
  // ============================================================================

  private parseSlidePersistAtom(reader: BinaryReader, _header: RecordHeader): void {
    // SlidePersistAtom contains slide ID and other info
    reader.readUInt32LE(); // persistIdRef
    reader.skip(4); // reserved
    reader.readUInt32LE(); // cTexts
    const slideId = reader.readUInt32LE();
    reader.skip(4); // reserved

    this.context.currentSlideId = slideId;
    
    if (!this.context.slides.has(slideId)) {
      this.context.slides.set(slideId, {
        slideId,
        texts: [],
        title: '',
        notes: '',
        shapes: [],
        tables: [],
        images: [],
      });
    }
  }

  private parseSlideAtom(reader: BinaryReader, _header: RecordHeader): void {
    // SlideAtom contains slide layout and master slide reference
    // Read and skip all fields - they advance the reader position
    reader.readInt32LE();  // geom
    reader.readBytes(8);   // rgPlaceholderTypes
    reader.readUInt32LE(); // masterIdRef
    reader.readUInt32LE(); // notesIdRef
    reader.readUInt16LE(); // slideFlags
    reader.skip(2);
  }

  // ============================================================================
  // IMAGE PARSING
  // ============================================================================

  private parseBlipJPEG(reader: BinaryReader, header: RecordHeader): void {
    // Skip the header (UID and tag)
    const uidSize = header.recInstance === 0x46A || header.recInstance === 0x6E2 ? 16 : 17;
    if (header.recLen <= uidSize) return;
    
    reader.skip(uidSize);
    
    const imageData = reader.readBytes(header.recLen - uidSize);
    if (imageData.length > 100) {
      this.context.pictures.push({
        name: `image_${this.context.pictures.length + 1}.jpg`,
        type: 'image/jpeg',
        size: imageData.length,
        extension: 'jpg',
        data: this.arrayToBase64(imageData),
      });
    }
  }

  private parseBlipPNG(reader: BinaryReader, header: RecordHeader): void {
    // Skip the header (UID and tag)
    const uidSize = header.recInstance === 0x6E0 ? 16 : 17;
    if (header.recLen <= uidSize) return;
    
    reader.skip(uidSize);
    
    const imageData = reader.readBytes(header.recLen - uidSize);
    if (imageData.length > 100) {
      this.context.pictures.push({
        name: `image_${this.context.pictures.length + 1}.png`,
        type: 'image/png',
        size: imageData.length,
        extension: 'png',
        data: this.arrayToBase64(imageData),
      });
    }
  }

  private parseBlipGeneric(reader: BinaryReader, header: RecordHeader): void {
    // Generic handler for EMF, WMF, PICT, DIB, TIFF
    const extensions: Record<number, { ext: string; mime: string }> = {
      [RecordType.OfficeArtBlipEMF]: { ext: 'emf', mime: 'image/emf' },
      [RecordType.OfficeArtBlipWMF]: { ext: 'wmf', mime: 'image/wmf' },
      [RecordType.OfficeArtBlipPICT]: { ext: 'pict', mime: 'image/pict' },
      [RecordType.OfficeArtBlipDIB]: { ext: 'bmp', mime: 'image/bmp' },
      [RecordType.OfficeArtBlipTIFF]: { ext: 'tiff', mime: 'image/tiff' },
    };

    const info = extensions[header.recType] || { ext: 'bin', mime: 'application/octet-stream' };
    
    if (header.recLen <= 16) return;
    
    // Skip UID
    reader.skip(16);
    
    const imageData = reader.readBytes(header.recLen - 16);
    if (imageData.length > 100) {
      this.context.pictures.push({
        name: `image_${this.context.pictures.length + 1}.${info.ext}`,
        type: info.mime,
        size: imageData.length,
        extension: info.ext,
        data: this.arrayToBase64(imageData),
      });
    }
  }

  // ============================================================================
  // METADATA PARSING
  // ============================================================================

  private parseDocumentAtom(reader: BinaryReader, _header: RecordHeader): void {
    // DocumentAtom contains document-level settings
    const slideSizeX = reader.readInt32LE();
    const slideSizeY = reader.readInt32LE();
    // Skip remaining fields - we just need slide dimensions
    reader.readInt32LE();  // notesWidth
    reader.readInt32LE();  // notesHeight
    reader.readInt32LE();  // serverZoom
    reader.readUInt32LE(); // notesMasterPersistIdRef
    reader.readUInt32LE(); // handoutMasterPersistIdRef
    reader.readUInt16LE(); // firstSlideNumber
    reader.readUInt16LE(); // slideSizeType
    reader.readUInt8();    // saveWithFonts
    reader.readUInt8();    // omitTitlePlace
    reader.readUInt8();    // rightToLeft
    reader.readUInt8();    // showComments

    // Convert EMUs to inches for format
    const widthInches = slideSizeX / 914400;
    const heightInches = slideSizeY / 914400;
    
    this.context.metadata.presentationFormat = `${widthInches.toFixed(1)} x ${heightInches.toFixed(1)} inches`;
  }

  // ============================================================================
  // UTILITIES
  // ============================================================================

  private cleanText(text: string): string {
    return text
      .replace(/\r\n/g, '\n')
      .replace(/\r/g, '\n')
      .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/g, '')
      .replace(/\t/g, ' ')
      .replace(/\n+/g, '\n')
      .replace(/ +/g, ' ')
      .trim();
  }

  private isSystemString(text: string): boolean {
    const systemPatterns = [
      /^[\x00-\x1F]+$/,
      /^[A-F0-9]{8,}$/i,
      /^Root Entry$/i,
      /^PowerPoint Document$/i,
      /^Current User$/i,
      /^SummaryInformation$/i,
      /^DocumentSummaryInformation$/i,
      /^Pictures$/i,
      /^_VBA_PROJECT/i,
      /^[*]PPT/,
      /^\d+$/,
    ];
    return systemPatterns.some(p => p.test(text));
  }

  private arrayToBase64(array: Uint8Array): string {
    let binary = '';
    const chunkSize = 8192;
    for (let i = 0; i < array.length; i += chunkSize) {
      const chunk = array.slice(i, i + chunkSize);
      for (let j = 0; j < chunk.length; j++) {
        binary += String.fromCharCode(chunk[j]);
      }
    }
    return btoa(binary);
  }
}

// ============================================================================
// ADDITIONAL TEXT EXTRACTION METHODS
// ============================================================================

/**
 * Extract text by scanning for Unicode patterns
 * This catches text that might be missed by record parsing
 */
function extractTextByScanning(data: Uint8Array): string[] {
  const texts: string[] = [];
  const minLength = 3;
  
  // Scan for UTF-16LE strings (Unicode)
  let currentText = '';
  for (let i = 0; i < data.length - 1; i += 2) {
    const charCode = data[i] | (data[i + 1] << 8);
    
    // Valid printable character or whitespace
    if ((charCode >= 32 && charCode < 127) || 
        (charCode >= 160 && charCode < 65536) ||
        charCode === 9 || charCode === 10 || charCode === 13) {
      currentText += String.fromCharCode(charCode);
    } else {
      if (currentText.length >= minLength) {
        const cleaned = currentText.trim();
        if (cleaned.length >= minLength && !isGarbageText(cleaned)) {
          texts.push(cleaned);
        }
      }
      currentText = '';
    }
  }
  
  if (currentText.length >= minLength) {
    const cleaned = currentText.trim();
    if (cleaned.length >= minLength && !isGarbageText(cleaned)) {
      texts.push(cleaned);
    }
  }

  // Scan for ASCII strings
  currentText = '';
  for (let i = 0; i < data.length; i++) {
    const byte = data[i];
    
    if ((byte >= 32 && byte < 127) || byte === 9 || byte === 10 || byte === 13) {
      currentText += String.fromCharCode(byte);
    } else {
      if (currentText.length >= minLength) {
        const cleaned = currentText.trim();
        if (cleaned.length >= minLength && !isGarbageText(cleaned)) {
          texts.push(cleaned);
        }
      }
      currentText = '';
    }
  }
  
  if (currentText.length >= minLength) {
    const cleaned = currentText.trim();
    if (cleaned.length >= minLength && !isGarbageText(cleaned)) {
      texts.push(cleaned);
    }
  }

  return texts;
}

/**
 * Check if text is likely garbage/metadata
 */
function isGarbageText(text: string): boolean {
  const garbagePatterns = [
    /^[A-F0-9]{8,}$/i,
    /^[_\-\.]+$/,
    /^Root Entry$/i,
    /^PowerPoint Document$/i,
    /^Current User$/i,
    /^SummaryInformation$/i,
    /^DocumentSummaryInformation$/i,
    /^Pictures$/i,
    /^\d+$/,
    /^[a-z]$/i,
    /^\s*$/,
    /^[\x00-\x1F]+$/,
    /^[*\[\]{}|\\\/]+$/,
    /^PPT/i,
    /^Microsoft Office/i,
    /^Arial$/i,
    /^Times New Roman$/i,
    /^Calibri$/i,
    /^Tahoma$/i,
    /^Verdana$/i,
    /^[A-Z]{1,2}\d{1,2}$/i,
  ];
  
  return garbagePatterns.some(pattern => pattern.test(text));
}

/**
 * Parse OLE Summary Information stream for metadata
 */
function parseSummaryInformation(data: Uint8Array): Partial<PresentationMetadata> {
  const metadata: Partial<PresentationMetadata> = {};
  
  try {
    const reader = new BinaryReader(data);
    
    // Skip byte order (2), version (2), system identifier (4), CLSID (16)
    reader.skip(24);
    
    // Number of property sets
    const numPropertySets = reader.readUInt32LE();
    if (numPropertySets === 0) return metadata;
    
    // Skip FMTID (16) and offset (4) for first property set
    reader.skip(20);
    
    // Property set header
    reader.readUInt32LE(); // size
    const numProperties = reader.readUInt32LE();
    
    // Property ID/Offset pairs
    const properties: Array<{ id: number; offset: number }> = [];
    for (let i = 0; i < numProperties && i < 100; i++) {
      const id = reader.readUInt32LE();
      const offset = reader.readUInt32LE();
      properties.push({ id, offset });
    }
    
    // Read property values
    const baseOffset = reader.position - (numProperties * 8) - 8;
    
    for (const prop of properties) {
      reader.seek(baseOffset + prop.offset);
      const type = reader.readUInt32LE();
      
      // VT_LPSTR (30) or VT_LPWSTR (31)
      if (type === 30 || type === 31) {
        const strLen = reader.readUInt32LE();
        const str = type === 31 
          ? reader.readUnicodeString(strLen * 2)
          : reader.readAsciiString(strLen);
        
        // Property IDs from OLE specification
        switch (prop.id) {
          case 2: metadata.title = str; break;
          case 3: metadata.subject = str; break;
          case 4: metadata.creator = str; break;
          case 5: metadata.keywords = str; break;
          case 6: metadata.description = str; break;
          case 8: metadata.lastModifiedBy = str; break;
          case 9: metadata.revision = str; break;
          case 18: metadata.application = str; break;
        }
      }
    }
  } catch {
    // Ignore parsing errors
  }
  
  return metadata;
}

/**
 * Parse Document Summary Information for additional metadata
 */
function parseDocumentSummaryInformation(data: Uint8Array): Partial<PresentationMetadata> {
  const metadata: Partial<PresentationMetadata> = {};
  
  try {
    const reader = new BinaryReader(data);
    
    // Skip header similar to Summary Information
    reader.skip(24);
    
    const numPropertySets = reader.readUInt32LE();
    if (numPropertySets === 0) return metadata;
    
    reader.skip(20);
    
    reader.readUInt32LE(); // size
    const numProperties = reader.readUInt32LE();
    
    const properties: Array<{ id: number; offset: number }> = [];
    for (let i = 0; i < numProperties && i < 100; i++) {
      const id = reader.readUInt32LE();
      const offset = reader.readUInt32LE();
      properties.push({ id, offset });
    }
    
    const baseOffset = reader.position - (numProperties * 8) - 8;
    
    for (const prop of properties) {
      reader.seek(baseOffset + prop.offset);
      const type = reader.readUInt32LE();
      
      if (type === 30 || type === 31) {
        const strLen = reader.readUInt32LE();
        const str = type === 31 
          ? reader.readUnicodeString(strLen * 2)
          : reader.readAsciiString(strLen);
        
        switch (prop.id) {
          case 2: metadata.category = str; break;
          case 14: metadata.manager = str; break;
          case 15: metadata.company = str; break;
        }
      } else if (type === 3) { // VT_I4
        const value = reader.readInt32LE();
        switch (prop.id) {
          case 4: metadata.totalSlides = value; break;
          case 6: metadata.totalParagraphs = value; break;
          case 7: metadata.totalWords = value; break;
        }
      }
    }
  } catch {
    // Ignore parsing errors
  }
  
  return metadata;
}

// ============================================================================
// SLIDE CREATION FROM TEXT
// ============================================================================

/**
 * Create slide objects from extracted texts
 */
function createSlidesFromTexts(texts: string[]): SlideContent[] {
  if (texts.length === 0) {
    return [{
      slideNumber: 1,
      title: 'No Content Found',
      textContent: ['No text content could be extracted from this file.'],
      notes: '',
      shapes: [],
      images: [],
      tables: [],
    }];
  }

  // Filter out garbage and deduplicate
  const cleanTexts = [...new Set(texts)]
    .filter(t => t.length >= 2 && !isGarbageText(t))
    .sort((a, b) => {
      // Sort by length (shorter texts like titles first)
      if (a.length < 50 && b.length >= 50) return -1;
      if (b.length < 50 && a.length >= 50) return 1;
      return 0;
    });

  if (cleanTexts.length === 0) {
    return [{
      slideNumber: 1,
      title: 'Extracted Content',
      textContent: texts.slice(0, 10),
      notes: '',
      shapes: [],
      images: [],
      tables: [],
    }];
  }

  // Group texts into slides
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

  for (const text of cleanTexts) {
    // Heuristic: Short text at start of slide is likely a title
    if (!currentSlide.title && text.length < 100 && text.length > 1) {
      currentSlide.title = text;
    } else {
      currentSlide.textContent.push(text);
    }

    // Start new slide after accumulating enough content
    if (currentSlide.textContent.length >= 8 || 
        (currentSlide.textContent.length >= 3 && text.length < 30 && currentSlide.textContent.length > 0)) {
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

  // Ensure slides have numbers and titles
  slides.forEach((slide, index) => {
    slide.slideNumber = index + 1;
    if (!slide.title) {
      slide.title = `Slide ${index + 1}`;
    }
  });

  return slides.length > 0 ? slides : [{
    slideNumber: 1,
    title: 'Extracted Content',
    textContent: cleanTexts,
    notes: '',
    shapes: [],
    images: [],
    tables: [],
  }];
}

// ============================================================================
// MAIN EXPORT FUNCTION
// ============================================================================

/**
 * Parse a PPT file and extract all data
 */
export async function parsePPT(file: File): Promise<ExtractedPresentation> {
  const buffer = await file.arrayBuffer();
  const data = new Uint8Array(buffer);
  
  try {
    // Parse as OLE Compound Document
    const cfb = CFB.read(data, { type: 'array' });
    
    // Initialize metadata
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
    const summaryInfo = cfb.find('\x05SummaryInformation');
    if (summaryInfo && summaryInfo.content) {
      const summaryMeta = parseSummaryInformation(new Uint8Array(summaryInfo.content));
      metadata = { ...metadata, ...summaryMeta };
    }

    // Parse Document Summary Information
    const docSummary = cfb.find('\x05DocumentSummaryInformation');
    if (docSummary && docSummary.content) {
      const docMeta = parseDocumentSummaryInformation(new Uint8Array(docSummary.content));
      metadata = { ...metadata, ...docMeta };
    }

    // Get all texts and media
    let allTexts: string[] = [];
    let allMedia: MediaInfo[] = [];

    // Parse the main PowerPoint Document stream
    const pptDoc = cfb.find('PowerPoint Document');
    if (pptDoc && pptDoc.content) {
      const pptData = new Uint8Array(pptDoc.content);
      
      // Use the PPT Parser
      const parser = new PPTParser(pptData);
      const context = parser.parse();
      
      // Collect texts from context
      allTexts.push(...context.documentTexts);
      allMedia.push(...context.pictures);
      
      // Also use scanning as backup
      const scannedTexts = extractTextByScanning(pptData);
      allTexts.push(...scannedTexts);
      
      // Update metadata from context
      if (context.metadata.presentationFormat) {
        metadata.presentationFormat = context.metadata.presentationFormat;
      }
    }

    // Parse Pictures stream if available
    const picturesStream = cfb.find('Pictures');
    if (picturesStream && picturesStream.content) {
      const picturesData = new Uint8Array(picturesStream.content);
      const picturesParser = new PPTParser(picturesData);
      const picturesContext = picturesParser.parse();
      allMedia.push(...picturesContext.pictures);
    }

    // Deduplicate texts
    const uniqueTexts = [...new Set(allTexts)];
    
    // Create slides
    const slides = createSlidesFromTexts(uniqueTexts);
    
    // Update total slides in metadata
    metadata.totalSlides = slides.length;
    
    // Count words
    let wordCount = 0;
    for (const slide of slides) {
      wordCount += slide.title.split(/\s+/).filter(w => w.length > 0).length;
      for (const text of slide.textContent) {
        wordCount += text.split(/\s+/).filter(w => w.length > 0).length;
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
      media: allMedia,
      themes: [],
      masterSlides: [],
      customProperties: {
        parsedWith: 'full-ppt-parser',
        cfbVersion: 'cfb',
      },
    };

  } catch (error) {
    console.error('PPT parsing error:', error);
    
    // Fallback: Try basic text scanning
    const fallbackTexts = extractTextByScanning(data);
    const slides = createSlidesFromTexts(fallbackTexts);
    
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
        description: 'Legacy PowerPoint format',
        application: 'Microsoft PowerPoint (Legacy)',
        appVersion: '',
        company: '',
        manager: '',
        totalSlides: slides.length,
        totalWords: 0,
        totalParagraphs: 0,
        presentationFormat: '',
        template: '',
      },
      slides,
      media: [],
      themes: [],
      masterSlides: [],
      customProperties: {
        parsedWith: 'fallback-scanner',
        error: error instanceof Error ? error.message : 'Unknown error',
      },
    };
  }
}
