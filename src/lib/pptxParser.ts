/**
 * PPTX Parser - Extract data from PowerPoint files
 * PPTX files are ZIP archives containing XML files
 */

import JSZip from 'jszip';
import type { 
  ExtractedPresentation, 
  SlideContent, 
  PresentationMetadata, 
  MediaInfo, 
  ShapeInfo,
  TableInfo,
  ThemeInfo 
} from '../types';

/**
 * Parse a PPTX file and extract all data
 */
export async function parsePPTX(file: File): Promise<ExtractedPresentation> {
  const zip = await JSZip.loadAsync(file);
  
  const metadata = await extractMetadata(zip);
  const slides = await extractSlides(zip);
  const media = await extractMedia(zip);
  const themes = await extractThemes(zip);
  const masterSlides = await extractMasterSlides(zip);
  const customProperties = await extractCustomProperties(zip);
  
  return {
    id: crypto.randomUUID(),
    fileName: file.name,
    fileSize: file.size,
    fileType: 'pptx',
    extractedAt: new Date().toISOString(),
    metadata: {
      ...metadata,
      totalSlides: slides.length,
    },
    slides,
    media,
    themes,
    masterSlides,
    customProperties,
  };
}

/**
 * Extract metadata from docProps/core.xml and docProps/app.xml
 */
async function extractMetadata(zip: JSZip): Promise<PresentationMetadata> {
  const metadata: PresentationMetadata = {
    title: '',
    subject: '',
    creator: '',
    lastModifiedBy: '',
    created: '',
    modified: '',
    revision: '',
    category: '',
    keywords: '',
    description: '',
    application: '',
    appVersion: '',
    company: '',
    manager: '',
    totalSlides: 0,
    totalWords: 0,
    totalParagraphs: 0,
    presentationFormat: '',
    template: '',
  };

  // Parse core.xml (Dublin Core metadata)
  const coreXml = await zip.file('docProps/core.xml')?.async('text');
  if (coreXml) {
    const parser = new DOMParser();
    const doc = parser.parseFromString(coreXml, 'text/xml');
    
    metadata.title = getTextContent(doc, 'dc:title') || getTextContent(doc, 'title') || '';
    metadata.subject = getTextContent(doc, 'dc:subject') || getTextContent(doc, 'subject') || '';
    metadata.creator = getTextContent(doc, 'dc:creator') || getTextContent(doc, 'creator') || '';
    metadata.lastModifiedBy = getTextContent(doc, 'cp:lastModifiedBy') || getTextContent(doc, 'lastModifiedBy') || '';
    metadata.created = getTextContent(doc, 'dcterms:created') || getTextContent(doc, 'created') || '';
    metadata.modified = getTextContent(doc, 'dcterms:modified') || getTextContent(doc, 'modified') || '';
    metadata.revision = getTextContent(doc, 'cp:revision') || getTextContent(doc, 'revision') || '';
    metadata.category = getTextContent(doc, 'cp:category') || getTextContent(doc, 'category') || '';
    metadata.keywords = getTextContent(doc, 'cp:keywords') || getTextContent(doc, 'keywords') || '';
    metadata.description = getTextContent(doc, 'dc:description') || getTextContent(doc, 'description') || '';
  }

  // Parse app.xml (Application metadata)
  const appXml = await zip.file('docProps/app.xml')?.async('text');
  if (appXml) {
    const parser = new DOMParser();
    const doc = parser.parseFromString(appXml, 'text/xml');
    
    metadata.application = getTextContent(doc, 'Application') || '';
    metadata.appVersion = getTextContent(doc, 'AppVersion') || '';
    metadata.company = getTextContent(doc, 'Company') || '';
    metadata.manager = getTextContent(doc, 'Manager') || '';
    metadata.totalSlides = parseInt(getTextContent(doc, 'Slides') || '0', 10);
    metadata.totalWords = parseInt(getTextContent(doc, 'Words') || '0', 10);
    metadata.totalParagraphs = parseInt(getTextContent(doc, 'Paragraphs') || '0', 10);
    metadata.presentationFormat = getTextContent(doc, 'PresentationFormat') || '';
    metadata.template = getTextContent(doc, 'Template') || '';
  }

  return metadata;
}

/**
 * Extract slide content from ppt/slides/
 */
async function extractSlides(zip: JSZip): Promise<SlideContent[]> {
  const slides: SlideContent[] = [];
  const slideFiles: { path: string; num: number }[] = [];

  // Find all slide files
  zip.forEach((path) => {
    const match = path.match(/ppt\/slides\/slide(\d+)\.xml$/);
    if (match) {
      slideFiles.push({ path, num: parseInt(match[1], 10) });
    }
  });

  // Sort by slide number
  slideFiles.sort((a, b) => a.num - b.num);

  // Parse each slide
  for (const slideFile of slideFiles) {
    const slideXml = await zip.file(slideFile.path)?.async('text');
    if (slideXml) {
      const slide = parseSlideXml(slideXml, slideFile.num);
      
      // Try to get notes for this slide
      const notesPath = `ppt/notesSlides/notesSlide${slideFile.num}.xml`;
      const notesXml = await zip.file(notesPath)?.async('text');
      if (notesXml) {
        slide.notes = parseNotesXml(notesXml);
      }
      
      slides.push(slide);
    }
  }

  return slides;
}

/**
 * Parse individual slide XML
 */
function parseSlideXml(xml: string, slideNumber: number): SlideContent {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xml, 'text/xml');
  
  const textContent: string[] = [];
  const shapes: ShapeInfo[] = [];
  const images: MediaInfo[] = [];
  const tables: TableInfo[] = [];
  let title = '';

  // Extract all text from a:t elements
  const textElements = doc.getElementsByTagName('a:t');
  for (let i = 0; i < textElements.length; i++) {
    const text = textElements[i].textContent?.trim();
    if (text) {
      textContent.push(text);
    }
  }

  // Try to identify title (usually in p:ph with type="title" or "ctrTitle")
  const phElements = doc.getElementsByTagName('p:ph');
  for (let i = 0; i < phElements.length; i++) {
    const type = phElements[i].getAttribute('type');
    if (type === 'title' || type === 'ctrTitle') {
      const parent = findParentWithTag(phElements[i], 'p:sp');
      if (parent) {
        const titleTexts = parent.getElementsByTagName('a:t');
        const titleParts: string[] = [];
        for (let j = 0; j < titleTexts.length; j++) {
          const t = titleTexts[j].textContent?.trim();
          if (t) titleParts.push(t);
        }
        title = titleParts.join(' ');
      }
    }
  }

  // If no title found, use first text content
  if (!title && textContent.length > 0) {
    title = textContent[0];
  }

  // Extract shapes (p:sp elements)
  const spElements = doc.getElementsByTagName('p:sp');
  for (let i = 0; i < spElements.length; i++) {
    const sp = spElements[i];
    const shapeText: string[] = [];
    const textEls = sp.getElementsByTagName('a:t');
    for (let j = 0; j < textEls.length; j++) {
      const t = textEls[j].textContent?.trim();
      if (t) shapeText.push(t);
    }
    
    // Get shape type from nvSpPr if available
    const nvSpPr = sp.getElementsByTagName('p:nvSpPr')[0];
    let shapeType = 'Shape';
    if (nvSpPr) {
      const nvPr = nvSpPr.getElementsByTagName('p:nvPr')[0];
      if (nvPr) {
        const ph = nvPr.getElementsByTagName('p:ph')[0];
        if (ph) {
          shapeType = ph.getAttribute('type') || 'Shape';
        }
      }
    }

    shapes.push({
      type: shapeType,
      text: shapeText.join(' '),
    });
  }

  // Extract tables (a:tbl elements)
  const tblElements = doc.getElementsByTagName('a:tbl');
  for (let i = 0; i < tblElements.length; i++) {
    const tbl = tblElements[i];
    const rows = tbl.getElementsByTagName('a:tr');
    const tableData: string[][] = [];
    
    for (let r = 0; r < rows.length; r++) {
      const cells = rows[r].getElementsByTagName('a:tc');
      const rowData: string[] = [];
      for (let c = 0; c < cells.length; c++) {
        const cellTexts = cells[c].getElementsByTagName('a:t');
        const cellContent: string[] = [];
        for (let t = 0; t < cellTexts.length; t++) {
          const text = cellTexts[t].textContent?.trim();
          if (text) cellContent.push(text);
        }
        rowData.push(cellContent.join(' '));
      }
      tableData.push(rowData);
    }

    if (tableData.length > 0) {
      tables.push({
        rows: tableData.length,
        columns: tableData[0]?.length || 0,
        cells: tableData,
      });
    }
  }

  // Extract image references (p:pic elements)
  const picElements = doc.getElementsByTagName('p:pic');
  for (let i = 0; i < picElements.length; i++) {
    const blipElements = picElements[i].getElementsByTagName('a:blip');
    for (let j = 0; j < blipElements.length; j++) {
      const embed = blipElements[j].getAttribute('r:embed');
      if (embed) {
        images.push({
          name: `Image reference: ${embed}`,
          type: 'image',
          size: 0,
          extension: '',
        });
      }
    }
  }

  return {
    slideNumber,
    title,
    textContent,
    notes: '',
    shapes,
    images,
    tables,
  };
}

/**
 * Parse notes XML
 */
function parseNotesXml(xml: string): string {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xml, 'text/xml');
  
  const notes: string[] = [];
  const textElements = doc.getElementsByTagName('a:t');
  for (let i = 0; i < textElements.length; i++) {
    const text = textElements[i].textContent?.trim();
    if (text) {
      notes.push(text);
    }
  }
  
  // Filter out slide number placeholders
  return notes.filter(n => !/^\d+$/.test(n)).join('\n');
}

/**
 * Extract media files from ppt/media/
 */
async function extractMedia(zip: JSZip): Promise<MediaInfo[]> {
  const media: MediaInfo[] = [];
  
  const mediaFolder = zip.folder('ppt/media');
  if (mediaFolder) {
    const files: { name: string; file: JSZip.JSZipObject }[] = [];
    
    mediaFolder.forEach((relativePath, file) => {
      if (!file.dir) {
        files.push({ name: relativePath, file });
      }
    });

    for (const { name, file } of files) {
      const extension = name.split('.').pop()?.toLowerCase() || '';
      const data = await file.async('base64');
      const size = (await file.async('uint8array')).length;
      
      let type = 'unknown';
      if (['jpg', 'jpeg', 'png', 'gif', 'bmp', 'tiff', 'webp'].includes(extension)) {
        type = 'image';
      } else if (['mp4', 'avi', 'mov', 'wmv', 'webm'].includes(extension)) {
        type = 'video';
      } else if (['mp3', 'wav', 'ogg', 'wma', 'm4a'].includes(extension)) {
        type = 'audio';
      }
      
      media.push({
        name,
        type,
        size,
        data,
        extension,
      });
    }
  }
  
  return media;
}

/**
 * Extract theme information
 */
async function extractThemes(zip: JSZip): Promise<ThemeInfo[]> {
  const themes: ThemeInfo[] = [];
  const themeFiles: string[] = [];
  
  zip.forEach((path) => {
    if (path.match(/ppt\/theme\/theme\d+\.xml$/)) {
      themeFiles.push(path);
    }
  });

  for (const themePath of themeFiles) {
    const themeXml = await zip.file(themePath)?.async('text');
    if (themeXml) {
      const parser = new DOMParser();
      const doc = parser.parseFromString(themeXml, 'text/xml');
      
      const name = doc.getElementsByTagName('a:theme')[0]?.getAttribute('name') || 'Theme';
      const colors: string[] = [];
      const fonts: string[] = [];
      
      // Extract color scheme
      const clrScheme = doc.getElementsByTagName('a:clrScheme')[0];
      if (clrScheme) {
        const colorElements = clrScheme.children;
        for (let i = 0; i < colorElements.length; i++) {
          const colorName = colorElements[i].tagName.replace('a:', '');
          const srgbClr = colorElements[i].getElementsByTagName('a:srgbClr')[0];
          const sysClr = colorElements[i].getElementsByTagName('a:sysClr')[0];
          
          if (srgbClr) {
            colors.push(`${colorName}: #${srgbClr.getAttribute('val')}`);
          } else if (sysClr) {
            colors.push(`${colorName}: ${sysClr.getAttribute('lastClr') || sysClr.getAttribute('val')}`);
          }
        }
      }
      
      // Extract font scheme
      const fontScheme = doc.getElementsByTagName('a:fontScheme')[0];
      if (fontScheme) {
        const majorFont = fontScheme.getElementsByTagName('a:majorFont')[0];
        const minorFont = fontScheme.getElementsByTagName('a:minorFont')[0];
        
        if (majorFont) {
          const latin = majorFont.getElementsByTagName('a:latin')[0];
          if (latin) fonts.push(`Major: ${latin.getAttribute('typeface')}`);
        }
        if (minorFont) {
          const latin = minorFont.getElementsByTagName('a:latin')[0];
          if (latin) fonts.push(`Minor: ${latin.getAttribute('typeface')}`);
        }
      }
      
      themes.push({ name, colors, fonts });
    }
  }
  
  return themes;
}

/**
 * Extract master slide names
 */
async function extractMasterSlides(zip: JSZip): Promise<string[]> {
  const masters: string[] = [];
  
  zip.forEach((path) => {
    const match = path.match(/ppt\/slideMasters\/slideMaster(\d+)\.xml$/);
    if (match) {
      masters.push(`Master Slide ${match[1]}`);
    }
  });
  
  return masters;
}

/**
 * Extract custom properties
 */
async function extractCustomProperties(zip: JSZip): Promise<Record<string, string>> {
  const props: Record<string, string> = {};
  
  const customXml = await zip.file('docProps/custom.xml')?.async('text');
  if (customXml) {
    const parser = new DOMParser();
    const doc = parser.parseFromString(customXml, 'text/xml');
    
    const properties = doc.getElementsByTagName('property');
    for (let i = 0; i < properties.length; i++) {
      const name = properties[i].getAttribute('name');
      const value = properties[i].textContent;
      if (name && value) {
        props[name] = value;
      }
    }
  }
  
  return props;
}

/**
 * Helper: Get text content from XML element by tag name
 */
function getTextContent(doc: Document, tagName: string): string {
  const elements = doc.getElementsByTagName(tagName);
  if (elements.length > 0) {
    return elements[0].textContent || '';
  }
  return '';
}

/**
 * Helper: Find parent element with specific tag
 */
function findParentWithTag(element: Element, tagName: string): Element | null {
  let current = element.parentElement;
  while (current) {
    if (current.tagName === tagName) {
      return current;
    }
    current = current.parentElement;
  }
  return null;
}
