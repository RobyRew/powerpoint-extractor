/**
 * PowerPoint Data Types
 */

export interface SlideContent {
  slideNumber: number;
  title: string;
  textContent: string[];
  notes: string;
  shapes: ShapeInfo[];
  images: MediaInfo[];
  tables: TableInfo[];
}

export interface ShapeInfo {
  type: string;
  text: string;
  position?: { x: number; y: number };
  size?: { width: number; height: number };
}

export interface MediaInfo {
  name: string;
  type: string;
  size: number;
  data?: string; // Base64 encoded
  extension: string;
}

export interface TableInfo {
  rows: number;
  columns: number;
  cells: string[][];
}

export interface PresentationMetadata {
  title: string;
  subject: string;
  creator: string;
  lastModifiedBy: string;
  created: string;
  modified: string;
  revision: string;
  category: string;
  keywords: string;
  description: string;
  application: string;
  appVersion: string;
  company: string;
  manager: string;
  totalSlides: number;
  totalWords: number;
  totalParagraphs: number;
  presentationFormat: string;
  template: string;
}

export interface ExtractedPresentation {
  id: string;
  fileName: string;
  fileSize: number;
  fileType: 'pptx' | 'ppt';
  extractedAt: string;
  metadata: PresentationMetadata;
  slides: SlideContent[];
  media: MediaInfo[];
  themes: ThemeInfo[];
  masterSlides: string[];
  customProperties: Record<string, string>;
}

export interface ThemeInfo {
  name: string;
  colors: string[];
  fonts: string[];
}

export interface ExportFormat {
  id: string;
  name: string;
  extension: string;
  icon: string;
  description: string;
}

export const EXPORT_FORMATS: ExportFormat[] = [
  { id: 'json', name: 'JSON', extension: '.json', icon: 'FileJson', description: 'Full data in JSON format' },
  { id: 'xml', name: 'XML', extension: '.xml', icon: 'FileCode', description: 'Structured XML format' },
  { id: 'csv', name: 'CSV', extension: '.csv', icon: 'FileSpreadsheet', description: 'Spreadsheet compatible' },
  { id: 'txt', name: 'Text', extension: '.txt', icon: 'FileText', description: 'Plain text content' },
  { id: 'html', name: 'HTML', extension: '.html', icon: 'Globe', description: 'Web page format' },
  { id: 'pdf', name: 'PDF', extension: '.pdf', icon: 'FileText', description: 'Document format' },
];
