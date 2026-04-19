import type { AppSettings } from '@/types'

export const APP_NAME = 'PowerPoint Extractor'
export const APP_VERSION = '2.0.0'

export const STORAGE_KEYS = {
  settings: 'pptext_settings',
} as const

export const DEFAULT_SETTINGS: AppSettings = {
  theme: 'system',
  language: 'en',
  defaultExportFormats: ['json'],
  includeNotes: true,
  includeMedia: true,
  includeMetadata: true,
  includeThemes: true,
  includeShapes: true,
  includeTables: true,
  includeCustomProperties: true,
  autoExpandSlides: false,
  compactView: false,
  maxInitialSlides: 20,
  confirmBeforeDelete: true,
  confirmBeforeClear: true,
  showFileSize: true,
  showSlideCount: true,
  pdfPageSize: 'a4',
  pdfOrientation: 'portrait',
  csvDelimiter: ',',
  jsonIndent: 2,
  htmlIncludeStyles: true,
  sortSlidesBy: 'number',
  mediaPreviewSize: 'medium',
  exportFilenamePattern: 'both',
}

export const ACCEPTED_EXTENSIONS = ['.ppt', '.pptx'] as const
export const ACCEPTED_MIME_TYPES = [
  'application/vnd.ms-powerpoint',
  'application/vnd.openxmlformats-officedocument.presentationml.presentation',
] as const
