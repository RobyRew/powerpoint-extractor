export interface Translations {
  // App
  appName: string
  appDescription: string

  // Navigation
  extractor: string
  viewer: string
  export_: string
  settings: string

  // Actions
  save: string
  cancel: string
  delete: string
  close: string
  clear: string
  confirm: string
  done: string
  apply: string
  reset: string
  search: string
  expandAll: string
  collapseAll: string

  // File operations
  dropFiles: string
  dropFilesDescription: string
  browseFiles: string
  supportedFormats: string
  processing: string
  extractData: string
  extractedData: string
  clearAll: string
  clearAllConfirm: string
  removeFile: string
  viewData: string
  noFiles: string
  fileAdded: string
  filesCleared: string

  // Export
  exportData: string
  exportFormat: string
  exportSuccess: string
  exportError: string
  exportAll: string
  exportSelected: string
  exportMedia: string
  downloadAllZip: string
  selectFormats: string
  noDataToExport: string
  presentations: string

  // Data viewer
  slides: string
  metadata: string
  themes: string
  media: string
  content: string
  speakerNotes: string
  tables: string
  shapes: string
  fonts: string
  colors: string
  masterSlides: string
  customProperties: string
  untitledSlide: string
  textBlocks: string
  hasNotes: string
  noThemes: string
  noMedia: string

  // Metadata fields
  title: string
  subject: string
  creator: string
  lastModifiedBy: string
  created: string
  modified: string
  category: string
  keywords: string
  description: string
  application: string
  appVersion: string
  company: string
  totalSlides: string
  totalWords: string
  format: string

  // Settings
  appearance: string
  theme: string
  systemTheme: string
  lightTheme: string
  darkTheme: string
  oledTheme: string
  language: string
  english: string
  spanish: string
  german: string
  french: string
  compactView: string
  compactViewDesc: string
  autoExpandSlides: string
  autoExpandSlidesDesc: string
  showFileSize: string
  showSlideCount: string
  maxInitialSlides: string
  confirmBeforeDelete: string
  confirmBeforeClear: string
  mediaPreviewSize: string
  previewSmall: string
  previewMedium: string
  previewLarge: string
  sortSlidesBy: string
  sortByNumber: string
  sortByTitle: string

  // Export settings
  exportSettings: string
  defaultExportFormats: string
  includeNotes: string
  includeMedia: string
  includeMetadata: string
  includeThemes: string
  includeShapes: string
  includeTables: string
  includeCustomProperties: string
  pdfPageSize: string
  pdfOrientation: string
  portrait: string
  landscape: string
  csvDelimiter: string
  jsonIndent: string
  htmlIncludeStyles: string
  htmlIncludeStylesDesc: string
  exportFilenamePattern: string
  filenameOriginal: string
  filenameTimestamp: string
  filenameBoth: string

  // Data management
  dataManagement: string
  storageUsed: string
  clearAppData: string
  clearAppDataConfirm: string
  resetSettings: string
  resetSettingsConfirm: string
  about: string
  version: string
  privacyNote: string

  // Features
  featureExtract: string
  featureExtractDesc: string
  featureExport: string
  featureExportDesc: string
  featureMedia: string
  featureMediaDesc: string

  // Command palette
  commandPalette: string
  navigation: string
  actions: string
  noResults: string

  // Errors
  invalidFileType: string
  errorProcessing: string

  // Behavior
  behavior: string
  extraction: string

  // images (kept for parser compat)
  images: string
  notes: string
  text: string
}
