/**
 * Translations System
 * Multi-language support for PowerPoint Extractor
 */

export type Language = 'en' | 'es' | 'de' | 'fr';

export interface Translations {
  // App
  appName: string;
  appDescription: string;
  
  // Actions
  upload: string;
  extract: string;
  export: string;
  clear: string;
  clearAll: string;
  remove: string;
  view: string;
  download: string;
  close: string;
  cancel: string;
  confirm: string;
  
  // File handling
  dropFiles: string;
  selectFiles: string;
  supportedFormats: string;
  processing: string;
  processed: string;
  noFiles: string;
  fileSize: string;
  
  // Data extraction
  extractData: string;
  extractedData: string;
  slides: string;
  metadata: string;
  media: string;
  themes: string;
  text: string;
  tables: string;
  shapes: string;
  notes: string;
  images: string;
  
  // Export
  exportAs: string;
  exportAll: string;
  exportSelected: string;
  exportFormat: string;
  
  // Settings
  settings: string;
  theme: string;
  language: string;
  appearance: string;
  light: string;
  dark: string;
  oled: string;
  neumorphic: string;
  storage: string;
  clearData: string;
  clearDataDescription: string;
  clearDataConfirm: string;
  storageUsed: string;
  version: string;
  
  // Messages
  errorProcessing: string;
  noDataExtracted: string;
  exportSuccess: string;
  
  // Footer
  madeWith: string;
  by: string;
  supportsPPT: string;
}

// English translations
const en: Translations = {
  // App
  appName: 'PowerPoint Extractor',
  appDescription: 'Extract data from PPT & PPTX files',
  
  // Actions
  upload: 'Upload',
  extract: 'Extract',
  export: 'Export',
  clear: 'Clear',
  clearAll: 'Clear All',
  remove: 'Remove',
  view: 'View',
  download: 'Download',
  close: 'Close',
  cancel: 'Cancel',
  confirm: 'Confirm',
  
  // File handling
  dropFiles: 'Drop PowerPoint files here',
  selectFiles: 'Select Files',
  supportedFormats: 'Supports PPT & PPTX',
  processing: 'Processing...',
  processed: 'Processed',
  noFiles: 'No files uploaded',
  fileSize: 'Size',
  
  // Data extraction
  extractData: 'Extract Data',
  extractedData: 'Extracted Data',
  slides: 'Slides',
  metadata: 'Metadata',
  media: 'Media',
  themes: 'Themes',
  text: 'Text',
  tables: 'Tables',
  shapes: 'Shapes',
  notes: 'Notes',
  images: 'Images',
  
  // Export
  exportAs: 'Export as',
  exportAll: 'Export All',
  exportSelected: 'Export Selected',
  exportFormat: 'Export Format',
  
  // Settings
  settings: 'Settings',
  theme: 'Theme',
  language: 'Language',
  appearance: 'Appearance',
  light: 'Light',
  dark: 'Dark',
  oled: 'OLED',
  neumorphic: 'Neumorphic',
  storage: 'Storage',
  clearData: 'Clear Data',
  clearDataDescription: 'Remove all cached data and preferences',
  clearDataConfirm: 'Are you sure you want to clear all data? This action cannot be undone.',
  storageUsed: 'Storage used',
  version: 'Version',
  
  // Messages
  errorProcessing: 'Error processing file',
  noDataExtracted: 'No data could be extracted',
  exportSuccess: 'Export successful',
  
  // Footer
  madeWith: 'Made with',
  by: 'by',
  supportsPPT: 'Supports PPT (97-2003) and PPTX (2007+) formats',
};

// Spanish translations
const es: Translations = {
  // App
  appName: 'PowerPoint Extractor',
  appDescription: 'Extrae datos de archivos PPT y PPTX',
  
  // Actions
  upload: 'Subir',
  extract: 'Extraer',
  export: 'Exportar',
  clear: 'Limpiar',
  clearAll: 'Limpiar Todo',
  remove: 'Eliminar',
  view: 'Ver',
  download: 'Descargar',
  close: 'Cerrar',
  cancel: 'Cancelar',
  confirm: 'Confirmar',
  
  // File handling
  dropFiles: 'Suelta archivos PowerPoint aquí',
  selectFiles: 'Seleccionar Archivos',
  supportedFormats: 'Soporta PPT y PPTX',
  processing: 'Procesando...',
  processed: 'Procesado',
  noFiles: 'No hay archivos subidos',
  fileSize: 'Tamaño',
  
  // Data extraction
  extractData: 'Extraer Datos',
  extractedData: 'Datos Extraídos',
  slides: 'Diapositivas',
  metadata: 'Metadatos',
  media: 'Medios',
  themes: 'Temas',
  text: 'Texto',
  tables: 'Tablas',
  shapes: 'Formas',
  notes: 'Notas',
  images: 'Imágenes',
  
  // Export
  exportAs: 'Exportar como',
  exportAll: 'Exportar Todo',
  exportSelected: 'Exportar Seleccionado',
  exportFormat: 'Formato de Exportación',
  
  // Settings
  settings: 'Configuración',
  theme: 'Tema',
  language: 'Idioma',
  appearance: 'Apariencia',
  light: 'Claro',
  dark: 'Oscuro',
  oled: 'OLED',
  neumorphic: 'Neumórfico',
  storage: 'Almacenamiento',
  clearData: 'Borrar Datos',
  clearDataDescription: 'Eliminar todos los datos en caché y preferencias',
  clearDataConfirm: '¿Estás seguro de que quieres borrar todos los datos? Esta acción no se puede deshacer.',
  storageUsed: 'Almacenamiento usado',
  version: 'Versión',
  
  // Messages
  errorProcessing: 'Error al procesar el archivo',
  noDataExtracted: 'No se pudieron extraer datos',
  exportSuccess: 'Exportación exitosa',
  
  // Footer
  madeWith: 'Hecho con',
  by: 'por',
  supportsPPT: 'Soporta formatos PPT (97-2003) y PPTX (2007+)',
};

// German translations
const de: Translations = {
  // App
  appName: 'PowerPoint Extractor',
  appDescription: 'Daten aus PPT & PPTX-Dateien extrahieren',
  
  // Actions
  upload: 'Hochladen',
  extract: 'Extrahieren',
  export: 'Exportieren',
  clear: 'Löschen',
  clearAll: 'Alles Löschen',
  remove: 'Entfernen',
  view: 'Ansehen',
  download: 'Herunterladen',
  close: 'Schließen',
  cancel: 'Abbrechen',
  confirm: 'Bestätigen',
  
  // File handling
  dropFiles: 'PowerPoint-Dateien hier ablegen',
  selectFiles: 'Dateien Auswählen',
  supportedFormats: 'Unterstützt PPT & PPTX',
  processing: 'Verarbeitung...',
  processed: 'Verarbeitet',
  noFiles: 'Keine Dateien hochgeladen',
  fileSize: 'Größe',
  
  // Data extraction
  extractData: 'Daten Extrahieren',
  extractedData: 'Extrahierte Daten',
  slides: 'Folien',
  metadata: 'Metadaten',
  media: 'Medien',
  themes: 'Themen',
  text: 'Text',
  tables: 'Tabellen',
  shapes: 'Formen',
  notes: 'Notizen',
  images: 'Bilder',
  
  // Export
  exportAs: 'Exportieren als',
  exportAll: 'Alles Exportieren',
  exportSelected: 'Ausgewählte Exportieren',
  exportFormat: 'Exportformat',
  
  // Settings
  settings: 'Einstellungen',
  theme: 'Thema',
  language: 'Sprache',
  appearance: 'Erscheinungsbild',
  light: 'Hell',
  dark: 'Dunkel',
  oled: 'OLED',
  neumorphic: 'Neumorphisch',
  storage: 'Speicher',
  clearData: 'Daten Löschen',
  clearDataDescription: 'Alle zwischengespeicherten Daten und Einstellungen entfernen',
  clearDataConfirm: 'Sind Sie sicher, dass Sie alle Daten löschen möchten? Diese Aktion kann nicht rückgängig gemacht werden.',
  storageUsed: 'Speicher verwendet',
  version: 'Version',
  
  // Messages
  errorProcessing: 'Fehler bei der Verarbeitung der Datei',
  noDataExtracted: 'Es konnten keine Daten extrahiert werden',
  exportSuccess: 'Export erfolgreich',
  
  // Footer
  madeWith: 'Gemacht mit',
  by: 'von',
  supportsPPT: 'Unterstützt PPT (97-2003) und PPTX (2007+) Formate',
};

// French translations
const fr: Translations = {
  // App
  appName: 'PowerPoint Extractor',
  appDescription: 'Extraire les données des fichiers PPT et PPTX',
  
  // Actions
  upload: 'Télécharger',
  extract: 'Extraire',
  export: 'Exporter',
  clear: 'Effacer',
  clearAll: 'Tout Effacer',
  remove: 'Supprimer',
  view: 'Voir',
  download: 'Télécharger',
  close: 'Fermer',
  cancel: 'Annuler',
  confirm: 'Confirmer',
  
  // File handling
  dropFiles: 'Déposez les fichiers PowerPoint ici',
  selectFiles: 'Sélectionner des Fichiers',
  supportedFormats: 'Supporte PPT et PPTX',
  processing: 'Traitement...',
  processed: 'Traité',
  noFiles: 'Aucun fichier téléchargé',
  fileSize: 'Taille',
  
  // Data extraction
  extractData: 'Extraire les Données',
  extractedData: 'Données Extraites',
  slides: 'Diapositives',
  metadata: 'Métadonnées',
  media: 'Médias',
  themes: 'Thèmes',
  text: 'Texte',
  tables: 'Tableaux',
  shapes: 'Formes',
  notes: 'Notes',
  images: 'Images',
  
  // Export
  exportAs: 'Exporter en',
  exportAll: 'Tout Exporter',
  exportSelected: 'Exporter la Sélection',
  exportFormat: "Format d'Exportation",
  
  // Settings
  settings: 'Paramètres',
  theme: 'Thème',
  language: 'Langue',
  appearance: 'Apparence',
  light: 'Clair',
  dark: 'Sombre',
  oled: 'OLED',
  neumorphic: 'Neumorphique',
  storage: 'Stockage',
  clearData: 'Effacer les Données',
  clearDataDescription: 'Supprimer toutes les données en cache et les préférences',
  clearDataConfirm: 'Êtes-vous sûr de vouloir effacer toutes les données ? Cette action ne peut pas être annulée.',
  storageUsed: 'Stockage utilisé',
  version: 'Version',
  
  // Messages
  errorProcessing: 'Erreur lors du traitement du fichier',
  noDataExtracted: "Aucune donnée n'a pu être extraite",
  exportSuccess: 'Exportation réussie',
  
  // Footer
  madeWith: 'Fait avec',
  by: 'par',
  supportsPPT: 'Supporte les formats PPT (97-2003) et PPTX (2007+)',
};

export const translations: Record<Language, Translations> = {
  en,
  es,
  de,
  fr,
};

export const languageNames: Record<Language, string> = {
  en: 'English',
  es: 'Español',
  de: 'Deutsch',
  fr: 'Français',
};

export function getBrowserLanguage(): Language {
  const browserLang = navigator.language.split('-')[0];
  if (browserLang in translations) {
    return browserLang as Language;
  }
  return 'en';
}
