/**
 * PowerPoint Extractor - Main Application
 * Extract data from PPT and PPTX files
 */

import { useState, useEffect, useCallback } from 'react';
import { Header, Footer, Settings, DropZone, FileList, DataViewer, ExportPanel } from './components';
import { parsePPTX, parsePPT } from './lib';
import type { ThemeId } from './styles/themes';
import { THEMES } from './styles/themes';
import type { ExtractedPresentation } from './types';
import { useI18n } from './context';

function AppContent() {
  const { t } = useI18n();
  const [theme, setTheme] = useState<ThemeId>('light');
  const [files, setFiles] = useState<File[]>([]);
  const [extractedData, setExtractedData] = useState<ExtractedPresentation[]>([]);
  const [processingFile, setProcessingFile] = useState<string | null>(null);
  const [viewingPresentation, setViewingPresentation] = useState<ExtractedPresentation | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [settingsOpen, setSettingsOpen] = useState(false);

  // Load theme from localStorage
  useEffect(() => {
    const savedTheme = localStorage.getItem('pptx-extractor-theme') as ThemeId | null;
    if (savedTheme && THEMES.find(t => t.id === savedTheme)) {
      setTheme(savedTheme);
    }
  }, []);

  // Apply theme class
  useEffect(() => {
    THEMES.forEach(t => {
      document.documentElement.classList.remove(t.className);
    });
    const themeConfig = THEMES.find(t => t.id === theme) || THEMES[0];
    document.documentElement.classList.add(themeConfig.className);
    localStorage.setItem('pptx-extractor-theme', theme);
  }, [theme]);

  // Process files
  const processFiles = useCallback(async (newFiles: File[]) => {
    for (const file of newFiles) {
      // Skip if already processed
      if (extractedData.find(p => p.fileName === file.name)) continue;
      
      setProcessingFile(file.name);
      setError(null);

      try {
        const isPPTX = file.name.toLowerCase().endsWith('.pptx');
        const data = isPPTX 
          ? await parsePPTX(file) 
          : await parsePPT(file);
        
        setExtractedData(prev => [...prev, data]);
      } catch (err) {
        console.error(`Error processing ${file.name}:`, err);
        setError(`${t.errorProcessing}: ${file.name}: ${err instanceof Error ? err.message : 'Unknown error'}`);
      }
    }
    setProcessingFile(null);
  }, [extractedData, t]);

  // Handle file selection
  const handleFilesSelected = useCallback((selectedFiles: File[]) => {
    const newFiles = selectedFiles.filter(
      f => !files.find(existing => existing.name === f.name)
    );
    
    if (newFiles.length > 0) {
      setFiles(prev => [...prev, ...newFiles]);
      processFiles(newFiles);
    }
  }, [files, processFiles]);

  // Handle file removal
  const handleRemoveFile = useCallback((index: number) => {
    const fileToRemove = files[index];
    setFiles(prev => prev.filter((_, i) => i !== index));
    setExtractedData(prev => prev.filter(p => p.fileName !== fileToRemove.name));
  }, [files]);

  // Handle clear all
  const handleClearAll = useCallback(() => {
    setFiles([]);
    setExtractedData([]);
    setError(null);
  }, []);

  return (
    <div className="min-h-screen flex flex-col bg-[rgb(var(--background))]">
      <Header onSettingsClick={() => setSettingsOpen(true)} />

      <main className="flex-1 container mx-auto px-4 py-8">
        <div className="max-w-4xl mx-auto space-y-6">
          {/* Hero Section */}
          <div className="text-center mb-8">
            <h2 className="text-2xl md:text-3xl font-bold text-[rgb(var(--foreground))] mb-2">
              {t.extractData}
            </h2>
            <p className="text-[rgb(var(--muted-foreground))]">
              {t.supportedFormats}
            </p>
          </div>

          {/* Drop Zone */}
          <DropZone 
            onFilesSelected={handleFilesSelected}
            isProcessing={!!processingFile}
          />

          {/* Error Message */}
          {error && (
            <div className="p-4 rounded-lg bg-[rgb(var(--destructive)/0.1)] border border-[rgb(var(--destructive)/0.3)]">
              <p className="text-sm text-[rgb(var(--destructive))]">{error}</p>
            </div>
          )}

          {/* File List */}
          <FileList
            files={files}
            extractedData={extractedData}
            processingFile={processingFile}
            onRemoveFile={handleRemoveFile}
            onViewData={setViewingPresentation}
          />

          {/* Export Panel */}
          <ExportPanel presentations={extractedData} />

          {/* Clear All Button */}
          {files.length > 0 && (
            <div className="flex justify-center">
              <button
                onClick={handleClearAll}
                className="btn btn-ghost text-sm text-[rgb(var(--muted-foreground))]"
              >
                {t.clearAll}
              </button>
            </div>
          )}

          {/* Features Info */}
          {files.length === 0 && (
            <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mt-8">
              <div className="card p-4 text-center">
                <div className="w-12 h-12 mx-auto mb-3 rounded-full bg-[rgb(var(--secondary))] flex items-center justify-center">
                  <span className="text-2xl">📊</span>
                </div>
                <h3 className="font-semibold mb-1">{t.extractedData}</h3>
                <p className="text-sm text-[rgb(var(--muted-foreground))]">
                  {t.text}, {t.metadata}, {t.themes}, {t.notes}, {t.tables}
                </p>
              </div>
              <div className="card p-4 text-center">
                <div className="w-12 h-12 mx-auto mb-3 rounded-full bg-[rgb(var(--secondary))] flex items-center justify-center">
                  <span className="text-2xl">📁</span>
                </div>
                <h3 className="font-semibold mb-1">{t.exportFormat}</h3>
                <p className="text-sm text-[rgb(var(--muted-foreground))]">
                  JSON, XML, CSV, TXT, HTML, PDF
                </p>
              </div>
              <div className="card p-4 text-center">
                <div className="w-12 h-12 mx-auto mb-3 rounded-full bg-[rgb(var(--secondary))] flex items-center justify-center">
                  <span className="text-2xl">🖼️</span>
                </div>
                <h3 className="font-semibold mb-1">{t.media}</h3>
                <p className="text-sm text-[rgb(var(--muted-foreground))]">
                  {t.images}
                </p>
              </div>
            </div>
          )}
        </div>
      </main>

      <Footer />

      {/* Settings Panel */}
      <Settings
        isOpen={settingsOpen}
        onClose={() => setSettingsOpen(false)}
        theme={theme}
        onThemeChange={setTheme}
      />

      {/* Data Viewer Modal */}
      {viewingPresentation && (
        <DataViewer
          presentation={viewingPresentation}
          onClose={() => setViewingPresentation(null)}
        />
      )}
    </div>
  );
}

// Main App with I18n Provider
import { I18nProvider } from './context';

function App() {
  return (
    <I18nProvider>
      <AppContent />
    </I18nProvider>
  );
}

export default App;
