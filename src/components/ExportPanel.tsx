/**
 * Export Panel Component
 */

import { useState } from 'react';
import { 
  Download, 
  FileJson, 
  FileCode, 
  FileSpreadsheet, 
  FileText, 
  Globe, 
  Image as ImageIcon,
  Package,
  Loader2,
  Check
} from 'lucide-react';
import type { ExtractedPresentation } from '../types';
import { 
  exportToJSON, 
  exportToXML, 
  exportToCSV, 
  exportToText, 
  exportToHTML,
  exportToPDF,
  downloadFile,
  downloadMediaAsZip,
  downloadAllAsZip
} from '../lib';

interface ExportPanelProps {
  presentations: ExtractedPresentation[];
}

const exportFormats = [
  { id: 'json', name: 'JSON', icon: FileJson, description: 'Full structured data' },
  { id: 'xml', name: 'XML', icon: FileCode, description: 'XML format' },
  { id: 'csv', name: 'CSV', icon: FileSpreadsheet, description: 'Spreadsheet format' },
  { id: 'txt', name: 'Text', icon: FileText, description: 'Plain text' },
  { id: 'html', name: 'HTML', icon: Globe, description: 'Web page' },
  { id: 'pdf', name: 'PDF', icon: FileText, description: 'Document' },
];

export function ExportPanel({ presentations }: ExportPanelProps) {
  const [selectedFormats, setSelectedFormats] = useState<Set<string>>(new Set(['json']));
  const [isExporting, setIsExporting] = useState(false);
  const [exportSuccess, setExportSuccess] = useState<string | null>(null);

  const hasMedia = presentations.some(p => p.media.length > 0);
  const totalMedia = presentations.reduce((acc, p) => acc + p.media.length, 0);

  const toggleFormat = (formatId: string) => {
    const newSelected = new Set(selectedFormats);
    if (newSelected.has(formatId)) {
      newSelected.delete(formatId);
    } else {
      newSelected.add(formatId);
    }
    setSelectedFormats(newSelected);
  };

  const handleExportSingle = async (formatId: string) => {
    setIsExporting(true);
    const timestamp = new Date().toISOString().split('T')[0];

    try {
      switch (formatId) {
        case 'json':
          downloadFile(exportToJSON(presentations), `export-${timestamp}.json`, 'application/json');
          break;
        case 'xml':
          downloadFile(exportToXML(presentations), `export-${timestamp}.xml`, 'application/xml');
          break;
        case 'csv':
          downloadFile(exportToCSV(presentations), `export-${timestamp}.csv`, 'text/csv');
          break;
        case 'txt':
          downloadFile(exportToText(presentations), `export-${timestamp}.txt`, 'text/plain');
          break;
        case 'html':
          downloadFile(exportToHTML(presentations), `export-${timestamp}.html`, 'text/html');
          break;
        case 'pdf':
          const pdf = exportToPDF(presentations);
          pdf.save(`export-${timestamp}.pdf`);
          break;
      }
      setExportSuccess(formatId);
      setTimeout(() => setExportSuccess(null), 2000);
    } catch (error) {
      console.error('Export error:', error);
    } finally {
      setIsExporting(false);
    }
  };

  const handleExportMedia = async () => {
    setIsExporting(true);
    try {
      await downloadMediaAsZip(presentations);
      setExportSuccess('media');
      setTimeout(() => setExportSuccess(null), 2000);
    } catch (error) {
      console.error('Media export error:', error);
    } finally {
      setIsExporting(false);
    }
  };

  const handleExportAll = async () => {
    if (selectedFormats.size === 0) return;
    
    setIsExporting(true);
    try {
      await downloadAllAsZip(presentations, Array.from(selectedFormats));
      setExportSuccess('all');
      setTimeout(() => setExportSuccess(null), 2000);
    } catch (error) {
      console.error('Export all error:', error);
    } finally {
      setIsExporting(false);
    }
  };

  if (presentations.length === 0) return null;

  return (
    <div className="card card-elevated p-4">
      <div className="flex items-center gap-3 mb-4">
        <Download className="w-5 h-5 text-[rgb(var(--primary))]" />
        <h3 className="font-semibold text-[rgb(var(--foreground))]">
          Export Data
        </h3>
        <span className="text-sm text-[rgb(var(--muted-foreground))]">
          {presentations.length} presentation{presentations.length !== 1 ? 's' : ''}
        </span>
      </div>

      {/* Format Selection */}
      <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-6 gap-2 mb-4">
        {exportFormats.map(format => {
          const Icon = format.icon;
          const isSelected = selectedFormats.has(format.id);
          const isSuccess = exportSuccess === format.id;

          return (
            <button
              key={format.id}
              onClick={() => toggleFormat(format.id)}
              disabled={isExporting}
              className={`
                p-3 rounded-lg border transition-all text-center
                ${isSelected 
                  ? 'border-[rgb(var(--primary))] bg-[rgb(var(--primary)/0.1)]' 
                  : 'border-[rgb(var(--border))] hover:border-[rgb(var(--muted-foreground))]'
                }
                ${isSuccess ? 'bg-[rgb(var(--success)/0.1)] border-[rgb(var(--success))]' : ''}
              `}
            >
              <Icon className={`w-5 h-5 mx-auto mb-1 ${isSelected ? 'text-[rgb(var(--primary))]' : ''}`} />
              <span className="text-sm font-medium">{format.name}</span>
              {isSuccess && <Check className="w-4 h-4 mx-auto mt-1 text-[rgb(var(--success))]" />}
            </button>
          );
        })}
      </div>

      {/* Actions */}
      <div className="flex flex-wrap gap-2">
        {/* Individual exports */}
        {Array.from(selectedFormats).map(formatId => {
          const format = exportFormats.find(f => f.id === formatId);
          if (!format) return null;
          const Icon = format.icon;

          return (
            <button
              key={formatId}
              onClick={() => handleExportSingle(formatId)}
              disabled={isExporting}
              className="btn btn-secondary text-sm"
            >
              <Icon className="w-4 h-4" />
              Export {format.name}
            </button>
          );
        })}

        {/* Media export */}
        {hasMedia && (
          <button
            onClick={handleExportMedia}
            disabled={isExporting}
            className="btn btn-secondary text-sm"
          >
            <ImageIcon className="w-4 h-4" />
            Export Media ({totalMedia})
          </button>
        )}

        {/* Export all as ZIP */}
        {selectedFormats.size > 0 && (
          <button
            onClick={handleExportAll}
            disabled={isExporting || selectedFormats.size === 0}
            className="btn btn-primary text-sm ml-auto"
          >
            {isExporting ? (
              <Loader2 className="w-4 h-4 animate-spin" />
            ) : (
              <Package className="w-4 h-4" />
            )}
            Download All as ZIP
          </button>
        )}
      </div>

      {/* Export info */}
      <p className="text-xs text-[rgb(var(--muted-foreground))] mt-4">
        Select formats above, then click individual export buttons or download everything as a ZIP file.
        {hasMedia && ' Media files will be included in a separate folder in the ZIP.'}
      </p>
    </div>
  );
}
