/**
 * File List Component - Shows uploaded files
 */

import { X, FileType, Loader2, CheckCircle } from 'lucide-react';
import type { ExtractedPresentation } from '../types';

interface FileListProps {
  files: File[];
  extractedData: ExtractedPresentation[];
  processingFile: string | null;
  onRemoveFile: (index: number) => void;
  onViewData: (presentation: ExtractedPresentation) => void;
}

export function FileList({ 
  files, 
  extractedData, 
  processingFile, 
  onRemoveFile,
  onViewData 
}: FileListProps) {
  if (files.length === 0) return null;

  const getFileStatus = (fileName: string) => {
    if (processingFile === fileName) return 'processing';
    if (extractedData.find(p => p.fileName === fileName)) return 'done';
    return 'pending';
  };

  const formatFileSize = (bytes: number): string => {
    if (bytes === 0) return '0 B';
    const k = 1024;
    const sizes = ['B', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return `${parseFloat((bytes / Math.pow(k, i)).toFixed(1))} ${sizes[i]}`;
  };

  return (
    <div className="card card-elevated p-4">
      <div className="flex items-center justify-between mb-4">
        <h3 className="font-semibold text-[rgb(var(--foreground))]">
          Uploaded Files ({files.length})
        </h3>
        <span className="text-sm text-[rgb(var(--muted-foreground))]">
          {extractedData.length} of {files.length} processed
        </span>
      </div>

      <div className="space-y-2">
        {files.map((file, index) => {
          const status = getFileStatus(file.name);
          const extracted = extractedData.find(p => p.fileName === file.name);
          const isPPTX = file.name.toLowerCase().endsWith('.pptx');

          return (
            <div
              key={`${file.name}-${index}`}
              className={`
                flex items-center gap-3 p-3 rounded-lg
                ${status === 'done' 
                  ? 'bg-[rgb(var(--success)/0.1)]' 
                  : 'bg-[rgb(var(--secondary))]'
                }
              `}
            >
              <div className={`
                w-10 h-10 rounded-lg flex items-center justify-center shrink-0
                ${isPPTX ? 'bg-orange-500' : 'bg-blue-500'}
              `}>
                <FileType className="w-5 h-5 text-white" />
              </div>

              <div className="flex-1 min-w-0">
                <p className="font-medium text-[rgb(var(--foreground))] truncate">
                  {file.name}
                </p>
                <div className="flex items-center gap-2 text-xs text-[rgb(var(--muted-foreground))]">
                  <span>{formatFileSize(file.size)}</span>
                  <span>•</span>
                  <span className="uppercase">{isPPTX ? 'PPTX' : 'PPT'}</span>
                  {extracted && (
                    <>
                      <span>•</span>
                      <span>{extracted.slides.length} slides</span>
                    </>
                  )}
                </div>
              </div>

              <div className="flex items-center gap-2">
                {status === 'processing' && (
                  <div className="flex items-center gap-2 text-sm text-[rgb(var(--muted-foreground))]">
                    <Loader2 className="w-4 h-4 animate-spin" />
                    <span>Processing...</span>
                  </div>
                )}

                {status === 'done' && extracted && (
                  <>
                    <CheckCircle className="w-5 h-5 text-[rgb(var(--success))]" />
                    <button
                      onClick={() => onViewData(extracted)}
                      className="btn btn-secondary text-sm py-1.5 px-3"
                    >
                      View
                    </button>
                  </>
                )}

                {status === 'pending' && (
                  <span className="text-sm text-[rgb(var(--muted-foreground))]">
                    Pending
                  </span>
                )}

                <button
                  onClick={() => onRemoveFile(index)}
                  className="p-1.5 rounded hover:bg-[rgb(var(--destructive)/0.1)] transition-colors"
                  title="Remove file"
                >
                  <X className="w-4 h-4 text-[rgb(var(--muted-foreground))]" />
                </button>
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}
