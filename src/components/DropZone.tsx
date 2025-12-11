/**
 * File Upload / Drop Zone Component
 */

import { useState, useRef, useCallback } from 'react';
import { Upload, FileUp, X, AlertCircle } from 'lucide-react';

interface DropZoneProps {
  onFilesSelected: (files: File[]) => void;
  isProcessing: boolean;
}

export function DropZone({ onFilesSelected, isProcessing }: DropZoneProps) {
  const [isDragging, setIsDragging] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const validateFiles = (files: FileList | File[]): File[] => {
    const validFiles: File[] = [];
    const invalidFiles: string[] = [];

    Array.from(files).forEach(file => {
      const ext = file.name.toLowerCase().split('.').pop();
      if (ext === 'ppt' || ext === 'pptx') {
        validFiles.push(file);
      } else {
        invalidFiles.push(file.name);
      }
    });

    if (invalidFiles.length > 0) {
      setError(`Invalid files: ${invalidFiles.join(', ')}. Only .ppt and .pptx files are supported.`);
    } else {
      setError(null);
    }

    return validFiles;
  };

  const handleDragEnter = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(true);
  }, []);

  const handleDragLeave = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);
  }, []);

  const handleDragOver = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
  }, []);

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);

    const files = e.dataTransfer.files;
    const validFiles = validateFiles(files);
    if (validFiles.length > 0) {
      onFilesSelected(validFiles);
    }
  }, [onFilesSelected]);

  const handleFileInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (files) {
      const validFiles = validateFiles(files);
      if (validFiles.length > 0) {
        onFilesSelected(validFiles);
      }
    }
    // Reset input
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  };

  const handleClick = () => {
    fileInputRef.current?.click();
  };

  return (
    <div className="w-full">
      <div
        className={`
          relative border-2 border-dashed rounded-xl p-8 md:p-12 text-center
          transition-all duration-200 cursor-pointer
          ${isDragging 
            ? 'border-[rgb(var(--primary))] bg-[rgb(var(--primary)/0.05)]' 
            : 'border-[rgb(var(--border))] hover:border-[rgb(var(--muted-foreground))]'
          }
          ${isProcessing ? 'pointer-events-none opacity-50' : ''}
        `}
        onDragEnter={handleDragEnter}
        onDragLeave={handleDragLeave}
        onDragOver={handleDragOver}
        onDrop={handleDrop}
        onClick={handleClick}
      >
        <input
          ref={fileInputRef}
          type="file"
          accept=".ppt,.pptx"
          multiple
          onChange={handleFileInputChange}
          className="hidden"
        />

        <div className="flex flex-col items-center gap-4">
          <div className={`
            w-16 h-16 rounded-full flex items-center justify-center
            ${isDragging 
              ? 'bg-[rgb(var(--primary))] text-[rgb(var(--primary-foreground))]' 
              : 'bg-[rgb(var(--secondary))]'
            }
          `}>
            {isDragging ? (
              <FileUp className="w-8 h-8" />
            ) : (
              <Upload className="w-8 h-8" />
            )}
          </div>

          <div>
            <p className="text-lg font-medium text-[rgb(var(--foreground))]">
              {isDragging ? 'Drop files here' : 'Drop PowerPoint files here'}
            </p>
            <p className="text-sm text-[rgb(var(--muted-foreground))] mt-1">
              or click to browse
            </p>
          </div>

          <div className="flex items-center gap-2 text-xs text-[rgb(var(--muted-foreground))]">
            <span className="badge">PPT</span>
            <span className="badge">PPTX</span>
            <span>Multiple files supported</span>
          </div>
        </div>
      </div>

      {error && (
        <div className="mt-4 p-4 rounded-lg bg-[rgb(var(--destructive)/0.1)] border border-[rgb(var(--destructive)/0.3)] flex items-start gap-3">
          <AlertCircle className="w-5 h-5 text-[rgb(var(--destructive))] shrink-0 mt-0.5" />
          <div>
            <p className="text-sm font-medium text-[rgb(var(--destructive))]">Invalid Files</p>
            <p className="text-sm text-[rgb(var(--muted-foreground))]">{error}</p>
          </div>
          <button
            onClick={(e) => {
              e.stopPropagation();
              setError(null);
            }}
            className="ml-auto p-1 hover:bg-[rgb(var(--destructive)/0.1)] rounded"
          >
            <X className="w-4 h-4 text-[rgb(var(--destructive))]" />
          </button>
        </div>
      )}
    </div>
  );
}
